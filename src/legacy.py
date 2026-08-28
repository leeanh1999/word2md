"""Support for the binary Office 97-2003 formats (.doc and .xls).

Neither mammoth nor openpyxl can read them, so this module upgrades a legacy
file to its modern equivalent first. Several backends are tried in order and
the app degrades gracefully instead of failing outright:

    .doc  ->  Microsoft Word (COM)  ->  LibreOffice  ->  built-in OLE text
    .xls  ->  xlrd (pure Python)    ->  Microsoft Excel (COM)  ->  LibreOffice

Only the last .doc fallback loses formatting, and it says so in a warning.
"""

from __future__ import annotations

import atexit
import shutil
import struct
import subprocess
import sys
import tempfile
import threading
from dataclasses import dataclass
from pathlib import Path

LEGACY_WORD = {".doc"}
LEGACY_EXCEL = {".xls"}
LEGACY_EXTENSIONS = LEGACY_WORD | LEGACY_EXCEL

ZIP_MAGIC = b"PK\x03\x04"
OLE_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"
RTF_MAGIC = b"{\\rt"

# Word: wdFormatDocumentDefault, Excel: xlOpenXMLWorkbook.
WD_FORMAT_DOCX = 16
WD_EXPORT_FORMAT_PDF = 17
XL_FORMAT_XLSX = 51
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3

SOFFICE_TIMEOUT = 180

_SOFFICE_CANDIDATES = (
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/usr/bin/soffice",
    "/usr/bin/libreoffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
)


class LegacyConversionError(RuntimeError):
    """Every backend refused to read the file."""


# ------------------------------------------------------------------ sniffing


@dataclass(frozen=True)
class FileKind:
    """What a file really is, regardless of its extension."""

    container: str  # "ole" | "zip" | "rtf" | "unknown"
    suggested_suffix: str | None = None

    @property
    def is_modern(self) -> bool:
        return self.container == "zip"


def sniff(source: Path) -> FileKind:
    """Legacy extensions lie often: .doc files are frequently .docx or RTF."""
    try:
        with source.open("rb") as handle:
            head = handle.read(8)
    except OSError:
        return FileKind("unknown")

    if head.startswith(ZIP_MAGIC):
        suffix = ".docx" if source.suffix.lower() in LEGACY_WORD else ".xlsx"
        return FileKind("zip", suffix)
    if head.startswith(OLE_MAGIC):
        return FileKind("ole")
    if head.startswith(RTF_MAGIC):
        return FileKind("rtf", ".rtf")
    return FileKind("unknown")


# ------------------------------------------------------------------ backends


def has_pywin32() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # noqa: F401
    except Exception:  # noqa: BLE001
        return False
    return True


def find_soffice() -> str | None:
    found = shutil.which("soffice") or shutil.which("soffice.exe")
    if found:
        return found
    for candidate in _SOFFICE_CANDIDATES:
        if Path(candidate).exists():
            return candidate
    return None


def has_msoffice(application: str) -> bool:
    """Look up the ProgID in the registry - no COM call, nothing launched."""
    if not has_pywin32():
        return False
    try:
        import winreg

        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, application))
    except Exception:  # noqa: BLE001
        return False
    return True


def available_backends() -> list[str]:
    """Human-readable list of what this machine can do, for diagnostics."""
    backends = []
    try:
        import xlrd  # noqa: F401

        backends.append("xlrd (.xls)")
    except Exception:  # noqa: BLE001
        pass
    if has_msoffice("Word.Application"):
        backends.append("Microsoft Word (COM)")
    if has_msoffice("Excel.Application"):
        backends.append("Microsoft Excel (COM)")
    if find_soffice():
        backends.append("LibreOffice")
    try:
        import olefile  # noqa: F401

        backends.append("OLE text (.doc, chỉ văn bản)")
    except Exception:  # noqa: BLE001
        pass
    return backends


def can_read_xls() -> bool:
    try:
        import xlrd  # noqa: F401
    except Exception:  # noqa: BLE001
        return False
    return True


# ------------------------------------------------------------------- cache

# Upgrading through Office costs seconds, and the same document is often
# touched twice (outline preview, then export). Results are cached for the
# lifetime of the process, keyed on the file's identity.
_cache: dict[tuple[str, int, int], Path] = {}
_cache_dir: Path | None = None
_cache_lock = threading.Lock()


def _workspace() -> Path:
    global _cache_dir
    if _cache_dir is None or not _cache_dir.exists():
        _cache_dir = Path(tempfile.mkdtemp(prefix="word2md-"))
        atexit.register(shutil.rmtree, _cache_dir, ignore_errors=True)
    return _cache_dir


def _identity(source: Path) -> tuple[str, int, int]:
    stat = source.stat()
    return (str(source.resolve()).lower(), stat.st_mtime_ns, stat.st_size)


def clear_cache() -> None:
    with _cache_lock:
        _cache.clear()
    global _cache_dir
    if _cache_dir is not None:
        shutil.rmtree(_cache_dir, ignore_errors=True)
        _cache_dir = None


# ---------------------------------------------------------------- upgrading


class LegacyUpgrader:
    """Converts legacy files to modern ones, reusing one Office instance.

    Starting Word or Excel costs a couple of seconds, so a batch keeps the
    same instance alive until `close()`.
    """

    def __init__(self, temp_dir: str | Path | None = None):
        self._temp = Path(temp_dir) if temp_dir else None
        self._word = None
        self._excel = None
        self._counter = 0
        self._com_ready = False

    # ------------------------------------------------------------ lifecycle

    def __enter__(self) -> "LegacyUpgrader":
        return self

    def __exit__(self, *_exc) -> None:
        self.close()

    def close(self) -> None:
        for attribute in ("_word", "_excel"):
            app = getattr(self, attribute, None)
            if app is None:
                continue
            try:
                app.Quit()
            except Exception:  # noqa: BLE001
                pass
            setattr(self, attribute, None)
        self._uninitialize_com()

    def _temp_target(self, source: Path, suffix: str) -> Path:
        self._counter += 1
        folder = self._temp or _workspace()
        folder.mkdir(parents=True, exist_ok=True)
        return folder / f"{self._counter:03d}_{source.stem}{suffix}"

    # ---------------------------------------------------------------- entry

    def upgrade(self, source: Path) -> tuple[Path, list[str]]:
        """Return a path readable by the modern pipeline, plus any warnings.

        Raises LegacyConversionError when nothing on this machine can read it.
        """
        source = Path(source).resolve()
        suffix = source.suffix.lower()
        kind = sniff(source)

        if kind.is_modern:
            target = self._cached(source, kind.suggested_suffix or ".docx", shutil.copy2)
            return target, [
                f"File có đuôi {suffix} nhưng thực chất là định dạng mới "
                f"({kind.suggested_suffix}); đã xử lý trực tiếp."
            ]

        is_word = suffix in LEGACY_WORD or kind.container == "rtf"
        target_suffix = ".docx" if is_word else ".xlsx"
        backends = self._backends(is_word)
        if not backends:
            raise LegacyConversionError(
                f"Không đọc được {source.name}: máy chưa cài Microsoft Office "
                "hoặc LibreOffice để đọc định dạng cũ."
            )

        attempts: list[str] = []

        def run(original: Path, target: Path) -> None:
            for name, runner in backends:
                try:
                    runner(original, target)
                except Exception as exc:  # noqa: BLE001 - try the next backend
                    attempts.append(f"{name}: {type(exc).__name__}: {exc}")
                    continue
                if target.exists() and target.stat().st_size > 0:
                    return
                attempts.append(f"{name}: không tạo được file kết quả.")
            raise LegacyConversionError(
                f"Không đọc được {source.name}. Đã thử: " + " | ".join(attempts)
            )

        return self._cached(source, target_suffix, run), []

    def to_pdf(self, source: Path, target: Path) -> list[str]:
        """Export a .docx as PDF, through whichever engine this machine has."""
        source, target = Path(source).resolve(), Path(target)
        target.parent.mkdir(parents=True, exist_ok=True)
        backends = []
        if has_msoffice("Word.Application"):
            backends.append(("Microsoft Word", self._word_to_pdf))
        if find_soffice():
            backends.append(("LibreOffice", self._soffice_convert))
        if not backends:
            raise LegacyConversionError(
                "Không xuất được PDF: máy chưa cài Microsoft Word hoặc LibreOffice."
            )

        attempts: list[str] = []
        for name, runner in backends:
            try:
                runner(source, target)
            except Exception as exc:  # noqa: BLE001 - try the next backend
                attempts.append(f"{name}: {type(exc).__name__}: {exc}")
                continue
            if target.exists() and target.stat().st_size > 0:
                if name == "LibreOffice":
                    return [
                        "PDF được xuất bằng LibreOffice; mục lục tự động có thể "
                        "còn trống cho tới khi mở lại bằng Word."
                    ]
                return []
            attempts.append(f"{name}: không tạo được file kết quả.")
        raise LegacyConversionError(
            "Không xuất được PDF. Đã thử: " + " | ".join(attempts)
        )

    def _cached(self, source: Path, suffix: str, produce) -> Path:
        key = _identity(source)
        with _cache_lock:
            hit = _cache.get(key)
        if hit is not None and hit.exists():
            return hit

        target = self._temp_target(source, suffix)
        produce(source, target)
        with _cache_lock:
            _cache[key] = target
        return target

    def _backends(self, is_word: bool):
        application = "Word.Application" if is_word else "Excel.Application"
        backends = []
        if has_msoffice(application):
            backends.append(
                (
                    "Microsoft Office",
                    self._word_convert if is_word else self._excel_convert,
                )
            )
        if find_soffice():
            backends.append(("LibreOffice", self._soffice_convert))
        return backends

    # ------------------------------------------------------------------ COM

    def _initialize_com(self) -> None:
        """COM must be initialised on whichever worker thread drives Office."""
        if self._com_ready:
            return
        import pythoncom

        try:
            pythoncom.CoInitialize()
        except Exception:  # noqa: BLE001 - already initialised on this thread
            pass
        self._com_ready = True

    def _uninitialize_com(self) -> None:
        if not self._com_ready:
            return
        self._com_ready = False
        try:
            import pythoncom

            pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001
            pass

    def _word_app(self):
        if self._word is None:
            import win32com.client

            self._initialize_com()
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            try:
                app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
            except Exception:  # noqa: BLE001 - not every build exposes this
                pass
            self._word = app
        return self._word

    def _word_convert(self, source: Path, target: Path) -> None:
        app = self._word_app()
        document = app.Documents.Open(
            str(source),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )
        try:
            document.SaveAs2(str(target), FileFormat=WD_FORMAT_DOCX)
        finally:
            try:
                document.Close(0)
            except Exception:  # noqa: BLE001
                pass

    def _word_to_pdf(self, source: Path, target: Path) -> None:
        app = self._word_app()
        document = app.Documents.Open(
            str(source),
            ConfirmConversions=False,
            ReadOnly=False,  # a table of contents has to be filled in first
            AddToRecentFiles=False,
            Visible=False,
        )
        try:
            try:
                document.Fields.Update()
            except Exception:  # noqa: BLE001 - a document may have no fields
                pass
            document.ExportAsFixedFormat(
                str(target), ExportFormat=WD_EXPORT_FORMAT_PDF
            )
        finally:
            try:
                document.Close(0)
            except Exception:  # noqa: BLE001
                pass

    def _excel_app(self):
        if self._excel is None:
            import win32com.client

            self._initialize_com()
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            try:
                app.AutomationSecurity = MSO_AUTOMATION_SECURITY_FORCE_DISABLE
            except Exception:  # noqa: BLE001
                pass
            self._excel = app
        return self._excel

    def _excel_convert(self, source: Path, target: Path) -> None:
        app = self._excel_app()
        workbook = app.Workbooks.Open(
            str(source), UpdateLinks=0, ReadOnly=True, AddToMru=False
        )
        try:
            workbook.SaveAs(str(target), FileFormat=XL_FORMAT_XLSX)
        finally:
            try:
                workbook.Close(False)
            except Exception:  # noqa: BLE001
                pass

    # ---------------------------------------------------------- LibreOffice

    def _soffice_convert(self, source: Path, target: Path) -> None:
        soffice = find_soffice()
        if soffice is None:
            raise LegacyConversionError("Không tìm thấy LibreOffice.")

        outdir = target.parent / f"{target.stem}_lo"
        outdir.mkdir(parents=True, exist_ok=True)
        fmt = target.suffix.lower().lstrip(".") or "docx"
        profile = (outdir / "profile").as_uri()

        subprocess.run(
            [
                soffice,
                "--headless",
                "--norestore",
                "--invisible",
                f"-env:UserInstallation={profile}",
                "--convert-to",
                fmt,
                "--outdir",
                str(outdir),
                str(source),
            ],
            check=True,
            timeout=SOFFICE_TIMEOUT,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
        )

        produced = outdir / f"{source.stem}.{fmt}"
        if not produced.exists():
            candidates = list(outdir.glob(f"*.{fmt}"))
            if not candidates:
                raise LegacyConversionError("LibreOffice không xuất ra file nào.")
            produced = candidates[0]
        shutil.move(str(produced), str(target))


# ------------------------------------------------- .doc plain-text fallback


def _doc_streams(path: Path) -> tuple[bytes, bytes]:
    import olefile

    with olefile.OleFileIO(str(path)) as ole:
        if not ole.exists("WordDocument"):
            raise LegacyConversionError("Không tìm thấy stream WordDocument.")
        document = ole.openstream("WordDocument").read()

        # FibBase.flags bit 0x0200 selects which table stream is current.
        flags = struct.unpack_from("<H", document, 0x0A)[0]
        preferred = "1Table" if flags & 0x0200 else "0Table"
        name = preferred if ole.exists(preferred) else None
        if name is None:
            name = next((n for n in ("1Table", "0Table") if ole.exists(n)), None)
        if name is None:
            raise LegacyConversionError("Không tìm thấy table stream.")
        table = ole.openstream(name).read()

    return document, table


def _parse_piece_table(clx: bytes) -> list[tuple[int, int, int, bool]]:
    """Walk the Clx to find the Pcdt, then decode its PlcPcd."""
    position = 0
    while position < len(clx):
        tag = clx[position]
        if tag == 0x01:  # Prc - formatting run we do not care about
            size = struct.unpack_from("<h", clx, position + 1)[0]
            position += 3 + size
        elif tag == 0x02:  # Pcdt - the piece table itself
            length = struct.unpack_from("<I", clx, position + 1)[0]
            return _parse_plcpcd(clx[position + 5 : position + 5 + length])
        else:
            break
    raise LegacyConversionError("Không đọc được bảng piece table.")


def _parse_plcpcd(plc: bytes) -> list[tuple[int, int, int, bool]]:
    count = (len(plc) - 4) // 12
    if count <= 0:
        raise LegacyConversionError("Piece table rỗng.")

    character_positions = struct.unpack_from(f"<{count + 1}I", plc, 0)
    base = 4 * (count + 1)
    pieces = []
    for index in range(count):
        raw_fc = struct.unpack_from("<I", plc, base + index * 8 + 2)[0]
        compressed = bool(raw_fc & 0x40000000)
        offset = raw_fc & 0x3FFFFFFF
        if compressed:
            offset //= 2
        pieces.append(
            (
                character_positions[index],
                character_positions[index + 1],
                offset,
                compressed,
            )
        )
    return pieces


def _strip_field_instructions(text: str) -> str:
    """Drop the code between \\x13 and \\x14, keep the result up to \\x15."""
    output = []
    skipping = False
    for char in text:
        if char == "\x13":
            skipping = True
        elif char == "\x14":
            skipping = False
        elif char == "\x15":
            skipping = False
        elif not skipping:
            output.append(char)
    return "".join(output)


_LINE_BREAKS = {"\r", "\x07", "\x0b", "\x0c", "\x0e"}


def extract_doc_text(source: Path) -> str:
    """Best-effort plain text from a Word 97-2003 binary document."""
    document, table = _doc_streams(Path(source))

    fc_clx, lcb_clx = struct.unpack_from("<II", document, 0x01A2)
    if lcb_clx <= 0 or fc_clx + lcb_clx > len(table):
        raise LegacyConversionError("Vị trí Clx không hợp lệ.")

    pieces = _parse_piece_table(table[fc_clx : fc_clx + lcb_clx])
    character_count = struct.unpack_from("<I", document, 0x4C)[0]  # ccpText

    chunks = []
    for start, end, offset, compressed in pieces:
        if start >= character_count:
            break
        length = min(end, character_count) - start
        if length <= 0:
            continue
        if compressed:
            chunks.append(document[offset : offset + length].decode("cp1252", "replace"))
        else:
            chunks.append(
                document[offset : offset + length * 2].decode("utf-16-le", "replace")
            )

    text = _strip_field_instructions("".join(chunks))

    cleaned = []
    for char in text:
        if char in _LINE_BREAKS:
            cleaned.append("\n")
        elif char in ("\t", "\n"):
            cleaned.append(char)
        elif char == "\xa0":
            cleaned.append(" ")
        elif ord(char) < 0x20:
            continue
        else:
            cleaned.append(char)

    paragraphs = [line.strip() for line in "".join(cleaned).split("\n")]
    return "\n\n".join(paragraph for paragraph in paragraphs if paragraph)


def doc_text_to_markdown(source: Path) -> tuple[str, list[str]]:
    """Fallback conversion used when no Office/LibreOffice backend exists."""
    text = extract_doc_text(Path(source))
    if not text.strip():
        raise LegacyConversionError("Không trích xuất được nội dung văn bản.")
    warning = (
        "Không có Microsoft Word hoặc LibreOffice trên máy nên chỉ trích xuất được "
        "văn bản thuần: tiêu đề, bảng biểu và định dạng đã bị mất."
    )
    return text + "\n", [warning]


__all__ = [
    "LEGACY_EXCEL",
    "LEGACY_EXTENSIONS",
    "LEGACY_WORD",
    "FileKind",
    "LegacyConversionError",
    "LegacyUpgrader",
    "available_backends",
    "can_read_xls",
    "clear_cache",
    "doc_text_to_markdown",
    "extract_doc_text",
    "find_soffice",
    "has_msoffice",
    "sniff",
]
