"""Core conversion engine: Word (.docx) and Excel (.xlsx) -> Markdown (.md)."""

from __future__ import annotations

import datetime as _dt
import mimetypes
import threading
import time
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Callable, Iterable, Sequence

import mammoth
import pandas as pd

from .html_to_markdown import html_to_markdown
from .legacy import (
    LEGACY_EXCEL,
    LEGACY_EXTENSIONS,
    LEGACY_WORD,
    LegacyConversionError,
    LegacyUpgrader,
    can_read_xls,
    doc_text_to_markdown,
    sniff,
)
from .outline import (
    OutlineNode,
    build_outline,
    extract_sections,
    index_nodes,
    section_text,
    slugify_title,
    top_level_selection,
)

WORD_EXTENSIONS = {".docx", ".doc"}
EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}
SUPPORTED_EXTENSIONS = WORD_EXTENSIONS | EXCEL_EXTENSIONS

STATUS_SUCCESS = "success"
STATUS_ERROR = "error"
STATUS_SKIPPED = "skipped"
STATUS_CANCELLED = "cancelled"

# Word styles mammoth does not map out of the box.
DOCX_STYLE_MAP = """
p[style-name='Title'] => h1:fresh
p[style-name='Subtitle'] => h2:fresh
p[style-name='Quote'] => blockquote > p:fresh
p[style-name='Intense Quote'] => blockquote > p:fresh
p[style-name='Caption'] => p > em:fresh
p[style-name='Footer'] => p:fresh
p[style-name='Header'] => p:fresh
p[style-name='Code'] => pre:separator('\\n')
r[style-name='Code Char'] => code
r[style-name='Strong'] => strong
"""


@dataclass
class ConversionOptions:
    extract_images: bool = True
    overwrite: bool = False
    first_row_as_header: bool = True
    add_title_heading: bool = True
    sheet_heading_level: int = 2
    promote_headings: bool = True
    split_sections: bool = False


@dataclass
class ConversionResult:
    source: Path
    status: str
    output: Path | None = None
    message: str = ""
    warnings: list[str] = field(default_factory=list)
    duration: float = 0.0

    @property
    def ok(self) -> bool:
        return self.status == STATUS_SUCCESS


ProgressCallback = Callable[[int, int, ConversionResult], None]


# --------------------------------------------------------------- file lookup


def is_supported(path: Path) -> bool:
    return path.suffix.lower() in SUPPORTED_EXTENSIONS and not path.name.startswith("~$")


def collect_files(paths: Iterable[str | Path], recursive: bool = True) -> list[Path]:
    """Expand a mix of files and folders into a sorted list of convertible files."""
    found: list[Path] = []
    seen: set[Path] = set()

    for raw in paths:
        path = Path(raw).expanduser()
        candidates: Iterable[Path]
        if path.is_dir():
            candidates = path.rglob("*") if recursive else path.glob("*")
        else:
            candidates = [path]

        for candidate in candidates:
            if not candidate.is_file() or not is_supported(candidate):
                continue
            try:
                key = candidate.resolve()
            except OSError:
                key = candidate.absolute()
            if key not in seen:
                seen.add(key)
                found.append(candidate)

    return sorted(found, key=lambda p: str(p).lower())


def target_path(source: Path, output_dir: Path, overwrite: bool = False) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    target = output_dir / f"{source.stem}.md"
    if overwrite:
        return target
    counter = 1
    while target.exists():
        target = output_dir / f"{source.stem} ({counter}).md"
        counter += 1
    return target


# ------------------------------------------------------------------ word


def _image_converter(image_dir: Path, warnings: list[str]):
    """Save embedded images next to the .md file and link them relatively."""
    state = {"count": 0}

    @mammoth.images.img_element
    def convert_image(image):
        try:
            state["count"] += 1
            extension = mimetypes.guess_extension(image.content_type or "") or ".png"
            if extension == ".jpe":
                extension = ".jpg"
            image_dir.mkdir(parents=True, exist_ok=True)
            name = f"image{state['count']}{extension}"
            with image.open() as stream:
                (image_dir / name).write_bytes(stream.read())
            return {
                "src": f"{image_dir.name}/{name}",
                "alt": getattr(image, "alt_text", "") or "",
            }
        except Exception as exc:  # noqa: BLE001 - one bad image must not kill the file
            warnings.append(f"Không trích xuất được ảnh: {exc}")
            return {"src": ""}

    return convert_image


def docx_to_markdown(
    source: Path,
    options: ConversionOptions | None = None,
    image_dir: Path | None = None,
    title: str | None = None,
) -> tuple[str, list[str]]:
    """Return (markdown, warnings) for a .docx file.

    `title` overrides the H1 derived from the file name, which matters when
    reading a temporary file upgraded from .doc.
    """
    options = options or ConversionOptions()
    warnings: list[str] = []

    kwargs = {"style_map": DOCX_STYLE_MAP}
    if options.extract_images and image_dir is not None:
        kwargs["convert_image"] = _image_converter(image_dir, warnings)
    elif not options.extract_images:
        kwargs["convert_image"] = mammoth.images.img_element(lambda _image: {"src": ""})

    with source.open("rb") as handle:
        result = mammoth.convert_to_html(handle, **kwargs)

    warnings.extend(message.message for message in result.messages)
    markdown = html_to_markdown(
        result.value, first_row_as_header=options.first_row_as_header
    )

    if options.add_title_heading and not markdown.lstrip().startswith("# "):
        markdown = f"# {title or source.stem}\n\n{markdown}"

    return markdown, warnings


def word_to_markdown(
    source: Path,
    options: ConversionOptions | None = None,
    image_dir: Path | None = None,
    upgrader: LegacyUpgrader | None = None,
) -> tuple[str, list[str]]:
    """Convert .docx directly, or .doc after upgrading it to .docx."""
    options = options or ConversionOptions()
    source = Path(source)
    if source.suffix.lower() not in LEGACY_WORD:
        return docx_to_markdown(source, options, image_dir)

    owns = upgrader is None
    upgrader = upgrader or LegacyUpgrader()
    try:
        try:
            upgraded, warnings = upgrader.upgrade(source)
        except LegacyConversionError as exc:
            markdown, warnings = _doc_text_fallback(source, options, exc)
            return markdown, warnings
        markdown, more = docx_to_markdown(upgraded, options, image_dir, source.stem)
        return markdown, warnings + more
    finally:
        if owns:
            upgrader.close()


def _doc_text_fallback(
    source: Path, options: ConversionOptions, cause: Exception
) -> tuple[str, list[str]]:
    """No Office and no LibreOffice: salvage the text straight from the OLE file."""
    try:
        markdown, warnings = doc_text_to_markdown(source)
    except Exception:  # noqa: BLE001 - report the original, more useful failure
        raise cause from None
    if options.add_title_heading and not markdown.lstrip().startswith("# "):
        markdown = f"# {source.stem}\n\n{markdown}"
    return markdown, warnings


def ensure_docx(
    source: Path, upgrader: LegacyUpgrader | None = None
) -> tuple[Path, str, list[str]]:
    """Return (path to a .docx, display title, warnings) for any Word file."""
    source = Path(source)
    if source.suffix.lower() not in LEGACY_WORD:
        return source, source.stem, []
    owns = upgrader is None
    upgrader = upgrader or LegacyUpgrader()
    try:
        upgraded, warnings = upgrader.upgrade(source)
    finally:
        if owns:
            upgrader.close()
    return upgraded, source.stem, warnings


# --------------------------------------------------- word section extraction


@dataclass
class DocumentOutline:
    """A parsed .docx ready for partial extraction."""

    source: Path
    markdown: str
    roots: list[OutlineNode]
    warnings: list[str] = field(default_factory=list)

    @property
    def nodes(self) -> dict[str, OutlineNode]:
        return index_nodes(self.roots)

    def preview(self, node_ids: Sequence[str], promote: bool = True) -> str:
        return extract_sections(
            self.markdown, node_ids, roots=self.roots, promote=promote
        )


def load_outline(
    source: str | Path,
    options: ConversionOptions | None = None,
    upgrader: LegacyUpgrader | None = None,
) -> DocumentOutline:
    """Read a Word file and expose its heading tree, like the Navigation Pane.

    Images are skipped here: nothing should be written to disk until the user
    actually exports. The final conversion re-runs with images enabled, and the
    heading structure - hence every node id - is identical either way. A .doc
    is upgraded first, and the result is cached so the export does not pay for
    a second round-trip through Word.
    """
    source = Path(source)
    if source.suffix.lower() not in WORD_EXTENSIONS:
        raise ValueError(
            f"Chỉ hỗ trợ đọc mục lục từ file Word (nhận được {source.suffix})."
        )

    options = replace(
        options or ConversionOptions(), extract_images=False, add_title_heading=False
    )
    readable, title, warnings = ensure_docx(source, upgrader)
    markdown, more = docx_to_markdown(readable, options, title=title)
    return DocumentOutline(
        source=source,
        markdown=markdown,
        roots=build_outline(markdown),
        warnings=warnings + more,
    )


def convert_docx_sections(
    source: str | Path,
    output_dir: str | Path,
    node_ids: Sequence[str],
    options: ConversionOptions | None = None,
    upgrader: LegacyUpgrader | None = None,
) -> list[ConversionResult]:
    """Export only the chosen sections of a Word document (.docx or .doc).

    With `options.split_sections` each selected section becomes its own file,
    otherwise they are concatenated into a single document.
    """
    options = options or ConversionOptions()
    source = Path(source)
    output_dir = Path(output_dir)
    started = time.perf_counter()

    def failed(message: str) -> list[ConversionResult]:
        return [
            ConversionResult(
                source=source,
                status=STATUS_ERROR,
                message=message,
                duration=time.perf_counter() - started,
            )
        ]

    if not node_ids:
        return failed("Chưa chọn mục nào để trích xuất.")

    try:
        output_dir.mkdir(parents=True, exist_ok=True)
        image_dir = output_dir / f"{source.stem}_images"
        convert_options = replace(options, add_title_heading=False)
        readable, title, warnings = ensure_docx(source, upgrader)
        markdown, more = docx_to_markdown(readable, convert_options, image_dir, title)
        warnings = warnings + more
        roots = build_outline(markdown)
        selected = top_level_selection(roots, node_ids)
        if not selected:
            return failed("Các mục đã chọn không còn tồn tại trong tài liệu.")
    except LegacyConversionError as exc:
        return failed(str(exc))
    except PermissionError:
        return failed("Không có quyền truy cập (file đang mở?).")
    except Exception as exc:  # noqa: BLE001
        return failed(f"{type(exc).__name__}: {exc}")

    if not options.split_sections:
        body = extract_sections(
            markdown, node_ids, roots=roots, promote=options.promote_headings
        )
        return [
            _write_section(
                source, output_dir, source.stem, body, options, warnings, started
            )
        ]

    shift = 0
    if options.promote_headings:
        levels = [node.level for node in selected if node.level > 0]
        shift = min(levels) - 1 if levels else 0

    results: list[ConversionResult] = []
    for position, node in enumerate(selected, start=1):
        body = section_text(
            markdown, node, promote=options.promote_headings, shift=shift
        )
        name = f"{source.stem} - {position:02d} {slugify_title(node.title)}"
        results.append(
            _write_section(
                source, output_dir, name, body, options, warnings, started, node
            )
        )
    return results


def _write_section(
    source: Path,
    output_dir: Path,
    stem: str,
    body: str,
    options: ConversionOptions,
    warnings: Sequence[str],
    started: float,
    node: OutlineNode | None = None,
) -> ConversionResult:
    label = source.name if node is None else f"{source.name} › {node.title}"
    if not body.strip():
        return ConversionResult(
            source=source,
            status=STATUS_SKIPPED,
            message=f"{label}: mục rỗng, không tạo file.",
            duration=time.perf_counter() - started,
        )

    if options.add_title_heading and not body.lstrip().startswith("# "):
        body = f"# {node.title if node is not None else source.stem}\n\n{body}"

    try:
        destination = target_path(
            Path(f"{stem}.md"), output_dir, overwrite=options.overwrite
        )
        destination.write_text(body.rstrip() + "\n", encoding="utf-8", newline="\n")
    except Exception as exc:  # noqa: BLE001
        return ConversionResult(
            source=source,
            status=STATUS_ERROR,
            message=f"{type(exc).__name__}: {exc}",
            duration=time.perf_counter() - started,
        )

    return ConversionResult(
        source=source,
        status=STATUS_SUCCESS,
        output=destination,
        message=f"Đã trích xuất: {label}",
        warnings=list(warnings),
        duration=time.perf_counter() - started,
    )


# ----------------------------------------------------------------- excel


def _clean_columns(frame: pd.DataFrame) -> pd.DataFrame:
    columns = []
    for index, name in enumerate(frame.columns):
        text = "" if name is None else str(name)
        if text.startswith("Unnamed:") or text.strip().lower() == "nan":
            text = ""
        columns.append(text.strip() or f"Cột {index + 1}")
    frame = frame.copy()
    frame.columns = columns
    return frame


def _format_value(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if value is pd.NaT:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, (_dt.datetime, pd.Timestamp)):
        if value.hour == value.minute == value.second == 0:
            return value.strftime("%Y-%m-%d")
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, _dt.date):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, _dt.time):
        return value.strftime("%H:%M:%S")

    text = str(value).strip()
    return text.replace("\r\n", "<br>").replace("\n", "<br>").replace("|", "\\|")


def _read_workbook(source: Path, engine: str) -> dict:
    return pd.read_excel(source, sheet_name=None, engine=engine, dtype=object)


def excel_to_markdown(
    source: Path,
    options: ConversionOptions | None = None,
    upgrader: LegacyUpgrader | None = None,
) -> tuple[str, list[str]]:
    """Convert any workbook, including the binary .xls format.

    xlrd reads .xls natively and is preferred: it needs no Office install and
    does not spawn a second process. Excel or LibreOffice only step in when
    xlrd chokes on the file.
    """
    options = options or ConversionOptions()
    source = Path(source)
    warnings: list[str] = []
    sheets = None
    legacy = source.suffix.lower() in LEGACY_EXCEL

    if legacy and not sniff(source).is_modern and can_read_xls():
        try:
            sheets = _read_workbook(source, "xlrd")
        except Exception as exc:  # noqa: BLE001 - fall through to Excel
            warnings.append(f"xlrd không đọc được ({exc}); thử qua Excel/LibreOffice.")

    if sheets is None and legacy:
        owns = upgrader is None
        upgrader = upgrader or LegacyUpgrader()
        try:
            upgraded, more = upgrader.upgrade(source)
            warnings.extend(more)
            sheets = _read_workbook(upgraded, "openpyxl")
        finally:
            if owns:
                upgrader.close()

    if sheets is None:
        sheets = _read_workbook(source, "openpyxl")

    return _sheets_to_markdown(sheets, source.stem, options, warnings)


def _sheets_to_markdown(
    sheets: dict, title: str, options: ConversionOptions, warnings: list[str]
) -> tuple[str, list[str]]:
    """Render one Markdown table per sheet into a single document."""
    parts: list[str] = []
    if options.add_title_heading:
        parts.append(f"# {title}")

    heading = "#" * max(1, min(6, options.sheet_heading_level))

    if not sheets:
        warnings.append("Workbook không có sheet nào.")

    for name, frame in sheets.items():
        parts.append(f"{heading} {name}")
        frame = frame.dropna(axis=0, how="all").dropna(axis=1, how="all")
        if frame.empty:
            parts.append("> Sheet không có dữ liệu.")
            warnings.append(f"Sheet '{name}' trống.")
            continue
        frame = _clean_columns(frame)
        formatted = frame.map(_format_value)
        try:
            # disable_numparse keeps values byte-for-byte: tabulate would
            # otherwise turn "007" into 7 and long IDs into scientific notation.
            table = formatted.to_markdown(
                index=False, tablefmt="pipe", disable_numparse=True
            )
        except Exception as exc:  # noqa: BLE001 - fall back rather than lose the sheet
            warnings.append(f"Sheet '{name}': dùng bảng dự phòng ({exc}).")
            table = _fallback_table(formatted)
        parts.append(table)

    return "\n\n".join(parts).strip() + "\n", warnings


def _fallback_table(frame: pd.DataFrame) -> str:
    header = [str(column) for column in frame.columns]
    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(["---"] * len(header)) + " |",
    ]
    for row in frame.itertuples(index=False, name=None):
        lines.append("| " + " | ".join(str(cell) for cell in row) + " |")
    return "\n".join(lines)


# --------------------------------------------------------------- dispatcher


def convert_file(
    source: str | Path,
    output_dir: str | Path,
    options: ConversionOptions | None = None,
    upgrader: LegacyUpgrader | None = None,
) -> ConversionResult:
    """Convert one file. Never raises: failures come back as an error result."""
    options = options or ConversionOptions()
    source = Path(source)
    output_dir = Path(output_dir)
    started = time.perf_counter()

    def finish(status: str, message: str = "", output: Path | None = None, warns=None):
        return ConversionResult(
            source=source,
            status=status,
            output=output,
            message=message,
            warnings=list(warns or []),
            duration=time.perf_counter() - started,
        )

    try:
        if not source.exists():
            return finish(STATUS_ERROR, "Không tìm thấy file.")
        if not source.is_file():
            return finish(STATUS_SKIPPED, "Không phải file.")

        suffix = source.suffix.lower()
        if suffix not in SUPPORTED_EXTENSIONS:
            return finish(STATUS_SKIPPED, f"Bỏ qua định dạng {suffix or '(không rõ)'}.")

        destination = target_path(source, output_dir, overwrite=options.overwrite)

        if suffix in WORD_EXTENSIONS:
            image_dir = destination.parent / f"{destination.stem}_images"
            markdown, warnings = word_to_markdown(source, options, image_dir, upgrader)
        else:
            markdown, warnings = excel_to_markdown(source, options, upgrader)

        destination.write_text(markdown, encoding="utf-8", newline="\n")
        return finish(STATUS_SUCCESS, "Chuyển đổi thành công.", destination, warnings)

    except LegacyConversionError as exc:
        return finish(STATUS_ERROR, str(exc))
    except PermissionError:
        return finish(STATUS_ERROR, "Không có quyền truy cập (file đang mở?).")
    except FileNotFoundError:
        return finish(STATUS_ERROR, "Không tìm thấy file.")
    except Exception as exc:  # noqa: BLE001 - keep the batch alive
        return finish(STATUS_ERROR, f"{type(exc).__name__}: {exc}")


def convert_many(
    sources: Sequence[str | Path],
    output_dir: str | Path,
    options: ConversionOptions | None = None,
    on_progress: ProgressCallback | None = None,
    cancel_event: threading.Event | None = None,
) -> list[ConversionResult]:
    """Convert a batch, reporting progress after each file.

    One LegacyUpgrader is shared across the batch so Word or Excel is started
    at most once, and only if a legacy file actually turns up.
    """
    options = options or ConversionOptions()
    results: list[ConversionResult] = []
    total = len(sources)
    needs_legacy = any(
        Path(source).suffix.lower() in LEGACY_EXTENSIONS for source in sources
    )
    upgrader = LegacyUpgrader() if needs_legacy else None

    try:
        for index, source in enumerate(sources, start=1):
            if cancel_event is not None and cancel_event.is_set():
                result = ConversionResult(
                    source=Path(source), status=STATUS_CANCELLED, message="Đã huỷ."
                )
            else:
                result = convert_file(source, output_dir, options, upgrader)

            results.append(result)
            if on_progress is not None:
                try:
                    on_progress(index, total, result)
                except Exception:  # noqa: BLE001 - UI callbacks must not break the batch
                    pass
    finally:
        if upgrader is not None:
            upgrader.close()

    return results


def summarize(results: Sequence[ConversionResult]) -> dict[str, int]:
    summary = {
        STATUS_SUCCESS: 0,
        STATUS_ERROR: 0,
        STATUS_SKIPPED: 0,
        STATUS_CANCELLED: 0,
    }
    for result in results:
        summary[result.status] = summary.get(result.status, 0) + 1
    return summary
