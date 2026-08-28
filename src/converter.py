"""Core conversion engine: Word/Excel -> Markdown, and Markdown -> Word."""

from __future__ import annotations

import datetime as _dt
import io
import mimetypes
import re
import shutil
import tempfile
import threading
import time
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Callable, Iterable, Sequence
from urllib.parse import quote, unquote

import mammoth
import pandas as pd

from .attachments import PreparedDocument, prepare_attachments
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
from .md_to_docx import MARKDOWN_EXTENSIONS, DocxSettings, markdown_to_docx
from .md_to_xlsx import XlsxSettings, markdown_to_xlsx
from .pdf import PDF_EXTENSIONS, pdf_to_markdown
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
SUPPORTED_EXTENSIONS = (
    WORD_EXTENSIONS | EXCEL_EXTENSIONS | MARKDOWN_EXTENSIONS | PDF_EXTENSIONS
)
# What a Markdown file can be converted into.
MARKDOWN_TARGETS = {".docx", ".xlsx", ".pdf"}

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
    extract_attachments: bool = True
    overwrite: bool = False
    first_row_as_header: bool = True
    add_title_heading: bool = True
    sheet_heading_level: int = 2
    promote_headings: bool = True
    split_sections: bool = False
    # What a Markdown file turns into: ".docx" (Word) or ".xlsx" (Excel).
    markdown_target: str = ".docx"
    # Only used by the Markdown -> Word direction.
    docx: DocxSettings = field(default_factory=DocxSettings)
    # Only used by the Markdown -> Excel direction.
    xlsx: XlsxSettings = field(default_factory=XlsxSettings)


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


def output_suffix(source: Path, markdown_target: str = ".docx") -> str:
    """Markdown goes back to Office; everything else comes out as Markdown."""
    if source.suffix.lower() not in MARKDOWN_EXTENSIONS:
        return ".md"
    return markdown_target if markdown_target in MARKDOWN_TARGETS else ".docx"


def target_path(
    source: Path, output_dir: Path, overwrite: bool = False, suffix: str | None = None
) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    suffix = suffix or output_suffix(source)
    target = output_dir / f"{source.stem}{suffix}"
    if overwrite:
        return target
    counter = 1
    while target.exists():
        target = output_dir / f"{source.stem} ({counter}){suffix}"
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
    attachment_dir: Path | None = None,
) -> tuple[str, list[str]]:
    """Return (markdown, warnings) for a .docx file.

    `title` overrides the H1 derived from the file name, which matters when
    reading a temporary file upgraded from .doc.
    """
    options = options or ConversionOptions()
    warnings: list[str] = []

    def convert(handle):
        kwargs = {"style_map": DOCX_STYLE_MAP}
        if options.extract_images and image_dir is not None:
            kwargs["convert_image"] = _image_converter(image_dir, warnings)
        elif not options.extract_images:
            kwargs["convert_image"] = mammoth.images.img_element(
                lambda _image: {"src": ""}
            )
        with handle:
            return mammoth.convert_to_html(handle, **kwargs)

    prepared = PreparedDocument()
    if options.extract_attachments:
        prepared = prepare_attachments(source, attachment_dir)
        warnings.extend(prepared.warnings)

    if prepared.data is None:
        result = convert(source.open("rb"))
    else:
        try:
            result = convert(io.BytesIO(prepared.data))
        except Exception as exc:  # noqa: BLE001 - the document matters more
            warnings.append(f"Bỏ qua liên kết file đính kèm ({exc}); đọc bản gốc.")
            result = convert(source.open("rb"))

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
    attachment_dir: Path | None = None,
) -> tuple[str, list[str]]:
    """Convert .docx directly, or .doc after upgrading it to .docx."""
    options = options or ConversionOptions()
    source = Path(source)
    if source.suffix.lower() not in LEGACY_WORD:
        return docx_to_markdown(source, options, image_dir, None, attachment_dir)

    owns = upgrader is None
    upgrader = upgrader or LegacyUpgrader()
    try:
        try:
            upgraded, warnings = upgrader.upgrade(source)
        except LegacyConversionError as exc:
            markdown, warnings = _doc_text_fallback(source, options, exc)
            return markdown, warnings
        markdown, more = docx_to_markdown(
            upgraded, options, image_dir, source.stem, attachment_dir
        )
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

    if not node_ids:
        return _section_failed(source, "Chưa chọn mục nào để trích xuất.", started)

    # Reading the outline means converting the whole document, images and
    # attachments included. They go to a scratch folder first so that only the
    # ones the chosen sections actually point at reach the output folder.
    staging = Path(tempfile.mkdtemp(prefix="word2md-sections-"))
    try:
        return _export_sections(
            source, output_dir, node_ids, options, upgrader, staging, started
        )
    finally:
        shutil.rmtree(staging, ignore_errors=True)


def section_stems(
    roots: Sequence[OutlineNode],
    node_ids: Sequence[str],
    source: Path,
    split: bool,
) -> list[str]:
    """The file stems a section export would write, without exporting anything.

    Lets a caller check for name clashes before the work starts.
    """
    selected = top_level_selection(roots, node_ids)
    if not selected:
        return []
    if split or len(selected) == 1:
        return [slugify_title(node.title, source.stem) for node in selected]
    # Several sections merged into one file have no single heading to be named
    # after, so the document keeps its own name.
    return [source.stem]


def _section_failed(
    source: Path, message: str, started: float
) -> list[ConversionResult]:
    return [
        ConversionResult(
            source=source,
            status=STATUS_ERROR,
            message=message,
            duration=time.perf_counter() - started,
        )
    ]


def _export_sections(
    source: Path,
    output_dir: Path,
    node_ids: Sequence[str],
    options: ConversionOptions,
    upgrader: LegacyUpgrader | None,
    staging: Path,
    started: float,
) -> list[ConversionResult]:
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
        image_dir = staging / f"{source.stem}_images"
        attachment_dir = staging / f"{source.stem}_attachments"
        convert_options = replace(options, add_title_heading=False)
        readable, title, warnings = ensure_docx(source, upgrader)
        markdown, more = docx_to_markdown(
            readable, convert_options, image_dir, title, attachment_dir
        )
        warnings = warnings + more
        roots = build_outline(markdown)
        selected = top_level_selection(roots, node_ids)
        if not selected:
            return _section_failed(
                source, "Các mục đã chọn không còn tồn tại trong tài liệu.", started
            )
    except LegacyConversionError as exc:
        return _section_failed(source, str(exc), started)
    except PermissionError:
        return _section_failed(
            source, "Không có quyền truy cập (file đang mở?).", started
        )
    except Exception as exc:  # noqa: BLE001
        return _section_failed(source, f"{type(exc).__name__}: {exc}", started)

    staged = {
        image_dir.name: (image_dir, "_images"),
        attachment_dir.name: (attachment_dir, "_attachments"),
    }

    if not options.split_sections:
        body = extract_sections(
            markdown, node_ids, roots=roots, promote=options.promote_headings
        )
        alone = selected[0] if len(selected) == 1 else None
        stem = section_stems(roots, node_ids, source, split=False)[0]
        return [
            _write_section(
                source,
                output_dir,
                stem,
                body,
                options,
                warnings,
                started,
                alone,
                staged,
            )
        ]

    results: list[ConversionResult] = []
    stems = section_stems(roots, node_ids, source, split=True)
    for node, stem in zip(selected, stems):
        # Each file stands on its own, so its heading is promoted all the way
        # to H1 rather than by the amount the whole selection shares.
        body = section_text(markdown, node, promote=options.promote_headings)
        results.append(
            _write_section(
                source,
                output_dir,
                stem,
                body,
                options,
                warnings,
                started,
                node,
                staged,
            )
        )
    return results


# Markdown link target: the path inside ![alt](…) or [label](…).
_ASSET_LINK = re.compile(r"\]\(([^()\s]+)\)")


def _relocate_assets(
    body: str, destination: Path, staged: dict[str, tuple[Path, str]]
) -> str:
    """Copy the assets `body` points at into folders named after `destination`.

    Only the referenced files are copied, so exporting one section does not
    drag the whole document's images and attachments along with it.
    """
    folders: dict[str, str] = {}

    def replace(match: re.Match) -> str:
        folder, separator, name = match.group(1).partition("/")
        staged_dir, suffix = staged.get(folder, (None, ""))
        if staged_dir is None or not separator:
            return match.group(0)
        origin = staged_dir / unquote(name)
        if not origin.is_file():
            return match.group(0)

        target = folders.get(folder)
        if target is None:
            target = f"{destination.stem}{suffix}"
            folders[folder] = target
            (destination.parent / target).mkdir(parents=True, exist_ok=True)
        shutil.copy2(origin, destination.parent / target / origin.name)
        return f"]({quote(f'{target}/{origin.name}', safe='/')})"

    return _ASSET_LINK.sub(replace, body)


def _write_section(
    source: Path,
    output_dir: Path,
    stem: str,
    body: str,
    options: ConversionOptions,
    warnings: Sequence[str],
    started: float,
    node: OutlineNode | None = None,
    staged: dict[str, tuple[Path, str]] | None = None,
) -> ConversionResult:
    label = source.name if node is None else f"{source.name} › {node.title}"
    if not body.strip():
        return ConversionResult(
            source=source,
            status=STATUS_SKIPPED,
            message=f"{label}: mục rỗng, không tạo file.",
            duration=time.perf_counter() - started,
        )

    # Any leading heading already names the section; a second one would only
    # repeat it.
    if options.add_title_heading and not body.lstrip().startswith("#"):
        body = f"# {node.title if node is not None else source.stem}\n\n{body}"

    try:
        destination = target_path(
            Path(f"{stem}.md"), output_dir, overwrite=options.overwrite, suffix=".md"
        )
        if staged:
            body = _relocate_assets(body, destination, staged)
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


def markdown_to_pdf(
    source: Path,
    destination: Path,
    options: ConversionOptions | None = None,
    upgrader: LegacyUpgrader | None = None,
) -> list[str]:
    """Write a .md file as PDF, by way of the Word document it describes.

    Every option of the Markdown -> Word direction therefore applies to the
    page as well: font, size, paper, spacing, page breaks and the table of
    contents.
    """
    options = options or ConversionOptions()
    source, destination = Path(source), Path(destination)
    owns = upgrader is None
    upgrader = upgrader or LegacyUpgrader()
    staging = Path(tempfile.mkdtemp(prefix="word2md-pdf-"))
    try:
        interim = staging / f"{source.stem}.docx"
        warnings = markdown_to_docx(
            source,
            interim,
            embed_images=options.extract_images,
            add_title_heading=options.add_title_heading,
            settings=options.docx,
        )
        return warnings + upgrader.to_pdf(interim, destination)
    finally:
        shutil.rmtree(staging, ignore_errors=True)
        if owns:
            upgrader.close()


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

        destination = target_path(
            source,
            output_dir,
            overwrite=options.overwrite,
            suffix=output_suffix(source, options.markdown_target),
        )

        if suffix in MARKDOWN_EXTENSIONS:
            if destination.suffix.lower() == ".pdf":
                warnings = markdown_to_pdf(source, destination, options, upgrader)
                return finish(
                    STATUS_SUCCESS, "Chuyển đổi thành công.", destination, warnings
                )
            if destination.suffix.lower() == ".xlsx":
                warnings = markdown_to_xlsx(
                    source,
                    destination,
                    embed_images=options.extract_images,
                    add_title_heading=options.add_title_heading,
                    settings=options.xlsx,
                )
            else:
                warnings = markdown_to_docx(
                    source,
                    destination,
                    embed_images=options.extract_images,
                    add_title_heading=options.add_title_heading,
                    settings=options.docx,
                )
            return finish(STATUS_SUCCESS, "Chuyển đổi thành công.", destination, warnings)

        if suffix in PDF_EXTENSIONS:
            markdown, warnings = pdf_to_markdown(
                source,
                image_dir=destination.parent / f"{destination.stem}_images",
                extract_images=options.extract_images,
                title=source.stem if options.add_title_heading else None,
            )
        elif suffix in WORD_EXTENSIONS:
            image_dir = destination.parent / f"{destination.stem}_images"
            attachment_dir = destination.parent / f"{destination.stem}_attachments"
            markdown, warnings = word_to_markdown(
                source, options, image_dir, upgrader, attachment_dir
            )
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
    suffixes = [Path(source).suffix.lower() for source in sources]
    needs_office = any(suffix in LEGACY_EXTENSIONS for suffix in suffixes) or (
        options.markdown_target == ".pdf"
        and any(suffix in MARKDOWN_EXTENSIONS for suffix in suffixes)
    )
    upgrader = LegacyUpgrader() if needs_office else None

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
