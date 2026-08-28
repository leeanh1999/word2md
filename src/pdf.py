"""PDF, in both directions.

A PDF says where ink goes, not what the text means, so the two directions work
in completely different ways:

    .md  -> a Word document (`md_to_docx.py`) -> exported as PDF by Word or
            LibreOffice, so every Word option applies to the page as well.
    .pdf -> read here with pdfplumber: lines are grouped back into paragraphs,
            font sizes give the headings, ruled areas come back as tables, and
            embedded images are written out next to the Markdown.

Word can also reflow a PDF into a .docx, but it took nearly seven minutes for a
250 KB file in testing - it stops on a "Word will now convert your PDF" notice
that no automation flag turns off - so reading stays in Python, where the same
file takes about a second.
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

from .legacy import LegacyConversionError, find_soffice, has_msoffice

PDF_EXTENSIONS = {".pdf"}
PDF_MAGIC = b"%PDF-"

# A gap taller than this many line heights starts a new paragraph.
_PARAGRAPH_GAP = 1.55
# How much bigger than the body text a line has to be to count as a heading.
# Word sets Heading 1 only a point above a 13 pt body, so the margin is small.
_HEADING_RATIO = 1.04
# A heading is a short line; a long one is just emphasised text.
_HEADING_MAX_WORDS = 20
# A heading does not end mid-sentence.
_SENTENCE_END = ".,;:"

_BULLET_RE = re.compile(r"^([•●▪◦‣⁃*+]|[-–—])\s+(?P<rest>.+)$")
_ORDERED_RE = re.compile(r"^(?P<number>\d{1,3})[.)]\s+(?P<rest>.+)$")
# "2.1.3 Tên mục": the numbering says how deep the heading is.
_NUMBERED_RE = re.compile(r"^(\d{1,3}(?:\.\d{1,3})+)\s+\S")
_CONTROL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

_IMAGE_EXTENSIONS = {"png", "jpg", "jpeg", "gif", "tiff", "bmp", "webp"}


class PdfConversionError(LegacyConversionError):
    """This machine cannot do the PDF conversion that was asked for."""


# --------------------------------------------------------------- what we have


def has_pypdf() -> bool:
    try:
        import pypdf  # noqa: F401
    except Exception:  # noqa: BLE001
        return False
    return True


def has_pdfplumber() -> bool:
    try:
        import pdfplumber  # noqa: F401
    except Exception:  # noqa: BLE001
        return False
    return True


def has_word() -> bool:
    return has_msoffice("Word.Application")


def can_read_pdf() -> bool:
    return has_pdfplumber()


def can_write_pdf() -> bool:
    """Only a real Office engine can lay a page out."""
    return has_word() or find_soffice() is not None


def pdf_backends() -> list[str]:
    found = []
    if has_pdfplumber():
        found.append("pdfplumber (đọc PDF)")
    if has_pypdf():
        found.append("pypdf (ảnh trong PDF)")
    if has_word():
        found.append("Microsoft Word (xuất PDF)")
    if find_soffice():
        found.append("LibreOffice (xuất PDF)")
    return found


def pdf_support_note() -> str:
    """One line telling the user what PDF support looks like here."""
    if can_read_pdf() and can_write_pdf():
        return "PDF: đọc trực tiếp, xuất qua Microsoft Word hoặc LibreOffice."
    if can_read_pdf():
        return "PDF: đọc được; muốn xuất PDF cần Microsoft Word hoặc LibreOffice."
    if can_write_pdf():
        return "PDF: xuất được; chưa đọc được PDF (thiếu pdfplumber)."
    return "PDF: cần cài pdfplumber để đọc và Microsoft Word/LibreOffice để xuất."


def looks_like_pdf(source: str | Path) -> bool:
    try:
        with open(source, "rb") as handle:
            return handle.read(len(PDF_MAGIC)) == PDF_MAGIC
    except OSError:
        return False


# ------------------------------------------------------------------- reading


@dataclass
class _Line:
    text: str
    size: float
    bold: bool
    top: float
    bottom: float

    @property
    def height(self) -> float:
        return max(1.0, self.bottom - self.top)


@dataclass
class _Page:
    number: int
    items: list = field(default_factory=list)  # ("line", _Line) | ("table", rows)


def _open(source: Path):
    try:
        import pdfplumber
    except Exception as exc:  # noqa: BLE001
        raise PdfConversionError(
            "Không đọc được PDF: máy chưa cài pdfplumber "
            "(pip install -r requirements.txt)."
        ) from exc

    try:
        return pdfplumber.open(str(source))
    except Exception as exc:  # noqa: BLE001 - a locked or broken file
        message = str(exc)
        if "password" in message.lower() or "encrypt" in message.lower():
            raise PdfConversionError("PDF được bảo vệ bằng mật khẩu.") from exc
        raise PdfConversionError(f"Không mở được PDF: {message}") from exc


def _clean(text: str) -> str:
    return re.sub(r"[ \t]+", " ", _CONTROL_CHARS.sub(" ", text)).strip()


def _inside(word: dict, boxes) -> bool:
    """Whether an object sits inside one of the given boxes."""
    if "top" not in word or "x0" not in word:
        return False
    middle = (word["top"] + word.get("bottom", word["top"])) / 2
    centre = (word["x0"] + word.get("x1", word["x0"])) / 2
    return any(
        left <= centre <= right and top <= middle <= bottom
        for left, top, right, bottom in boxes
    )


def _line_from_chars(text: str, chars, top: float, bottom: float) -> _Line | None:
    """One line of text, with the size and weight most of it is set in."""
    text = _clean(text)
    if not text:
        return None
    sizes: Counter = Counter()
    bold = 0
    for char in chars:
        sizes[round(float(char.get("size") or 0), 1)] += 1
        name = str(char.get("fontname") or "").lower()
        if "bold" in name or "black" in name or "heavy" in name:
            bold += 1
    return _Line(
        text=text,
        size=max(sizes) if sizes else 0.0,
        bold=bool(chars) and bold * 2 > len(chars),
        top=top,
        bottom=bottom,
    )


def _table_rows(table) -> list[list[str]]:
    try:
        extracted = table.extract()
    except Exception:  # noqa: BLE001 - a table we cannot read is skipped
        return []
    rows = []
    for row in extracted:
        cells = [_clean(cell or "").replace("\n", "<br>") for cell in row]
        if any(cells):
            rows.append(cells)
    return rows


def _page_items(page) -> _Page:
    """One page as lines of text and tables, in the order they are printed."""
    try:
        tables = page.find_tables()
    except Exception:  # noqa: BLE001 - tables are a bonus, text is not
        tables = []

    boxes = [table.bbox for table in tables]
    # Whatever a table already covers must not be read a second time as text.
    body = page.filter(lambda obj: not _inside(obj, boxes)) if boxes else page
    try:
        found = body.extract_text_lines(strip=True, return_chars=True)
    except Exception:  # noqa: BLE001 - a page we cannot read leaves no lines
        found = []

    items: list[tuple[float, str, object]] = []
    for entry in found:
        line = _line_from_chars(
            entry.get("text", ""),
            entry.get("chars") or [],
            float(entry.get("top", 0.0)),
            float(entry.get("bottom", 0.0)),
        )
        if line is not None:
            items.append((line.top, "line", line))
    for table in tables:
        rows = _table_rows(table)
        if rows:
            items.append((table.bbox[1], "table", rows))

    items.sort(key=lambda item: item[0])
    return _Page(number=page.page_number, items=[(kind, payload) for _, kind, payload in items])


def _body_size(pages: list[_Page]) -> float:
    """The size most of the text is set in - the baseline for headings."""
    counter: Counter = Counter()
    for page in pages:
        for kind, payload in page.items:
            if kind == "line" and payload.size:
                counter[round(payload.size, 1)] += len(payload.text)
    if not counter:
        return 0.0
    return counter.most_common(1)[0][0]


def _heading_sizes(pages: list[_Page], body: float) -> dict[float, int]:
    """Rank the sizes bigger than the body text, largest first, as H1..H6.

    Ranking beats measuring: a document whose headings are only a point or two
    above its body text still has a hierarchy, and it comes out as one.
    """
    if not body:
        return {}
    bigger = sorted(
        {
            round(payload.size, 1)
            for page in pages
            for kind, payload in page.items
            if kind == "line" and payload.size >= body * _HEADING_RATIO
        },
        reverse=True,
    )
    return {size: min(6, index + 1) for index, size in enumerate(bigger)}


def _line_kind(line: _Line, levels: dict[float, int]) -> tuple[str, object]:
    """What a line of text is: a heading, a list item, or plain text."""
    bullet = _BULLET_RE.match(line.text)
    if bullet:
        return "bullet", bullet.group("rest")
    ordered = _ORDERED_RE.match(line.text)

    short = len(line.text.split()) <= _HEADING_MAX_WORDS
    level = levels.get(round(line.size, 1), 0)
    if not level and line.bold and short and line.text[-1] not in _SENTENCE_END:
        # Body-size text, bold and alone on its line: a run-in heading, one step
        # below the smallest real one.
        level = min(6, (max(levels.values()) if levels else 0) + 1)
    if level and short:
        # Its own numbering knows the depth better than the type size does.
        numbered = _NUMBERED_RE.match(line.text)
        if numbered:
            level = max(level, min(6, numbered.group(1).count(".") + 1))
        return "heading", level

    if ordered:
        return "ordered", (int(ordered.group("number")), ordered.group("rest"))
    return "text", None


class _Markdown:
    """Collects blocks and renders them as Markdown."""

    def __init__(self):
        self.blocks: list[str] = []
        self._paragraph: list[str] = []
        self._list = False

    def flush(self) -> None:
        if self._paragraph:
            self.blocks.append(" ".join(self._paragraph))
            self._paragraph = []

    def text(self, line: str) -> None:
        self._paragraph.append(line)
        self._list = False

    def block(self, text: str) -> None:
        self.flush()
        if text:
            self.blocks.append(text)
        self._list = False

    def item(self, text: str) -> None:
        """One list item, kept in the same block as the item before it."""
        self.flush()
        if self._list and self.blocks:
            self.blocks[-1] += f"\n{text}"
        else:
            self.blocks.append(text)
        self._list = True

    def render(self) -> str:
        self.flush()
        return "\n\n".join(block for block in self.blocks if block.strip())


def _table_markdown(rows: list[list[str]]) -> str:
    width = max(len(row) for row in rows)
    padded = [row + [""] * (width - len(row)) for row in rows]
    header = [cell.replace("|", r"\|") for cell in padded[0]]
    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(["---"] * width) + " |",
    ]
    for row in padded[1:]:
        lines.append("| " + " | ".join(cell.replace("|", r"\|") for cell in row) + " |")
    return "\n".join(lines)


def _render(
    pages: list[_Page], levels: dict[float, int], images: dict[int, list[str]]
) -> str:
    out = _Markdown()
    previous: _Line | None = None
    list_state = ""  # "bullet", "ordered" or ""

    for page in pages:
        for kind, payload in page.items:
            if kind == "table":
                out.block(_table_markdown(payload))
                previous, list_state = None, ""
                continue

            line: _Line = payload
            what, extra = _line_kind(line, levels)

            if what == "heading":
                out.block("#" * min(6, int(extra)) + f" {line.text}")
                previous, list_state = line, ""
                continue

            if what == "bullet":
                out.item(f"- {extra}")
                previous, list_state = line, "bullet"
                continue

            if what == "ordered":
                number, rest = extra
                out.item(f"{number}. {rest}")
                previous, list_state = line, "ordered"
                continue

            # Plain text: a big enough vertical gap means a new paragraph, and
            # so does a line that continues nothing.
            if previous is not None and list_state == "":
                gap = line.top - previous.bottom
                if gap > previous.height * _PARAGRAPH_GAP:
                    out.flush()
            else:
                out.flush()
            list_state = ""
            out.text(line.text)
            previous = line

        for link in images.get(page.number, []):
            out.block(link)
        previous, list_state = None, ""

    return out.render()


def _extract_images(source: Path, image_dir: Path, warnings: list[str]) -> dict[int, list[str]]:
    """Write every embedded image next to the Markdown, keyed by page number."""
    try:
        from pypdf import PdfReader
    except Exception:  # noqa: BLE001
        warnings.append("Chưa cài pypdf nên không tách được ảnh trong PDF.")
        return {}

    found: dict[int, list[str]] = {}
    try:
        reader = PdfReader(str(source))
    except Exception as exc:  # noqa: BLE001 - the text still went through
        warnings.append(f"Không đọc được ảnh trong PDF ({exc}).")
        return {}

    count = 0
    for index, page in enumerate(reader.pages, start=1):
        try:
            embedded = list(page.images)
        except Exception as exc:  # noqa: BLE001 - one bad page, not one bad file
            warnings.append(f"Trang {index}: không tách được ảnh ({exc}).")
            continue
        for image in embedded:
            count += 1
            suffix = Path(image.name or "").suffix.lower().lstrip(".")
            if suffix not in _IMAGE_EXTENSIONS:
                suffix = "png"
            target = image_dir / f"image{count}.{suffix}"
            try:
                image_dir.mkdir(parents=True, exist_ok=True)
                target.write_bytes(image.data)
            except Exception as exc:  # noqa: BLE001
                warnings.append(f"Không lưu được ảnh {target.name} ({exc}).")
                continue
            link = f"![]({image_dir.name}/{target.name})"
            found.setdefault(index, []).append(link)
    return found


def pdf_to_markdown(
    source: str | Path,
    image_dir: Path | None = None,
    extract_images: bool = True,
    title: str | None = None,
) -> tuple[str, list[str]]:
    """Read a PDF as Markdown, returning (markdown, warnings)."""
    source = Path(source)
    warnings: list[str] = []
    if not looks_like_pdf(source):
        warnings.append(f"{source.name} không có chữ ký PDF; vẫn thử đọc.")

    with _open(source) as pdf:
        pages = [_page_items(page) for page in pdf.pages]

    if not pages:
        raise PdfConversionError("PDF không có trang nào.")

    images: dict[int, list[str]] = {}
    if extract_images and image_dir is not None:
        images = _extract_images(source, Path(image_dir), warnings)

    body = _body_size(pages)
    markdown = _render(pages, _heading_sizes(pages, body), images)

    if not markdown.strip():
        raise PdfConversionError(
            "Không trích xuất được văn bản: PDF này có thể chỉ là ảnh scan."
        )
    if title and not markdown.lstrip().startswith("# "):
        markdown = f"# {title}\n\n{markdown}"

    return markdown + "\n", warnings


__all__ = [
    "PDF_EXTENSIONS",
    "PdfConversionError",
    "can_read_pdf",
    "can_write_pdf",
    "has_pdfplumber",
    "has_pypdf",
    "has_word",
    "looks_like_pdf",
    "pdf_backends",
    "pdf_support_note",
    "pdf_to_markdown",
]
