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
# A cell holding this many line breaks is not a cell but a whole column, merged
# because the table rules its columns and not its rows.
_COLLAPSED_CELL_BREAKS = 2
# Two words share a row while their tops sit within this share of a line height.
_ROW_TOLERANCE = 0.6
# A page holding less text than this was never really read: it is a picture of
# a page, and OCR is the only way into it.
_SCANNED_PAGE_CHARS = 12
# In OCR output, a gap this many line heights wide is a column, not a space.
_OCR_COLUMN_GAP = 1.2
# A column has to hold its place in this share of a block's rows to be real.
_OCR_COLUMN_SHARE = 0.6
# Fewer aligned rows than this is a coincidence, not a table.
_OCR_MIN_TABLE_ROWS = 3
# How much taller than the body an OCR'd line has to be to count as a heading.
# The margin dwarfs the 1.04 a text PDF gets away with: a font size is exact,
# an OCR box is only as tall as the letters and marks that landed in it.
_OCR_HEADING_MARGIN = 1.2
# How wide a size band is, as a share of the body size, once a line is over
# that margin and the only question left is which heading level it is.
_OCR_SIZE_BAND = 0.25

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


def _column_bounds(table) -> list[float]:
    """Left-to-right x positions of the table's column rules."""
    if not table.columns:
        return []
    edges = {round(column.bbox[0], 1) for column in table.columns}
    edges.add(round(table.columns[-1].bbox[2], 1))
    return sorted(edges)


def _looks_collapsed(rows) -> bool:
    """Whether a whole column came back as one cell for want of row rules."""
    return any(
        (cell or "").count("\n") >= _COLLAPSED_CELL_BREAKS
        for row in rows
        for cell in row
    )


def _merge_wrapped(rows: list[list[str]]) -> list[list[str]]:
    """Fold a wrapped line back into the row above it.

    A line that starts with an empty first column is the tail of the row before
    it - a description too long to fit on one line - not a row of its own.
    """
    merged: list[list[str]] = []
    for row in rows:
        if merged and row and not row[0] and any(row):
            for index, cell in enumerate(row):
                if cell:
                    merged[-1][index] = f"{merged[-1][index]}\n{cell}".strip("\n")
            continue
        merged.append(list(row))
    return merged


def _rows_from_words(table) -> list[list[str]]:
    """Rebuild a table that rules its columns but not its rows.

    A bank statement or an invoice usually draws the column rules and leaves the
    rows to alignment alone. pdfplumber then reads the whole block as one row,
    so every date of the month arrives in a single cell. Slicing the printed
    words by the column rules puts each line back on a row of its own.
    """
    bounds = _column_bounds(table)
    if len(bounds) < 3:  # one column is nothing to rebuild
        return []

    page = table.page
    left, top, right, bottom = table.bbox
    try:
        # A hair of margin: a rule sitting exactly on the crop edge is dropped.
        words = page.crop(
            (
                max(0.0, left - 1),
                max(0.0, top - 1),
                min(float(page.width), right + 1),
                min(float(page.height), bottom + 1),
            )
        ).extract_words()
    except Exception:  # noqa: BLE001 - keep whatever pdfplumber already read
        return []

    def middle(word) -> float:
        return (float(word["top"]) + float(word["bottom"])) / 2

    rows: list[list[str]] = []
    current: list[list[str]] = []
    line: float | None = None
    for word in sorted(words, key=lambda item: (middle(item), item["x0"])):
        height = max(1.0, float(word["bottom"]) - float(word["top"]))
        if line is None or middle(word) - line > height * _ROW_TOLERANCE:
            if current:
                rows.append([" ".join(cell).strip() for cell in current])
            current = [[] for _ in range(len(bounds) - 1)]
            line = middle(word)
        centre = (float(word["x0"]) + float(word["x1"])) / 2
        for index in range(len(bounds) - 1):
            if bounds[index] <= centre <= bounds[index + 1]:
                current[index].append(word["text"])
                break
    if current:
        rows.append([" ".join(cell).strip() for cell in current])
    return _merge_wrapped(rows)


def _table_rows(table) -> list[list[str]]:
    try:
        extracted = table.extract()
    except Exception:  # noqa: BLE001 - a table we cannot read is skipped
        return []
    if _looks_collapsed(extracted):
        rebuilt = _rows_from_words(table)
        # Only take the rebuild if it actually pulled the block apart.
        if len(rebuilt) > len(extracted):
            extracted = rebuilt
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


# ----------------------------------------------------------- scanned pages


def _page_text_length(page: _Page) -> int:
    """How much text a page gave up - tables included.

    A page that is nothing but a table has no loose lines on it at all, and
    counting only those would send a perfectly readable page off to OCR.
    """
    total = 0
    for kind, payload in page.items:
        if kind == "line":
            total += len(payload.text)
        else:
            total += sum(len(cell) for row in payload for cell in row)
    return total


def _line_from_words(row) -> _Line | None:
    """One printed line, rebuilt from the words OCR found on it."""
    row = sorted(row, key=lambda word: word.x0)
    text = _clean(" ".join(word.text for word in row))
    if not text:
        return None
    heights = sorted(word.height for word in row)
    return _Line(
        text=text,
        # The median box: one tall letter must not promote the whole line.
        size=heights[len(heights) // 2],
        bold=False,  # OCR reports no weight
        top=min(word.top for word in row),
        bottom=max(word.bottom for word in row),
    )


def _word_rows(words) -> list[list]:
    """Group the recognised words back into the lines they were printed on.

    The middle of a word places it, not its top edge: a hyphen is a box three
    pixels tall sitting halfway down the line, and lining boxes up by their
    tops files it as a line of its own - which is how a bullet loses its
    bullet.
    """
    rows: list[list] = []
    row: list = []
    middle = 0.0
    for word in sorted(words, key=lambda item: ((item.top + item.bottom) / 2, item.x0)):
        centre = (word.top + word.bottom) / 2
        if row and centre - middle > row[0].height * _ROW_TOLERANCE:
            rows.append(sorted(row, key=lambda item: item.x0))
            row = []
        if not row:
            middle = centre
        row.append(word)
    if row:
        rows.append(sorted(row, key=lambda item: item.x0))
    return rows


def _row_gaps(row: list) -> list[tuple[float, float]]:
    """The stretches of a row too wide to be the space between two words."""
    if len(row) < 2:
        return []
    heights = sorted(word.height for word in row)
    height = heights[len(heights) // 2]
    return [
        (left.x1, right.x0)
        for left, right in zip(row, row[1:])
        if right.x0 - left.x1 > height * _OCR_COLUMN_GAP
    ]


def _shared_columns(rows: list[list]) -> list[float]:
    """The white lanes that run down every row - the column boundaries.

    Where a gap falls says less than how far it stretches: one row may break
    after a short word and the next after a long one, and the two still leave
    the same lane of white between the columns. So the lanes are intersected
    rather than averaged, and a lane counts as a column only when it survives
    in most of the rows. A paragraph leaves no lane at all and comes back
    empty, which is what keeps ordinary prose out of a table.
    """
    gaps = [_row_gaps(row) for row in rows]
    edges = sorted({value for row in gaps for span in row for value in span})
    if len(edges) < 2:
        return []

    needed = max(2, round(len(rows) * _OCR_COLUMN_SHARE))
    columns: list[float] = []
    lane: list[tuple[float, float]] = []
    for left, right in zip(edges, edges[1:]):
        middle = (left + right) / 2
        covered = sum(
            1 for row in gaps if any(start <= middle <= stop for start, stop in row)
        )
        if covered >= needed:
            lane.append((left, right))
        elif lane:
            columns.append((lane[0][0] + lane[-1][1]) / 2)
            lane = []
    if lane:
        columns.append((lane[0][0] + lane[-1][1]) / 2)
    return columns


def _split_rows(rows: list[list], columns: list[float]) -> list[list[str]]:
    """Cut every row at the column boundaries."""
    bounds = [float("-inf"), *columns, float("inf")]
    out: list[list[str]] = []
    for row in rows:
        cells: list[list[str]] = [[] for _ in range(len(bounds) - 1)]
        for word in row:
            centre = (word.x0 + word.x1) / 2
            for index in range(len(bounds) - 1):
                if bounds[index] <= centre < bounds[index + 1]:
                    cells[index].append(word.text)
                    break
        out.append([_clean(" ".join(cell)) for cell in cells])
    return _merge_wrapped(out)


def _ocr_items(words) -> list[tuple[str, object]]:
    """A scanned page as lines and tables, the same shape a read page has."""
    rows = _word_rows(words)
    items: list[tuple[str, object]] = []
    index = 0
    while index < len(rows):
        # How far the aligned block that starts here runs.
        run = index
        while run < len(rows) and _row_gaps(rows[run]):
            run += 1

        columns = (
            _shared_columns(rows[index:run])
            if run - index >= _OCR_MIN_TABLE_ROWS
            else []
        )
        if columns:
            items.append(("table", _split_rows(rows[index:run], columns)))
            index = run
            continue

        stop = max(run, index + 1)
        for row in rows[index:stop]:
            line = _line_from_words(row)
            if line is not None:
                items.append(("line", line))
        index = stop
    return items


def _weighted_median_size(lines: list[_Line]) -> float:
    """The size half the recognised text is set in, by the letter."""
    sizes = sorted((line.size, len(line.text)) for line in lines if line.size)
    if not sizes:
        return 0.0
    total = sum(weight for _, weight in sizes)
    seen = 0
    for size, weight in sizes:
        seen += weight
        if seen * 2 >= total:
            return size
    return sizes[-1][0]


def _quantise(lines: list[_Line]) -> None:
    """Pull the OCR sizes back to the few the page was really set in.

    OCR measures ink, not type. A box is as tall as what happens to be in it,
    so `hàng` stands taller than `caa` in the very same sentence, and in
    Vietnamese - where most words carry a mark - the spread swallows the step
    between body text and a small heading. Left alone, the size ranking reads
    that noise as a hierarchy six levels deep and scatters `###` through the
    middle of a paragraph.

    So everything within a margin of the middle size becomes the body, and
    only text that stands clearly above it stays a heading.
    """
    median = _weighted_median_size(lines)
    if not median:
        return
    band = median * _OCR_SIZE_BAND
    for line in lines:
        if line.size <= median * _OCR_HEADING_MARGIN:
            line.size = median
        else:
            line.size = round(line.size / band) * band


def _ocr_scanned_pages(
    source: Path,
    pages: list[_Page],
    language: str | None,
    warnings: list[str],
) -> list[int]:
    """Read every page that holds no text through OCR, in place.

    Returns the 1-based numbers of the pages that OCR actually filled in.
    """
    blank = [
        index
        for index, page in enumerate(pages)
        if _page_text_length(page) < _SCANNED_PAGE_CHARS
    ]
    if not blank:
        return []

    from .ocr import (
        OcrError,
        can_ocr,
        has_vietocr,
        ocr_languages,
        ocr_pdf_pages,
        ocr_support_note,
        pick_language,
    )

    if not can_ocr():
        warnings.append(
            f"{len(blank)} trang chỉ có ảnh, chưa OCR được. {ocr_support_note()}"
        )
        return []

    chosen = language or pick_language(ocr_languages())
    try:
        found = ocr_pdf_pages(source, blank, chosen)
    except OcrError as exc:
        warnings.append(str(exc))
        return []

    done: list[int] = []
    lines: list[_Line] = []
    for index, words in found.items():
        items = _ocr_items(words)
        if not items:
            continue
        pages[index] = _Page(number=pages[index].number, items=items)
        lines.extend(payload for kind, payload in items if kind == "line")
        done.append(pages[index].number)

    _quantise(lines)
    if done:
        reader = "VietOCR" if has_vietocr() else chosen
        warnings.append(f"Đã OCR {len(done)} trang ảnh scan (dò chữ {chosen}, đọc bằng {reader}).")
        if not has_vietocr() and chosen and not chosen.lower().startswith("vi"):
            warnings.append(
                f"Engine OCR không có gói nhận dạng tiếng Việt và thiếu mô hình VietOCR, "
                f"nên chữ có dấu đọc bằng {chosen} sẽ sai."
            )
    return done


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
    ocr: bool = True,
    ocr_language: str | None = None,
) -> tuple[str, list[str]]:
    """Read a PDF as Markdown, returning (markdown, warnings).

    A page that carries no text is a scan, and `ocr=True` reads it with the
    OCR engine Windows ships with. The words come back placed on the page, so
    the same layout pass runs over them and a scan keeps its headings,
    paragraphs, lists and tables.
    """
    source = Path(source)
    warnings: list[str] = []
    if not looks_like_pdf(source):
        warnings.append(f"{source.name} không có chữ ký PDF; vẫn thử đọc.")

    with _open(source) as pdf:
        pages = [_page_items(page) for page in pdf.pages]

    if not pages:
        raise PdfConversionError("PDF không có trang nào.")

    scanned = _ocr_scanned_pages(source, pages, ocr_language, warnings) if ocr else []

    images: dict[int, list[str]] = {}
    if extract_images and image_dir is not None:
        images = _extract_images(source, Path(image_dir), warnings)
        # The one picture on a scanned page is the page itself; the text of it
        # is already in the Markdown, so linking the scan as well says nothing.
        for number in scanned:
            images.pop(number, None)

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
