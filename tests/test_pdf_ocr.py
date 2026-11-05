"""Tests for the two PDFs that used to come out wrong.

A table that rules its columns and not its rows collapsed into a single row -
every date of a bank statement in one cell - and a scanned page came out empty
because there was no text in it to read. Both are built here from scratch: the
PDFs are written byte by byte so the tests need no sample files, and the scan
is the same page rasterised into a picture.

The OCR tests skip themselves on a machine with no recogniser installed.
"""

from __future__ import annotations

import io
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.ocr import (  # noqa: E402
    can_ocr,
    has_pypdfium2,
    has_vietocr,
    ocr_languages,
    ocr_support_note,
    pick_language,
)
from src.pdf import can_read_pdf, pdf_to_markdown  # noqa: E402

CAN_READ = can_read_pdf()
CAN_OCR = can_ocr()

# The columns of the statement, and the rows printed between them.
COLUMNS = [40, 130, 220, 400, 540]
HEADINGS = ["Ngay giao dich", "Ngay he thong", "Chi tiet", "So tien (VND)"]
ENTRIES = [
    ["03/08", "03/08", "Thanh toan hoa don", "1.200.000"],
    ["04/08", "04/08", "Chuyen khoan den", "5.000.000"],
    ["06/08", "06/08", "Rut tien ATM", "2.000.000"],
    ["08/08", "09/08", "Phi dich vu", "11.000"],
]


def _page(content: bytes) -> bytes:
    """A one-page PDF around a content stream, with a working xref table."""
    objects = [
        b"<</Type/Catalog/Pages 2 0 R>>",
        b"<</Type/Pages/Kids[3 0 R]/Count 1>>",
        b"<</Type/Page/Parent 2 0 R/MediaBox[0 0 595 842]"
        b"/Resources<</Font<</F1 4 0 R>>>>/Contents 5 0 R>>",
        b"<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>",
        b"<</Length %d>>\nstream\n" % len(content) + content + b"\nendstream",
    ]
    body, offsets = b"%PDF-1.4\n", []
    for number, obj in enumerate(objects, start=1):
        offsets.append(len(body))
        body += b"%d 0 obj\n" % number + obj + b"\nendobj\n"
    start = len(body)
    body += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objects) + 1)
    for offset in offsets:
        body += b"%010d 00000 n \n" % offset
    return body + b"trailer<</Size %d/Root 1 0 R>>\nstartxref\n%d\n%%%%EOF\n" % (
        len(objects) + 1,
        start,
    )


def statement_pdf(target: Path) -> Path:
    """A statement: column rules drawn, row rules left to alignment alone."""
    out = ["0.5 w"]
    top, bottom = 760, 640
    for x in COLUMNS:
        out.append(f"{x} {top} m {x} {bottom} l S")
    out.append(f"{COLUMNS[0]} {top} m {COLUMNS[-1]} {top} l S")
    out.append(f"{COLUMNS[0]} {bottom} m {COLUMNS[-1]} {bottom} l S")
    y = 745
    for row in [HEADINGS, *ENTRIES]:
        out.append("BT /F1 9 Tf")
        for x, cell in zip(COLUMNS, row):
            out.append(f"1 0 0 1 {x + 4} {y} Tm ({cell}) Tj")
        out.append("ET")
        y -= 22
    target.write_bytes(_page("\n".join(out).encode("latin-1")))
    return target


def prose_pdf(target: Path) -> Path:
    """A title, two numbered headings, two paragraphs and a bullet list."""
    blocks = [
        (18, ["BAO CAO KY THUAT QUY III"]),
        (13, ["1. Muc tieu"]),
        (
            10,
            [
                "Tai lieu nay mo ta ket qua kiem thu he thong trong quy vua qua,",
                "bao gom pham vi kiem thu va cac loi da phat hien trong ky.",
            ],
        ),
        (13, ["2. Pham vi"]),
        (
            10,
            [
                "- Kiem thu chuc nang tren ba moi truong khac nhau",
                "- Kiem thu hieu nang voi hai nghin nguoi dung dong thoi",
            ],
        ),
    ]
    out: list[str] = []
    y = 780
    for size, lines in blocks:
        y -= size + 8
        for line in lines:
            out.append(f"BT /F1 {size} Tf 1 0 0 1 60 {y} Tm ({line}) Tj ET")
            y -= size + 4
    target.write_bytes(_page("\n".join(out).encode("latin-1")))
    return target


def scan_of(source: Path, target: Path, dpi: int = 200) -> Path:
    """The same page as a picture: what a scanner would have handed over."""
    import pypdfium2

    document = pypdfium2.PdfDocument(str(source))
    try:
        image = document[0].render(scale=dpi / 72).to_pil().convert("RGB")
    finally:
        document.close()
    buffer = io.BytesIO()
    image.save(buffer, format="PDF", resolution=dpi)
    target.write_bytes(buffer.getvalue())
    return target


VIETNAMESE = [
    (44, "BÁO CÁO KỸ THUẬT QUÝ III"),
    (30, "1. Mục tiêu"),
    (24, "Tài liệu này mô tả kết quả kiểm thử hệ thống trong quý vừa qua,"),
    (24, "bao gồm phạm vi kiểm thử và các lỗi đã phát hiện trong kỳ."),
    (30, "2. Phạm vi"),
    (24, "- Kiểm thử chức năng trên ba môi trường khác nhau"),
    (24, "- Rà soát bảo mật theo danh mục nội bộ"),
]
# The real font only (msttcorefonts on Linux): VietOCR misreads look-alikes.
ARIAL = next((p for p in (Path("C:/Windows/Fonts/arial.ttf"), Path("/usr/share/fonts/truetype/msttcorefonts/Arial.ttf"), Path("/usr/share/fonts/truetype/msttcorefonts/arial.ttf")) if p.exists()), Path("C:/Windows/Fonts/arial.ttf"))


def vietnamese_scan(target: Path, dpi: int = 200) -> Path:
    """A page of accented Vietnamese, drawn straight to a picture.

    A hand-written PDF cannot carry the diacritics - Helvetica has no glyph
    for them - so this one is painted with a real font instead. The marks are
    the point: they make an OCR box taller than the letters alone, which is
    what used to turn half a paragraph into headings.
    """
    from PIL import Image, ImageDraw, ImageFont

    image = Image.new("RGB", (1654, 900), "white")
    draw = ImageDraw.Draw(image)
    y = 60
    for size, text in VIETNAMESE:
        draw.text(
            (120, y), text, font=ImageFont.truetype(str(ARIAL), size), fill="black"
        )
        y += size + 26
    buffer = io.BytesIO()
    image.save(buffer, format="PDF", resolution=dpi)
    target.write_bytes(buffer.getvalue())
    return target


def rows_of(markdown: str) -> list[list[str]]:
    """The cells of the first Markdown table in a document."""
    rows = []
    for line in markdown.splitlines():
        line = line.strip()
        if line.startswith("|") and set(line) - set("| -"):
            rows.append([cell.strip() for cell in line.strip("|").split("|")])
    return rows


class RowlessTableTests(unittest.TestCase):
    """A table ruled down its columns only must not collapse into one row."""

    def setUp(self):
        if not CAN_READ:
            self.skipTest("cần pdfplumber để đọc PDF")
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.markdown, _ = pdf_to_markdown(
            statement_pdf(self.dir / "sao kê.pdf"), extract_images=False
        )

    def tearDown(self):
        self.tmp.cleanup()

    def test_every_entry_gets_its_own_row(self):
        rows = rows_of(self.markdown)
        self.assertEqual(len(rows), len(ENTRIES) + 1, self.markdown)

    def test_no_cell_swallows_a_whole_column(self):
        for row in rows_of(self.markdown):
            for cell in row:
                self.assertNotIn("<br>", cell, self.markdown)

    def test_the_columns_keep_their_order(self):
        rows = rows_of(self.markdown)
        self.assertEqual(rows[0], HEADINGS)
        self.assertEqual(rows[1], ENTRIES[0])

    def test_a_dated_row_is_not_split_across_columns(self):
        rows = rows_of(self.markdown)
        self.assertEqual([row[0] for row in rows[1:]], [e[0] for e in ENTRIES])


class OcrSupportTests(unittest.TestCase):
    """What the module reports on a machine, whatever it has installed."""

    def test_the_support_note_is_one_finished_sentence(self):
        note = ocr_support_note()
        self.assertTrue(note.startswith("OCR:"), note)
        self.assertTrue(note.endswith("."), note)

    def test_the_languages_are_bcp_47_tags(self):
        for tag in ocr_languages():
            self.assertRegex(tag, r"^[a-zA-Z]{2,3}(-[A-Za-z0-9]+)*$")

    def test_vietnamese_is_preferred_when_it_is_there(self):
        self.assertEqual(pick_language(["en-US", "vi-VN"]), "vi-VN")
        self.assertEqual(pick_language(["fr-FR", "en-US"]), "en-US")
        self.assertEqual(pick_language(["fr-FR"]), "fr-FR")
        self.assertIsNone(pick_language([]))

    def test_a_scan_without_ocr_is_an_error_that_names_the_reason(self):
        if not CAN_READ or not has_pypdfium2():
            self.skipTest("cần pdfplumber và pypdfium2 để dựng bản scan")
        from src.pdf import PdfConversionError

        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            scan = scan_of(prose_pdf(folder / "a.pdf"), folder / "scan.pdf")
            with self.assertRaises(PdfConversionError) as caught:
                pdf_to_markdown(scan, extract_images=False, ocr=False)
        self.assertIn("scan", str(caught.exception))


@unittest.skipUnless(CAN_OCR and CAN_READ, "máy chưa có engine OCR nào")
class OcrLayoutTests(unittest.TestCase):
    """A scan has to come out shaped like the page it was made from."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        folder = Path(cls.tmp.name)
        cls.prose, _ = pdf_to_markdown(
            scan_of(prose_pdf(folder / "prose.pdf"), folder / "prose-scan.pdf"),
            extract_images=False,
        )
        cls.statement, cls.warnings = pdf_to_markdown(
            scan_of(statement_pdf(folder / "sk.pdf"), folder / "sk-scan.pdf"),
            extract_images=False,
        )

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_the_page_is_not_empty(self):
        self.assertTrue(self.prose.strip())

    def test_the_biggest_line_becomes_the_top_heading(self):
        self.assertRegex(self.prose, r"(?m)^# \S")

    def test_the_smaller_headings_sit_below_it(self):
        self.assertRegex(self.prose, r"(?m)^## 1\.")
        self.assertRegex(self.prose, r"(?m)^## 2\.")

    def test_the_bullet_list_stays_a_list(self):
        bullets = [line for line in self.prose.splitlines() if line.startswith("- ")]
        self.assertEqual(len(bullets), 2, self.prose)

    def test_a_wrapped_paragraph_is_one_block_not_two(self):
        blocks = [b for b in self.prose.split("\n\n") if b.startswith("Tai lieu")]
        self.assertEqual(len(blocks), 1, self.prose)

    def test_prose_is_never_mistaken_for_a_table(self):
        self.assertNotIn("|", self.prose)

    def test_a_scanned_table_comes_back_a_table(self):
        rows = rows_of(self.statement)
        self.assertEqual(len(rows), len(ENTRIES) + 1, self.statement)
        self.assertTrue(all(len(row) == len(HEADINGS) for row in rows), rows)

    def test_the_dates_land_in_the_first_column(self):
        rows = rows_of(self.statement)
        self.assertEqual([row[0] for row in rows[1:]], [e[0] for e in ENTRIES])

    def test_the_user_is_told_the_page_was_read_by_ocr(self):
        self.assertTrue(any("OCR" in warning for warning in self.warnings))

    def test_a_missing_vietnamese_recogniser_is_said_out_loud(self):
        if has_vietocr():
            self.skipTest("máy này đọc tiếng Việt bằng VietOCR")
        if any(tag.lower().startswith("vi") for tag in ocr_languages()):
            self.skipTest("máy này có sẵn gói tiếng Việt của Windows")
        self.assertTrue(
            any("tiếng Việt" in warning for warning in self.warnings), self.warnings
        )

    def test_the_warning_names_who_read_the_page(self):
        reader = "VietOCR" if has_vietocr() else "en"
        self.assertTrue(
            any(reader in warning for warning in self.warnings), self.warnings
        )


@unittest.skipUnless(CAN_OCR and CAN_READ, "máy chưa có engine OCR nào")
class AccentedScanTests(unittest.TestCase):
    """Diacritics must not be read as a change of type size.

    A mark above a letter makes the OCR box taller, so an accented paragraph
    measures as tall as a small heading. The structure is what is checked
    here, never the letters: without a Vietnamese recogniser installed the
    words themselves come back mangled, and that is the engine's business.
    """

    @classmethod
    def setUpClass(cls):
        if not ARIAL.exists():
            raise unittest.SkipTest("cần font Arial để vẽ trang tiếng Việt")
        cls.tmp = tempfile.TemporaryDirectory()
        cls.markdown, _ = pdf_to_markdown(
            vietnamese_scan(Path(cls.tmp.name) / "scan.pdf"), extract_images=False
        )

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_the_page_is_read_at_all(self):
        self.assertGreater(len(self.markdown.split()), 20, self.markdown)

    def test_the_accents_survive_where_vietocr_is_here(self):
        if not has_vietocr():
            self.skipTest("thiếu mô hình VietOCR")
        for phrase in ("Tài liệu này mô tả kết quả kiểm thử", "Rà soát bảo mật"):
            self.assertIn(phrase, self.markdown)

    def test_only_the_title_becomes_a_heading(self):
        headings = [
            line for line in self.markdown.splitlines() if line.startswith("#")
        ]
        self.assertEqual(len(headings), 1, self.markdown)
        self.assertTrue(headings[0].startswith("# "), headings)

    def test_a_paragraph_never_turns_into_a_heading(self):
        self.assertNotRegex(self.markdown, r"(?m)^#{2,} ")

    def test_a_bullet_keeps_its_bullet(self):
        bullets = [
            line for line in self.markdown.splitlines() if line.startswith("- ")
        ]
        self.assertEqual(len(bullets), 2, self.markdown)

    def test_a_dash_never_drifts_to_the_end_of_its_line(self):
        for line in self.markdown.splitlines():
            self.assertFalse(line.rstrip().endswith(" -"), self.markdown)

    def test_the_wrapped_paragraph_stays_one_block(self):
        prose = [
            block
            for block in self.markdown.split("\n\n")
            if not block.startswith("#") and "- " not in block
        ]
        # The two printed lines hold two dozen words between them; a paragraph
        # left broken in half would leave nothing longer than a dozen.
        longest = max((len(block.split()) for block in prose), default=0)
        self.assertGreaterEqual(longest, 18, self.markdown)


if __name__ == "__main__":
    unittest.main(verbosity=2)
