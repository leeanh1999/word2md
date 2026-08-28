"""Tests for both PDF directions.

Reading a PDF is pure Python and always runs. Writing one needs Word or
LibreOffice, so those tests - and the samples they need - skip themselves on a
machine that has neither.
"""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import src.pdf as pdf_module  # noqa: E402
from src.converter import (  # noqa: E402
    STATUS_SUCCESS,
    SUPPORTED_EXTENSIONS,
    ConversionOptions,
    collect_files,
    convert_file,
    output_suffix,
)
from src.legacy import LegacyUpgrader  # noqa: E402
from src.pdf import (  # noqa: E402
    PdfConversionError,
    can_read_pdf,
    can_write_pdf,
    looks_like_pdf,
    pdf_backends,
    pdf_support_note,
    pdf_to_markdown,
)
from make_samples import make_docx  # noqa: E402

CAN_WRITE = can_write_pdf()
CAN_READ = can_read_pdf()

# The smallest PDF that still says something: one page, one line of Helvetica.
MINIMAL_PDF = b"""%PDF-1.4
1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj
2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj
3 0 obj<</Type/Page/Parent 2 0 R/MediaBox[0 0 300 200]/Resources<</Font<</F1 4 0 R>>>>/Contents 5 0 R>>endobj
4 0 obj<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>endobj
5 0 obj<</Length 76>>stream
BT /F1 12 Tf 20 150 Td (Bao cao ky thuat) Tj ET
BT /F1 12 Tf 20 120 Td (Noi dung) Tj ET
endstream
endobj
trailer<</Root 1 0 R>>
"""


def write_minimal(target: Path) -> Path:
    target.write_bytes(MINIMAL_PDF)
    return target


class SupportTests(unittest.TestCase):
    def test_pdf_is_a_supported_extension(self):
        self.assertIn(".pdf", SUPPORTED_EXTENSIONS)

    def test_a_folder_scan_picks_pdfs_up(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            write_minimal(folder / "a.pdf")
            (folder / "b.md").write_text("# b\n", encoding="utf-8")
            found = {path.name for path in collect_files([folder])}
        self.assertEqual(found, {"a.pdf", "b.md"})

    def test_markdown_can_target_pdf(self):
        self.assertEqual(output_suffix(Path("a.md"), ".pdf"), ".pdf")
        self.assertEqual(output_suffix(Path("a.pdf"), ".pdf"), ".md", "a PDF reads as md")

    def test_the_support_note_says_something_useful(self):
        note = pdf_support_note()
        self.assertIn("PDF", note)
        self.assertTrue(note.endswith("."), note)

    def test_backends_are_listed(self):
        self.assertEqual(pdf_backends(), [b for b in pdf_backends() if b])

    def test_the_magic_number_is_checked(self):
        with tempfile.TemporaryDirectory() as tmp:
            good = write_minimal(Path(tmp) / "good.pdf")
            bad = Path(tmp) / "bad.pdf"
            bad.write_text("not a pdf at all", encoding="utf-8")
            self.assertTrue(looks_like_pdf(good))
            self.assertFalse(looks_like_pdf(bad))
            self.assertFalse(looks_like_pdf(Path(tmp) / "missing.pdf"))


@unittest.skipUnless(CAN_READ, "cần pdfplumber để đọc PDF")
class ReadTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)

    def tearDown(self):
        self.tmp.cleanup()

    def test_text_comes_out_in_reading_order(self):
        source = write_minimal(self.dir / "tối giản.pdf")
        markdown, warnings = pdf_to_markdown(source)
        self.assertIn("Bao cao ky thuat", markdown)
        self.assertIn("Noi dung", markdown)
        self.assertEqual(warnings, [])

    def test_a_title_is_added_when_the_document_has_no_heading(self):
        source = write_minimal(self.dir / "tối giản.pdf")
        markdown, _ = pdf_to_markdown(source, title="tối giản")
        self.assertTrue(markdown.startswith("# tối giản"), markdown[:40])

    def test_a_file_that_is_not_a_pdf_is_reported_not_crashed(self):
        source = self.dir / "giả.pdf"
        source.write_text("chỉ là văn bản", encoding="utf-8")
        with self.assertRaises(PdfConversionError):
            pdf_to_markdown(source)

    def test_a_missing_file_is_an_error_not_a_traceback(self):
        with self.assertRaises(PdfConversionError):
            pdf_to_markdown(self.dir / "không-có.pdf")

    def test_convert_file_writes_markdown_next_to_the_others(self):
        source = write_minimal(self.dir / "báo cáo.pdf")
        result = convert_file(source, self.dir / "out")
        self.assertEqual(result.status, STATUS_SUCCESS, result.message)
        self.assertEqual(result.output.name, "báo cáo.md")
        self.assertIn("Bao cao ky thuat", result.output.read_text(encoding="utf-8"))


@unittest.skipUnless(
    CAN_WRITE and CAN_READ, "cần Microsoft Word hoặc LibreOffice để xuất PDF"
)
class RoundTripTests(unittest.TestCase):
    """A Markdown file printed as PDF and read straight back."""

    MARKDOWN = (
        "# Bao cao quy III\n\n"
        "Doan mo dau cua tai lieu.\n\n"
        "## 1. Bang so lieu\n\n"
        "| Hang muc | Ke hoach |\n|---|---:|\n| Doanh thu | 1200 |\n\n"
        "## 2. Danh sach\n\n- mot\n- hai\n"
    )

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.dir = Path(cls.tmp.name)
        source = cls.dir / "bao cao.md"
        source.write_text(cls.MARKDOWN, encoding="utf-8")
        cls.result = convert_file(
            source, cls.dir / "pdf", ConversionOptions(markdown_target=".pdf")
        )
        cls.back = (
            convert_file(cls.result.output, cls.dir / "md")
            if cls.result.output
            else None
        )

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_the_pdf_is_written(self):
        self.assertEqual(self.result.status, STATUS_SUCCESS, self.result.message)
        self.assertEqual(self.result.output.suffix, ".pdf")
        self.assertTrue(looks_like_pdf(self.result.output))
        self.assertGreater(self.result.output.stat().st_size, 1000)

    def test_the_text_survives_the_trip(self):
        markdown = self.back.output.read_text(encoding="utf-8")
        self.assertIn("Doan mo dau cua tai lieu.", markdown)

    def test_the_headings_keep_their_hierarchy(self):
        markdown = self.back.output.read_text(encoding="utf-8")
        self.assertIn("# Bao cao quy III", markdown)
        self.assertIn("## 1. Bang so lieu", markdown)
        self.assertIn("## 2. Danh sach", markdown)

    def test_the_table_comes_back_as_a_table(self):
        markdown = self.back.output.read_text(encoding="utf-8")
        self.assertIn("| Hang muc | Ke hoach |", markdown)
        self.assertIn("| Doanh thu | 1200 |", markdown)

    def test_the_list_comes_back_as_a_list(self):
        markdown = self.back.output.read_text(encoding="utf-8")
        self.assertIn("- mot\n- hai", markdown)


@unittest.skipUnless(CAN_WRITE and CAN_READ, "cần Word/LibreOffice và pdfplumber")
class WordSampleTests(unittest.TestCase):
    """The full Word sample, printed as PDF and read back."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.dir = Path(cls.tmp.name)
        docx = make_docx(cls.dir / "test.docx")
        cls.pdf = cls.dir / "test.pdf"
        upgrader = LegacyUpgrader(cls.dir / "work")
        try:
            upgrader.to_pdf(docx, cls.pdf)
        finally:
            upgrader.close()
        cls.markdown, cls.warnings = pdf_to_markdown(
            cls.pdf, image_dir=cls.dir / "test_images", title="test"
        )

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_vietnamese_words_are_not_broken_apart(self):
        self.assertIn("Báo cáo kỹ thuật quý III", self.markdown)
        self.assertIn("Đoạn văn có chữ in đậm", self.markdown)

    def test_the_numbered_headings_are_headings(self):
        self.assertIn("## 1. Tổng quan", self.markdown)
        self.assertIn("## 2. Danh sách công việc", self.markdown)

    def test_deeper_numbering_gives_a_deeper_heading(self):
        self.assertIn("### A.1 Viết tắt", self.markdown)

    def test_the_bullet_list_is_one_block(self):
        self.assertIn("- Thu thập yêu cầu\n- Thiết kế kiến trúc", self.markdown)

    def test_the_numbered_list_keeps_its_numbers(self):
        self.assertIn("1. Chuẩn bị môi trường\n2. Chạy pipeline", self.markdown)

    def test_the_table_is_a_markdown_table_with_escaped_pipes(self):
        self.assertIn("| Hạng mục | Kế hoạch | Thực hiện | Ghi chú |", self.markdown)
        self.assertIn(r"Tốt \| ổn định", self.markdown)

    def test_the_table_text_is_not_repeated_as_prose(self):
        self.assertEqual(self.markdown.count("Doanh thu"), 1, self.markdown)

    def test_paragraphs_are_not_split_line_by_line(self):
        long_line = "Tài liệu này được sinh tự động để kiểm thử bộ chuyển đổi"
        self.assertIn(long_line, self.markdown)


@unittest.skipIf(CAN_WRITE, "chỉ kiểm tra máy không có Word/LibreOffice")
class NoEngineTests(unittest.TestCase):
    def test_asking_for_a_pdf_says_what_is_missing(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "a.md"
            source.write_text("# A\n", encoding="utf-8")
            result = convert_file(
                source, Path(tmp) / "out", ConversionOptions(markdown_target=".pdf")
            )
        self.assertNotEqual(result.status, STATUS_SUCCESS)
        self.assertIn("Microsoft Word", result.message)


class MissingLibraryTests(unittest.TestCase):
    """The message a machine without pdfplumber gets."""

    def test_reading_without_pdfplumber_explains_itself(self):
        import builtins

        real_import = builtins.__import__

        def blocked(name, *args, **kwargs):
            if name == "pdfplumber":
                raise ImportError("No module named 'pdfplumber'")
            return real_import(name, *args, **kwargs)

        builtins.__import__ = blocked
        try:
            with tempfile.TemporaryDirectory() as tmp:
                source = write_minimal(Path(tmp) / "a.pdf")
                with self.assertRaises(PdfConversionError) as caught:
                    pdf_to_markdown(source)
            self.assertIn("pdfplumber", str(caught.exception))
        finally:
            builtins.__import__ = real_import

    def test_the_support_note_names_what_is_missing(self):
        original = pdf_module.has_pdfplumber
        pdf_module.has_pdfplumber = lambda: False
        try:
            self.assertIn("pdfplumber", pdf_support_note())
        finally:
            pdf_module.has_pdfplumber = original


if __name__ == "__main__":
    unittest.main(verbosity=2)
