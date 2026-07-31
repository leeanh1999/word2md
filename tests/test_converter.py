"""Automated tests: run with `python -m unittest discover -s tests -v`."""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.converter import (  # noqa: E402
    STATUS_ERROR,
    STATUS_SKIPPED,
    STATUS_SUCCESS,
    ConversionOptions,
    collect_files,
    convert_file,
    convert_many,
    summarize,
    target_path,
)
from src.html_to_markdown import html_to_markdown  # noqa: E402
from make_samples import build_all  # noqa: E402


class HtmlToMarkdownTests(unittest.TestCase):
    def test_headings_and_inline(self):
        markdown = html_to_markdown(
            "<h1>Title</h1><h3>Sub</h3>"
            "<p>Plain <strong>bold</strong> and <em>italic</em> and "
            '<a href="https://x.dev">link</a>.</p>'
        )
        self.assertIn("# Title", markdown)
        self.assertIn("### Sub", markdown)
        self.assertIn("**bold**", markdown)
        self.assertIn("*italic*", markdown)
        self.assertIn("[link](https://x.dev)", markdown)

    def test_emphasis_keeps_spaces_outside_markers(self):
        markdown = html_to_markdown("<p>a<strong> bold </strong>b</p>")
        self.assertIn("a **bold** b", markdown)

    def test_nested_lists(self):
        markdown = html_to_markdown(
            "<ul><li>one<ul><li>one.a</li></ul></li><li>two</li></ul>"
        )
        lines = markdown.splitlines()
        self.assertEqual(lines[0], "- one")
        self.assertEqual(lines[1], "  - one.a")
        self.assertEqual(lines[2], "- two")

    def test_ordered_list_numbering(self):
        markdown = html_to_markdown("<ol><li>a</li><li>b</li><li>c</li></ol>")
        self.assertEqual(markdown.splitlines()[:3], ["1. a", "2. b", "3. c"])

    def test_table_with_header_cells(self):
        markdown = html_to_markdown(
            "<table><tr><th>A</th><th>B</th></tr>"
            "<tr><td>1</td><td>2</td></tr></table>"
        )
        lines = markdown.splitlines()
        self.assertEqual(lines[0], "| A | B |")
        self.assertEqual(lines[1], "| --- | --- |")
        self.assertEqual(lines[2], "| 1 | 2 |")

    def test_table_escapes_pipes_and_newlines(self):
        markdown = html_to_markdown(
            "<table><tr><td>a|b</td><td>x<br>y</td></tr>"
            "<tr><td>c</td><td>d</td></tr></table>"
        )
        self.assertIn(r"a\|b", markdown)
        self.assertIn("x<br>y", markdown)

    def test_ragged_table_is_padded(self):
        markdown = html_to_markdown(
            "<table><tr><td>a</td><td>b</td><td>c</td></tr>"
            "<tr><td>d</td></tr></table>"
        )
        for line in markdown.splitlines():
            self.assertEqual(line.count("|"), 4, line)

    def test_blockquote_and_code(self):
        markdown = html_to_markdown("<blockquote><p>quoted</p></blockquote>")
        self.assertIn("> quoted", markdown)
        markdown = html_to_markdown("<pre>line1\nline2</pre>")
        self.assertIn("```\nline1\nline2\n```", markdown)

    def test_entities_and_unclosed_tags(self):
        markdown = html_to_markdown("<p>a &amp; b &lt;c&gt; &quot;d&quot;</p>")
        self.assertIn('a & b <c> "d"', markdown)
        self.assertIn("dangling", html_to_markdown("<p>dangling"))

    def test_empty_input(self):
        self.assertEqual(html_to_markdown("").strip(), "")


class ConverterTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.samples_dir = Path(cls.tmp.name) / "samples"
        cls.out = Path(cls.tmp.name) / "out"
        cls.samples = build_all(cls.samples_dir, legacy=False)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_docx_conversion(self):
        result = convert_file(self.samples["docx"], self.out)
        self.assertEqual(result.status, STATUS_SUCCESS, result.message)
        text = result.output.read_text(encoding="utf-8")
        self.assertIn("# Báo cáo kỹ thuật quý III", text)
        self.assertIn("## 1. Tổng quan", text)
        self.assertIn("**chữ in đậm**", text)
        self.assertIn("*chữ in nghiêng*", text)
        self.assertIn("- Thu thập yêu cầu", text)
        self.assertIn("1. Chuẩn bị môi trường", text)
        self.assertIn("| --- |", text)
        self.assertIn("Doanh thu", text)
        self.assertIn("> Chất lượng", text)

    def test_xlsx_all_sheets_in_one_file(self):
        result = convert_file(self.samples["xlsx"], self.out)
        self.assertEqual(result.status, STATUS_SUCCESS, result.message)
        text = result.output.read_text(encoding="utf-8")
        self.assertIn("## Doanh thu", text)
        self.assertIn("## Nhân sự", text)
        self.assertIn("## Trống", text)
        self.assertIn("1500000", text)
        self.assertIn("Trần Văn A", text)
        self.assertIn("2021-03-01", text)
        self.assertIn("Sheet không có dữ liệu.", text)

    def test_corrupt_file_reports_error_without_raising(self):
        result = convert_file(self.samples["corrupt"], self.out)
        self.assertEqual(result.status, STATUS_ERROR)
        self.assertTrue(result.message)

    def test_missing_file(self):
        result = convert_file(self.out / "nope.docx", self.out)
        self.assertEqual(result.status, STATUS_ERROR)

    def test_unsupported_extension_is_skipped(self):
        stray = Path(self.tmp.name) / "note.txt"
        stray.write_text("hello", encoding="utf-8")
        self.assertEqual(convert_file(stray, self.out).status, STATUS_SKIPPED)

    def test_batch_continues_after_failure(self):
        batch = [
            self.samples["corrupt"],
            self.samples["docx"],
            self.samples["xlsx"],
        ]
        seen = []
        results = convert_many(
            batch, self.out, on_progress=lambda i, t, r: seen.append((i, t))
        )
        counts = summarize(results)
        self.assertEqual(counts[STATUS_SUCCESS], 2)
        self.assertEqual(counts[STATUS_ERROR], 1)
        self.assertEqual(seen, [(1, 3), (2, 3), (3, 3)])

    def test_progress_callback_exception_does_not_break_batch(self):
        def bad_callback(*_args):
            raise RuntimeError("boom")

        results = convert_many(
            [self.samples["docx"]], self.out, on_progress=bad_callback
        )
        self.assertEqual(results[0].status, STATUS_SUCCESS)

    def test_collect_files_scans_folders_and_ignores_noise(self):
        (self.samples_dir / "~$test.docx").write_bytes(b"lock")
        (self.samples_dir / "readme.txt").write_text("x", encoding="utf-8")
        nested = self.samples_dir / "sub"
        nested.mkdir(exist_ok=True)
        (nested / "deep.xlsx").write_bytes(self.samples["xlsx"].read_bytes())

        found = collect_files([self.samples_dir])
        names = sorted(p.name for p in found)
        self.assertIn("deep.xlsx", names)
        self.assertIn("test.docx", names)
        self.assertNotIn("~$test.docx", names)
        self.assertNotIn("readme.txt", names)

        shallow = collect_files([self.samples_dir], recursive=False)
        self.assertNotIn("deep.xlsx", [p.name for p in shallow])

    def test_collect_files_deduplicates(self):
        docx = self.samples["docx"]
        found = collect_files([docx, docx, self.samples_dir])
        self.assertEqual(sum(1 for p in found if p.name == "test.docx"), 1)

    def test_no_overwrite_creates_unique_names(self):
        folder = Path(self.tmp.name) / "unique"
        first = convert_file(self.samples["docx"], folder)
        second = convert_file(self.samples["docx"], folder)
        self.assertNotEqual(first.output, second.output)
        self.assertTrue(second.output.name.endswith("(1).md"))

    def test_overwrite_reuses_same_name(self):
        folder = Path(self.tmp.name) / "over"
        options = ConversionOptions(overwrite=True)
        first = convert_file(self.samples["docx"], folder, options)
        second = convert_file(self.samples["docx"], folder, options)
        self.assertEqual(first.output, second.output)

    def test_target_path_creates_output_dir(self):
        folder = Path(self.tmp.name) / "made" / "deep"
        target_path(Path("a.docx"), folder)
        self.assertTrue(folder.is_dir())

    def test_title_heading_toggle(self):
        folder = Path(self.tmp.name) / "titles"
        with_title = convert_file(
            self.samples["xlsx"], folder, ConversionOptions(add_title_heading=True)
        )
        without = convert_file(
            self.samples["xlsx"], folder, ConversionOptions(add_title_heading=False)
        )
        self.assertTrue(
            with_title.output.read_text(encoding="utf-8").startswith("# test")
        )
        self.assertTrue(
            without.output.read_text(encoding="utf-8").startswith("## Doanh thu")
        )


if __name__ == "__main__":
    unittest.main(verbosity=2)
