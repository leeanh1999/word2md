"""Tests for the navigation outline and partial extraction."""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.converter import (  # noqa: E402
    STATUS_ERROR,
    STATUS_SUCCESS,
    ConversionOptions,
    convert_docx_sections,
    load_outline,
)
from src.outline import (  # noqa: E402
    build_outline,
    extract_sections,
    index_nodes,
    iter_nodes,
    outline_to_text,
    section_text,
    shift_headings,
    slugify_title,
    top_level_selection,
)
from make_samples import build_all  # noqa: E402

DOC = "\n".join(
    [
        "Lời mở đầu.",
        "",
        "# Chương 1",
        "",
        "Nội dung chương 1.",
        "",
        "## 1.1 Mục con",
        "",
        "Chi tiết 1.1.",
        "",
        "### 1.1.1 Mục cháu",
        "",
        "Chi tiết sâu.",
        "",
        "## 1.2 Mục con khác",
        "",
        "Chi tiết 1.2.",
        "",
        "# Chương 2",
        "",
        "Nội dung chương 2.",
        "",
    ]
)


class BuildOutlineTests(unittest.TestCase):
    def setUp(self):
        self.roots = build_outline(DOC)
        self.nodes = index_nodes(self.roots)

    def test_preamble_becomes_its_own_node(self):
        self.assertEqual(self.roots[0].level, 0)
        self.assertTrue(self.roots[0].is_preamble)
        self.assertEqual(self.roots[0].node_id, "1")

    def test_tree_shape_matches_heading_levels(self):
        titles = [node.title for node in self.roots]
        self.assertEqual(titles[1:], ["Chương 1", "Chương 2"])
        chapter = self.nodes["2"]
        self.assertEqual([c.title for c in chapter.children], ["1.1 Mục con", "1.2 Mục con khác"])
        self.assertEqual([c.title for c in chapter.children[0].children], ["1.1.1 Mục cháu"])

    def test_ids_are_hierarchical(self):
        self.assertEqual(sorted(self.nodes), ["1", "2", "2.1", "2.1.1", "2.2", "3"])

    def test_section_ends_before_next_same_or_higher_heading(self):
        chapter_one = self.nodes["2"]
        lines = DOC.split("\n")
        self.assertEqual(lines[chapter_one.start], "# Chương 1")
        self.assertEqual(lines[chapter_one.end], "# Chương 2")

    def test_no_preamble_when_document_starts_with_heading(self):
        roots = build_outline("# A\n\ntext\n")
        self.assertEqual(len(roots), 1)
        self.assertFalse(roots[0].is_preamble)

    def test_document_without_headings_has_only_preamble(self):
        roots = build_outline("chỉ là văn bản thường\n")
        self.assertEqual(len(roots), 1)
        self.assertTrue(roots[0].is_preamble)

    def test_empty_document_has_no_nodes(self):
        self.assertEqual(build_outline(""), [])

    def test_hash_inside_code_fence_is_not_a_heading(self):
        markdown = "# Real\n\n```\n# not a heading\n```\n\n## Also real\n"
        titles = [node.title for node in iter_nodes(build_outline(markdown))]
        self.assertEqual(titles, ["Real", "Also real"])

    def test_outline_to_text_lists_every_node(self):
        text = outline_to_text(self.roots)
        self.assertEqual(len(text.splitlines()), len(self.nodes))
        self.assertIn("2.1.1", text)


class ExtractionTests(unittest.TestCase):
    def setUp(self):
        self.roots = build_outline(DOC)
        self.nodes = index_nodes(self.roots)

    def test_selecting_a_node_pulls_in_its_subtree(self):
        text = extract_sections(DOC, ["2"], roots=self.roots)
        self.assertIn("# Chương 1", text)
        self.assertIn("1.1.1 Mục cháu", text)
        self.assertNotIn("Chương 2", text)

    def test_selecting_a_leaf_excludes_siblings(self):
        text = extract_sections(DOC, ["2.2"], roots=self.roots, promote=False)
        self.assertIn("## 1.2 Mục con khác", text)
        self.assertNotIn("1.1 Mục con", text)

    def test_multiple_sections_keep_document_order(self):
        text = extract_sections(DOC, ["3", "2.1"], roots=self.roots, promote=False)
        self.assertLess(text.index("1.1 Mục con"), text.index("Chương 2"))

    def test_descendant_of_selected_parent_is_not_duplicated(self):
        text = extract_sections(DOC, ["2", "2.1"], roots=self.roots)
        self.assertEqual(text.count("1.1.1 Mục cháu"), 1)

    def test_promote_lifts_shallowest_selection_to_h1(self):
        text = extract_sections(DOC, ["2.1"], roots=self.roots, promote=True)
        self.assertIn("# 1.1 Mục con", text)
        self.assertNotIn("## 1.1 Mục con", text)
        self.assertIn("## 1.1.1 Mục cháu", text)

    def test_promote_uses_one_shift_across_mixed_levels(self):
        text = extract_sections(DOC, ["2.1", "3"], roots=self.roots, promote=True)
        self.assertIn("# Chương 2", text)
        self.assertIn("# 1.1 Mục con", text)
        self.assertIn("## 1.1.1 Mục cháu", text)

    def test_preamble_extraction(self):
        text = extract_sections(DOC, ["1"], roots=self.roots)
        self.assertEqual(text.strip(), "Lời mở đầu.")

    def test_unknown_ids_are_ignored(self):
        self.assertEqual(extract_sections(DOC, ["nope"], roots=self.roots), "")
        text = extract_sections(DOC, ["nope", "3"], roots=self.roots)
        self.assertIn("Chương 2", text)

    def test_top_level_selection_drops_covered_children(self):
        selected = top_level_selection(self.roots, ["2.1.1", "2", "3"])
        self.assertEqual([n.node_id for n in selected], ["2", "3"])

    def test_section_text_single_node(self):
        text = section_text(DOC, self.nodes["3"])
        self.assertTrue(text.startswith("# Chương 2"))

    def test_shift_headings_never_goes_above_h1(self):
        self.assertEqual(shift_headings("## a", 5), "# a")

    def test_shift_headings_leaves_fenced_content_alone(self):
        text = "## a\n\n```\n## inside\n```\n"
        self.assertIn("## inside", shift_headings(text, 1))

    def test_slugify_strips_illegal_characters(self):
        self.assertEqual(slugify_title('A/B: "c" <d>?'), "AB c d")
        self.assertEqual(slugify_title("   "), "section")
        self.assertLessEqual(len(slugify_title("x" * 200)), 80)


class DocxSectionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.samples = build_all(Path(cls.tmp.name) / "samples", legacy=False)
        cls.docx = cls.samples["docx"]
        cls.outline = load_outline(cls.docx)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def _out(self, name: str) -> Path:
        return Path(self.tmp.name) / name

    def test_outline_matches_word_navigation_pane(self):
        titles = [node.title for node in iter_nodes(self.outline.roots)]
        self.assertEqual(
            titles,
            [
                "Báo cáo kỹ thuật quý III",
                "1. Tổng quan",
                "2. Danh sách công việc",
                "3. Các bước thực hiện",
                "4. Bảng số liệu",
                "5. Trích dẫn",
                "Phụ lục",
                "A. Thuật ngữ",
                "A.1 Viết tắt",
                "A.2 Tham chiếu",
                "B. Liên hệ",
            ],
        )

    def test_outline_nesting_depth(self):
        nodes = self.outline.nodes
        appendix = next(n for n in self.outline.roots if n.title == "Phụ lục")
        glossary = appendix.children[0]
        self.assertEqual(glossary.title, "A. Thuật ngữ")
        self.assertEqual([c.title for c in glossary.children], ["A.1 Viết tắt", "A.2 Tham chiếu"])
        self.assertIn(glossary.children[0].node_id, nodes)

    def test_load_outline_writes_nothing(self):
        folder = self._out("untouched")
        folder.mkdir()
        load_outline(self.docx)
        self.assertEqual(list(folder.iterdir()), [])

    def test_combined_export(self):
        folder = self._out("combined")
        appendix = next(n for n in self.outline.roots if n.title == "Phụ lục")
        results = convert_docx_sections(self.docx, folder, [appendix.node_id])
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0].status, STATUS_SUCCESS, results[0].message)
        text = results[0].output.read_text(encoding="utf-8")
        self.assertIn("# Phụ lục", text)
        self.assertIn("A.1 Viết tắt", text)
        self.assertNotIn("Bảng số liệu", text)

    def test_split_export_creates_one_file_per_section(self):
        folder = self._out("split")
        report = self.outline.roots[0]
        ids = [child.node_id for child in report.children[:3]]
        options = ConversionOptions(split_sections=True)
        results = convert_docx_sections(self.docx, folder, ids, options)
        self.assertEqual(len(results), 3)
        self.assertTrue(all(r.status == STATUS_SUCCESS for r in results))
        names = sorted(p.name for p in folder.glob("*.md"))
        self.assertEqual(len(names), 3)
        self.assertTrue(names[0].startswith("test - 01 "))

    def test_section_ids_stay_valid_between_preview_and_export(self):
        folder = self._out("stable")
        node_id = self.outline.roots[0].children[1].node_id
        preview = self.outline.preview([node_id])
        results = convert_docx_sections(self.docx, folder, [node_id])
        exported = results[0].output.read_text(encoding="utf-8")
        self.assertEqual(preview.strip(), exported.strip())

    def test_empty_selection_is_an_error(self):
        results = convert_docx_sections(self.docx, self._out("none"), [])
        self.assertEqual(results[0].status, STATUS_ERROR)

    def test_unknown_selection_is_an_error(self):
        results = convert_docx_sections(self.docx, self._out("bad"), ["99"])
        self.assertEqual(results[0].status, STATUS_ERROR)

    def test_load_outline_rejects_non_word_files(self):
        with self.assertRaises(ValueError):
            load_outline(self.samples["xlsx"])

    def test_corrupt_docx_reports_error(self):
        results = convert_docx_sections(
            self.samples["corrupt"], self._out("corrupt"), ["1"]
        )
        self.assertEqual(results[0].status, STATUS_ERROR)


if __name__ == "__main__":
    unittest.main(verbosity=2)
