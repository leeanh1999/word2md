"""Tests for files attached to a Word document as OLE objects."""

from __future__ import annotations

import struct
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.attachments import prepare_attachments, read_ole10_native  # noqa: E402
from src.converter import (  # noqa: E402
    STATUS_SUCCESS,
    ConversionOptions,
    convert_file,
    load_outline,
)
from make_samples import ATTACHMENTS, _ole10_native, build_all  # noqa: E402


class Ole10NativeTests(unittest.TestCase):
    def test_reads_name_and_payload(self):
        stream = _ole10_native("bao cao.pdf", b"%PDF-1.4" + b"x" * 100)
        self.assertEqual(
            read_ole10_native(stream), ("bao cao.pdf", b"%PDF-1.4" + b"x" * 100)
        )

    def test_prefers_the_utf16_name_over_the_ansi_one(self):
        """The ANSI copy mangles anything outside the system code page."""
        stream = _ole10_native("báo cáo.pdf", b"%PDF-1.4")
        name, _ = read_ole10_native(stream)
        self.assertEqual(name, "báo cáo.pdf")

    def test_strips_the_directory_from_the_name(self):
        stream = _ole10_native(r"C:\tai lieu\bao cao.pdf", b"%PDF-1.4")
        self.assertEqual(read_ole10_native(stream)[0], "bao cao.pdf")

    def test_names_an_extensionless_file_after_its_content(self):
        stream = _ole10_native("khong duoi", b"%PDF-1.4 ...")
        self.assertEqual(read_ole10_native(stream)[0], "khong duoi.pdf")

    def test_rejects_garbage(self):
        self.assertIsNone(read_ole10_native(b"not an ole package"))
        self.assertIsNone(read_ole10_native(struct.pack("<IH", 2, 7)))


class AttachmentExtractionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.samples = build_all(Path(cls.tmp.name) / "samples", legacy=False)
        cls.source = cls.samples["attachments"]

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def _convert(self, name: str, options: ConversionOptions | None = None):
        out = Path(self.tmp.name) / name
        result = convert_file(self.source, out, options)
        self.assertEqual(result.status, STATUS_SUCCESS, result.message)
        return result, result.output.read_text(encoding="utf-8")

    def test_writes_every_attachment_in_its_own_format(self):
        result, markdown = self._convert("extract")
        folder = result.output.parent / "attachments_attachments"

        for name, data in ATTACHMENTS.items():
            with self.subTest(name=name):
                self.assertEqual((folder / name).read_bytes(), data)
        self.assertNotIn(".emf", markdown)

    def test_links_to_the_attachment_where_the_object_was(self):
        _, markdown = self._convert("links")
        self.assertIn(
            "[báo cáo.pdf](attachments_attachments/b%C3%A1o%20c%C3%A1o.pdf) "
            "là tệp đính kèm.",
            markdown,
        )
        self.assertIn(
            "[trang web.html](attachments_attachments/trang%20web.html) "
            "là tệp đính kèm.",
            markdown,
        )

    def test_the_rest_of_the_document_is_untouched(self):
        _, with_attachments = self._convert("full")
        _, plain = self._convert(
            "plain", ConversionOptions(extract_attachments=False)
        )
        body = [line for line in plain.splitlines() if not line.startswith("![")]
        self.assertEqual(
            [line for line in with_attachments.splitlines() if "](attachments" not in line],
            [line for line in body if "là tệp đính kèm" not in line],
        )

    def test_disabled_keeps_the_previous_behaviour(self):
        result, markdown = self._convert(
            "disabled", ConversionOptions(extract_attachments=False)
        )
        self.assertIn(".emf", markdown)
        self.assertFalse((result.output.parent / "attachments_attachments").exists())

    def test_images_can_be_skipped_while_attachments_are_kept(self):
        result, markdown = self._convert(
            "no-images", ConversionOptions(extract_images=False)
        )
        self.assertNotIn(".emf", markdown)
        self.assertTrue(
            (result.output.parent / "attachments_attachments" / "báo cáo.pdf").exists()
        )

    def test_the_outline_preview_writes_nothing(self):
        outline = load_outline(self.source)
        self.assertIn("báo cáo.pdf là tệp đính kèm.", outline.markdown)
        self.assertFalse(list(self.source.parent.glob("*_attachments")))

    def test_a_document_without_objects_is_not_rewritten(self):
        prepared = prepare_attachments(self.samples["docx"])
        self.assertIsNone(prepared.data)
        self.assertEqual(prepared.files, [])

    def test_a_broken_file_is_reported_rather_than_raised(self):
        prepared = prepare_attachments(self.samples["corrupt"])
        self.assertIsNone(prepared.data)


if __name__ == "__main__":
    unittest.main(verbosity=2)
