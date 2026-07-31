"""Tests for the binary Office 97-2003 formats (.doc and .xls).

Tests needing a real .doc/.xls skip themselves on machines without Microsoft
Office, since that is the only way to author those samples here.
"""

from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import src.legacy as legacy  # noqa: E402
from src.converter import (  # noqa: E402
    STATUS_ERROR,
    STATUS_SUCCESS,
    SUPPORTED_EXTENSIONS,
    ConversionOptions,
    collect_files,
    convert_docx_sections,
    convert_file,
    convert_many,
    load_outline,
)
from src.legacy import (  # noqa: E402
    LegacyConversionError,
    LegacyUpgrader,
    available_backends,
    can_read_xls,
    extract_doc_text,
    has_msoffice,
    sniff,
)
from make_samples import build_all  # noqa: E402

HAS_OFFICE = has_msoffice("Word.Application") and has_msoffice("Excel.Application")


class SniffTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.dir = Path(cls.tmp.name)

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def _write(self, name: str, data: bytes) -> Path:
        path = self.dir / name
        path.write_bytes(data)
        return path

    def test_zip_disguised_as_doc(self):
        kind = sniff(self._write("a.doc", b"PK\x03\x04rest"))
        self.assertTrue(kind.is_modern)
        self.assertEqual(kind.suggested_suffix, ".docx")

    def test_zip_disguised_as_xls(self):
        kind = sniff(self._write("a.xls", b"PK\x03\x04rest"))
        self.assertEqual(kind.suggested_suffix, ".xlsx")

    def test_ole_container(self):
        kind = sniff(self._write("a.doc", b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1more"))
        self.assertEqual(kind.container, "ole")
        self.assertFalse(kind.is_modern)

    def test_rtf_container(self):
        kind = sniff(self._write("a.doc", b"{\\rtf1\\ansi"))
        self.assertEqual(kind.container, "rtf")

    def test_unknown_container(self):
        self.assertEqual(sniff(self._write("a.doc", b"hello")).container, "unknown")

    def test_missing_file_is_unknown(self):
        self.assertEqual(sniff(self.dir / "nope.doc").container, "unknown")


class RegistrationTests(unittest.TestCase):
    def test_legacy_extensions_are_supported(self):
        self.assertIn(".doc", SUPPORTED_EXTENSIONS)
        self.assertIn(".xls", SUPPORTED_EXTENSIONS)

    def test_collect_files_picks_up_legacy(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            for name in ("a.doc", "b.xls", "c.docx", "d.txt", "~$e.doc"):
                (folder / name).write_bytes(b"x")
            names = sorted(p.name for p in collect_files([folder]))
            self.assertEqual(names, ["a.doc", "b.xls", "c.docx"])

    def test_available_backends_is_a_list_of_strings(self):
        backends = available_backends()
        self.assertIsInstance(backends, list)
        self.assertTrue(all(isinstance(item, str) for item in backends))


class MislabelledFileTests(unittest.TestCase):
    """A modern file with a legacy extension must not need Office at all."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.samples = build_all(Path(cls.tmp.name) / "samples", legacy=False)
        cls.out = Path(cls.tmp.name) / "out"

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_docx_named_doc(self):
        result = convert_file(self.samples["fake_doc"], self.out)
        self.assertEqual(result.status, STATUS_SUCCESS, result.message)
        text = result.output.read_text(encoding="utf-8")
        self.assertIn("| --- |", text)
        self.assertTrue(any("thực chất là định dạng mới" in w for w in result.warnings))

    def test_xlsx_named_xls(self):
        result = convert_file(self.samples["fake_xls"], self.out)
        self.assertEqual(result.status, STATUS_SUCCESS, result.message)
        self.assertIn("## Doanh thu", result.output.read_text(encoding="utf-8"))

    def test_title_uses_the_original_name(self):
        result = convert_file(self.samples["fake_doc"], self.out / "titled")
        self.assertEqual(result.output.stem.split(" (")[0], "mislabelled")


@unittest.skipUnless(HAS_OFFICE, "cần Microsoft Office để tạo file .doc/.xls mẫu")
class BinaryFormatTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory()
        cls.samples = build_all(Path(cls.tmp.name) / "samples")
        cls.out = Path(cls.tmp.name) / "out"

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()
        legacy.clear_cache()

    def test_doc_matches_docx_output(self):
        from_doc = convert_file(self.samples["doc"], self.out / "a")
        from_docx = convert_file(self.samples["docx"], self.out / "b")
        self.assertEqual(from_doc.status, STATUS_SUCCESS, from_doc.message)
        self.assertEqual(
            from_doc.output.read_text(encoding="utf-8"),
            from_docx.output.read_text(encoding="utf-8"),
        )

    def test_xls_matches_xlsx_output(self):
        from_xls = convert_file(self.samples["xls"], self.out / "c")
        from_xlsx = convert_file(self.samples["xlsx"], self.out / "d")
        self.assertEqual(from_xls.status, STATUS_SUCCESS, from_xls.message)
        self.assertEqual(
            from_xls.output.read_text(encoding="utf-8"),
            from_xlsx.output.read_text(encoding="utf-8"),
        )

    def test_doc_outline_matches_docx_outline(self):
        from_doc = load_outline(self.samples["doc"])
        from_docx = load_outline(self.samples["docx"])
        self.assertEqual(
            [n.title for n in from_doc.roots], [n.title for n in from_docx.roots]
        )
        self.assertEqual(sorted(from_doc.nodes), sorted(from_docx.nodes))

    def test_doc_section_extraction(self):
        outline = load_outline(self.samples["doc"])
        appendix = next(n for n in outline.roots if n.title == "Phụ lục")
        results = convert_docx_sections(
            self.samples["doc"], self.out / "sections", [appendix.node_id]
        )
        self.assertEqual(results[0].status, STATUS_SUCCESS, results[0].message)
        text = results[0].output.read_text(encoding="utf-8")
        self.assertIn("# Phụ lục", text)
        self.assertNotIn("Bảng số liệu", text)

    def test_mixed_batch_keeps_going(self):
        batch = [
            self.samples["corrupt"],
            self.samples["doc"],
            self.samples["xls"],
            self.samples["docx"],
        ]
        results = convert_many(batch, self.out / "batch")
        statuses = [r.status for r in results]
        self.assertEqual(statuses.count(STATUS_SUCCESS), 3)
        self.assertEqual(statuses.count(STATUS_ERROR), 1)

    def test_upgrade_result_is_cached(self):
        legacy.clear_cache()
        with LegacyUpgrader() as upgrader:
            first, _ = upgrader.upgrade(self.samples["doc"])
            second, _ = upgrader.upgrade(self.samples["doc"])
        self.assertEqual(first, second)

    def test_cache_survives_a_new_upgrader(self):
        with LegacyUpgrader() as upgrader:
            first, _ = upgrader.upgrade(self.samples["doc"])
        with LegacyUpgrader() as upgrader:
            second, _ = upgrader.upgrade(self.samples["doc"])
        self.assertEqual(first, second)

    def test_ole_text_extraction(self):
        text = extract_doc_text(self.samples["doc"])
        self.assertIn("Báo cáo kỹ thuật quý III", text)
        self.assertIn("Thu thập yêu cầu", text)
        self.assertIn("Phụ lục", text)
        self.assertNotIn("\x07", text)
        self.assertNotIn("\r", text)

    def test_plain_text_fallback_without_office(self):
        original_office, original_soffice = legacy.has_msoffice, legacy.find_soffice
        legacy.has_msoffice = lambda _app: False
        legacy.find_soffice = lambda: None
        legacy.clear_cache()
        try:
            result = convert_file(self.samples["doc"], self.out / "fallback")
        finally:
            legacy.has_msoffice = original_office
            legacy.find_soffice = original_soffice
            legacy.clear_cache()

        self.assertEqual(result.status, STATUS_SUCCESS, result.message)
        text = result.output.read_text(encoding="utf-8")
        self.assertIn("Báo cáo kỹ thuật quý III", text)
        self.assertTrue(any("văn bản thuần" in w for w in result.warnings))

    def test_xls_read_without_office(self):
        """xlrd alone must be enough for .xls - no Excel process involved."""
        original_office, original_soffice = legacy.has_msoffice, legacy.find_soffice
        legacy.has_msoffice = lambda _app: False
        legacy.find_soffice = lambda: None
        legacy.clear_cache()
        try:
            result = convert_file(self.samples["xls"], self.out / "xlrd-only")
        finally:
            legacy.has_msoffice = original_office
            legacy.find_soffice = original_soffice
            legacy.clear_cache()

        self.assertEqual(result.status, STATUS_SUCCESS, result.message)
        self.assertIn("Trần Văn A", result.output.read_text(encoding="utf-8"))


class NoBackendTests(unittest.TestCase):
    """Behaviour on a machine with neither Office nor LibreOffice."""

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.original = (legacy.has_msoffice, legacy.find_soffice)
        legacy.has_msoffice = lambda _app: False
        legacy.find_soffice = lambda: None
        legacy.clear_cache()

    def tearDown(self):
        legacy.has_msoffice, legacy.find_soffice = self.original
        legacy.clear_cache()
        self.tmp.cleanup()

    def test_upgrade_raises_a_helpful_error(self):
        source = self.dir / "x.doc"
        source.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1junk")
        with LegacyUpgrader() as upgrader:
            with self.assertRaises(LegacyConversionError) as caught:
                upgrader.upgrade(source)
        self.assertIn("LibreOffice", str(caught.exception))

    def test_convert_file_reports_the_error_instead_of_raising(self):
        source = self.dir / "x.doc"
        source.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1junk")
        result = convert_file(source, self.dir / "out")
        self.assertEqual(result.status, STATUS_ERROR)
        self.assertTrue(result.message)

    def test_mislabelled_file_still_works(self):
        docx = build_all(self.dir / "samples", legacy=False)["fake_doc"]
        result = convert_file(docx, self.dir / "out")
        self.assertEqual(result.status, STATUS_SUCCESS, result.message)


class XlsReaderTests(unittest.TestCase):
    def test_xlrd_is_installed(self):
        self.assertTrue(can_read_xls(), "xlrd cần thiết để đọc .xls không cần Office")


if __name__ == "__main__":
    unittest.main(verbosity=2)
