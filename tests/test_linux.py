"""Linux port: Tesseract TSV parsing, language choice, release asset names."""

from __future__ import annotations

import sys
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src import legacy, ocr, updater  # noqa: E402

HEADER = "level\tpage_num\tblock_num\tpar_num\tline_num\tword_num\tleft\ttop\twidth\theight\tconf\ttext"
TSV = "\n".join(
    [
        HEADER,
        "1\t1\t0\t0\t0\t0\t0\t0\t2480\t3508\t-1\t",
        "4\t1\t1\t1\t1\t0\t300\t300\t900\t60\t-1\t",
        "5\t1\t1\t1\t1\t1\t300\t300\t250\t60\t96.5\tTài",
        "5\t1\t1\t1\t1\t2\t600\t300\t250\t60\t91\tliệu",
        "5\t1\t1\t1\t1\t3\t900\t300\t20\t60\t3\t~",
        "5\t1\t1\t1\t1\t4\t950\t300\t20\t60\t95\t   ",
    ]
)


class TesseractTsvTests(unittest.TestCase):
    def test_words_are_scaled_back_to_points(self):
        words = ocr.parse_tesseract_tsv(TSV, scale=300 / 72)
        self.assertEqual([w.text for w in words], ["Tài", "liệu"])
        first = words[0]
        self.assertAlmostEqual(first.x0, 72.0)
        self.assertAlmostEqual(first.top, 72.0)
        self.assertAlmostEqual(first.x1, (300 + 250) * 72 / 300)
        self.assertAlmostEqual(first.bottom, 360 * 72 / 300)

    def test_garbage_input_gives_no_words(self):
        self.assertEqual(ocr.parse_tesseract_tsv("", 1.0), [])
        self.assertEqual(ocr.parse_tesseract_tsv("not\ta\theader\n1\t2", 1.0), [])


class PickLanguageTests(unittest.TestCase):
    def test_tesseract_prefers_vietnamese_with_english(self):
        self.assertEqual(ocr.pick_language(["eng", "vie"]), "vie+eng")
        self.assertEqual(ocr.pick_language(["eng"]), "eng")
        self.assertEqual(ocr.pick_language(["chi_sim"]), "chi_sim")

    def test_windows_tags_still_work(self):
        self.assertEqual(ocr.pick_language(["en-US", "vi-VN"]), "vi-VN")


class AssetNameTests(unittest.TestCase):
    def test_names_match_build_output(self):
        self.assertEqual(updater.asset_name("x64"), "word2md-x64.exe")
        self.assertEqual(updater.asset_name("linux-x64"), "word2md-linux-x64")

    def test_linux_picks_the_linux_asset(self):
        payload = {
            "tag_name": "v99.0.0",
            "assets": [
                {"name": n, "browser_download_url": f"https://x/{n}", "size": 1}
                for n in ("word2md-x64.exe", "word2md-linux-x64")
            ],
        }
        with mock.patch.object(updater.sys, "platform", "linux"), mock.patch.object(
            updater.platform, "machine", return_value="x86_64"
        ):
            update = updater.select_update(payload)
        self.assertTrue(update.url.endswith("word2md-linux-x64"))


class ExternalEnvTests(unittest.TestCase):
    """A frozen build must not hand its bundled libraries to soffice/tesseract."""

    def test_frozen_build_restores_the_original_library_path(self):
        env = {"LD_LIBRARY_PATH": "/tmp/_MEI123", "LD_LIBRARY_PATH_ORIG": "/opt/lib"}
        with mock.patch.dict(legacy.os.environ, env, clear=True), mock.patch.object(
            legacy.sys, "frozen", True, create=True
        ), mock.patch.object(legacy.sys, "platform", "linux"):
            self.assertEqual(legacy.external_env().get("LD_LIBRARY_PATH"), "/opt/lib")

    def test_frozen_build_without_an_original_drops_the_bundle_path(self):
        with mock.patch.dict(
            legacy.os.environ, {"LD_LIBRARY_PATH": "/tmp/_MEI123"}, clear=True
        ), mock.patch.object(legacy.sys, "frozen", True, create=True), mock.patch.object(
            legacy.sys, "platform", "linux"
        ):
            self.assertNotIn("LD_LIBRARY_PATH", legacy.external_env())

    def test_source_run_keeps_the_environment(self):
        with mock.patch.dict(legacy.os.environ, {"LD_LIBRARY_PATH": "/x"}, clear=True):
            if getattr(legacy.sys, "frozen", False):
                self.skipTest("frozen")
            self.assertEqual(legacy.external_env()["LD_LIBRARY_PATH"], "/x")


if __name__ == "__main__":
    unittest.main()
