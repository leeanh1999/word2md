"""Linux port: Tesseract TSV parsing, language choice, release asset names."""

from __future__ import annotations

import hashlib
import io
import sys
import tarfile
import tempfile
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


class RelaunchEnvironmentTests(unittest.TestCase):
    """The build started after an update must not inherit the old bundle's libraries."""

    def test_bundle_library_path_is_dropped(self):
        env = {
            "PATH": "/usr/bin",
            "LD_LIBRARY_PATH": "/tmp/_MEIold",
            "_PYI_LINUX_PROCESS_NAME": "word2md-linux-x",
            "_PYI_APPLICATION_HOME_DIR": "/tmp/_MEIold",
        }
        cleaned = updater.clean_environment(env, bundle="/tmp/_MEIold")
        self.assertEqual(cleaned, {"PATH": "/usr/bin"})

    def test_users_own_library_path_survives(self):
        env = {"LD_LIBRARY_PATH": "/tmp/_MEIold:/opt/lib", "LD_LIBRARY_PATH_ORIG": "/opt/lib"}
        cleaned = updater.clean_environment(env, bundle="/tmp/_MEIold")
        self.assertEqual(cleaned, {"LD_LIBRARY_PATH": "/opt/lib"})

    def test_without_a_bundle_nothing_else_changes(self):
        env = {"LD_LIBRARY_PATH": "/opt/lib"}
        self.assertEqual(updater.clean_environment(env, bundle=""), env)


def _tarball(files: dict[str, bytes]) -> bytes:
    buffer = io.BytesIO()
    with tarfile.open(fileobj=buffer, mode="w:gz") as tar:
        for name, data in files.items():
            info = tarfile.TarInfo(name)
            info.size, info.mode = len(data), 0o755
            tar.addfile(info, io.BytesIO(data))
    return buffer.getvalue()


class _Response(io.BytesIO):
    def __init__(self, data: bytes):
        super().__init__(data)
        self.headers = {"Content-Length": str(len(data))}

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


class ArchiveUpdateTests(unittest.TestCase):
    """On Linux the updater fetches the .tar.gz and takes the binary out of it."""

    NAMES = ("word2md-linux-x64", "word2md-linux-x64.tar.gz", "word2md-linux-x64.tar.gz.sha256")

    def _payload(self, names):
        return {
            "tag_name": "v99.0.0",
            "assets": [
                {"name": n, "browser_download_url": f"https://x/{n}", "size": 1} for n in names
            ],
        }

    def test_archive_is_preferred_over_the_bare_binary(self):
        update = updater.select_update(self._payload(self.NAMES), arch="linux-x64", current="1.0")
        self.assertTrue(update.url.endswith(".tar.gz"))
        self.assertTrue(update.checksum_url.endswith(".tar.gz.sha256"))
        self.assertEqual(update.archive_member, "word2md-linux-x64/word2md-linux-x64")

    def test_bare_binary_when_there_is_no_archive(self):
        update = updater.select_update(
            self._payload(("word2md-linux-x64",)), arch="linux-x64", current="1.0"
        )
        self.assertTrue(update.url.endswith("/word2md-linux-x64"))
        self.assertIsNone(update.archive_member)

    def test_windows_never_looks_for_an_archive(self):
        update = updater.select_update(
            self._payload(("word2md-x64.exe", "word2md-x64.tar.gz")), arch="x64", current="1.0"
        )
        self.assertTrue(update.url.endswith(".exe"))

    def _download(self, archive: bytes, checksum: str | None = None):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        dest = Path(tmp.name) / "word2md.new"
        update = updater.Update(
            version="99.0.0",
            url="https://x/word2md-linux-x64.tar.gz",
            size=len(archive),
            checksum_url="https://x/a.sha256" if checksum else None,
            archive_member="word2md-linux-x64/word2md-linux-x64",
        )

        def fake(request, timeout=None):
            if request.full_url.endswith(".sha256"):
                return _Response(f"{checksum}  a".encode())
            return _Response(archive)

        with mock.patch.object(updater.urllib.request, "urlopen", fake):
            result = updater.download(update, dest=dest)
        return dest, result

    def test_only_the_executable_comes_out(self):
        binary = b"\x7fELF" + b"x" * 1000
        archive = _tarball({
            "word2md-linux-x64/word2md-linux-x64": binary,
            "word2md-linux-x64/install-linux.sh": b"#!/bin/sh",
        })
        dest, result = self._download(archive, hashlib.sha256(archive).hexdigest())
        self.assertEqual(result, dest)
        self.assertEqual(dest.read_bytes(), binary)
        self.assertTrue(dest.stat().st_mode & 0o100)
        self.assertEqual(sorted(p.name for p in dest.parent.iterdir()), [dest.name])

    def test_wrong_checksum_is_refused(self):
        archive = _tarball({"word2md-linux-x64/word2md-linux-x64": b"x"})
        with self.assertRaises(updater.UpdateError):
            self._download(archive, "0" * 64)

    def test_archive_without_the_executable_is_refused(self):
        archive = _tarball({"word2md-linux-x64/other": b"x"})
        with self.assertRaises(updater.UpdateError) as caught:
            self._download(archive)
        self.assertIn("thiếu", str(caught.exception))

    def test_corrupt_archive_is_refused(self):
        with self.assertRaises(updater.UpdateError):
            self._download(b"not a tarball at all")


if __name__ == "__main__":
    unittest.main()
