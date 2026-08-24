"""Tests for the in-app updater. No network: the release payload is a fixture."""

from __future__ import annotations

import hashlib
import io
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src import updater  # noqa: E402
from src.updater import Update, UpdateError  # noqa: E402

DOWNLOAD = "https://github.com/leeanh1999/word2md/releases/download"


def release(tag: str = "v9.9.9", names=("word2md-x64.exe",), **extra) -> dict:
    payload = {
        "tag_name": tag,
        "body": "Ghi chú phát hành",
        "assets": [
            {
                "name": name,
                "size": 1024,
                "browser_download_url": f"{DOWNLOAD}/{tag}/{name}",
            }
            for name in names
        ],
    }
    payload.update(extra)
    return payload


class VersionTests(unittest.TestCase):
    def test_parse(self):
        self.assertEqual(updater.parse_version("v1.3.1"), (1, 3, 1))
        self.assertEqual(updater.parse_version("1.10"), (1, 10))
        self.assertEqual(updater.parse_version("v2.0.0-beta.1"), (2, 0, 0))
        self.assertEqual(updater.parse_version(""), (0,))
        self.assertEqual(updater.parse_version("nightly"), (0,))

    def test_ordering(self):
        self.assertTrue(updater.is_newer("1.4.0", "1.3.1"))
        self.assertTrue(updater.is_newer("v1.10.0", "1.9.9"))
        self.assertFalse(updater.is_newer("1.3.1", "1.3.1"))
        self.assertFalse(updater.is_newer("1.3.0", "1.3.1"))
        # A malformed tag must never look newer than a real version.
        self.assertFalse(updater.is_newer("junk", "1.0.0"))


class SelectUpdateTests(unittest.TestCase):
    def test_picks_the_asset_for_this_architecture(self):
        payload = release(names=("word2md-x64.exe", "word2md-arm64.exe"))
        update = updater.select_update(payload, arch="arm64", current="1.3.1")
        self.assertIsNotNone(update)
        self.assertEqual(update.version, "9.9.9")
        self.assertTrue(update.url.endswith("word2md-arm64.exe"))
        self.assertEqual(update.notes, "Ghi chú phát hành")
        self.assertIsNone(update.checksum_url)

    def test_checksum_asset_is_paired_up(self):
        payload = release(names=("word2md-x64.exe", "word2md-x64.exe.sha256"))
        update = updater.select_update(payload, arch="x64", current="1.3.1")
        self.assertTrue(update.checksum_url.endswith(".sha256"))

    def test_nothing_to_do(self):
        cases = {
            "same version": release(tag="v1.3.1"),
            "older": release(tag="v1.0.0"),
            "draft": release(draft=True),
            "no asset for this arch": release(names=("word2md-arm64.exe",)),
            "empty": {},
        }
        for label, payload in cases.items():
            with self.subTest(label):
                self.assertIsNone(
                    updater.select_update(payload, arch="x64", current="1.3.1")
                )


class _FakeResponse(io.BytesIO):
    def __init__(self, data: bytes, headers: dict | None = None):
        super().__init__(data)
        self.headers = headers or {}

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()
        return False


class DownloadTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.addCleanup(self.tmp.cleanup)
        self.payload = b"MZ" + b"x" * 5000
        self.digest = hashlib.sha256(self.payload).hexdigest()

    def _update(self, checksum: bool = False) -> Update:
        return Update(
            version="9.9.9",
            url=f"{DOWNLOAD}/v9.9.9/word2md-x64.exe",
            size=len(self.payload),
            checksum_url=f"{DOWNLOAD}/v9.9.9/word2md-x64.exe.sha256" if checksum else None,
        )

    def _urlopen(self, sums: str | None = None):
        def fake(request, timeout=None):
            if request.full_url.endswith(".sha256"):
                return _FakeResponse((sums or "").encode())
            return _FakeResponse(
                self.payload, {"Content-Length": str(len(self.payload))}
            )

        return fake

    def test_writes_the_file_and_reports_progress(self):
        dest = self.dir / "word2md-x64.exe.new"
        seen = []
        with mock.patch.object(updater.urllib.request, "urlopen", self._urlopen()):
            out = updater.download(
                self._update(), dest=dest, on_progress=lambda d, t: seen.append((d, t))
            )
        self.assertEqual(out, dest)
        self.assertEqual(dest.read_bytes(), self.payload)
        self.assertEqual(seen[-1], (len(self.payload), len(self.payload)))

    def test_matching_checksum_passes(self):
        dest = self.dir / "ok.exe"
        sums = f"{self.digest}  word2md-x64.exe\n"
        with mock.patch.object(updater.urllib.request, "urlopen", self._urlopen(sums)):
            updater.download(self._update(checksum=True), dest=dest)
        self.assertTrue(dest.exists())

    def test_bad_checksum_is_refused_and_cleaned_up(self):
        dest = self.dir / "bad.exe"
        with mock.patch.object(
            updater.urllib.request, "urlopen", self._urlopen("00" * 32)
        ):
            with self.assertRaises(UpdateError):
                updater.download(self._update(checksum=True), dest=dest)
        self.assertFalse(dest.exists())

    def test_cancel_stops_and_removes_the_partial_file(self):
        dest = self.dir / "cancelled.exe"
        with mock.patch.object(updater.urllib.request, "urlopen", self._urlopen()):
            with self.assertRaises(UpdateError):
                updater.download(self._update(), dest=dest, cancel=lambda: True)
        self.assertFalse(dest.exists())


class ApplyTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.addCleanup(self.tmp.cleanup)
        self.exe = self.dir / "word2md-x64.exe"
        self.exe.write_bytes(b"old")
        self.staged = updater.staging_path(self.exe)
        self.staged.write_bytes(b"new")

    def test_swap_keeps_the_predecessor_as_backup(self):
        updater.apply_update(self.staged, exe=self.exe, restart=False)
        self.assertEqual(self.exe.read_bytes(), b"new")
        self.assertEqual(updater.backup_path(self.exe).read_bytes(), b"old")
        self.assertFalse(self.staged.exists())

    def test_failed_swap_puts_the_old_executable_back(self):
        with mock.patch.object(Path, "replace", side_effect=OSError("boom")):
            with self.assertRaises(UpdateError):
                updater.apply_update(self.staged, exe=self.exe, restart=False)
        self.assertEqual(self.exe.read_bytes(), b"old")
        self.assertFalse(updater.backup_path(self.exe).exists())

    def test_missing_download_is_refused(self):
        with self.assertRaises(UpdateError):
            updater.apply_update(self.dir / "nope.exe", exe=self.exe, restart=False)
        self.assertEqual(self.exe.read_bytes(), b"old")

    def test_running_from_source_cannot_self_update(self):
        with mock.patch.object(updater, "current_exe", return_value=None):
            with self.assertRaises(UpdateError):
                updater.apply_update(self.staged, restart=False)

    def test_sweep_removes_backup_and_staged_files(self):
        updater.backup_path(self.exe).write_bytes(b"older")
        updater.sweep_leftovers(self.exe)
        self.assertFalse(updater.backup_path(self.exe).exists())
        self.assertFalse(self.staged.exists())
        self.assertTrue(self.exe.exists())  # the live executable is untouched

    def test_sweep_survives_a_locked_leftover(self):
        updater.backup_path(self.exe).write_bytes(b"locked")
        with mock.patch.object(Path, "unlink", side_effect=OSError("in use")):
            updater.sweep_leftovers(self.exe)  # must not raise


if __name__ == "__main__":
    unittest.main(verbosity=2)
