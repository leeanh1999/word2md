"""Self-update against the project's GitHub Releases.

The app ships as a single PyInstaller executable, so an update is one file
swap: download the asset built for this architecture, put it where the running
executable lives, and restart into it.

Windows refuses to overwrite an executable that is running, but it does allow
renaming one. That is the whole trick here - the running exe is renamed out of
the way, the new one takes its place, and the leftover is deleted on the next
launch by `sweep_leftovers`.
"""

from __future__ import annotations

import hashlib
import json
import os
import platform
import re
import shutil
import subprocess
import sys
import tarfile
import tempfile
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path

from . import __version__

REPO = "leeanh1999/word2md"
LATEST_RELEASE_URL = f"https://api.github.com/repos/{REPO}/releases/latest"
RELEASES_PAGE = f"https://github.com/{REPO}/releases/latest"

# What platform.machine() reports -> the suffix build.py puts in the file name.
ARCH_SUFFIX = {
    "AMD64": "x64",
    "x86_64": "x64",
    "ARM64": "arm64",
    "aarch64": "arm64",
}

USER_AGENT = f"word2md/{__version__}"
TIMEOUT = 10.0

# The renamed predecessor, waiting to be deleted on the next launch.
BACKUP_SUFFIX = ".old"
# The freshly downloaded executable, before it takes over.
STAGED_SUFFIX = ".new"


class UpdateError(RuntimeError):
    """Anything that stops an update from being downloaded or applied."""


@dataclass(frozen=True)
class Update:
    """One newer release, with the asset that matches this machine."""

    version: str
    url: str
    size: int
    notes: str = ""
    checksum_url: str | None = None
    # Set when `url` is a .tar.gz: the path of the executable inside it.
    archive_member: str | None = None

    @property
    def size_text(self) -> str:
        return f"{self.size / 1_048_576:.1f} MB" if self.size else "?"


# --------------------------------------------------------------- versions


def parse_version(tag: str) -> tuple[int, ...]:
    """`v1.10.2-beta` -> `(1, 10, 2)`; anything unparsable sorts as `(0,)`."""
    head = str(tag or "").strip().lstrip("vV").split("-")[0].split("+")[0]
    parts: list[int] = []
    for chunk in head.split("."):
        digits = "".join(c for c in chunk if c.isdigit())
        if not digits:
            break
        parts.append(int(digits))
    return tuple(parts) or (0,)


def is_newer(candidate: str, current: str = __version__) -> bool:
    return parse_version(candidate) > parse_version(current)


def current_arch() -> str:
    """The build this machine needs: `x64`/`arm64` on Windows, `linux-x64` on Linux."""
    arch = ARCH_SUFFIX.get(platform.machine(), "x64")
    return f"linux-{arch}" if sys.platform.startswith("linux") else arch


def archive_name(arch: str) -> str | None:
    """The .tar.gz build.py packs for a Linux `arch`, or None elsewhere."""
    return f"word2md-{arch}.tar.gz" if arch.startswith("linux-") else None


def asset_name(arch: str) -> str:
    """The release asset build.py produces for `arch` (see build.target_name)."""
    return f"word2md-{arch}" if arch.startswith("linux-") else f"word2md-{arch}.exe"


# ------------------------------------------------------------------ check


def select_update(
    payload: dict, arch: str | None = None, current: str = __version__
) -> Update | None:
    """Turn a GitHub release payload into an `Update`, or None.

    None means there is nothing to do: the release is not newer, it is a draft,
    or it carries no executable for this architecture (the ARM64 build is
    published by a separate job and can legitimately be missing).
    """
    if not payload or payload.get("draft"):
        return None

    tag = payload.get("tag_name") or ""
    if not is_newer(tag, current):
        return None

    arch = arch or current_arch()
    binary = asset_name(arch)
    assets = {a.get("name"): a for a in payload.get("assets") or []}

    # On Linux the .tar.gz (smaller) is preferred; the bare binary is the
    # fallback, and is what releases before the archive existed carry.
    member = None
    wanted = archive_name(arch)
    if wanted and wanted in assets:
        member = f"{binary}/{binary}"
    else:
        wanted = binary
    asset = assets.get(wanted)
    if not asset:
        return None

    checksum = assets.get(f"{wanted}.sha256")
    return Update(
        version=tag.lstrip("vV"),
        url=asset["browser_download_url"],
        size=int(asset.get("size") or 0),
        notes=(payload.get("body") or "").strip(),
        checksum_url=checksum["browser_download_url"] if checksum else None,
        archive_member=member,
    )


def plain_notes(notes: str, limit: int = 700) -> str:
    """Release notes as plain text, because a message box has no Markdown.

    The notes come from CHANGELOG.md, so they arrive full of `###`, `**` and
    link syntax that a dialog would show literally.
    """
    lines: list[str] = []
    for raw in (notes or "").splitlines():
        line = raw.strip()
        if line.startswith("#"):
            line = line.lstrip("#").strip() + ":"
        elif line.startswith(("- ", "* ")):
            line = "• " + line[2:]
        if line or (lines and lines[-1]):
            lines.append(line)

    # Inline markup is cleaned on the joined text, because a wrapped **bold**
    # phrase has its opening and closing markers on different lines.
    text = "\n".join(lines).strip()
    text = re.sub(r"\[([^\]]+)\]\(([^)\s]+)\)", r"\1 (\2)", text)  # links
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text, flags=re.S)  # bold
    text = re.sub(r"(?<!\w)\*(.+?)\*(?!\w)", r"\1", text, flags=re.S)  # italics
    text = re.sub(r"(?<!\w)_([^_]+)_(?!\w)", r"\1", text)  # italics, underscored
    text = text.replace("`", "")
    if len(text) > limit:
        text = text[:limit].rsplit(" ", 1)[0].rstrip(" ,.;:-") + "…"
    return text


def _request(url: str) -> urllib.request.Request:
    return urllib.request.Request(
        url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": USER_AGENT,
        },
    )


def check_for_update(timeout: float = TIMEOUT) -> Update | None:
    """Ask GitHub for the latest release. Raises `UpdateError` on a bad trip."""
    try:
        with urllib.request.urlopen(_request(LATEST_RELEASE_URL), timeout=timeout) as r:
            payload = json.load(r)
    except urllib.error.HTTPError as exc:
        if exc.code == 404:  # no release published yet
            return None
        raise UpdateError(f"GitHub trả về lỗi {exc.code}.") from exc
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        raise UpdateError(f"Không kết nối được tới GitHub: {exc}") from exc
    except ValueError as exc:  # json.JSONDecodeError is a subclass
        raise UpdateError("GitHub trả về dữ liệu không đọc được.") from exc

    return select_update(payload)


# --------------------------------------------------------------- download


def current_exe() -> Path | None:
    """The running executable, or None when running from source."""
    if not getattr(sys, "frozen", False):
        return None
    return Path(sys.executable).resolve()


def staging_path(exe: Path) -> Path:
    """Where the download lands: beside the exe, so the swap is a same-volume rename."""
    return exe.with_name(exe.name + STAGED_SUFFIX)


def backup_path(exe: Path) -> Path:
    return exe.with_name(exe.name + BACKUP_SUFFIX)


def sweep_leftovers(exe: Path | None = None) -> None:
    """Delete the previous version left behind by an update. Never raises.

    Called at startup, when nothing holds the old file open any more.
    """
    exe = exe or current_exe()
    if exe is None:
        return
    for path in (backup_path(exe), staging_path(exe), _archive_path(staging_path(exe))):
        try:
            path.unlink(missing_ok=True)
        except OSError:  # still locked, or not ours - try again next launch
            pass


def _archive_path(dest: Path) -> Path:
    """Where a downloaded .tar.gz waits before its executable is taken out."""
    return dest.with_name(dest.name + ".tar.gz")


def _extract_member(archive: Path, member: str, dest: Path) -> None:
    """Copy one regular file out of `archive` to `dest`, and nothing else.

    Only the named member is read, through extractfile, so no path inside the
    archive can make it write anywhere but `dest`.
    """
    try:
        with tarfile.open(archive, "r:gz") as tar:
            try:
                info = tar.getmember(member)
            except KeyError as exc:
                raise UpdateError(f"Bản cập nhật thiếu {member}.") from exc
            if not info.isfile():
                raise UpdateError(f"{member} trong bản cập nhật không phải file.")
            source = tar.extractfile(info)
            if source is None:
                raise UpdateError(f"Không đọc được {member} trong bản cập nhật.")
            with source, dest.open("wb") as out:
                shutil.copyfileobj(source, out, 262_144)
    except (tarfile.TarError, EOFError, OSError) as exc:
        dest.unlink(missing_ok=True)
        raise UpdateError(f"File cập nhật bị hỏng: {exc}") from exc


def _fetch_checksum(url: str, timeout: float) -> str | None:
    """The expected SHA-256, or None if the release does not publish one."""
    try:
        with urllib.request.urlopen(_request(url), timeout=timeout) as resp:
            text = resp.read(4096).decode("utf-8", "replace")
    except (urllib.error.URLError, TimeoutError, OSError):
        return None
    # Accept both a bare digest and the `<digest>  <name>` sha256sum format.
    token = text.strip().split()[0] if text.strip() else ""
    return token.lower() if len(token) == 64 else None


def download(
    update: Update,
    dest: Path | None = None,
    on_progress=None,
    cancel=None,
    timeout: float = 30.0,
) -> Path:
    """Fetch the new executable to `dest` and verify it. Returns the path.

    A .tar.gz update is verified as downloaded, then only its executable is
    taken out to `dest`.

    `on_progress(done, total)` is called as bytes arrive; `cancel()` is polled
    between chunks and aborts the download when it returns True.
    """
    if dest is None:
        exe = current_exe()
        dest = staging_path(exe) if exe else Path(tempfile.gettempdir()) / "word2md.exe"

    expected = (
        _fetch_checksum(update.checksum_url, timeout) if update.checksum_url else None
    )
    digest = hashlib.sha256()
    # A .tar.gz lands beside dest first; dest only ever holds the executable.
    fetched = _archive_path(dest) if update.archive_member else dest
    done = 0

    try:
        dest.parent.mkdir(parents=True, exist_ok=True)
        with urllib.request.urlopen(_request(update.url), timeout=timeout) as resp:
            total = int(resp.headers.get("Content-Length") or update.size or 0)
            with fetched.open("wb") as out:
                while True:
                    if cancel is not None and cancel():
                        raise UpdateError("Đã huỷ tải bản cập nhật.")
                    chunk = resp.read(262_144)
                    if not chunk:
                        break
                    out.write(chunk)
                    digest.update(chunk)
                    done += len(chunk)
                    if on_progress is not None:
                        on_progress(done, total)
    except PermissionError as exc:
        raise UpdateError(
            f"Không có quyền ghi vào {dest.parent}. Hãy chạy với quyền quản trị "
            "hoặc chuyển app sang thư mục khác."
        ) from exc
    except UpdateError:
        fetched.unlink(missing_ok=True)
        raise
    except (urllib.error.URLError, TimeoutError, OSError) as exc:
        fetched.unlink(missing_ok=True)
        raise UpdateError(f"Tải bản cập nhật thất bại: {exc}") from exc

    if expected and digest.hexdigest() != expected:
        fetched.unlink(missing_ok=True)
        raise UpdateError("File tải về không khớp mã kiểm tra SHA-256.")

    if done == 0:
        fetched.unlink(missing_ok=True)
        raise UpdateError("File tải về rỗng.")

    if update.archive_member:
        try:
            _extract_member(fetched, update.archive_member, dest)
        finally:
            fetched.unlink(missing_ok=True)

    if os.name != "nt":
        # A Linux download arrives as plain data; the swap must stay runnable.
        dest.chmod(dest.stat().st_mode | 0o755)

    return dest


# ------------------------------------------------------------------ apply


def apply_update(new_exe: Path, exe: Path | None = None, restart: bool = True) -> Path:
    """Swap `new_exe` in for the running executable, then relaunch it.

    Does not return when `restart` is on: the process exits so the freshly
    installed executable takes over.
    """
    exe = exe or current_exe()
    if exe is None:
        raise UpdateError(
            "Chỉ bản đã đóng gói (.exe / Linux binary) mới tự cập nhật được. "
            f"Hãy tải bản mới tại {RELEASES_PAGE}"
        )
    if not new_exe.is_file():
        raise UpdateError("Không tìm thấy file cập nhật đã tải.")

    backup = backup_path(exe)
    try:
        backup.unlink(missing_ok=True)
    except OSError:
        pass

    try:
        exe.rename(backup)  # allowed even though this exe is running
    except OSError as exc:
        raise UpdateError(f"Không thay được file đang chạy: {exc}") from exc

    try:
        new_exe.replace(exe)
    except OSError as exc:
        backup.rename(exe)  # put things back exactly as they were
        raise UpdateError(f"Không ghi được bản mới: {exc}") from exc

    if restart:
        relaunch(exe)
    return exe


# The one-file bootloader hands these to the app it starts, so the app can
# find the folder it was unpacked into. A child that inherits them runs out of
# *this* executable's unpacked folder instead of its own, which for a different
# build means half-loaded packages ("cannot import name 'ops' from
# 'pandas._libs'") and a dead successor.
BOOTLOADER_ENV = (
    "_PYI_APPLICATION_HOME_DIR",
    "_PYI_ARCHIVE_FILE",
    "_PYI_PARENT_PROCESS_LEVEL",
    "_PYI_SPLASH_IPC",
    "_MEIPASS2",  # PyInstaller 5 and earlier
)


def clean_environment(env: dict | None = None, bundle: str | None = None) -> dict:
    """A copy of the environment with the bootloader's own variables removed.

    On Linux the bootloader also points LD_LIBRARY_PATH at this build's
    unpacked folder (keeping the user's value in LD_LIBRARY_PATH_ORIG). The
    successor must not inherit that: the folder is deleted as soon as this
    process exits, and while it lasts it would hand the old libraries to every
    program the new build starts (soffice, tesseract).
    """
    env = dict(os.environ if env is None else env)
    for name in BOOTLOADER_ENV:
        env.pop(name, None)
    for name in [key for key in env if key.startswith("_PYI_")]:
        env.pop(name)

    bundle = bundle if bundle is not None else getattr(sys, "_MEIPASS", None)
    for name in ("LD_LIBRARY_PATH", "LIBPATH"):
        original = env.pop(f"{name}_ORIG", None)
        if original is not None:
            env[name] = original
        elif bundle and name in env:
            kept = [p for p in env[name].split(os.pathsep) if p and p != bundle]
            if kept:
                env[name] = os.pathsep.join(kept)
            else:
                env.pop(name)
    return env


def relaunch(exe: Path) -> None:
    """Start the new executable and leave. Never returns."""
    flags = getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(
        subprocess, "CREATE_NEW_PROCESS_GROUP", 0
    )
    try:
        subprocess.Popen(
            [str(exe)],
            close_fds=True,
            cwd=str(exe.parent),
            env=clean_environment(),
            creationflags=flags if os.name == "nt" else 0,
        )
    except OSError as exc:
        raise UpdateError(f"Không khởi động lại được app: {exc}") from exc
    os._exit(0)  # skip atexit/Tk teardown - the successor is already running
