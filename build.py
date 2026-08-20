"""Package the app into a single executable with PyInstaller.

    python build.py            # build dist/word2md-x64.exe on an Intel/AMD PC
    python build.py --clean    # remove this architecture's artefacts first
    python build.py --onedir   # folder build (faster start-up)

The executable is named after the architecture it runs on, so the x64 and the
ARM64 build can sit in dist/ side by side.

PyInstaller is not a cross-compiler and --target-arch is macOS only, so the
architecture of the output is always the architecture of the Python running
this script. An ARM64 executable therefore has to be built on Windows on ARM
with ARM64 Python - either a real device or the windows-11-arm CI runner, see
.github/workflows/build.yml.

Equivalent raw command on Windows:

    pyinstaller --noconfirm --clean --onefile --noconsole --name word2md-x64 \
        --collect-data customtkinter --collect-all tkinterdnd2 \
        --hidden-import tabulate --hidden-import openpyxl \
        main.py
"""

from __future__ import annotations

import argparse
import platform
import shutil
import struct
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
APP_NAME = "word2md"
ICON = ROOT / "assets" / "icon.ico"

EXCLUDES = [
    "matplotlib",
    "scipy",
    "IPython",
    "notebook",
    "pytest",
    "sqlalchemy",
    "PyQt5",
    "PySide2",
    "test",
    "unittest",
]

# The short names used in file names, keyed by what platform.machine() reports.
ARCH_ALIASES = {
    "amd64": "x64",
    "x86_64": "x64",
    "arm64": "arm64",
    "aarch64": "arm64",
    "x86": "x86",
    "i386": "x86",
    "i686": "x86",
}

# Per-architecture bootloaders inside the PyInstaller wheel. A missing folder
# means pip installed the wheel for the wrong architecture.
BOOTLOADERS = {
    "x64": "Windows-64bit-intel",
    "arm64": "Windows-64bit-arm",
    "x86": "Windows-32bit-intel",
}

# IMAGE_FILE_MACHINE_*, as reported by IsWow64Process2 and stored in PE headers.
MACHINE_CODES = {0x8664: "x64", 0xAA64: "arm64", 0x014C: "x86"}


def interpreter_arch() -> str:
    """What this Python builds for, which is what PyInstaller will emit."""
    machine = platform.machine().lower()
    return ARCH_ALIASES.get(machine, machine or "unknown")


def native_arch() -> str | None:
    """What the machine really is; it differs from the above under emulation.

    An x64 Python on Windows on ARM reports AMD64 from platform.machine() and
    from PROCESSOR_ARCHITECTURE alike, so the kernel has to be asked directly.
    Returns None when the answer is unavailable.
    """
    if sys.platform != "win32":
        return None
    import ctypes
    from ctypes import wintypes

    process = wintypes.USHORT()
    native = wintypes.USHORT()
    try:
        ok = ctypes.windll.kernel32.IsWow64Process2(
            ctypes.windll.kernel32.GetCurrentProcess(),
            ctypes.byref(process),
            ctypes.byref(native),
        )
    except (AttributeError, OSError):  # older than Windows 10 1709
        return None
    if not ok:
        return None
    return MACHINE_CODES.get(native.value)


def pe_machine(path: Path) -> str | None:
    """Read the architecture out of the PE header of a built executable.

    The only way to be sure the build really is what its name claims, rather
    than an x64 binary produced by an emulated interpreter.
    """
    try:
        with path.open("rb") as handle:
            if handle.read(2) != b"MZ":
                return None
            handle.seek(0x3C)
            offset = struct.unpack("<I", handle.read(4))[0]
            handle.seek(offset)
            if handle.read(4) != b"PE\0\0":
                return None
            machine = struct.unpack("<H", handle.read(2))[0]
    except (OSError, struct.error):
        return None
    return MACHINE_CODES.get(machine)


def target_name(arch: str) -> str:
    return f"{APP_NAME}-{arch}"


def check_toolchain(arch: str) -> None:
    """Fail early when the PyInstaller wheel cannot produce this architecture."""
    import PyInstaller

    expected = BOOTLOADERS.get(arch)
    if expected is None or sys.platform != "win32":
        return
    folder = Path(PyInstaller.__file__).parent / "bootloader" / expected
    if folder.is_dir():
        return

    available = sorted(
        d.name
        for d in (Path(PyInstaller.__file__).parent / "bootloader").iterdir()
        if d.is_dir() and d.name != "images"
    )
    raise SystemExit(
        f"PyInstaller không có bootloader {expected} (chỉ có: {', '.join(available)}).\n"
        f"Bản PyInstaller đã cài không dành cho {arch}. Hãy cài lại trong môi "
        f"trường Python {arch}:\n"
        "    python -m pip install --force-reinstall --no-cache-dir pyinstaller"
    )


def remove_artefacts(name: str) -> None:
    """Delete one architecture's output, leaving the other architecture alone."""
    for path in (ROOT / "build" / name, ROOT / "dist" / name):
        if path.is_dir():
            shutil.rmtree(path, ignore_errors=True)
        else:
            path.unlink(missing_ok=True)
    (ROOT / "dist" / f"{name}.exe").unlink(missing_ok=True)
    (ROOT / f"{name}.spec").unlink(missing_ok=True)


def build(
    onefile: bool = True,
    clean: bool = False,
    console: bool = False,
    arch: str | None = None,
) -> Path:
    import PyInstaller.__main__

    arch = arch or interpreter_arch()
    name = target_name(arch)
    check_toolchain(arch)

    if clean:
        remove_artefacts(name)

    args = [
        str(ROOT / "main.py"),
        "--name",
        name,
        "--noconfirm",
        "--clean",
        "--onefile" if onefile else "--onedir",
        "--console" if console else "--noconsole",
        "--distpath",
        str(ROOT / "dist"),
        "--workpath",
        str(ROOT / "build"),
        "--specpath",
        str(ROOT),
        # customtkinter ships JSON themes and tkinterdnd2 ships the tkdnd
        # binaries; neither is discoverable by static analysis.
        "--collect-data",
        "customtkinter",
        "--collect-all",
        "tkinterdnd2",
        # python-docx builds every .docx from a template that ships as package
        # data, invisible to static analysis.
        "--collect-data",
        "docx",
        # pandas imports its optional dependencies lazily: tabulate inside
        # DataFrame.to_markdown(), and the Excel engines on first read.
        "--hidden-import",
        "tabulate",
        "--hidden-import",
        "openpyxl",
        "--hidden-import",
        "openpyxl.cell._writer",
        "--hidden-import",
        "xlrd",
        # .doc / .xls support: OLE parsing plus Office automation.
        "--hidden-import",
        "olefile",
    ]
    if sys.platform == "win32":
        args += [
            "--hidden-import",
            "pythoncom",
            "--hidden-import",
            "pywintypes",
            "--hidden-import",
            "win32com",
            "--hidden-import",
            "win32com.client",
        ]
    for module in EXCLUDES:
        args += ["--exclude-module", module]
    if ICON.exists():
        args += ["--icon", str(ICON)]

    print("pyinstaller " + " ".join(args[1:]) + "\n")
    PyInstaller.__main__.run(args)

    suffix = ".exe" if sys.platform == "win32" else ""
    return ROOT / "dist" / (f"{name}{suffix}" if onefile else name)


def megabytes(path: Path) -> float:
    if path.is_file():
        return path.stat().st_size / 1_048_576
    total = sum(f.stat().st_size for f in path.rglob("*") if f.is_file())
    return total / 1_048_576


def resolve_arch(requested: str) -> str:
    """Refuse a request this interpreter cannot satisfy, instead of lying."""
    current = interpreter_arch()
    machine = native_arch()

    if requested != "auto" and requested != current:
        message = [
            f"Không build được bản {requested} bằng Python {current}: "
            "PyInstaller không cross-compile (--target-arch chỉ dành cho macOS).",
        ]
        if requested == "arm64":
            message.append(
                "Bản ARM64 phải build trên máy Windows on ARM với Python ARM64, "
                "hoặc trên runner windows-11-arm (.github/workflows/build.yml)."
            )
            if machine == "arm64":
                message.append(
                    "Máy này là ARM64 nhưng đang chạy Python x64 (giả lập). "
                    "Cài Python ARM64 rồi tạo lại venv là build được ngay."
                )
        raise SystemExit("\n".join(message))

    if machine and machine != current:
        print(
            f"CẢNH BÁO: máy là {machine} nhưng Python là {current}; "
            f"kết quả sẽ là bản {current} chạy qua giả lập.\n"
        )
    return current


def main() -> int:
    parser = argparse.ArgumentParser(description="Build the word2md executable")
    parser.add_argument("--onedir", action="store_true", help="Folder build")
    parser.add_argument(
        "--clean", action="store_true", help="Wipe this architecture's artefacts"
    )
    parser.add_argument("--console", action="store_true", help="Keep a console window")
    parser.add_argument(
        "--arch",
        choices=("auto", "x64", "arm64", "x86"),
        default="auto",
        help="Kiến trúc đích; phải trùng với Python đang chạy (mặc định: auto)",
    )
    args = parser.parse_args()

    arch = resolve_arch(args.arch)
    print(f"Target: {target_name(arch)}  (Python {platform.machine()})\n")

    result = build(
        onefile=not args.onedir,
        clean=args.clean,
        console=args.console,
        arch=arch,
    )

    if not result.exists():
        print(f"\nBuild FAILED: {result} not found", file=sys.stderr)
        return 1

    print(f"\nBuild OK: {result}  ({megabytes(result):.1f} MB)")

    executable = result if result.is_file() else result / f"{result.name}.exe"
    built = pe_machine(executable)
    if built is None:
        return 0
    if built != arch:
        print(
            f"Build FAILED: {executable.name} là bản {built}, không phải {arch}.",
            file=sys.stderr,
        )
        return 1
    print(f"Verified: PE header = {built}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
