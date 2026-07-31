"""word2md - entry point.

Run without arguments to open the desktop GUI, or pass files/folders to
convert from the command line:

    python main.py report.docx data.xlsx -o output
    python main.py .\\documents -o .\\md

Partial extraction follows the document's navigation outline:

    python main.py report.docx --list-sections
    python main.py report.docx --sections 2,3.1 -o output
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

from src import __app_name__, __version__


def _configure_stdio() -> None:
    """Make CLI output survive a --noconsole build.

    PyInstaller's windowed bootloader leaves sys.stdout as None when there is
    no console, and hands over a cp1252 stream when launched from one - both
    blow up on Vietnamese text before any work gets done.
    """
    if sys.platform == "win32":
        try:
            import ctypes

            ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        except Exception:  # noqa: BLE001
            pass

    for name in ("stdout", "stderr"):
        stream = getattr(sys, name, None)
        if stream is None:
            setattr(sys, name, open(os.devnull, "w", encoding="utf-8"))
            continue
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is None:
            continue
        try:
            reconfigure(encoding="utf-8", errors="replace")
        except Exception:  # noqa: BLE001
            pass


def _run_cli(argv: list[str]) -> int:
    from src.converter import (
        STATUS_ERROR,
        STATUS_SUCCESS,
        WORD_EXTENSIONS,
        ConversionOptions,
        collect_files,
        convert_many,
        summarize,
    )

    parser = argparse.ArgumentParser(
        prog="word2md", description=f"{__app_name__} v{__version__}"
    )
    parser.add_argument(
        "inputs", nargs="*", help="File hoặc thư mục cần chuyển đổi"
    )
    parser.add_argument(
        "--backends",
        action="store_true",
        help="Liệt kê backend đọc được .doc/.xls trên máy này rồi thoát",
    )
    parser.add_argument("-o", "--output", default="output", help="Thư mục lưu file .md")
    parser.add_argument("--no-recursive", action="store_true", help="Không quét thư mục con")
    parser.add_argument("--no-images", action="store_true", help="Bỏ qua ảnh trong Word")
    parser.add_argument("--overwrite", action="store_true", help="Ghi đè file trùng tên")
    parser.add_argument("--no-title", action="store_true", help="Không thêm tiêu đề H1")

    sections = parser.add_argument_group("Trích xuất một phần (chỉ với 1 file Word)")
    sections.add_argument(
        "--list-sections", action="store_true", help="In mục lục kèm mã mục rồi thoát"
    )
    sections.add_argument(
        "--sections",
        metavar="ID[,ID…]",
        help="Chỉ xuất các mục theo mã, ví dụ 2,3.1 — dùng 'all' để chọn hết",
    )
    sections.add_argument(
        "--split-sections", action="store_true", help="Mỗi mục một file .md"
    )
    sections.add_argument(
        "--no-promote", action="store_true", help="Giữ nguyên bậc tiêu đề gốc"
    )

    parser.add_argument("--version", action="version", version=__version__)
    args = parser.parse_args(argv)

    if args.backends:
        from src.legacy import available_backends

        found = available_backends()
        print("Backend đọc định dạng cũ (.doc/.xls) trên máy này:")
        for backend in found or ["(không có)"]:
            print(f"  - {backend}")
        return 0

    if not args.inputs:
        parser.error("thiếu file hoặc thư mục cần chuyển đổi")

    files = collect_files(args.inputs, recursive=not args.no_recursive)
    if not files:
        print("Không tìm thấy file Word/Excel nào.", file=sys.stderr)
        return 1

    options = ConversionOptions(
        extract_images=not args.no_images,
        overwrite=args.overwrite,
        add_title_heading=not args.no_title,
        promote_headings=not args.no_promote,
        split_sections=args.split_sections,
    )

    if args.list_sections or args.sections:
        words = [f for f in files if f.suffix.lower() in WORD_EXTENSIONS]
        if len(words) != 1:
            print(
                "Trích xuất theo mục cần đúng một file Word "
                f"(đang có {len(words)}).",
                file=sys.stderr,
            )
            return 1
        return _run_sections(words[0], args, options)

    def on_progress(index, total, result):
        mark = "OK " if result.status == STATUS_SUCCESS else "ERR"
        target = result.output.name if result.output else result.message
        print(f"[{index}/{total}] {mark} {result.source.name} -> {target}")
        for warning in result.warnings:
            print(f"          ! {warning}")

    results = convert_many(files, Path(args.output), options, on_progress=on_progress)
    counts = summarize(results)
    print(f"\nThành công: {counts[STATUS_SUCCESS]} | Lỗi: {counts[STATUS_ERROR]}")
    return 0 if counts[STATUS_ERROR] == 0 else 2


def _run_sections(source: Path, args, options) -> int:
    from src.converter import STATUS_SUCCESS, convert_docx_sections, load_outline
    from src.outline import iter_nodes, outline_to_text

    try:
        outline = load_outline(source, options)
    except Exception as exc:  # noqa: BLE001
        print(f"Không đọc được {source.name}: {exc}", file=sys.stderr)
        return 1

    if not outline.roots:
        print(f"{source.name} không có tiêu đề Heading 1–6.", file=sys.stderr)
        return 1

    if args.list_sections:
        print(f"Mục lục của {source.name}:\n")
        print(outline_to_text(outline.roots))
        return 0

    if args.sections.strip().lower() == "all":
        node_ids = [node.node_id for node in outline.roots]
    else:
        node_ids = [part.strip() for part in args.sections.split(",") if part.strip()]

    known = {node.node_id for node in iter_nodes(outline.roots)}
    unknown = [node_id for node_id in node_ids if node_id not in known]
    if unknown:
        print(
            f"Không tìm thấy mã mục: {', '.join(unknown)}. "
            "Dùng --list-sections để xem danh sách.",
            file=sys.stderr,
        )
        return 1

    results = convert_docx_sections(source, Path(args.output), node_ids, options)
    errors = 0
    for result in results:
        if result.status == STATUS_SUCCESS:
            print(f"OK  {result.output}")
        else:
            errors += result.status != "skipped"
            print(f"ERR {result.message}", file=sys.stderr)
    return 0 if errors == 0 else 2


def main() -> int:
    _configure_stdio()
    if len(sys.argv) > 1:
        return _run_cli(sys.argv[1:])

    from src.gui import App

    App().run()
    return 0


if __name__ == "__main__":
    sys.exit(main())
