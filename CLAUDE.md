# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

Windows desktop app (customtkinter) + CLI that converts Word (.docx/.doc), Excel (.xlsx/.xlsm/.xls) and PDF to Markdown, and Markdown back to Word, Excel or PDF. Python 3.13, Windows-first (COM via pywin32, Windows.Media.Ocr via winrt). UI strings are Vietnamese.

## Commands

```powershell
# Setup
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt

# Run: no args -> GUI, args -> CLI
.venv\Scripts\python.exe main.py
.venv\Scripts\python.exe main.py report.docx notes.md --to excel -o output
.venv\Scripts\python.exe main.py --backends          # which .doc/.xls backends exist here

# Tests (stdlib unittest, not pytest; same command CI runs)
.venv\Scripts\python.exe -m unittest discover -s tests -t tests -v
# One module / class / test (run from tests/ via -t tests so `make_samples` imports)
.venv\Scripts\python.exe -m unittest discover -s tests -t tests -p test_pdf.py
.venv\Scripts\python.exe -m unittest test_converter.HtmlToMarkdownTests.test_headings_and_inline   # with cwd=tests
# GUI smoke test is a plain script (no TestCase), needs a desktop session
.venv\Scripts\python.exe tests\test_gui_smoke.py

# Build single-file exe -> dist\word2md-<arch>.exe
.venv\Scripts\python.exe build.py --clean            # --arch arm64 only on ARM64 Python
```

Tests generate their own fixtures via `tests/make_samples.py` (`build_all()` into a temp dir); `tests/samples/` is gitignored. No linter is configured.

## Architecture

- `main.py` — dispatches GUI vs CLI; CLI exit codes: 0 ok, 1 bad args/not found, 2 some files failed.
- `src/converter.py` — the core engine shared by GUI and CLI. `convert_file()` never raises (errors come back as a `ConversionResult` with `STATUS_ERROR`); `convert_many()` shares one `LegacyUpgrader` per batch so Word/Excel start at most once, and supports progress callbacks + `cancel_event`. Routing is by extension (`WORD_EXTENSIONS`, `EXCEL_EXTENSIONS`, `PDF_EXTENSIONS`, `MARKDOWN_EXTENSIONS`); `.md` targets are chosen by `--to` / `MARKDOWN_TARGETS`.
- **Word → MD**: `attachments.py` first unwraps OLE-embedded files on an in-memory copy of the docx (falls back to the original on any failure) → `mammoth` with `DOCX_STYLE_MAP` → own HTML renderer `html_to_markdown.py` (mammoth's MD writer drops tables).
- **MD → Word**: `md_to_docx.py` has its own Markdown parser (no third-party one) and writes via python-docx. Styles (`Code`, `Code Char`) are chosen to match `DOCX_STYLE_MAP` so `.docx → .md → .docx` round-trips. `md_to_xlsx.py` reuses that parser.
- **MD → PDF**: builds the docx, then has Word (COM) or LibreOffice print it. No pure-Python fallback by design.
- **PDF → MD**: `pdf.py` (pdfplumber layout heuristics: paragraph gaps, ranked font sizes → heading levels, ruled tables, column-only-ruled table repair; images via pypdf). Scanned pages go to `ocr.py`.
- **OCR** (`ocr.py`): Windows.Media.Ocr gives word boxes; VietOCR re-reads each box for Vietnamese. VietOCR runs as ONNX via onnxruntime (`vietocr_onnx.py`, model in `src/models/vietocr/`, bundled into the exe) — never torch at runtime. Constraints: onnxruntime must be imported **before** any winrt import (winrt ships an old `msvcp140.dll`; wrong order = access violation), and `build.py` forces the same order in PyInstaller's analysis via a temp `sitecustomize`. OCR runs on one long-lived MTA thread because pywin32 COM needs STA.
- **Legacy formats** (`legacy.py`): `.doc` → Word COM → LibreOffice → OLE text extraction; `.xls` → xlrd → Excel COM → LibreOffice. Detects mislabelled files by magic bytes; upgrade results cached by (path, mtime, size).
- **Sections** (`outline.py`, `section_dialog.py`): outline from Heading 1–6; hierarchical IDs (`2.1.3`) shared between GUI and CLI `--sections`. Selecting a section takes its whole branch; exported assets are filtered to what each section uses.
- `gui.py` — two tabs (to-MD, from-MD), each with its own queue/options; conversion runs on a background thread.
- `updater.py` — self-update from GitHub Releases: download + sha256 verify, rename running exe to `.old`, start successor with PyInstaller bootloader env vars (`_PYI_*`) stripped.

## Linux

- `run.sh` / `build.sh` / `install-linux.sh` mirror the `.bat` files; `build.py` names the output `dist/word2md-linux-<arch>` and packs it into `dist/word2md-linux-<arch>.tar.gz` (binary + `.sha256` + `install-linux.sh`), the only Linux release asset. The updater downloads that archive (`updater.archive_name()`, member `<name>/<name>`), falling back to a bare binary for older releases; keep the names in sync.
- OCR backend is picked by `ocr.ocr_backend()`: Windows.Media.Ocr if winrt imports, else the `tesseract` CLI (TSV, `parse_tesseract_tsv`). Tesseract languages are traineddata names (`vie+eng`), not BCP-47 tags. VietOCR refine runs on both.
- Needs system `python3-tk`; `.doc` and MD → PDF need LibreOffice (`soffice`). CI job `linux` builds on ubuntu-22.04.

## Release / dependency rules

- Version lives in `src/__init__.py` (`__version__`). Pushing a `v*` tag runs tests, builds x64 + ARM64 in CI, and publishes a release whose notes are that tag's section of `CHANGELOG.md` (Keep a Changelog format) — update `CHANGELOG.md` before tagging; the in-app updater shows those notes.
- `requirements.txt` pins `pandas` and `cryptography` per architecture because the ARM64 runner can only install packages with `win_arm64` wheels. Check ARM64 wheel availability before bumping or adding dependencies.
- Regenerating the ONNX model: `pip install --no-deps vietocr==0.3.13`, `pip install -r requirements-export.txt`, then `python tools/export_vietocr_onnx.py`. `--no-deps` is required (vietocr pins an incompatible pillow). torch/vietocr/onnx are excluded from the exe in `build.py` — keep it that way.
- New runtime-imported packages may need `--hidden-import`/`--collect-*` entries in `build.py`.
