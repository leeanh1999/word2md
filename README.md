# word2md

Desktop app that converts **Word (.docx, .doc)** and **Excel (.xlsx, .xlsm, .xls)**
files to **Markdown (.md)** — and **Markdown back to Word (.docx)** — with
drag-and-drop and batch processing.

## Features

- `customtkinter` interface (light / dark / follow system), with **one tab per
  direction** — *Word / Excel → Markdown* and *Markdown → Word* — each with its own
  file queue and its own options.
- Drag & drop files or whole folders onto the window (via `tkinterdnd2`), or use the
  picker buttons.
- Pick a folder → every convertible file inside is found recursively and sent to the
  tab that handles it.
- Progress bar plus a per-file status list (pending / working / OK / error).
- Conversion runs on a background thread with a cancel button; the UI never freezes.
- One bad file does **not** stop the batch: the error is recorded and the next file
  is processed.
- Choose the output folder, with an option to overwrite or auto-number `name (1).md`.
- **Extract parts of a Word document by section** (see below).
- **Markdown back to Word**: drop a `.md` file in and get a real `.docx` with Word
  headings, lists, tables and embedded images — with a **font picker** (every font
  installed on the machine), font size, paper size, line spacing, a page break before
  each `H1` and an optional automatic table of contents.
- **Files attached to a Word document** (PDF, HTML, ZIP…) are written out in their real
  format instead of the `.emf` icon Word shows for them.
- **Updates itself**: the app checks GitHub Releases at startup, and one click
  downloads the build for this architecture, swaps it in and restarts (see below).
- Command-line mode for automation.

## Section extraction (Navigation Pane)

The **"Extract sections…"** toolbar button — or the **"Sections…"** button on the row
of any `.docx` file — opens a window that mirrors Word's Navigation Pane:

- A tree of `Heading 1`–`Heading 6` with collapse / expand and a checkbox per section.
- A heading filter box and *Select all / Clear / Expand / Collapse* buttons.
- A **Markdown preview** pane that updates as you tick sections, showing the section
  count, line count and how many files will be produced.
- Content before the first heading shows up as a `(Preamble)` section.

Selecting a section takes its **entire branch**: the heading, its body and every
subsection, up to the next heading at the same or a higher level. That is why ticking
a parent automatically ticks its children and locks them (they are already included in
the export).

Two export options:

- **Split each section into its own `.md` file** — instead of merging everything into
  one file.
- **Promote selected headings to H1** — extracting a `Heading 3` section then produces
  a file starting with `#` instead of `###`; subsections are shifted by the same amount,
  so the relative hierarchy is preserved. With *split*, every file is promoted on its
  own, so each one starts at `#`.

Section IDs are hierarchical (`2`, `2.1`, `2.1.3`) and are shared between the GUI and
the CLI.

### What a section export writes

An exported section is a self-contained folder-mate of the document it came from:

- The `.md` file is **named after the heading** (`Phụ lục.md`), not after the document.
  Merging several sections into one file is the exception — there is no single heading
  to take the name from, so the document keeps its own.
- Its images and attachments go to `<heading>_images/` and `<heading>_attachments/`,
  and **only the ones that section actually uses** are written. Reading the outline
  still converts the whole document, but the assets land in a scratch folder first and
  are copied out per file, so exporting one appendix no longer drops every picture in
  the document next to it.
- A section that uses neither creates no folder at all.

### Word → Markdown

`mammoth` reads `.docx` into semantic HTML, then the `src/html_to_markdown.py` module
renders it to Markdown. (Mammoth's built-in Markdown writer is deprecated and drops
tables, so the project renders its own output to keep tables.)

Preserved: `h1`–`h6` headings, paragraphs, bold / italic / strikethrough, ordered and
unordered lists (including nesting), tables, blockquotes, code blocks, links and images.
Embedded images are written to a `<filename>_images/` folder and referenced by relative
path.

### Attached files (OLE objects)

A file attached to a Word document — a PDF, an HTML page, a ZIP, a spreadsheet — is
stored twice: as an `.emf` icon that Word draws on the page, and as the file itself,
wrapped in an OLE container under `word/embeddings/`. Mammoth only ever sees the icon,
so attachments used to come out as meaningless `.emf` images.

`src/attachments.py` unwraps the container, writes each attachment to a
`<filename>_attachments/` folder **in its original format and under its original
name**, and rewrites the document so mammoth reads a hyperlink where the object was:

```markdown
[báo cáo.pdf](report_attachments/b%C3%A1o%20c%C3%A1o.pdf) is attached.
```

Details:

- The file name comes from the UTF-16 copy Word stores in the `\x01Ole10Native`
  stream, so Vietnamese names survive; the ANSI copy next to it mangles them.
- An object with no name, or none the container reveals, is named after its content
  signature (`%PDF-` → `.pdf`, `PK` → `.zip`, …).
- Office files embedded as OOXML parts (`Excel.Sheet.12` and friends) are copied out
  as-is; old binary containers become `.doc` / `.xls` / `.ppt`.
- An object shown **as content** rather than as an icon keeps its picture and gains a
  link after it. An object nothing can unwrap — an equation, say — is left exactly as
  it was.
- A **linked** object is not copied: the link points at the original path.
- Rewriting the document is done on an in-memory copy; if anything about it upsets
  mammoth, the original file is read instead and the conversion still completes.

Turn it off with `--no-attachments`, or the *"Tách file đính kèm"* checkbox.

### Markdown → Word

Dropping a `.md` file in reverses the direction: `src/md_to_docx.py` parses the
Markdown itself — no third-party parser — and writes the document with `python-docx`.

Converted: ATX and setext headings, paragraphs with hard breaks, bold / italic /
strikethrough / inline code, links and autolinks, images, ordered and unordered lists
(nested, with task-list checkboxes), tables with column alignment, block quotes, fenced
and indented code blocks, horizontal rules, `<sup>` / `<sub>`, and backslash escapes.
YAML front matter is read as metadata and fills the document properties instead of
appearing in the text.

What the output looks like is up to you. The defaults are **Times New Roman 13 pt**,
A4, line spacing 1.15 — the font a document gets when nothing is chosen:

| Option | GUI | CLI | Default |
|---|---|---|---|
| Font | dropdown of installed fonts (typing is allowed) | `--font` | Times New Roman |
| Size | dropdown | `--font-size` | 13 pt |
| Paper | dropdown | `--page-size` | A4 (also Letter, Legal) |
| Line spacing | dropdown | `--line-spacing` | 1.15 |
| Page break before every H1 | checkbox | `--page-break-h1` | off |
| Table of contents | checkbox | `--toc` | off |
| Embed images | checkbox | `--no-images` to skip | on |
| H1 from the file name | checkbox | `--no-title` to skip | on |

The chosen font is written into the theme as well as into `Normal`, so the built-in
heading styles follow it instead of quietly staying on the template's Calibri. Code
blocks keep Consolas — a monospace face is the point of them.

**Everything is black.** Word's built-in headings are blue through
`themeColor="accent1"`, which outranks any colour set beside it, so the theme reference
is stripped from every style the document uses — headings, title, quotes and hyperlinks
alike (a link keeps its underline). The accent-coloured rules under `Title` and
`Intense Quote` are repainted black for the same reason.

The table of contents is Word's own `TOC` field, and the document is marked
`updateFields`, so Word offers to fill it in the first time the file is opened.

Details worth knowing:

- **Images are embedded, not linked.** A relative `![alt](doc_images/x.png)` is resolved
  against the folder of the `.md` file and inserted into the document, scaled down to
  the page width when it is too wide. A remote `https://…` image is not downloaded: its
  alt text is written instead and a warning is reported. `--no-images` skips embedding
  entirely.
- **Links to attachments keep working.** `[báo cáo.pdf](report_attachments/….pdf)`
  becomes a Word hyperlink with the same relative target, so a document converted to
  Markdown and back still points at the files next to it.
- **Every list gets its own Word numbering instance**, so a second ordered list starts
  at 1 again instead of continuing the first, and `3. …` still starts at 3.
- **Styles are chosen so the trip is reversible.** Code blocks use a `Code` paragraph
  style and inline code a `Code Char` character style — exactly the names
  `DOCX_STYLE_MAP` maps back to `<pre>` and `<code>` — so `.docx → .md → .docx` returns
  the structure it started with.
- Pages are A4, and the document is given an H1 from the file name only when the
  Markdown has none of its own (`--no-title` turns that off).

### Excel → Markdown

`pandas` + `openpyxl` read **all** sheets. Each sheet becomes a `## Sheet name` section
with a Markdown table (`DataFrame.to_markdown()` via `tabulate`), all merged into
**one** `.md` file. Dates are normalised to `YYYY-MM-DD`, empty cells stay empty, `|`
is escaped, and newlines inside a cell become `<br>`. Values are emitted verbatim
(`disable_numparse`) so codes like `007` or long IDs are not mangled.

## Legacy formats: `.doc` and `.xls`

No pure-Python library reads these Office 97–2003 binary formats completely, so the app
tries several backends in order and **always degrades gracefully** instead of failing:

| Format | 1st choice | 2nd choice | 3rd choice |
|---|---|---|---|
| `.doc` | Microsoft Word (COM) | LibreOffice | OLE text extraction |
| `.xls` | `xlrd` (pure Python) | Microsoft Excel (COM) | LibreOffice |

For `.xls`, `xlrd` reads the file directly, so **no Office install is required** and
Excel never has to start — the output matches the equivalent `.xlsx` file exactly.

For `.doc`, if Word or LibreOffice is available the file is upgraded to `.docx` and then
goes through the standard pipeline, keeping headings / lists / tables — the result is
**byte-identical** to converting the original `.docx`. If neither is installed, the app
parses the OLE structure itself (the Word 97 piece table) to pull out plain text and
clearly warns that formatting was lost.

Check which backends the machine has:

```powershell
.venv\Scripts\python.exe main.py --backends
```

Two details worth knowing:

- **Mislabelled files are detected automatically.** Plenty of real-world `.doc` files
  are actually renamed `.docx` or RTF. The app reads magic bytes rather than trusting
  the extension, so those work even without Office.
- **Upgrade results are cached** by (path, mtime, size) for the whole session, so
  browsing the outline and then exporting does not run Word twice. An entire batch
  shares a single Word/Excel process.

## Project layout

```
word2md/
├── main.py                    # Entry point: no args -> GUI, args -> CLI
├── build.py                   # PyInstaller packaging script (x64 / ARM64)
├── build.bat / run.bat        # Windows helpers
├── .github/workflows/build.yml # CI: builds both x64 and ARM64
├── requirements.txt           # Runtime dependencies
├── requirements-dev.txt       # + pyinstaller
├── src/
│   ├── converter.py           # Core engine
│   ├── attachments.py         # OLE objects -> real files + links
│   ├── html_to_markdown.py    # HTML (mammoth) -> Markdown
│   ├── md_to_docx.py          # Markdown -> Word (python-docx)
│   ├── legacy.py              # .doc/.xls backends + OLE parser
│   ├── outline.py             # Outline tree + partial extraction
│   ├── section_dialog.py      # Section picker window
│   ├── updater.py             # In-app update from GitHub Releases
│   └── gui.py                 # customtkinter interface
├── tests/
│   ├── make_samples.py        # Generates test.docx / test.xlsx / corrupt.docx
│   ├── test_converter.py      # Unit tests for the engine and HTML -> Markdown
│   ├── test_md_to_docx.py     # Unit tests for Markdown -> Word (+ round trips)
│   ├── test_attachments.py    # Unit tests for OLE attachments
│   ├── test_outline.py        # Unit tests for the outline and extraction
│   ├── test_legacy.py         # Unit tests for .doc/.xls
│   ├── test_updater.py        # Unit tests for the in-app updater
│   └── test_gui_smoke.py      # Builds a real window, runs a batch + section picker
└── output/                    # Default output folder
```

## Install

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
```

## Run

```powershell
# GUI
.venv\Scripts\python.exe main.py
# or: run.bat

# Command line
.venv\Scripts\python.exe main.py report.docx data.xlsx -o output
.venv\Scripts\python.exe main.py .\documents -o .\md --overwrite

# The other way round: Markdown -> Word
.venv\Scripts\python.exe main.py notes.md -o output
.venv\Scripts\python.exe main.py .\md -o .\docx --overwrite
.venv\Scripts\python.exe main.py notes.md --font Arial --font-size 12 --toc

# Section extraction
.venv\Scripts\python.exe main.py report.docx --list-sections
.venv\Scripts\python.exe main.py report.docx --sections 2,3.1 -o output
.venv\Scripts\python.exe main.py report.docx --sections all --split-sections
```

`--list-sections` prints the outline tree with the section IDs:

```
1         Q3 Technical Report  [H1, 32 lines]
1.1         1. Overview  [H2, 4 lines]
1.2         2. Task list  [H2, 6 lines]
```

CLI options: `-o/--output`, `--no-recursive`, `--no-images`, `--no-attachments`,
`--overwrite`, `--no-title`, `--backends`, `--font`, `--font-size`, `--page-size`,
`--line-spacing`, `--page-break-h1`, `--toc`, `--list-sections`,
`--sections ID[,ID…]`, `--split-sections`, `--no-promote`. The `--sections` group accepts exactly one Word
file. Exit codes: `0` success, `1` bad arguments / file not found, `2` some files
failed.

## Tests

```powershell
.venv\Scripts\python.exe tests\make_samples.py      # generate sample files
.venv\Scripts\python.exe -m unittest discover -s tests -t tests -v
.venv\Scripts\python.exe tests\test_gui_smoke.py    # needs a desktop session
```

## Building the `.exe`

```powershell
.venv\Scripts\python.exe build.py --clean
# or: build.bat
```

Result: `dist\word2md-x64.exe` (~39 MB, single file, no console window). The filename
carries the architecture, so the x64 and ARM64 builds can sit side by side in `dist\`;
`--clean` only removes the build for the architecture currently being built.

The equivalent PyInstaller command:

```powershell
pyinstaller --noconfirm --clean --onefile --noconsole --name word2md-x64 `
    --collect-data customtkinter --collect-all tkinterdnd2 --collect-data docx `
    --hidden-import tabulate --hidden-import openpyxl `
    main.py
```

`--collect-data customtkinter` bundles the theme JSON files, `--collect-all tkinterdnd2`
bundles the `tkdnd` native library, `--collect-data docx` bundles the blank template
every generated `.docx` is built from, and `tabulate` must be declared explicitly
because `pandas` only imports it at runtime inside `to_markdown()`.

Drop an icon at `assets/icon.ico` and `build.py` will pick it up automatically.

### Windows ARM64 build

PyInstaller does **not** cross-compile (`--target-arch` is macOS-only), so the `.exe`
always matches the architecture of the Python running `build.py`. An ARM64 build cannot
be produced on an x64 machine. `build.py --arch arm64` fails fast instead of quietly
emitting an x64 binary, and after the build it reads the PE header to verify the actual
architecture.

Two ways to get an ARM64 build:

```powershell
# 1. On a Windows on ARM machine, with ARM64 Python (not the emulated x64 build)
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv\Scripts\python.exe build.py --clean --arch arm64   # -> dist\word2md-arm64.exe
```

2. Via CI: `.github/workflows/build.yml` builds both, with the ARM64 job on a
   `windows-11-arm` runner. That runner is **free for public repos only**; private
   repos need real hardware or a larger ARM64 runner.

## Updating in place

The app updates itself from the GitHub Releases of this repository — no installer,
no admin rights.

- **On startup** it asks the API for the latest release, two seconds after the window
  is up, on a background thread. Nothing is shown unless there is a newer version; a
  timeout or an offline machine is swallowed silently.
- **The "Kiểm tra cập nhật" button** does the same check out loud, and reports being
  up to date or unreachable.
- Only the asset built for the running architecture (`word2md-x64.exe` or
  `word2md-arm64.exe`) is offered, and only when its tag parses as a higher version
  than `src/__init__.py`'s `__version__`.
- Accepting downloads it next to the running executable, showing progress on the
  main progress bar, and verifies it against the `.sha256` asset the release job
  publishes. A mismatch or a cancelled download deletes the partial file.
- Then the app **replaces itself and restarts**: Windows will not let a running
  executable be overwritten, but it does allow renaming one, so the running exe is
  renamed to `word2md-<arch>.exe.old`, the download takes its place, the new
  executable is started and this process exits. The `.old` file is still locked until
  the old process is gone, so it is deleted at the *next* launch by `sweep_leftovers`.
- The successor is started with `_PYI_APPLICATION_HOME_DIR` and the other bootloader
  variables **stripped from its environment**. Inheriting them makes the new build run
  out of the old one's unpacked `_MEIxxxxx` folder, which dies on a half-loaded
  package (`cannot import name 'ops' from 'pandas._libs'`).
- Declining the restart keeps the download; the button turns into **"Cài v<version>"**
  so it installs without downloading again.
- A swap that fails halfway puts the original executable back, and a conversion in
  progress blocks the restart.

Running from source (`python main.py`) skips all of this and just points at the
releases page — `git pull` is the update there.

```
src/updater.py     check_for_update() -> download() -> apply_update() -> relaunch()
```

Two situations need a manual download:

- The app lives somewhere the user cannot write, such as `C:\Program Files`. The
  download fails with a permission message before anything is touched.
- The release has no build for that architecture yet (the ARM64 job is separate).

## Releases

Pushing a `v*` tag runs the tests, builds both executables and publishes them as a
GitHub Release:

```powershell
git tag v1.4.1
git push origin v1.4.1
```

The release job waits for both architectures, so a release never ships with one of
them missing.

Every dependency ships a `win_arm64` wheel for CPython 3.13, including `pywin32` (so
`.doc`/`.xls` over COM still work) and `tkinterdnd2` (which bundles `tkdnd/win-arm64`,
keeping drag-and-drop native). Only `pandas` publishes `win_arm64` wheels from 3.0
onwards, so `requirements.txt` pins per architecture:

```
pandas==2.3.3; platform_machine != "ARM64"
pandas>=3.0.5; platform_machine == "ARM64"
```

All 148 tests pass on pandas 3.0.5 and the Markdown output is identical to pandas 2.3.3.

Without an ARM64 build, the x64 one still runs on Windows on ARM through the Prism
emulation layer, just slower.

## Known limitations

- Tables and images inside `.doc` are only preserved when Microsoft Word or LibreOffice
  is installed; the pure-Python fallback extracts text only.
- Section extraction relies on the `Heading 1`–`Heading 6` styles. Documents that fake
  headings with large or bold text produce no outline — Word's own Navigation Pane is
  empty for them too.
- Merged cells in Word tables are flattened, because Markdown has no colspan/rowspan.
- Pictures that are genuinely `.emf` / `.wmf` in the document (pasted charts and
  drawings) are still written out in that format: they are images, not attachments,
  and converting vector metafiles would need a rendering engine.
- Files currently open in Word/Excel may raise a permission error; close them and retry.
- Markdown → Word covers the Markdown this app produces plus common CommonMark; raw
  HTML blocks, footnotes, definition lists and reference-style links are not
  interpreted. Table column alignment reaches the `.docx` but not the Markdown coming
  back, and a fenced block's language is dropped — Word has nowhere to keep it.
- A folder containing both the sources and their `output\` subfolder converts the
  results too on the next run: `.md` is now an input format as well.
