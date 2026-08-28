# Changelog

Every released version of word2md, newest first. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project follows
[Semantic Versioning](https://semver.org/).

Pushing a `v*` tag builds both executables and publishes them as a GitHub Release,
using the section below that matches the tag as the release notes.

Only the newest release keeps its executables on GitHub. Older ones are removed:
none of them can update itself, so downloading one is a dead end.

## [1.5.0] - 2026-08-28

Every file in the queue converts on its own now, PDF joins the formats in both
directions, and Markdown can come out as a workbook as well as a document.

### Added

- **A "Chuyển" button on every row**, in both directions. It converts that one file with
  the open tab's options and leaves the rest of the queue alone; the button at the bottom
  still runs the whole tab. While anything is converting, the per-row buttons ("Chuyển",
  "Mục…" and the remove ✕) are greyed out.
- **A prompt when the result already exists.** Converting the same file a second time
  asks what to do: *Yes* overwrites the old file, *No* writes it under a new name
  (`tên (1).md`), *Cancel* converts nothing. It covers the row buttons, the batch button
  and section extraction. Ticking **"Ghi đè file trùng tên"** answers *overwrite* up
  front, so nothing is asked.
- **Markdown → Excel.** The *Markdown → Word / Excel* tab has an output-format picker
  ("Định dạng xuất"): **Word (.docx)** by default, or **Excel (.xlsx)**. On the command
  line it is `--to word` / `--to excel`. Picking Excel greys out the Word-only options,
  which have nothing to say about a workbook.
- The new `src/md_to_xlsx.py` reuses the Markdown parser of the Word direction and writes
  the workbook with `openpyxl`, mirroring what *Excel → Markdown* produces: one worksheet
  per top-level section (a lone `#` title steps aside so its `##` sections become the
  sheets), Markdown tables as real cells with a bold shaded header, column alignment,
  sized columns and frozen headers, and plain numbers stored as numbers while `007` or
  `12%` stay text. Headings, paragraphs, lists, quotes and code go down column A in
  document order, and a cell that is only a link becomes a real hyperlink.
- **Images are embedded in the workbook**, scaled to fit a 640×480 box, each on a row as
  tall as the picture and with its alt text as a caption. This needs Pillow, now a
  dependency (`pillow==12.3.0`); without it, or for a remote, missing or unreadable
  image, the path is kept as text and a warning says why. An image inside a table cell
  stays a path: a picture cannot live in a cell. `--no-images` and the *"Nhúng ảnh"*
  checkbox turn embedding off, and that checkbox stays enabled for Excel.
- The **font and size** chosen in the tab apply to the workbook as well (`--font`,
  `--font-size`): they are written onto every cell, with headings a few points above the
  body text and code blocks keeping Consolas at the body size. Those two dropdowns stay
  enabled when the target is Excel; only the page options grey out.
- Cells are **wrapped and top-aligned**, so long text grows the row instead of running
  across the sheet, and a `<br>` inside a table cell shows as a second line. A column
  with no table in it gets up to 90 characters of width for prose; a table's own width
  wins where they share a column.

- **PDF, in both directions.** A PDF dropped into the *Word / Excel / PDF → Markdown*
  tab is read by the new `src/pdf.py` with `pdfplumber`: lines are grouped back into
  paragraphs, the type sizes are ranked into a heading hierarchy (numbering like `2.1.3`
  overriding the guess), ruled areas come back as Markdown tables whose text is not read
  twice, bullet and numbered lines become lists, and embedded images are written out with
  `pypdf` and linked. A scan with no text, a protected file or a file that is not a PDF is
  reported as an error rather than an empty result.
- The output-format picker offers **PDF (.pdf)** as well (`--to pdf`): the Markdown is
  converted to a Word document and Word - or LibreOffice - prints it, so **every option of
  the Markdown → Word direction applies to the page unchanged**, table of contents
  included (its fields are updated before the export). Without Word or LibreOffice the
  conversion fails with a message saying what to install.
- New dependencies for the reading side: `pdfplumber==0.11.10` and `pypdf==6.16.2`.
  Writing a PDF needs no new package - that is Word's or LibreOffice's job. `pdfminer.six`
  pulls in `cryptography`, which is pinned to 46.0.3 on ARM64: that is the last version
  publishing a `win_arm64` wheel, and the ARM64 runner cannot build it from source.

### Changed

- A one-file run reports in the status bar instead of opening a "Xong" message box.
  Errors still open a dialog.
- `converter.section_stems()` predicts the file names a section export will write, so the
  clash prompt can list them before any work starts. `_export_sections()` uses the same
  helper, so the names cannot drift apart.

### Fixed

- Control characters in a document - `\x0b`, which is what Word leaves behind for a
  soft line break - made the Excel export fail with openpyxl's `IllegalCharacterError`.
  They are stripped now, and `\x0b` / `\x0c` become real newlines.
- Text starting with `=` was written as a **formula**, so Excel showed `#NAME?` instead of
  the sentence. Such a cell is written as text now.
- A cell longer than Excel's 32 767-character limit is cut, with a warning, instead of
  producing a workbook Excel complains about.

## [1.4.2] - 2026-08-24

### Changed

- The update prompt shows the release notes as plain text. They come from this
  changelog, so the dialog used to display `###`, `**` and link syntax literally;
  headings, bullets, emphasis and links are now rendered for a message box, and
  the text is cut on a word boundary rather than mid-word.
- Release notes are published from this file instead of from generated commit
  subjects.

## [1.4.1] - 2026-08-24

The app updates itself now.

### Added

- **In-app update.** The app asks GitHub for the latest release two seconds after
  the window appears, on a background thread, and says nothing unless there is a
  newer version — an offline machine or a timeout is swallowed. The **"Kiểm tra cập
  nhật"** button in the header runs the same check out loud.
- Only the asset built for the running architecture (`word2md-x64.exe` or
  `word2md-arm64.exe`) is offered, and only when its tag parses as a higher version
  than the running one.
- The download lands beside the running executable, reports progress on the main
  progress bar, and is verified against the release's `.sha256` before anything is
  replaced. A mismatch, an empty file or a cancelled download deletes the partial
  file and changes nothing.
- **The app replaces itself and restarts.** Windows will not overwrite a running
  executable but will rename one, so the running exe becomes
  `word2md-<arch>.exe.old`, the download takes its place, the new build starts and
  the old process exits. The `.old` file stays locked until that process is gone, so
  it is deleted at the next launch.
- Postponing the restart keeps the download: the button becomes **"Cài v<version>"**
  and installs it without downloading again.
- A conversion in progress blocks the restart, and a swap that fails halfway puts
  the original executable back.
- The release workflow now publishes a `.sha256` next to each executable.

### Fixed

- The freshly installed executable died on startup with `cannot import name 'ops'
  from 'pandas._libs'`. A one-file build tells the app where it was unpacked through
  `_PYI_APPLICATION_HOME_DIR` and friends; the successor inherited them, skipped
  unpacking itself and ran out of the *old* build's `_MEIxxxxx` folder — full of the
  previous version's packages, and about to be deleted. Those variables are now
  stripped from the environment the successor starts with.

### Notes

- Running from source (`python main.py`) never self-updates; it points at the
  releases page instead.
- An app installed somewhere the user cannot write, such as `C:\Program Files`,
  reports a permission error before touching anything.
- 1.4.0 carried the same feature and was published briefly, but its executables
  could not restart after updating. It was replaced by 1.4.1 and removed.

## [1.3.1] - 2026-08-24

### Fixed

- A Word list starting below level one round-tripped as `- - text`, so the rebuilt
  `.docx` grew a stray empty bullet above the real content. Such a nested list is
  hoisted into the outer item's own level instead, keeping deeper nesting at its
  relative depth. Numbering instances are created on demand, so a list whose every
  item is hoisted leaves none behind.

## [1.3.0] - 2026-08-20

### Added

- **Markdown back to Word.** Dropping a `.md` file in produces a real `.docx`:
  headings, emphasis, links, images, nested lists, tables, quotes, fenced code,
  rules, `<sup>`/`<sub>`, task lists and YAML front matter, written with
  `python-docx` by a Markdown parser of the project's own.
- **One tab per direction**, each with its own queue and options; drag & drop routes
  every file to the tab that can convert it.
- Output settings for *Markdown → Word*: font (every face Tk can see), size, paper
  size, line spacing, a page break before each `H1`, and an optional table of
  contents field.

### Changed

- Round-tripping is preserved on purpose: code blocks take the `Code` paragraph
  style and inline code the `Code Char` character style, the names that map back to
  `<pre>` and `<code>`, so `.docx → .md → .docx` returns the structure it started
  with. Every list gets its own numbering instance, so a second ordered list starts
  at 1 again.
- Text comes out black throughout. Word's built-in headings are blue through
  `themeColor="accent1"`, which outranks a colour set beside it, so the theme
  reference is stripped from every style the document uses — and the chosen font is
  written into the theme, or headings would quietly stay on the template's Calibri.
- `python-docx` moved from `requirements-dev.txt` to `requirements.txt`.

## [1.2.0] - 2026-08-10

### Changed

- **A section export gets only its own assets.** Extracting one section used to
  write every image and attachment in the document next to it, so a one-page
  appendix could drop dozens of unrelated pictures in the output folder. Assets now
  land in a scratch folder and each output file copies out only the ones its own
  text points at, rewriting the links as it goes. A section that references nothing
  creates no folder.
- Exported files are named after the section heading (`Phụ lục.md`,
  `Phụ lục_images/`) instead of `test - 01 Phụ lục.md`. Merging several sections
  into one file has no single heading to take a name from, so it keeps the
  document's name.

### Fixed

- A split file is promoted to `H1` on its own, rather than by the shift the whole
  selection shares, as the README already promised.
- A body that starts with a heading no longer gets a duplicate title prepended
  above it.

## [1.1.0] - 2026-08-10

### Added

- **Attachments come out in their real format.** Word stores an attached file twice:
  as the `.emf` icon it draws on the page, and as the file itself in an OLE
  container under `word/embeddings/`. Mammoth only ever saw the icon, so every
  attachment used to arrive as a meaningless `.emf` image. The container is now
  unwrapped, each attachment is written to a `<name>_attachments/` folder under its
  original name, and the document is rewritten in memory so mammoth reads a
  hyperlink where the object was.
- **CI publishes the release.** The workflow already built both executables on a
  tag but left them as workflow artifacts, so every release was uploaded by hand. A
  release job now waits for both architectures and attaches them — requiring both
  keeps a release from ever shipping with one missing.

### Fixed

- Rewriting the source rather than the HTML keeps the link in place inside tables,
  headers and footnotes. If the rewrite upsets mammoth, the original file is read
  instead and the conversion still completes.

## [1.0.0] - 2026-07-31

### Added

- Desktop app converting **Word** (`.docx`, `.doc`) and **Excel** (`.xlsx`, `.xlsm`,
  `.xls`) to **Markdown**, with drag & drop, batch processing, a per-file status
  list and a cancellable background worker.
- **Section extraction** following the document's navigation outline: a
  `Heading 1`–`Heading 6` tree with a filter, a live Markdown preview, whole-branch
  selection, one file per section, and heading promotion.
- Legacy `.doc`/`.xls` support through Microsoft Office or LibreOffice, with a
  pure-Python fallback.
- Command-line mode for automation, and x64 / ARM64 one-file executables built by
  PyInstaller.

Older versions have no release to download any more, so each one links to the tag
it was built from. 1.0.0 predates tagging and links to its commit.

[1.4.2]: https://github.com/leeanh1999/word2md/releases/tag/v1.4.2
[1.4.1]: https://github.com/leeanh1999/word2md/tree/v1.4.1
[1.3.1]: https://github.com/leeanh1999/word2md/tree/v1.3.1
[1.3.0]: https://github.com/leeanh1999/word2md/tree/v1.3.0
[1.2.0]: https://github.com/leeanh1999/word2md/tree/v1.2.0
[1.1.0]: https://github.com/leeanh1999/word2md/tree/v1.1.0
[1.0.0]: https://github.com/leeanh1999/word2md/commit/060357976a07e3ec3c67f668d4848f19dbce120e
