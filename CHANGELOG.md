# Changelog

Every released version of word2md, newest first. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project follows
[Semantic Versioning](https://semver.org/).

Pushing a `v*` tag builds both executables and publishes them as a GitHub Release,
using the section below that matches the tag as the release notes.

Only the newest release keeps its executables on GitHub. Older ones are removed:
none of them can update itself, so downloading one is a dead end.

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
