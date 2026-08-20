"""customtkinter desktop front-end for the Word/Excel <-> Markdown converter."""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from . import __app_name__, __version__
from .converter import (
    EXCEL_EXTENSIONS,
    MARKDOWN_EXTENSIONS,
    STATUS_CANCELLED,
    STATUS_ERROR,
    STATUS_SKIPPED,
    STATUS_SUCCESS,
    WORD_EXTENSIONS,
    ConversionOptions,
    ConversionResult,
    collect_files,
    convert_docx_sections,
    convert_many,
    load_outline,
)
from .legacy import can_read_xls, find_soffice, has_msoffice
from .md_to_docx import (
    DEFAULT_FONT,
    DEFAULT_FONT_SIZE,
    DEFAULT_LINE_SPACING,
    PAGE_SIZES,
    DocxSettings,
)
from .section_dialog import DocumentChooser, SectionDialog


def legacy_support_note() -> str:
    """One line telling the user what .doc/.xls support looks like here."""
    word = has_msoffice("Word.Application") or find_soffice()
    excel = can_read_xls() or has_msoffice("Excel.Application") or find_soffice()
    if word and excel:
        return "Định dạng cũ .doc/.xls: giữ đầy đủ cấu trúc."
    if excel:
        return (
            "Định dạng cũ: .xls đọc trực tiếp, .doc chỉ lấy được văn bản thuần "
            "(cần Microsoft Word hoặc LibreOffice để giữ định dạng)."
        )
    return "Định dạng cũ .doc/.xls: cần cài Microsoft Office hoặc LibreOffice."


def available_fonts() -> list[str]:
    """Every font Tk can see, with the usual suspects kept at the top.

    Needs a Tk root to exist already, so it is called while building the UI.
    """
    preferred = [
        DEFAULT_FONT,
        "Arial",
        "Calibri",
        "Cambria",
        "Georgia",
        "Segoe UI",
        "Tahoma",
        "Verdana",
    ]
    try:
        from tkinter import font as tkfont

        # A leading "@" marks the vertical-writing clone of a CJK font.
        installed = sorted(
            {name for name in tkfont.families() if not name.startswith("@")}
        )
    except Exception:  # noqa: BLE001 - a font list is not worth crashing over
        installed = []

    head = [name for name in preferred if name in installed]
    return head + [name for name in installed if name not in head] or preferred


try:  # Drag & drop is a nice-to-have; the app stays usable without it.
    from tkinterdnd2 import DND_FILES, TkinterDnD

    DND_AVAILABLE = True
except Exception:  # noqa: BLE001
    DND_FILES = None
    TkinterDnD = None
    DND_AVAILABLE = False

STATUS_STYLE = {
    STATUS_SUCCESS: ("OK", "#2fa572"),
    STATUS_ERROR: ("LỖI", "#d64545"),
    STATUS_SKIPPED: ("BỎ QUA", "#b0851f"),
    STATUS_CANCELLED: ("HUỶ", "#7a7a7a"),
    "pending": ("CHỜ", "#7a7a7a"),
}

OFFICE_EXTENSIONS = WORD_EXTENSIONS | EXCEL_EXTENSIONS

OFFICE_FILE_TYPES = [
    ("Word & Excel", "*.docx *.doc *.xlsx *.xlsm *.xls"),
    ("Word", "*.docx *.doc"),
    ("Excel", "*.xlsx *.xlsm *.xls"),
    ("Tất cả", "*.*"),
]
MARKDOWN_FILE_TYPES = [
    ("Markdown", "*.md *.markdown *.mdown *.mkd"),
    ("Tất cả", "*.*"),
]

TAB_TO_MARKDOWN = "Word / Excel  →  Markdown"
TAB_TO_WORD = "Markdown  →  Word"

FONT_SIZES = ["10", "11", "12", "13", "14", "16", "18"]
LINE_SPACINGS = ["1.0", "1.15", "1.5", "2.0"]

# The flat, outlined look shared by every secondary button.
OUTLINE_BUTTON = {
    "fg_color": "transparent",
    "border_width": 1,
    "text_color": ("#1f6aa5", "#5aa9e6"),
}


class _DnDCTk(ctk.CTk):
    """CTk window with tkinterdnd2 support mixed in."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = None
        if DND_AVAILABLE:
            try:
                self.TkdndVersion = TkinterDnD._require(self)
            except Exception:  # noqa: BLE001
                self.TkdndVersion = None


def _make_root() -> ctk.CTk:
    if DND_AVAILABLE:
        try:
            return _DnDCTk()
        except Exception:  # noqa: BLE001
            pass
    return ctk.CTk()


class FileRow(ctk.CTkFrame):
    """One line in the queue list: name, path and live status badge."""

    def __init__(self, master, path: Path, on_remove, on_sections=None):
        super().__init__(master, fg_color="transparent")
        self.path = path
        self.grid_columnconfigure(1, weight=1)

        self.badge = ctk.CTkLabel(
            self,
            text=STATUS_STYLE["pending"][0],
            width=72,
            corner_radius=6,
            fg_color=STATUS_STYLE["pending"][1],
            text_color="#ffffff",
            font=ctk.CTkFont(size=11, weight="bold"),
        )
        self.badge.grid(row=0, column=0, rowspan=2, padx=(6, 10), pady=6)

        ctk.CTkLabel(
            self,
            text=path.name,
            anchor="w",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=0, column=1, sticky="ew")

        self.detail = ctk.CTkLabel(
            self,
            text=str(path.parent),
            anchor="w",
            text_color=("#5c5c5c", "#9a9a9a"),
            font=ctk.CTkFont(size=11),
        )
        self.detail.grid(row=1, column=1, sticky="ew")

        if on_sections is not None and path.suffix.lower() in WORD_EXTENSIONS:
            ctk.CTkButton(
                self,
                text="Mục…",
                width=58,
                height=26,
                font=ctk.CTkFont(size=11),
                command=lambda: on_sections(self.path),
                **OUTLINE_BUTTON,
            ).grid(row=0, column=2, rowspan=2, padx=(6, 0))

        self.remove_button = ctk.CTkButton(
            self,
            text="✕",
            width=28,
            fg_color="transparent",
            hover_color=("#d9d9d9", "#404040"),
            text_color=("#5c5c5c", "#9a9a9a"),
            command=lambda: on_remove(self.path),
        )
        self.remove_button.grid(row=0, column=3, rowspan=2, padx=(6, 6))

    def set_status(self, status: str, detail: str) -> None:
        label, color = STATUS_STYLE.get(status, STATUS_STYLE["pending"])
        self.badge.configure(text=label, fg_color=color)
        self.detail.configure(text=detail)

    def set_busy(self) -> None:
        self.badge.configure(text="...", fg_color="#3b8ed0")
        self.detail.configure(text="Đang xử lý…")


class FileQueue(ctk.CTkFrame):
    """One tab: the pickers, the file list, and room for its own options."""

    def __init__(
        self,
        master,
        app: "App",
        extensions: set[str],
        file_types: list[tuple[str, str]],
        hint: str,
        on_sections=None,
    ):
        super().__init__(master, fg_color="transparent")
        self.app = app
        self.extensions = extensions
        self.file_types = file_types
        self.on_sections = on_sections
        self.files: list[Path] = []
        self.rows: dict[Path, FileRow] = {}

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self.toolbar = ctk.CTkFrame(self, fg_color="transparent")
        self.toolbar.grid(row=0, column=0, sticky="ew", pady=(4, 6))
        ctk.CTkButton(
            self.toolbar, text="Chọn file…", width=130, command=self.pick_files
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            self.toolbar, text="Chọn thư mục…", width=140, command=self.pick_folder
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(
            self.toolbar,
            text="Xoá danh sách",
            width=130,
            command=self.clear,
            **OUTLINE_BUTTON,
        ).pack(side="left", padx=(0, 8))

        self.count_label = ctk.CTkLabel(
            self.toolbar, text="0 file", text_color=("#5c5c5c", "#9a9a9a")
        )
        self.count_label.pack(side="right")

        self.list_frame = ctk.CTkScrollableFrame(self, label_text="Hàng đợi chuyển đổi")
        self.list_frame.grid(row=1, column=0, sticky="nsew")
        self.list_frame.grid_columnconfigure(0, weight=1)

        self.placeholder = ctk.CTkLabel(
            self.list_frame,
            text=hint,
            text_color=("#8a8a8a", "#7a7a7a"),
            justify="center",
        )
        self.placeholder.grid(row=0, column=0, pady=50)

        self.options = ctk.CTkFrame(self)
        self.options.grid(row=2, column=0, sticky="ew", pady=(8, 4))

    # ------------------------------------------------------------- files

    def accepts(self, path: Path) -> bool:
        return path.suffix.lower() in self.extensions

    def pick_files(self) -> None:
        chosen = filedialog.askopenfilenames(
            title="Chọn file cần chuyển đổi", filetypes=self.file_types
        )
        if chosen:
            self.app.add_paths(chosen, queue=self)

    def pick_folder(self) -> None:
        folder = filedialog.askdirectory(title="Chọn thư mục chứa file cần chuyển đổi")
        if folder:
            self.app.add_paths([folder], queue=self)

    def add(self, paths) -> int:
        added = 0
        for path in paths:
            if path in self.rows or not self.accepts(path):
                continue
            self.files.append(path)
            row = FileRow(self.list_frame, path, self.remove, self.on_sections)
            row.grid(row=len(self.rows), column=0, sticky="ew", pady=2)
            self.rows[path] = row
            added += 1
        self.refresh()
        return added

    def remove(self, path: Path) -> None:
        if self.app.is_running():
            return
        row = self.rows.pop(path, None)
        if row is not None:
            row.destroy()
        if path in self.files:
            self.files.remove(path)
        for index, remaining in enumerate(self.files):
            self.rows[remaining].grid_configure(row=index)
        self.refresh()

    def clear(self) -> None:
        if self.app.is_running():
            return
        for row in self.rows.values():
            row.destroy()
        self.rows.clear()
        self.files.clear()
        self.refresh()
        self.app.status_label.configure(text="Đã xoá danh sách.")

    def mark_pending(self) -> None:
        for row in self.rows.values():
            row.set_status("pending", str(row.path.parent))

    def refresh(self) -> None:
        self.count_label.configure(text=f"{len(self.files)} file")
        if self.files:
            self.placeholder.grid_remove()
        else:
            self.placeholder.grid()


class App:
    def __init__(self) -> None:
        ctk.set_appearance_mode("system")
        ctk.set_default_color_theme("blue")

        self.root = _make_root()
        self.root.title(f"{__app_name__} v{__version__}")
        self.root.geometry("1000x780")
        self.root.minsize(880, 660)

        self.output_dir = ctk.StringVar(value=str(Path.cwd() / "output"))
        self.overwrite = ctk.BooleanVar(value=False)

        # Word / Excel -> Markdown
        self.extract_images = ctk.BooleanVar(value=True)
        self.extract_attachments = ctk.BooleanVar(value=True)
        self.add_title = ctk.BooleanVar(value=True)

        # Markdown -> Word
        self.embed_images = ctk.BooleanVar(value=True)
        self.docx_add_title = ctk.BooleanVar(value=True)
        self.page_break_h1 = ctk.BooleanVar(value=False)
        self.table_of_contents = ctk.BooleanVar(value=False)
        self.font_name = ctk.StringVar(value=DEFAULT_FONT)
        self.font_size = ctk.StringVar(value=f"{DEFAULT_FONT_SIZE:g}")
        self.page_size = ctk.StringVar(value="A4")
        self.line_spacing = ctk.StringVar(value=f"{DEFAULT_LINE_SPACING:g}")

        self.worker: threading.Thread | None = None
        self.cancel_event = threading.Event()
        self.events: queue.Queue = queue.Queue()
        self.last_output_dir: Path | None = None

        self._build_ui()
        self._enable_dnd()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(80, self._drain_events)

    # ------------------------------------------------------------------ ui

    def _build_ui(self) -> None:
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self.root, corner_radius=0, height=68)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            header,
            text="Word / Excel  ⇄  Markdown",
            font=ctk.CTkFont(size=22, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=20, pady=(14, 0))
        ctk.CTkLabel(
            header,
            text="Mỗi chiều chuyển đổi một tab, với tuỳ chọn riêng.",
            text_color=("#5c5c5c", "#9a9a9a"),
        ).grid(row=1, column=0, sticky="w", padx=20, pady=(0, 12))
        ctk.CTkOptionMenu(
            header,
            values=["System", "Light", "Dark"],
            width=110,
            command=lambda value: ctk.set_appearance_mode(value.lower()),
        ).grid(row=0, column=1, rowspan=2, padx=20)

        self.tabs = ctk.CTkTabview(self.root, command=self._on_tab_changed)
        self.tabs.grid(row=1, column=0, sticky="nsew", padx=16, pady=(8, 0))
        for name in (TAB_TO_MARKDOWN, TAB_TO_WORD):
            tab = self.tabs.add(name)
            tab.grid_columnconfigure(0, weight=1)
            tab.grid_rowconfigure(0, weight=1)

        self.queues: dict[str, FileQueue] = {
            TAB_TO_MARKDOWN: self._build_to_markdown_tab(),
            TAB_TO_WORD: self._build_to_word_tab(),
        }
        self._build_footer()

    def _drop_hint(self, formats: str) -> str:
        action = (
            "Kéo & thả file hoặc thư mục vào đây"
            if DND_AVAILABLE
            else "Dùng nút “Chọn file…” hoặc “Chọn thư mục…” ở trên"
        )
        return f"{action}\n\nHỗ trợ: {formats}"

    def _build_to_markdown_tab(self) -> FileQueue:
        pane = FileQueue(
            self.tabs.tab(TAB_TO_MARKDOWN),
            self,
            OFFICE_EXTENSIONS,
            OFFICE_FILE_TYPES,
            self._drop_hint(", ".join(sorted(OFFICE_EXTENSIONS)))
            + f"\n{legacy_support_note()}",
            on_sections=self.open_section_extractor,
        )
        pane.grid(row=0, column=0, sticky="nsew")

        self.sections_button = ctk.CTkButton(
            pane.toolbar,
            text="Trích xuất theo mục…",
            width=180,
            command=self.open_section_extractor,
            **OUTLINE_BUTTON,
        )
        self.sections_button.pack(side="left")

        checks = ctk.CTkFrame(pane.options, fg_color="transparent")
        checks.pack(fill="x", padx=12, pady=10)
        ctk.CTkCheckBox(
            checks, text="Trích xuất ảnh trong Word", variable=self.extract_images
        ).pack(side="left", padx=(0, 18))
        ctk.CTkCheckBox(
            checks, text="Tách file đính kèm", variable=self.extract_attachments
        ).pack(side="left", padx=(0, 18))
        ctk.CTkCheckBox(
            checks, text="Thêm tiêu đề từ tên file", variable=self.add_title
        ).pack(side="left")
        return pane

    def _build_to_word_tab(self) -> FileQueue:
        pane = FileQueue(
            self.tabs.tab(TAB_TO_WORD),
            self,
            MARKDOWN_EXTENSIONS,
            MARKDOWN_FILE_TYPES,
            self._drop_hint(", ".join(sorted(MARKDOWN_EXTENSIONS)))
            + "\nẢnh và file đính kèm được nhúng lại theo đường dẫn tương đối.",
        )
        pane.grid(row=0, column=0, sticky="nsew")

        layout = ctk.CTkFrame(pane.options, fg_color="transparent")
        layout.pack(fill="x", padx=12, pady=(10, 4))
        for column in (1, 3, 5, 7):
            layout.grid_columnconfigure(column, weight=1)

        ctk.CTkLabel(layout, text="Font:").grid(row=0, column=0, sticky="w")
        ctk.CTkComboBox(
            layout, values=available_fonts(), variable=self.font_name, width=190
        ).grid(row=0, column=1, sticky="ew", padx=(6, 16))

        ctk.CTkLabel(layout, text="Cỡ chữ:").grid(row=0, column=2, sticky="w")
        ctk.CTkComboBox(
            layout, values=FONT_SIZES, variable=self.font_size, width=80
        ).grid(row=0, column=3, sticky="ew", padx=(6, 16))

        ctk.CTkLabel(layout, text="Khổ giấy:").grid(row=0, column=4, sticky="w")
        ctk.CTkOptionMenu(
            layout, values=sorted(PAGE_SIZES), variable=self.page_size, width=100
        ).grid(row=0, column=5, sticky="ew", padx=(6, 16))

        ctk.CTkLabel(layout, text="Giãn dòng:").grid(row=0, column=6, sticky="w")
        ctk.CTkOptionMenu(
            layout, values=LINE_SPACINGS, variable=self.line_spacing, width=90
        ).grid(row=0, column=7, sticky="ew", padx=(6, 0))

        checks = ctk.CTkFrame(pane.options, fg_color="transparent")
        checks.pack(fill="x", padx=12, pady=(0, 10))
        ctk.CTkCheckBox(checks, text="Nhúng ảnh", variable=self.embed_images).pack(
            side="left", padx=(0, 18)
        )
        ctk.CTkCheckBox(
            checks, text="Thêm tiêu đề từ tên file", variable=self.docx_add_title
        ).pack(side="left", padx=(0, 18))
        ctk.CTkCheckBox(
            checks, text="Ngắt trang trước tiêu đề cấp 1", variable=self.page_break_h1
        ).pack(side="left", padx=(0, 18))
        ctk.CTkCheckBox(
            checks, text="Chèn mục lục", variable=self.table_of_contents
        ).pack(side="left")
        return pane

    def _build_footer(self) -> None:
        output = ctk.CTkFrame(self.root)
        output.grid(row=2, column=0, sticky="ew", padx=16, pady=(10, 6))
        output.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(output, text="Thư mục lưu:").grid(
            row=0, column=0, padx=(12, 8), pady=12, sticky="w"
        )
        ctk.CTkEntry(output, textvariable=self.output_dir).grid(
            row=0, column=1, sticky="ew", pady=12
        )
        ctk.CTkButton(output, text="Đổi…", width=70, command=self.pick_output).grid(
            row=0, column=2, padx=8, pady=12
        )
        ctk.CTkButton(
            output, text="Mở", width=60, command=self.open_output, **OUTLINE_BUTTON
        ).grid(row=0, column=3, padx=(0, 8), pady=12)
        ctk.CTkCheckBox(
            output, text="Ghi đè file trùng tên", variable=self.overwrite
        ).grid(row=0, column=4, padx=(4, 12), pady=12)

        footer = ctk.CTkFrame(self.root, fg_color="transparent")
        footer.grid(row=3, column=0, sticky="ew", padx=16, pady=(0, 16))
        footer.grid_columnconfigure(0, weight=1)

        self.progress = ctk.CTkProgressBar(footer)
        self.progress.set(0)
        self.progress.grid(row=0, column=0, sticky="ew", padx=(0, 12))

        self.convert_button = ctk.CTkButton(
            footer,
            text="Chuyển đổi",
            width=150,
            height=38,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.start_conversion,
        )
        self.convert_button.grid(row=0, column=1, rowspan=2)

        self.status_label = ctk.CTkLabel(
            footer,
            text="Sẵn sàng.",
            anchor="w",
            text_color=("#5c5c5c", "#9a9a9a"),
        )
        self.status_label.grid(row=1, column=0, sticky="ew", pady=(6, 0))

    def _enable_dnd(self) -> None:
        if not DND_AVAILABLE or getattr(self.root, "TkdndVersion", None) is None:
            return
        widgets = [self.root]
        for pane in self.queues.values():
            widgets += [pane.list_frame, pane.placeholder]
        for widget in widgets:
            try:
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._on_drop)
            except Exception:  # noqa: BLE001
                pass

    def _on_drop(self, event) -> None:
        try:
            paths = self.root.tk.splitlist(event.data)
        except Exception:  # noqa: BLE001
            paths = [event.data]
        self.add_paths(paths)

    def _on_tab_changed(self) -> None:
        self.progress.set(0)

    # --------------------------------------------------------------- files

    @property
    def current(self) -> FileQueue:
        return self.queues.get(self.tabs.get(), self.queues[TAB_TO_MARKDOWN])

    @property
    def files(self) -> list[Path]:
        return self.current.files

    @property
    def rows(self) -> dict[Path, FileRow]:
        return self.current.rows

    def clear_files(self) -> None:
        self.current.clear()

    def add_paths(self, paths, queue: FileQueue | None = None) -> None:
        """Queue files, sending each one to the tab that can convert it."""
        if self.is_running():
            return
        discovered = collect_files(paths)
        if not discovered:
            self.status_label.configure(
                text="Không tìm thấy file .docx/.xlsx/.md hợp lệ trong lựa chọn."
            )
            return

        targets = [queue] if queue is not None else list(self.queues.values())
        added: dict[FileQueue, int] = {}
        for pane in targets:
            count = pane.add([path for path in discovered if pane.accepts(path)])
            if count:
                added[pane] = count

        if not added:
            self.status_label.configure(text="Những file này thuộc tab còn lại.")
            return

        # Bring whichever tab actually received something to the front.
        busiest = max(added, key=added.get)
        if busiest is not self.current:
            for name, pane in self.queues.items():
                if pane is busiest:
                    self.tabs.set(name)
        self.status_label.configure(text=f"Đã thêm {sum(added.values())} file.")

    def pick_output(self) -> None:
        folder = filedialog.askdirectory(title="Chọn thư mục lưu kết quả")
        if folder:
            self.output_dir.set(folder)

    def open_output(self) -> None:
        folder = Path(self.last_output_dir or self.output_dir.get())
        try:
            folder.mkdir(parents=True, exist_ok=True)
            if sys.platform == "win32":
                os.startfile(folder)  # noqa: S606
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(folder)])
            else:
                subprocess.Popen(["xdg-open", str(folder)])
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Lỗi", f"Không mở được thư mục:\n{exc}")

    # ------------------------------------------------------------- options

    def _docx_settings(self) -> DocxSettings:
        return DocxSettings(
            font=self.font_name.get(),
            font_size=_to_float(self.font_size.get(), DEFAULT_FONT_SIZE),
            page_size=self.page_size.get(),
            line_spacing=_to_float(self.line_spacing.get(), DEFAULT_LINE_SPACING),
            page_break_before_h1=self.page_break_h1.get(),
            table_of_contents=self.table_of_contents.get(),
        )

    def options_for(self, pane: FileQueue) -> ConversionOptions:
        if pane is self.queues[TAB_TO_WORD]:
            return ConversionOptions(
                extract_images=self.embed_images.get(),
                overwrite=self.overwrite.get(),
                add_title_heading=self.docx_add_title.get(),
                docx=self._docx_settings(),
            )
        return ConversionOptions(
            extract_images=self.extract_images.get(),
            extract_attachments=self.extract_attachments.get(),
            overwrite=self.overwrite.get(),
            add_title_heading=self.add_title.get(),
        )

    # -------------------------------------------------- section extraction

    def open_section_extractor(self, path: Path | None = None) -> None:
        """Open Word's navigation outline for a document and export a subset."""
        if self.is_running():
            messagebox.showinfo("Đang bận", "Hãy đợi lô hiện tại chạy xong.")
            return

        if path is None:
            path = self._choose_word_document()
        if path is None:
            return
        if path.suffix.lower() not in WORD_EXTENSIONS:
            messagebox.showinfo(
                "Không hỗ trợ",
                "Trích xuất theo mục chỉ áp dụng cho file Word (.docx).",
            )
            return

        self.sections_button.configure(state="disabled")
        self.status_label.configure(text=f"Đang đọc mục lục của {path.name}…")

        def worker() -> None:
            try:
                outline = load_outline(Path(path))
            except Exception as exc:  # noqa: BLE001
                self.events.put(("outline-error", path, exc))
            else:
                self.events.put(("outline", outline))

        threading.Thread(target=worker, daemon=True).start()

    def _choose_word_document(self) -> Path | None:
        pane = self.queues[TAB_TO_MARKDOWN]
        candidates = [f for f in pane.files if f.suffix.lower() in WORD_EXTENSIONS]
        if len(candidates) == 1:
            return candidates[0]
        if len(candidates) > 1:
            return DocumentChooser(self.root, candidates).show()
        chosen = filedialog.askopenfilename(
            title="Chọn file Word để trích xuất theo mục",
            filetypes=[("Word", "*.docx *.doc"), ("Tất cả", "*.*")],
        )
        return Path(chosen) if chosen else None

    def _show_section_dialog(self, outline) -> None:
        self.sections_button.configure(state="normal")
        if not outline.roots:
            messagebox.showinfo(
                "Không có mục lục",
                f"{outline.source.name} không dùng Heading 1–6 nên không có mục "
                "nào để chọn.\nHãy dùng nút “Chuyển đổi” để xuất toàn bộ tài liệu.",
            )
            self.status_label.configure(text="Tài liệu không có tiêu đề.")
            return

        choice = SectionDialog(self.root, outline).show()
        if not choice:
            self.status_label.configure(text="Đã huỷ trích xuất theo mục.")
            return

        output_dir = self._prepare_output_dir()
        if output_dir is None:
            return

        options = ConversionOptions(
            extract_images=self.extract_images.get(),
            extract_attachments=self.extract_attachments.get(),
            overwrite=self.overwrite.get(),
            add_title_heading=self.add_title.get(),
            promote_headings=choice["promote"],
            split_sections=choice["split"],
        )

        self.status_label.configure(text="Đang trích xuất…")
        results = convert_docx_sections(
            choice["source"], output_dir, choice["node_ids"], options
        )
        self._report(results)

    # ---------------------------------------------------------- conversion

    def is_running(self) -> bool:
        return self.worker is not None and self.worker.is_alive()

    _is_running = is_running  # the name the smoke test uses

    def _prepare_output_dir(self) -> Path | None:
        output_dir = Path(self.output_dir.get().strip() or "output").expanduser()
        try:
            output_dir.mkdir(parents=True, exist_ok=True)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("Lỗi thư mục", f"Không tạo được thư mục lưu:\n{exc}")
            return None
        self.last_output_dir = output_dir
        return output_dir

    def start_conversion(self) -> None:
        if self.is_running():
            self.cancel_event.set()
            self.status_label.configure(text="Đang huỷ…")
            return

        pane = self.current
        if not pane.files:
            messagebox.showinfo(
                "Chưa có file", "Hãy thêm ít nhất một file vào tab đang mở."
            )
            return

        output_dir = self._prepare_output_dir()
        if output_dir is None:
            return

        options = self.options_for(pane)
        pane.mark_pending()
        self.progress.set(0)
        self.cancel_event.clear()
        self.convert_button.configure(text="Huỷ", fg_color="#b04545", hover_color="#8f3737")
        self.status_label.configure(text="Bắt đầu chuyển đổi…")

        files = list(pane.files)
        self.worker = threading.Thread(
            target=self._run_batch, args=(files, output_dir, options), daemon=True
        )
        self.worker.start()

    def _run_batch(self, files, output_dir: Path, options: ConversionOptions) -> None:
        def on_progress(index: int, total: int, result: ConversionResult) -> None:
            self.events.put(("progress", index, total, result))
            if index < total:
                self.events.put(("busy", files[index]))

        if files:
            self.events.put(("busy", files[0]))
        try:
            results = convert_many(
                files,
                output_dir,
                options,
                on_progress=on_progress,
                cancel_event=self.cancel_event,
            )
        except Exception as exc:  # noqa: BLE001 - worker must never die silently
            self.events.put(("fatal", exc))
            return
        self.events.put(("done", results))

    def _row_for(self, path: Path) -> FileRow | None:
        for pane in self.queues.values():
            row = pane.rows.get(path)
            if row is not None:
                return row
        return None

    def _drain_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                self._handle_event(event)
        except queue.Empty:
            pass
        finally:
            self.root.after(80, self._drain_events)

    def _handle_event(self, event) -> None:
        kind = event[0]

        if kind == "busy":
            row = self._row_for(event[1])
            if row is not None:
                row.set_busy()
            return

        if kind == "progress":
            _, index, total, result = event
            self.progress.set(index / total if total else 1)
            row = self._row_for(result.source)
            if row is not None:
                row.set_status(result.status, self._describe(result))
            self.status_label.configure(
                text=f"Đang xử lý {index}/{total} — {result.source.name}"
            )
            return

        if kind == "outline":
            self._show_section_dialog(event[1])
            return

        if kind == "outline-error":
            self.sections_button.configure(state="normal")
            self.status_label.configure(text="Không đọc được mục lục.")
            messagebox.showerror(
                "Lỗi đọc tài liệu", f"{Path(event[1]).name}:\n{event[2]}"
            )
            return

        if kind == "fatal":
            self._finish()
            messagebox.showerror("Lỗi nghiêm trọng", str(event[1]))
            return

        if kind == "done":
            results = event[1]
            self._finish()
            self._report(results)

    def _describe(self, result: ConversionResult) -> str:
        if result.ok and result.output is not None:
            text = f"→ {result.output.name}  ({result.duration:.2f}s)"
            if result.warnings:
                text += f"  •  {len(result.warnings)} cảnh báo"
            return text
        return result.message or str(result.source.parent)

    def _finish(self) -> None:
        self.worker = None
        self.convert_button.configure(
            text="Chuyển đổi",
            fg_color=ctk.ThemeManager.theme["CTkButton"]["fg_color"],
            hover_color=ctk.ThemeManager.theme["CTkButton"]["hover_color"],
        )

    def _report(self, results) -> None:
        ok = sum(1 for r in results if r.status == STATUS_SUCCESS)
        failed = [r for r in results if r.status == STATUS_ERROR]
        cancelled = sum(1 for r in results if r.status == STATUS_CANCELLED)
        skipped = sum(1 for r in results if r.status == STATUS_SKIPPED)

        self.progress.set(1 if not cancelled else self.progress.get())
        summary = f"Hoàn tất: {ok} thành công, {len(failed)} lỗi"
        if skipped:
            summary += f", {skipped} bỏ qua"
        if cancelled:
            summary += f", {cancelled} bị huỷ"
        self.status_label.configure(text=summary + f". Lưu tại: {self.last_output_dir}")

        if failed:
            detail = "\n".join(f"• {r.source.name}: {r.message}" for r in failed[:10])
            if len(failed) > 10:
                detail += f"\n… và {len(failed) - 10} file khác."
            messagebox.showwarning("Có file lỗi", f"{summary}\n\n{detail}")
        elif ok:
            messagebox.showinfo("Xong", f"{summary}.\n\nThư mục: {self.last_output_dir}")

    # -------------------------------------------------------------- lifecycle

    def _on_close(self) -> None:
        if self.is_running():
            if not messagebox.askokcancel("Đang chạy", "Huỷ tiến trình và thoát?"):
                return
            self.cancel_event.set()
        self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()


def _to_float(value: str, fallback: float) -> float:
    try:
        return float(str(value).replace(",", "."))
    except (TypeError, ValueError):
        return fallback


def main() -> int:
    App().run()
    return 0
