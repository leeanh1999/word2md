"""Dialog that mirrors Word's Navigation Pane so the user can export only
the sections they care about.
"""

from __future__ import annotations

from pathlib import Path

import customtkinter as ctk

from .converter import DocumentOutline
from .outline import OutlineNode, iter_nodes, top_level_selection

MAX_PREVIEW_CHARS = 60_000
INDENT_PER_LEVEL = 18


def ancestor_ids(node_id: str) -> list[str]:
    parts = node_id.split(".")
    return [".".join(parts[:count]) for count in range(1, len(parts))]


class _OutlineRow(ctk.CTkFrame):
    """One heading in the tree: expander, checkbox, rank badge."""

    def __init__(self, master, node: OutlineNode, dialog: "SectionDialog"):
        super().__init__(master, fg_color="transparent")
        self.node = node
        self.dialog = dialog
        self.depth = node.node_id.count(".")

        self.grid_columnconfigure(2, weight=1)

        pad = 6 + self.depth * INDENT_PER_LEVEL
        self.expander = ctk.CTkButton(
            self,
            text="▾" if node.children else "",
            width=20,
            height=20,
            fg_color="transparent",
            hover_color=("#d9d9d9", "#404040"),
            text_color=("#5c5c5c", "#9a9a9a"),
            command=self.dialog_toggle,
        )
        self.expander.grid(row=0, column=0, padx=(pad, 0))
        if not node.children:
            self.expander.configure(state="disabled", hover=False)

        rank = "¶" if node.is_preamble else f"H{node.level}"
        self.rank = ctk.CTkLabel(
            self,
            text=rank,
            width=26,
            text_color=("#8a8a8a", "#7a7a7a"),
            font=ctk.CTkFont(size=10, weight="bold"),
        )
        self.rank.grid(row=0, column=1, padx=(2, 4))

        self.variable = ctk.BooleanVar(value=False)
        weight = "bold" if node.level in (0, 1) else "normal"
        self.checkbox = ctk.CTkCheckBox(
            self,
            text=_ellipsize(node.title, 70),
            variable=self.variable,
            checkbox_width=18,
            checkbox_height=18,
            font=ctk.CTkFont(size=12, weight=weight),
            command=self.on_toggle,
        )
        self.checkbox.grid(row=0, column=2, sticky="w", pady=2)

        self.count = ctk.CTkLabel(
            self,
            text=f"{node.line_count} dòng",
            text_color=("#8a8a8a", "#7a7a7a"),
            font=ctk.CTkFont(size=10),
        )
        self.count.grid(row=0, column=3, padx=8)

    def dialog_toggle(self) -> None:
        self.dialog.toggle_expanded(self.node.node_id)

    def on_toggle(self) -> None:
        self.dialog.set_checked(self.node.node_id, self.variable.get())

    def set_expander(self, expanded: bool) -> None:
        if self.node.children:
            self.expander.configure(text="▾" if expanded else "▸")

    def set_checked(self, value: bool) -> None:
        self.variable.set(value)

    def set_locked(self, locked: bool) -> None:
        """A section covered by a checked ancestor is implied, not editable."""
        self.checkbox.configure(state="disabled" if locked else "normal")
        self.rank.configure(
            text_color=("#b0b0b0", "#5c5c5c") if locked else ("#8a8a8a", "#7a7a7a")
        )


class SectionDialog(ctk.CTkToplevel):
    def __init__(self, master, outline: DocumentOutline):
        super().__init__(master)
        self.outline = outline
        self.nodes = outline.nodes
        self.order = list(iter_nodes(outline.roots))
        self.result: dict | None = None

        self.checked: dict[str, bool] = {node.node_id: False for node in self.order}
        self.expanded: dict[str, bool] = {node.node_id: True for node in self.order}
        self.rows: dict[str, _OutlineRow] = {}

        self.split = ctk.BooleanVar(value=False)
        self.promote = ctk.BooleanVar(value=True)
        self.filter_text = ctk.StringVar(value="")

        self.title(f"Trích xuất theo mục — {outline.source.name}")
        self.geometry("1080x660")
        self.minsize(860, 520)
        self.transient(master)

        self._build_ui()
        self._populate()
        self._refresh()

        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.bind("<Escape>", lambda _event: self.cancel())
        self.after(120, self._grab)

    def _grab(self) -> None:
        try:
            self.grab_set()
            self.focus_force()
        except Exception:  # noqa: BLE001
            pass

    # ------------------------------------------------------------------ ui

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=3, uniform="panes")
        self.grid_columnconfigure(1, weight=2, uniform="panes")
        self.grid_rowconfigure(1, weight=1)

        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.grid(row=0, column=0, columnspan=2, sticky="ew", padx=14, pady=(12, 4))
        toolbar.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            toolbar,
            text="Chọn các mục cần xuất",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).grid(row=0, column=0, sticky="w")

        search = ctk.CTkEntry(
            toolbar, placeholder_text="Lọc theo tiêu đề…", textvariable=self.filter_text
        )
        search.grid(row=0, column=1, sticky="ew", padx=14)
        self.filter_text.trace_add("write", lambda *_: self._refresh_visibility())

        actions = (
            ("Chọn tất cả", self.select_all),
            ("Bỏ chọn", self.select_none),
            ("Mở rộng", lambda: self.set_all_expanded(True)),
            ("Thu gọn", lambda: self.set_all_expanded(False)),
        )
        for offset, (text, command) in enumerate(actions):
            ctk.CTkButton(
                toolbar,
                text=text,
                width=92,
                height=28,
                fg_color="transparent",
                border_width=1,
                text_color=("#1f6aa5", "#5aa9e6"),
                command=command,
            ).grid(row=0, column=2 + offset, padx=(0, 6))

        self.tree = ctk.CTkScrollableFrame(self, label_text="Mục lục tài liệu")
        self.tree.grid(row=1, column=0, sticky="nsew", padx=(14, 7), pady=6)
        self.tree.grid_columnconfigure(0, weight=1)

        preview_frame = ctk.CTkFrame(self)
        preview_frame.grid(row=1, column=1, sticky="nsew", padx=(7, 14), pady=6)
        preview_frame.grid_rowconfigure(1, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            preview_frame,
            text="Xem trước Markdown",
            font=ctk.CTkFont(size=13, weight="bold"),
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(10, 4))
        self.preview = ctk.CTkTextbox(
            preview_frame, wrap="none", font=ctk.CTkFont(family="Consolas", size=12)
        )
        self.preview.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        options = ctk.CTkFrame(self, fg_color="transparent")
        options.grid(row=2, column=0, columnspan=2, sticky="ew", padx=14, pady=(0, 4))
        ctk.CTkCheckBox(
            options,
            text="Tách mỗi mục thành một file .md riêng",
            variable=self.split,
            command=self._refresh_preview,
        ).pack(side="left", padx=(0, 20))
        ctk.CTkCheckBox(
            options,
            text="Nâng bậc tiêu đề đã chọn lên H1",
            variable=self.promote,
            command=self._refresh_preview,
        ).pack(side="left")

        footer = ctk.CTkFrame(self, fg_color="transparent")
        footer.grid(row=3, column=0, columnspan=2, sticky="ew", padx=14, pady=(4, 14))
        footer.grid_columnconfigure(0, weight=1)

        self.summary = ctk.CTkLabel(
            footer, text="", anchor="w", text_color=("#5c5c5c", "#9a9a9a")
        )
        self.summary.grid(row=0, column=0, sticky="ew")

        ctk.CTkButton(
            footer,
            text="Huỷ",
            width=100,
            height=36,
            fg_color="transparent",
            border_width=1,
            text_color=("#1f6aa5", "#5aa9e6"),
            command=self.cancel,
        ).grid(row=0, column=1, padx=(0, 8))

        self.export_button = ctk.CTkButton(
            footer,
            text="Xuất",
            width=140,
            height=36,
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.confirm,
        )
        self.export_button.grid(row=0, column=2)

    def _populate(self) -> None:
        if not self.order:
            ctk.CTkLabel(
                self.tree,
                text="Tài liệu không có tiêu đề nào (Heading 1–6).\n"
                "Hãy dùng nút “Chuyển đổi” để xuất toàn bộ nội dung.",
                text_color=("#8a8a8a", "#7a7a7a"),
                justify="center",
            ).grid(row=0, column=0, pady=50)
            return

        for position, node in enumerate(self.order):
            row = _OutlineRow(self.tree, node, self)
            row.grid(row=position, column=0, sticky="ew")
            self.rows[node.node_id] = row

    # --------------------------------------------------------------- state

    def toggle_expanded(self, node_id: str) -> None:
        self.expanded[node_id] = not self.expanded.get(node_id, True)
        self.rows[node_id].set_expander(self.expanded[node_id])
        self._refresh_visibility()

    def set_all_expanded(self, value: bool) -> None:
        for node_id, row in self.rows.items():
            self.expanded[node_id] = value
            row.set_expander(value)
        self._refresh_visibility()

    def set_checked(self, node_id: str, value: bool) -> None:
        for descendant in self.nodes[node_id].walk():
            self.checked[descendant.node_id] = value
            row = self.rows.get(descendant.node_id)
            if row is not None:
                row.set_checked(value)
        self._refresh()

    def select_all(self) -> None:
        for node in self.outline.roots:
            self.set_checked(node.node_id, True)

    def select_none(self) -> None:
        for node_id in self.checked:
            self.checked[node_id] = False
            row = self.rows.get(node_id)
            if row is not None:
                row.set_checked(False)
        self._refresh()

    def selected_ids(self) -> list[str]:
        return [node_id for node_id, value in self.checked.items() if value]

    # ------------------------------------------------------------- refresh

    def _refresh(self) -> None:
        for node_id, row in self.rows.items():
            row.set_locked(
                any(self.checked.get(parent) for parent in ancestor_ids(node_id))
            )
        self._refresh_visibility()
        self._refresh_preview()

    def _refresh_visibility(self) -> None:
        needle = self.filter_text.get().strip().lower()
        matched: set[str] = set()
        if needle:
            for node in self.order:
                if needle in node.title.lower():
                    matched.add(node.node_id)
                    matched.update(ancestor_ids(node.node_id))

        for node_id, row in self.rows.items():
            if needle:
                visible = node_id in matched
            else:
                visible = all(
                    self.expanded.get(parent, True) for parent in ancestor_ids(node_id)
                )
            if visible:
                row.grid()
            else:
                row.grid_remove()

    def _refresh_preview(self) -> None:
        ids = self.selected_ids()
        text = (
            self.outline.preview(ids, promote=self.promote.get())
            if ids
            else "(Chưa chọn mục nào)"
        )
        truncated = text[:MAX_PREVIEW_CHARS]
        if len(text) > MAX_PREVIEW_CHARS:
            truncated += "\n\n… (rút gọn phần xem trước)"

        self.preview.configure(state="normal")
        self.preview.delete("1.0", "end")
        self.preview.insert("1.0", truncated)
        self.preview.configure(state="disabled")

        top = top_level_selection(self.outline.roots, ids)
        lines = text.count("\n") if ids else 0
        files = len(top) if self.split.get() else (1 if top else 0)
        self.summary.configure(
            text=f"Đã chọn {len(top)} mục · ~{lines} dòng · tạo {files} file .md"
        )
        self.export_button.configure(state="normal" if top else "disabled")

    # -------------------------------------------------------------- result

    def confirm(self) -> None:
        ids = self.selected_ids()
        if not ids:
            return
        self.result = {
            "source": self.outline.source,
            "node_ids": ids,
            "split": self.split.get(),
            "promote": self.promote.get(),
        }
        self._close()

    def cancel(self) -> None:
        self.result = None
        self._close()

    def _close(self) -> None:
        try:
            self.grab_release()
        except Exception:  # noqa: BLE001
            pass
        self.destroy()

    def show(self) -> dict | None:
        self.wait_window()
        return self.result


class DocumentChooser(ctk.CTkToplevel):
    """Pick which queued .docx to open in the section extractor."""

    def __init__(self, master, paths: list[Path]):
        super().__init__(master)
        self.paths = paths
        self.result: Path | None = None

        self.title("Chọn tài liệu Word")
        self.geometry("520x180")
        self.resizable(False, False)
        self.transient(master)
        self.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self,
            text="Trích xuất theo mục cho tài liệu nào?",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")

        self.choice = ctk.StringVar(value=paths[0].name)
        ctk.CTkOptionMenu(
            self, values=[path.name for path in paths], variable=self.choice
        ).grid(row=1, column=0, padx=20, sticky="ew")

        buttons = ctk.CTkFrame(self, fg_color="transparent")
        buttons.grid(row=2, column=0, padx=20, pady=20, sticky="e")
        ctk.CTkButton(
            buttons,
            text="Huỷ",
            width=90,
            fg_color="transparent",
            border_width=1,
            text_color=("#1f6aa5", "#5aa9e6"),
            command=self._cancel,
        ).pack(side="left", padx=(0, 8))
        ctk.CTkButton(buttons, text="Tiếp tục", width=110, command=self._ok).pack(
            side="left"
        )

        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.bind("<Escape>", lambda _event: self._cancel())
        self.after(120, self._grab)

    def _grab(self) -> None:
        try:
            self.grab_set()
            self.focus_force()
        except Exception:  # noqa: BLE001
            pass

    def _ok(self) -> None:
        name = self.choice.get()
        self.result = next((p for p in self.paths if p.name == name), None)
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()

    def show(self) -> Path | None:
        self.wait_window()
        return self.result


def _ellipsize(text: str, limit: int) -> str:
    text = " ".join(text.split())
    return text if len(text) <= limit else text[: limit - 1] + "…"
