"""Headless-ish GUI smoke test: build the window, run a real batch, close it.

Requires a desktop session (it creates a real Tk window briefly).
Run with: python tests/test_gui_smoke.py
"""

from __future__ import annotations

import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from make_samples import build_all  # noqa: E402


def check_section_dialog(app, docx: Path, out_dir: Path) -> None:
    """Drive the navigation-outline dialog without a human."""
    from src.converter import load_outline
    from src.section_dialog import SectionDialog

    outline = load_outline(docx)
    print(f"\nOutline of {docx.name}: {len(outline.nodes)} mục")

    dialog = SectionDialog(app.root, outline)
    dialog.update()

    appendix = next(n for n in outline.roots if n.title == "Phụ lục")
    glossary = appendix.children[0]

    dialog.set_checked(glossary.node_id, True)
    dialog.update()
    print(f"  summary: {dialog.summary.cget('text')}")

    preview = dialog.preview.get("1.0", "end")
    assert "# A. Thuật ngữ" in preview, preview[:200]
    assert "## A.1 Viết tắt" in preview, preview[:200]
    assert "Bảng số liệu" not in preview

    for child in glossary.children:
        assert dialog.checked[child.node_id], "children must follow the parent"
        assert dialog.rows[child.node_id].checkbox.cget("state") == "disabled"

    dialog.filter_text.set("liên hệ")
    dialog.update()
    visible = [nid for nid, row in dialog.rows.items() if row.winfo_manager()]
    print(f"  filter 'liên hệ' -> {len(visible)} hàng hiển thị")
    assert len(visible) < len(dialog.rows)
    dialog.filter_text.set("")
    dialog.update()

    dialog.set_all_expanded(False)
    dialog.update()
    dialog.set_all_expanded(True)
    dialog.update()

    dialog.confirm()
    assert dialog.result is not None
    assert dialog.result["node_ids"]
    print(f"  result: {dialog.result['node_ids']}")

    from src.converter import ConversionOptions, convert_docx_sections

    results = convert_docx_sections(
        docx,
        out_dir,
        dialog.result["node_ids"],
        ConversionOptions(
            promote_headings=dialog.result["promote"],
            split_sections=dialog.result["split"],
        ),
    )
    print(f"  exported: {[r.output.name for r in results if r.output]}")
    assert all(r.status == "success" for r in results)


def check_app_wiring(app, docx: Path, out_dir: Path) -> None:
    """Exercise App.open_section_extractor with the dialog stubbed out."""
    import src.gui as gui

    class FakeDialog:
        def __init__(self, master, outline):
            self.outline = outline

        def show(self):
            return {
                "source": self.outline.source,
                "node_ids": [self.outline.roots[0].children[0].node_id],
                "split": False,
                "promote": True,
            }

    original_dialog, original_report = gui.SectionDialog, app._report
    reported = []
    gui.SectionDialog = FakeDialog
    app._report = reported.append
    app.output_dir.set(str(out_dir))
    try:
        app.open_section_extractor(docx)
        deadline = time.time() + 30
        while not reported and time.time() < deadline:
            app.root.update()
            time.sleep(0.02)
    finally:
        gui.SectionDialog = original_dialog
        app._report = original_report

    assert reported, "App never produced a result"
    results = reported[0]
    assert results[0].status == "success", results[0].message
    print(f"\nApp wiring OK -> {results[0].output.name}")


def main() -> int:
    from src.gui import DND_AVAILABLE, App
    from src.gui import legacy_support_note as gui_legacy_note

    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        samples = build_all(tmp_path / "samples", legacy=False)
        out_dir = tmp_path / "out"

        app = App()
        app.root.withdraw()
        print(f"Window created. Drag & drop available: {DND_AVAILABLE}")
        print(f"tkdnd version: {getattr(app.root, 'TkdndVersion', None)}")
        print(f"Legacy support: {gui_legacy_note()}")

        app.add_paths([tmp_path / "samples"])
        app.output_dir.set(str(out_dir))
        app.root.update()
        print(f"Queued {len(app.files)} file(s): {[p.name for p in app.files]}")
        assert len(app.files) == len(samples), f"{len(app.files)} != {len(samples)}"

        # Replace the summary dialog so the run stays unattended.
        app._report = lambda results: print(
            "Summary: "
            + ", ".join(f"{r.source.name}={r.status}" for r in results)
        )

        app.root.update()
        app.start_conversion()

        deadline = time.time() + 60
        while app._is_running() and time.time() < deadline:
            app.root.update()
            time.sleep(0.02)

        for _ in range(30):
            app.root.update()
            time.sleep(0.02)

        produced = sorted(p.name for p in out_dir.glob("*.md"))
        print(f"Produced: {produced}")
        assert "test.md" in produced
        assert "mislabelled.md" in produced
        # Every sample converts except the deliberately corrupt one.
        assert len(produced) == len(samples) - 1, produced

        for path, row in app.rows.items():
            print(f"  row {path.name}: {row.badge.cget('text')} | {row.detail.cget('text')}")

        check_section_dialog(app, samples["docx"], tmp_path / "sections")
        check_app_wiring(app, samples["docx"], tmp_path / "wired")

        app.clear_files()
        app.root.update()
        assert not app.files

        app.root.destroy()

    print("\nGUI smoke test PASSED")
    return 0


if __name__ == "__main__":
    sys.exit(main())
