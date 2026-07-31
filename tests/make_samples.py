"""Generate sample .docx / .xlsx files used for manual and automated testing."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from docx import Document
from docx.shared import Pt

SAMPLES_DIR = Path(__file__).resolve().parent / "samples"


def make_docx(path: Path) -> Path:
    document = Document()

    document.add_heading("Báo cáo kỹ thuật quý III", level=1)
    document.add_paragraph(
        "Tài liệu này được sinh tự động để kiểm thử bộ chuyển đổi Word sang Markdown."
    )

    document.add_heading("1. Tổng quan", level=2)
    paragraph = document.add_paragraph("Đoạn văn có ")
    paragraph.add_run("chữ in đậm").bold = True
    paragraph.add_run(", ")
    paragraph.add_run("chữ in nghiêng").italic = True
    paragraph.add_run(" và văn bản thường. Ký tự đặc biệt: & < > \" ' | 100%.")

    document.add_heading("2. Danh sách công việc", level=2)
    for item in ["Thu thập yêu cầu", "Thiết kế kiến trúc", "Triển khai & kiểm thử"]:
        document.add_paragraph(item, style="List Bullet")

    document.add_heading("3. Các bước thực hiện", level=2)
    for step in ["Chuẩn bị môi trường", "Chạy pipeline", "Bàn giao kết quả"]:
        document.add_paragraph(step, style="List Number")

    document.add_heading("4. Bảng số liệu", level=2)
    table = document.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    headers = ["Hạng mục", "Kế hoạch", "Thực hiện", "Ghi chú"]
    for cell, text in zip(table.rows[0].cells, headers):
        run = cell.paragraphs[0].add_run(text)
        run.bold = True
        run.font.size = Pt(11)

    rows = [
        ("Doanh thu", "1.000", "1.250", "Vượt 25%"),
        ("Chi phí", "800", "770", "Tiết kiệm"),
        ("Lợi nhuận", "200", "480", "Tốt | ổn định"),
    ]
    for row in rows:
        cells = table.add_row().cells
        for cell, text in zip(cells, row):
            cell.text = text

    document.add_heading("5. Trích dẫn", level=2)
    document.add_paragraph(
        "Chất lượng không phải là một hành động, đó là một thói quen.",
        style="Intense Quote",
    )

    # A second top-level branch with three heading levels, so the navigation
    # outline is a real tree rather than a flat list.
    document.add_heading("Phụ lục", level=1)
    document.add_paragraph("Phần tham chiếu bổ sung.")

    document.add_heading("A. Thuật ngữ", level=2)
    document.add_paragraph("Định nghĩa các thuật ngữ dùng trong báo cáo.")

    document.add_heading("A.1 Viết tắt", level=3)
    document.add_paragraph("KPI, SLA, OKR.")

    document.add_heading("A.2 Tham chiếu", level=3)
    document.add_paragraph("Xem thêm tài liệu nội bộ số 42.")

    document.add_heading("B. Liên hệ", level=2)
    document.add_paragraph("Phòng Kỹ thuật — ext. 1234.")

    path.parent.mkdir(parents=True, exist_ok=True)
    document.save(path)
    return path


def make_xlsx(path: Path) -> Path:
    sales = pd.DataFrame(
        {
            "Tháng": ["01/2026", "02/2026", "03/2026"],
            "Doanh thu": [1500000, 1820000, 2100500],
            "Số đơn": [120, 143, 165],
            "Tăng trưởng": [0.0, 0.213, 0.154],
        }
    )
    staff = pd.DataFrame(
        {
            "Mã NV": ["NV001", "NV002", "NV003"],
            "Họ tên": ["Trần Văn A", "Lê Thị B", "Phạm Minh C"],
            "Ngày vào làm": pd.to_datetime(["2021-03-01", "2022-07-15", "2024-01-08"]),
            "Đang làm việc": [True, True, False],
            "Ghi chú": ["Tốt", None, "Ký tự | đặc biệt"],
        }
    )
    empty = pd.DataFrame()

    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        sales.to_excel(writer, sheet_name="Doanh thu", index=False)
        staff.to_excel(writer, sheet_name="Nhân sự", index=False)
        empty.to_excel(writer, sheet_name="Trống", index=False)
    return path


def make_corrupt(path: Path) -> Path:
    """A file with a valid extension but garbage content, to test error handling."""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"this is definitely not a zip container")
    return path


def make_mislabelled(source: Path, path: Path) -> Path:
    """A modern file wearing a legacy extension - very common in the wild."""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(source.read_bytes())
    return path


def make_legacy(docx: Path, xlsx: Path, directory: Path) -> dict[str, Path]:
    """Downgrade the samples to .doc/.xls via Office COM, when available.

    Returns an empty dict on machines without Microsoft Office; the tests that
    need real binary samples skip themselves in that case.
    """
    import sys

    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    from src.legacy import LegacyUpgrader, has_msoffice

    if not (has_msoffice("Word.Application") and has_msoffice("Excel.Application")):
        return {}

    WD_FORMAT_97 = 0
    XL_FORMAT_97 = 56
    created: dict[str, Path] = {}
    upgrader = LegacyUpgrader()
    try:
        upgrader._initialize_com()

        word = upgrader._word_app()
        document = word.Documents.Open(str(docx.resolve()), ReadOnly=True, Visible=False)
        try:
            target = directory / "test.doc"
            document.SaveAs2(str(target.resolve()), FileFormat=WD_FORMAT_97)
            created["doc"] = target
        finally:
            document.Close(0)

        excel = upgrader._excel_app()
        workbook = excel.Workbooks.Open(str(xlsx.resolve()), ReadOnly=True)
        try:
            target = directory / "test.xls"
            workbook.SaveAs(str(target.resolve()), FileFormat=XL_FORMAT_97)
            created["xls"] = target
        finally:
            workbook.Close(False)
    finally:
        upgrader.close()

    return created


def build_all(directory: Path = SAMPLES_DIR, legacy: bool = True) -> dict[str, Path]:
    docx = make_docx(directory / "test.docx")
    xlsx = make_xlsx(directory / "test.xlsx")
    samples = {
        "docx": docx,
        "xlsx": xlsx,
        "corrupt": make_corrupt(directory / "corrupt.docx"),
        "fake_doc": make_mislabelled(docx, directory / "mislabelled.doc"),
        "fake_xls": make_mislabelled(xlsx, directory / "mislabelled.xls"),
    }
    if legacy:
        samples.update(make_legacy(docx, xlsx, directory))
    return samples


if __name__ == "__main__":
    for kind, created in build_all().items():
        print(f"{kind:10} -> {created}")
