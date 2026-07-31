# word2md

Ứng dụng desktop chuyển đổi file **Word (.docx, .doc)** và **Excel (.xlsx, .xlsm,
.xls)** sang **Markdown (.md)**, hỗ trợ kéo–thả và xử lý hàng loạt.

## Tính năng

- Giao diện `customtkinter` (sáng / tối / theo hệ thống).
- Kéo & thả file hoặc cả thư mục vào cửa sổ (qua `tkinterdnd2`), hoặc dùng nút chọn.
- Chọn thư mục → tự quét đệ quy mọi file Word/Excel bên trong.
- Thanh tiến trình + danh sách trạng thái từng file (chờ / đang xử lý / OK / lỗi).
- Chuyển đổi chạy trên thread riêng, có nút huỷ, giao diện không bị treo.
- Một file lỗi **không** làm dừng cả lô: lỗi được ghi nhận và tiếp tục file kế tiếp.
- Chọn thư mục lưu, tuỳ chọn ghi đè hoặc tự đánh số `tên (1).md`.
- **Trích xuất một phần tài liệu Word theo mục lục** (xem bên dưới).
- Có sẵn chế độ dòng lệnh cho việc tự động hoá.

## Trích xuất theo mục (Navigation Pane)

Nút **“Trích xuất theo mục…”** trên thanh công cụ — hoặc nút **“Mục…”** ngay trên
dòng của file `.docx` — mở một cửa sổ tái hiện đúng khung Navigation Pane của Word:

- Cây tiêu đề `Heading 1`–`Heading 6` với thu gọn / mở rộng và ô tick từng mục.
- Ô lọc theo tiêu đề, các nút *Chọn tất cả / Bỏ chọn / Mở rộng / Thu gọn*.
- Khung **xem trước Markdown** cập nhật ngay theo lựa chọn, kèm số mục, số dòng
  và số file sẽ tạo.
- Phần nội dung nằm trước tiêu đề đầu tiên xuất hiện dưới dạng mục `(Phần mở đầu)`.

Chọn một mục là lấy trọn **cả nhánh** của nó: tiêu đề, nội dung và mọi mục con, cho
đến trước tiêu đề cùng cấp hoặc cao hơn kế tiếp. Vì thế khi một mục cha đã được tick,
các mục con của nó tự động được tick và khoá lại (đã nằm trong phần xuất rồi).

Hai tuỳ chọn khi xuất:

- **Tách mỗi mục thành một file `.md` riêng** — thay vì gộp tất cả vào một file.
- **Nâng bậc tiêu đề đã chọn lên H1** — trích một mục `Heading 3` sẽ cho ra file bắt
  đầu bằng `#` thay vì `###`, các mục con được nâng theo cùng một mức nên cấu trúc
  phân cấp tương đối vẫn giữ nguyên.

Mã mục có dạng phân cấp (`2`, `2.1`, `2.1.3`) và được dùng chung cho cả GUI lẫn CLI.

### Word → Markdown

`mammoth` đọc `.docx` thành HTML ngữ nghĩa, sau đó module `src/html_to_markdown.py`
render sang Markdown. (Bộ ghi Markdown sẵn có của mammoth đã bị deprecated và bỏ
qua bảng biểu, nên dự án tự render để giữ được bảng.)

Giữ nguyên: tiêu đề `h1`–`h6`, đoạn văn, in đậm / in nghiêng / gạch ngang, danh sách
có thứ tự và không thứ tự (kể cả lồng nhau), bảng biểu, blockquote, khối code,
liên kết và ảnh. Ảnh nhúng được ghi ra thư mục `<tên file>_images/` và tham chiếu
bằng đường dẫn tương đối.

### Excel → Markdown

`pandas` + `openpyxl` đọc **tất cả** sheet. Mỗi sheet trở thành một mục `## Tên sheet`
kèm một bảng Markdown (`DataFrame.to_markdown()` với `tabulate`), tất cả gộp vào
**một** file `.md`. Ngày tháng được chuẩn hoá `YYYY-MM-DD`, ô trống thành rỗng, ký tự
`|` được escape, xuống dòng trong ô thành `<br>`. Giá trị được xuất nguyên văn
(`disable_numparse`) nên mã số dạng `007` hay ID dài không bị biến dạng.

## Định dạng cũ: `.doc` và `.xls`

Hai định dạng nhị phân của Office 97–2003 không thư viện Python thuần nào đọc trọn
vẹn được, nên app thử lần lượt nhiều backend và **luôn xuống cấp êm** thay vì báo lỗi:

| Định dạng | Ưu tiên 1 | Ưu tiên 2 | Ưu tiên 3 |
|---|---|---|---|
| `.doc` | Microsoft Word (COM) | LibreOffice | Trích xuất văn bản từ OLE |
| `.xls` | `xlrd` (Python thuần) | Microsoft Excel (COM) | LibreOffice |

Với `.xls`, `xlrd` đọc trực tiếp nên **không cần cài Office** và không phải khởi động
Excel — dữ liệu ra giống hệt file `.xlsx` tương ứng.

Với `.doc`, nếu có Word hoặc LibreOffice thì file được nâng cấp lên `.docx` rồi đi qua
đúng pipeline chuẩn, giữ nguyên tiêu đề / danh sách / bảng biểu — kết quả **giống từng
byte** với việc convert file `.docx` gốc. Nếu máy không có cả hai, app tự parse cấu trúc
OLE (piece table của Word 97) để lấy văn bản thuần và ghi rõ cảnh báo là đã mất định dạng.

Xem máy hiện có backend nào:

```powershell
.venv\Scripts\python.exe main.py --backends
```

Hai chi tiết đáng lưu ý:

- **File gắn sai đuôi được nhận diện tự động.** Rất nhiều file `.doc` ngoài thực tế
  thật ra là `.docx` hoặc RTF đổi tên. App đọc magic bytes chứ không tin phần mở rộng,
  nên trường hợp này chạy được ngay cả khi không có Office.
- **Kết quả nâng cấp được cache** theo (đường dẫn, mtime, kích thước) trong suốt phiên
  làm việc, nên xem mục lục rồi mới xuất sẽ không phải chạy Word hai lần. Cả một lô
  file dùng chung đúng một tiến trình Word/Excel.

## Cấu trúc dự án

```
word2md/
├── main.py                    # Điểm vào: không tham số -> GUI, có tham số -> CLI
├── build.py                   # Script đóng gói PyInstaller (x64 / ARM64)
├── build.bat / run.bat        # Tiện ích cho Windows
├── .github/workflows/build.yml # CI: build cả bản x64 và ARM64
├── requirements.txt           # Phụ thuộc lúc chạy
├── requirements-dev.txt       # + pyinstaller, python-docx (sinh file mẫu)
├── src/
│   ├── converter.py           # Core engine
│   ├── html_to_markdown.py    # HTML (mammoth) -> Markdown
│   ├── legacy.py              # Backend đọc .doc/.xls + parser OLE
│   ├── outline.py             # Cây mục lục + trích xuất từng phần
│   ├── section_dialog.py      # Cửa sổ chọn mục
│   └── gui.py                 # Giao diện customtkinter
├── tests/
│   ├── make_samples.py        # Sinh test.docx / test.xlsx / corrupt.docx
│   ├── test_converter.py      # Unit test cho engine và HTML -> Markdown
│   ├── test_outline.py        # Unit test cho mục lục và trích xuất
│   ├── test_legacy.py         # Unit test cho .doc/.xls
│   └── test_gui_smoke.py      # Dựng cửa sổ thật, chạy một lô + cửa sổ chọn mục
└── output/                    # Thư mục lưu mặc định
```

## Cài đặt

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
```

## Chạy

```powershell
# Giao diện
.venv\Scripts\python.exe main.py
# hoặc: run.bat

# Dòng lệnh
.venv\Scripts\python.exe main.py bao-cao.docx so-lieu.xlsx -o output
.venv\Scripts\python.exe main.py .\tai-lieu -o .\md --overwrite

# Trích xuất theo mục
.venv\Scripts\python.exe main.py bao-cao.docx --list-sections
.venv\Scripts\python.exe main.py bao-cao.docx --sections 2,3.1 -o output
.venv\Scripts\python.exe main.py bao-cao.docx --sections all --split-sections
```

`--list-sections` in ra cây mục lục kèm mã mục:

```
1         Báo cáo kỹ thuật quý III  [H1, 32 dòng]
1.1         1. Tổng quan  [H2, 4 dòng]
1.2         2. Danh sách công việc  [H2, 6 dòng]
```

Tuỳ chọn CLI: `-o/--output`, `--no-recursive`, `--no-images`, `--overwrite`,
`--no-title`, `--backends`, `--list-sections`, `--sections ID[,ID…]`,
`--split-sections`, `--no-promote`. Nhóm `--sections` chỉ nhận đúng một file Word.
Mã thoát: `0` thành công, `1` sai tham số / không tìm thấy file, `2` có file lỗi.

## Kiểm thử

```powershell
.venv\Scripts\python.exe tests\make_samples.py      # sinh file mẫu
.venv\Scripts\python.exe -m unittest discover -s tests -t tests -v
.venv\Scripts\python.exe tests\test_gui_smoke.py    # cần phiên desktop
```

## Đóng gói `.exe`

```powershell
.venv\Scripts\python.exe build.py --clean
# hoặc: build.bat
```

Kết quả: `dist\word2md-x64.exe` (~39 MB, một file duy nhất, không cửa sổ console).
Tên file mang theo kiến trúc nên bản x64 và bản ARM64 nằm cạnh nhau trong `dist\`
được; `--clean` chỉ xoá đúng bản của kiến trúc đang build.

Lệnh PyInstaller tương đương:

```powershell
pyinstaller --noconfirm --clean --onefile --noconsole --name word2md-x64 `
    --collect-data customtkinter --collect-all tkinterdnd2 `
    --hidden-import tabulate --hidden-import openpyxl `
    main.py
```

`--collect-data customtkinter` mang theo các theme JSON, `--collect-all tkinterdnd2`
mang theo thư viện nhị phân `tkdnd`, và `tabulate` phải khai báo tường minh vì
`pandas` chỉ import nó lúc chạy bên trong `to_markdown()`.

Đặt icon tại `assets/icon.ico` thì `build.py` sẽ tự dùng.

### Bản Windows ARM64

PyInstaller **không cross-compile** (`--target-arch` chỉ dành cho macOS), nên kiến trúc
của `.exe` luôn là kiến trúc của Python đang chạy `build.py`. Không thể tạo bản ARM64
từ máy x64. `build.py --arch arm64` sẽ báo lỗi ngay thay vì âm thầm xuất ra bản x64, và
sau khi build xong nó đọc PE header của file để kiểm tra lại kiến trúc thật.

Hai cách lấy bản ARM64:

```powershell
# 1. Trên máy Windows on ARM, với Python ARM64 (không phải bản x64 chạy giả lập)
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.venv\Scripts\python.exe build.py --clean --arch arm64   # -> dist\word2md-arm64.exe
```

2. Qua CI: `.github/workflows/build.yml` build cả hai bản, bản ARM64 chạy trên runner
   `windows-11-arm`. Runner này **chỉ miễn phí cho repo public**; repo private thì phải
   build trên máy thật hoặc dùng larger runner ARM64.

Toàn bộ phụ thuộc đều có wheel `win_arm64` cho CPython 3.13, kể cả `pywin32` (nên
`.doc`/`.xls` qua COM vẫn chạy) và `tkinterdnd2` (có sẵn `tkdnd/win-arm64` nên kéo-thả
vẫn native). Riêng `pandas` chỉ phát hành wheel `win_arm64` từ 3.0 trở đi, vì vậy
`requirements.txt` ghim theo kiến trúc:

```
pandas==2.3.3; platform_machine != "ARM64"
pandas>=3.0.5; platform_machine == "ARM64"
```

Cả 81 test đều pass trên pandas 3.0.5 và output Markdown giống hệt bản pandas 2.3.3.

Nếu chưa có bản ARM64, bản x64 vẫn chạy được trên Windows on ARM qua lớp giả lập
Prism, chỉ chậm hơn.

## Giới hạn đã biết

- Bảng biểu và ảnh trong `.doc` chỉ giữ được khi máy có Microsoft Word hoặc
  LibreOffice; bản fallback thuần Python chỉ lấy văn bản.
- Trích xuất theo mục dựa trên các style `Heading 1`–`Heading 6`. Tài liệu chỉ dùng
  chữ to / in đậm để giả làm tiêu đề sẽ không có mục lục — Word Navigation Pane
  cũng để trống trong trường hợp này.
- Ô gộp (merged cell) trong bảng Word được trải phẳng vì Markdown không có khái niệm
  colspan/rowspan.
- File đang mở trong Word/Excel có thể báo lỗi quyền truy cập; đóng file rồi thử lại.
