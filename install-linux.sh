#!/usr/bin/env bash
# Install word2md for the current user: ~/.local/bin/word2md + app menu entry.
#
# From a GitHub release (script and binary downloaded into the same folder):
#   bash install-linux.sh
# From a source checkout, after ./build.sh:
#   ./install-linux.sh                       (uses dist/word2md-linux-<arch>)
# Any other binary:
#   bash install-linux.sh path/to/word2md-linux-x64
# Remove:
#   bash install-linux.sh --uninstall
set -euo pipefail

HERE="$(dirname "$(readlink -f "$0")")"
BIN_DIR="${HOME}/.local/bin"
APP_DIR="${XDG_DATA_HOME:-$HOME/.local/share}/applications"
DESKTOP="${APP_DIR}/word2md.desktop"

if [ "${1:-}" = "--uninstall" ]; then
    rm -f "${BIN_DIR}/word2md" "${BIN_DIR}/word2md.old" "${BIN_DIR}/word2md.new" "$DESKTOP"
    command -v update-desktop-database >/dev/null && update-desktop-database "$APP_DIR" || true
    echo "Đã gỡ word2md."
    exit 0
fi

case "$(uname -m)" in
    x86_64|amd64) ARCH=x64 ;;
    aarch64|arm64) ARCH=arm64 ;;
    *) ARCH="$(uname -m)" ;;
esac
NAME="word2md-linux-${ARCH}"

# The binary: the argument, else beside this script (release download),
# else in dist/ (source checkout).
if [ -n "${1:-}" ]; then
    SRC="$(readlink -f "$1")"
elif [ -f "${HERE}/${NAME}" ]; then
    SRC="${HERE}/${NAME}"
elif [ -f "${HERE}/dist/${NAME}" ]; then
    SRC="${HERE}/dist/${NAME}"
else
    echo "Không thấy ${NAME}." >&2
    echo "Tải ${NAME} từ trang release về cùng thư mục với script này," >&2
    echo "hoặc truyền đường dẫn: bash install-linux.sh path/to/${NAME}" >&2
    exit 1
fi
if [ ! -f "$SRC" ]; then
    echo "Không thấy file: $SRC" >&2
    exit 1
fi

# Verify against the release checksum when it was downloaded too.
SUM="${SRC}.sha256"
if [ -f "$SUM" ]; then
    expected="$(awk '{print $1; exit}' "$SUM")"
    actual="$(sha256sum "$SRC" | awk '{print $1}')"
    if [ "$expected" != "$actual" ]; then
        echo "Mã SHA-256 không khớp - file tải về bị lỗi, hãy tải lại." >&2
        exit 1
    fi
    echo "SHA-256: OK"
fi

mkdir -p "$BIN_DIR" "$APP_DIR"
install -m 755 "$SRC" "${BIN_DIR}/word2md"
cat > "$DESKTOP" <<DESKTOP_EOF
[Desktop Entry]
Type=Application
Name=word2md
Comment=Chuyển Word/Excel/PDF sang Markdown và ngược lại
Exec=${BIN_DIR}/word2md
Terminal=false
Categories=Office;Utility;
DESKTOP_EOF
command -v update-desktop-database >/dev/null && update-desktop-database "$APP_DIR" || true
echo "Đã cài: ${BIN_DIR}/word2md  (menu ứng dụng: word2md)"

# What the app needs from the system but does not carry itself.
missing=()
command -v tesseract >/dev/null || missing+=("tesseract-ocr tesseract-ocr-vie  (OCR PDF scan)")
if command -v tesseract >/dev/null && ! tesseract --list-langs 2>/dev/null | grep -qx vie; then
    missing+=("tesseract-ocr-vie  (OCR tiếng Việt)")
fi
command -v soffice >/dev/null || missing+=("libreoffice-writer  (đọc .doc, xuất PDF)")
if [ ${#missing[@]} -gt 0 ]; then
    echo
    echo "Chưa có (không bắt buộc, cài để dùng đủ tính năng):"
    printf '  sudo apt install %s\n' "${missing[@]}"
fi

case ":${PATH}:" in
    *":${BIN_DIR}:"*) ;;
    *) echo; echo "Lưu ý: ${BIN_DIR} chưa có trong PATH - đăng xuất/đăng nhập lại, hoặc chạy ${BIN_DIR}/word2md" ;;
esac
