#!/usr/bin/env bash
# Install a built binary for the current user: ~/.local/bin/word2md + menu entry.
#   ./install-linux.sh                      -> uses dist/word2md-linux-<arch>
#   ./install-linux.sh path/to/word2md-linux-x64
#   ./install-linux.sh --uninstall
set -euo pipefail
cd "$(dirname "$(readlink -f "$0")")"

BIN_DIR="${HOME}/.local/bin"
APP_DIR="${XDG_DATA_HOME:-$HOME/.local/share}/applications"
DESKTOP="${APP_DIR}/word2md.desktop"

if [ "${1:-}" = "--uninstall" ]; then
    rm -f "${BIN_DIR}/word2md" "$DESKTOP"
    echo "Đã gỡ word2md."
    exit 0
fi

case "$(uname -m)" in
    x86_64|amd64) ARCH=x64 ;;
    aarch64|arm64) ARCH=arm64 ;;
    *) ARCH="$(uname -m)" ;;
esac
SRC="${1:-dist/word2md-linux-${ARCH}}"
if [ ! -f "$SRC" ]; then
    echo "Không thấy $SRC - chạy ./build.sh trước hoặc truyền đường dẫn binary." >&2
    exit 1
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
echo "Đã cài: ${BIN_DIR}/word2md  (menu: word2md)"
