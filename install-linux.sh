#!/usr/bin/env bash
# Install word2md for the current user, with the system packages it needs.
#
#   bash install-linux.sh                 cài gói còn thiếu (hỏi trước) + cài app
#   bash install-linux.sh --yes           không hỏi lại (vẫn cần mật khẩu sudo)
#   bash install-linux.sh --no-deps       chỉ cài app, không đụng gói hệ thống
#   bash install-linux.sh FILE            cài binary ở đường dẫn FILE
#   bash install-linux.sh --uninstall     gỡ app (không gỡ gói hệ thống)
#
# The binary is looked for beside this script (a release download), then in
# dist/ (a source checkout). The app goes to ~/.local/bin without sudo, so the
# in-app updater can replace it later; only the system packages need sudo.
#
# System packages, installed only when missing:
#   tesseract + Vietnamese data   OCR of scanned PDFs
#   LibreOffice Writer            .doc input and Markdown -> PDF
# LibreOffice is installed as the full libreoffice-writer and only when no
# soffice exists: the -nogui variant conflicts with a desktop LibreOffice and
# would make the package manager remove it.
set -euo pipefail

HERE="$(dirname "$(readlink -f "$0")")"
BIN_DIR="${HOME}/.local/bin"
APP_DIR="${XDG_DATA_HOME:-$HOME/.local/share}/applications"
DESKTOP="${APP_DIR}/word2md.desktop"
ICON_DIR="${XDG_DATA_HOME:-$HOME/.local/share}/icons/hicolor/256x256/apps"

ASSUME_YES=0
WITH_DEPS=1
SRC_ARG=""
for arg in "$@"; do
    case "$arg" in
        --uninstall)
            rm -f "${BIN_DIR}/word2md" "${BIN_DIR}/word2md.old" "${BIN_DIR}/word2md.new" "$DESKTOP" \
                "${ICON_DIR}/word2md.png"
            command -v update-desktop-database >/dev/null && update-desktop-database "$APP_DIR" || true
            echo "Đã gỡ word2md (gói hệ thống như tesseract, LibreOffice được giữ nguyên)."
            exit 0
            ;;
        -y|--yes) ASSUME_YES=1 ;;
        --no-deps) WITH_DEPS=0 ;;
        -h|--help) sed -n '2,10p' "$0" | sed 's/^# \{0,1\}//'; exit 0 ;;
        -*) echo "Không hiểu tuỳ chọn: $arg (xem --help)" >&2; exit 1 ;;
        *) SRC_ARG="$arg" ;;
    esac
done

# ------------------------------------------------------------ the binary

case "$(uname -m)" in
    x86_64|amd64) ARCH=x64 ;;
    aarch64|arm64) ARCH=arm64 ;;
    *) ARCH="$(uname -m)" ;;
esac
NAME="word2md-linux-${ARCH}"

if [ -n "$SRC_ARG" ]; then
    SRC="$(readlink -f "$SRC_ARG")"
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
[ -f "$SRC" ] || { echo "Không thấy file: $SRC" >&2; exit 1; }

# Check the download before anything is installed.
if [ -f "${SRC}.sha256" ]; then
    expected="$(awk '{print $1; exit}' "${SRC}.sha256")"
    actual="$(sha256sum "$SRC" | awk '{print $1}')"
    if [ "$expected" != "$actual" ]; then
        echo "Mã SHA-256 không khớp - file tải về bị lỗi, hãy tải lại." >&2
        exit 1
    fi
    echo "SHA-256: OK"
fi

# ------------------------------------------------------ system packages

has_tesseract() { command -v tesseract >/dev/null; }
# tesseract 4 prints the list on stderr, 5 on stdout.
has_vie() { has_tesseract && tesseract --list-langs 2>&1 | grep -qx vie; }
has_soffice() { command -v soffice >/dev/null || command -v libreoffice >/dev/null; }

missing_features() {
    local out=()
    has_tesseract || out+=("OCR (tesseract)")
    has_vie || out+=("OCR tiếng Việt (tesseract vie)")
    has_soffice || out+=("LibreOffice (đọc .doc, xuất PDF)")
    if [ ${#out[@]} -gt 0 ]; then printf '%s\n' "${out[@]}"; fi
}

# Package names per package manager for what is missing right now.
PM=""
PKGS=()
plan_packages() {
    PKGS=()
    if command -v apt-get >/dev/null; then
        PM=apt
        has_tesseract || PKGS+=(tesseract-ocr)
        has_vie || PKGS+=(tesseract-ocr-vie)
        has_soffice || PKGS+=(libreoffice-writer)
    elif command -v dnf >/dev/null; then
        PM=dnf
        has_tesseract || PKGS+=(tesseract)
        has_vie || PKGS+=(tesseract-langpack-vie)
        has_soffice || PKGS+=(libreoffice-writer)
    elif command -v pacman >/dev/null; then
        PM=pacman
        has_tesseract || PKGS+=(tesseract)
        has_vie || PKGS+=(tesseract-data-vie)
        has_soffice || PKGS+=(libreoffice-fresh)
    else
        PM=""
    fi
}

run_root() {
    if [ "$(id -u)" -eq 0 ]; then
        "$@"
    elif command -v sudo >/dev/null; then
        sudo "$@"
    else
        echo "Cần quyền root để cài gói hệ thống nhưng không có sudo." >&2
        return 1
    fi
}

install_packages() {
    case "$PM" in
        apt)
            run_root apt-get update
            run_root env DEBIAN_FRONTEND=noninteractive apt-get install -y "${PKGS[@]}"
            ;;
        dnf) run_root dnf install -y "${PKGS[@]}" ;;
        pacman) run_root pacman -S --needed --noconfirm "${PKGS[@]}" ;;
    esac
}

if [ "$WITH_DEPS" -eq 1 ]; then
    plan_packages
    if [ ${#PKGS[@]} -eq 0 ] && [ -z "$(missing_features)" ]; then
        echo "Gói hệ thống: đã đủ (tesseract + tiếng Việt, LibreOffice)."
    elif [ -z "$PM" ]; then
        echo "Không nhận ra trình quản lý gói (apt/dnf/pacman). Hãy tự cài:"
        missing_features | sed 's/^/  - /'
    else
        echo "Thiếu gói hệ thống, sẽ cài bằng ${PM}:"
        printf '  %s\n' "${PKGS[@]}"
        answer=y
        if [ "$ASSUME_YES" -eq 0 ]; then
            if [ -t 0 ]; then
                read -r -p "Cài ngay? [Y/n] " answer || answer=n
            else
                answer=n
                echo "(Không có terminal để hỏi - bỏ qua; chạy lại với --yes để cài.)"
            fi
        fi
        case "${answer:-y}" in
            [Yy]*|"")
                if ! install_packages; then
                    echo "Cài gói hệ thống thất bại - vẫn cài app, các tính năng trên sẽ chưa dùng được." >&2
                fi
                ;;
            *) echo "Bỏ qua cài gói hệ thống." ;;
        esac
    fi
fi

# ------------------------------------------------------------- the app

mkdir -p "$BIN_DIR" "$APP_DIR"
install -m 755 "$SRC" "${BIN_DIR}/word2md"

# The logo, from beside the binary (the release archive, or dist/).
ICON_LINE=""
ICON_SRC="$(dirname "$SRC")/word2md.png"
if [ -f "$ICON_SRC" ]; then
    mkdir -p "$ICON_DIR"
    install -m 644 "$ICON_SRC" "${ICON_DIR}/word2md.png"
    # An absolute path, not a theme name: a running GNOME Shell caches theme
    # lookups and would keep showing no icon until the next login.
    ICON_LINE="Icon=${ICON_DIR}/word2md.png"
fi

# StartupWMClass matches the window's WM_CLASS (src/gui.py WM_CLASS), so the
# dock/taskbar shows this entry's icon for the running window.
cat > "$DESKTOP" <<DESKTOP_EOF
[Desktop Entry]
Type=Application
Name=word2md
Comment=Chuyển Word/Excel/PDF sang Markdown và ngược lại
Exec=${BIN_DIR}/word2md
${ICON_LINE}
StartupWMClass=Word2md
Terminal=false
Categories=Office;
DESKTOP_EOF
command -v update-desktop-database >/dev/null && update-desktop-database "$APP_DIR" || true
echo
echo "Đã cài: ${BIN_DIR}/word2md  (menu ứng dụng: word2md)"

still="$(missing_features)"
if [ -n "$still" ]; then
    echo "Chưa có (app vẫn chạy, riêng các tính năng này chưa dùng được):"
    echo "$still" | sed 's/^/  - /'
fi

case ":${PATH}:" in
    *":${BIN_DIR}:"*) echo "Chạy: word2md" ;;
    *) echo "Chạy: ${BIN_DIR}/word2md  (${BIN_DIR} chưa có trong PATH - đăng nhập lại để gõ được 'word2md')" ;;
esac
