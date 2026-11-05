#!/usr/bin/env bash
# Build dist/word2md-linux-<arch> (single ELF binary) with PyInstaller.
# Extra args go to build.py, e.g. ./build.sh --onedir
set -euo pipefail
cd "$(dirname "$(readlink -f "$0")")"

PYTHON="${PYTHON:-python3}"
if [ ! -x .venv/bin/python ]; then
    echo "[1/2] Creating virtual environment..."
    "$PYTHON" -m venv .venv
    .venv/bin/python -m pip install --upgrade pip
    .venv/bin/python -m pip install -r requirements-dev.txt
fi

echo "[2/2] Building executable..."
exec .venv/bin/python build.py --clean "$@"
