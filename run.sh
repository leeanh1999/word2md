#!/usr/bin/env bash
# Launch the app from source using the project virtual environment (Linux).
#   ./run.sh                     -> GUI
#   ./run.sh report.docx -o out  -> CLI
set -euo pipefail
cd "$(dirname "$(readlink -f "$0")")"

PYTHON="${PYTHON:-python3}"
if [ ! -x .venv/bin/python ]; then
    echo "Creating virtual environment..."
    "$PYTHON" -m venv .venv
    .venv/bin/python -m pip install --upgrade pip
    .venv/bin/python -m pip install -r requirements.txt
fi

exec .venv/bin/python main.py "$@"
