@echo off
REM Build word2md-<arch>.exe using the project virtual environment.
REM The architecture follows the Python in .venv; see build.py for ARM64.
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo [1/2] Creating virtual environment...
    python -m venv .venv || exit /b 1
    .venv\Scripts\python.exe -m pip install --upgrade pip
    .venv\Scripts\python.exe -m pip install -r requirements-dev.txt || exit /b 1
)

echo [2/2] Building executable...
.venv\Scripts\python.exe build.py --clean %*
endlocal
