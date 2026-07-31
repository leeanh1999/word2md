@echo off
REM Launch the app from source using the project virtual environment.
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment...
    python -m venv .venv || exit /b 1
    .venv\Scripts\python.exe -m pip install --upgrade pip
    .venv\Scripts\python.exe -m pip install -r requirements.txt || exit /b 1
)

if "%~1"=="" (
    start "" ".venv\Scripts\pythonw.exe" main.py
) else (
    .venv\Scripts\python.exe main.py %*
)
endlocal
