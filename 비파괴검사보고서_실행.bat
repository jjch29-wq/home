@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\pythonw.exe" (
    echo [ERROR] Python virtual environment was not found.
    echo Open this project in VS Code and check the Python environment.
    pause
    exit /b 1
)

start "" ".venv\Scripts\pythonw.exe" "home\src\비파괴검사보고서.py"
endlocal
