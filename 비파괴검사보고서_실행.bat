@echo off
setlocal
cd /d "%~dp0"

set PYTHONW=%~dp0.venv\Scripts\pythonw.exe
set SCRIPT=%~dp0home\src\run_ndt.py

if not exist "%PYTHONW%" (
    echo [ERROR] Python not found: %PYTHONW%
    pause
    exit /b 1
)

start "" "%PYTHONW%" "%SCRIPT%"
endlocal

