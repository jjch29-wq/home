@echo off
setlocal
cd /d "%~dp0"
set "APP_PYTHON=%~dp0.venv\Scripts\pythonw.exe"
if not exist "%APP_PYTHON%" exit /b 1
start "Central App" /D "%~dp0home\src" "%APP_PYTHON%" "%~dp0home\src\site_apps\central\main.py"
endlocal
