@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0..\..\..\..\.."
set "PYTHONPATH=%CD%\home\src"
set "PMI_CENTRAL_MODE=dev"
".venv\Scripts\python.exe" "home\src\site_apps\central\src\app.py" --pick-profile
endlocal

