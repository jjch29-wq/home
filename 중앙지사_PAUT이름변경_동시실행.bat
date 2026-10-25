@echo off
setlocal
cd /d "%~dp0"

set "APP_PYTHON=%~dp0.venv\Scripts\pythonw.exe"
if not exist "%APP_PYTHON%" (
    echo [ERROR] Python virtual environment was not found:
    echo %APP_PYTHON%
    pause
    exit /b 1
)

start "Central App" /D "%~dp0home\src" "%APP_PYTHON%" "%~dp0home\src\site_apps\central\main.py"
start "PAUT Rename" cmd /c call "%~dp0PAUT_이름변경_독립실행.bat"

endlocal
