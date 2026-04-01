@echo off
setlocal
cd /d "%~dp0"
echo Starting PMI/PAUT Report Manager...
python src/main.py
if %ERRORLEVEL% neq 0 (
    echo.
    echo Application failed with error code %ERRORLEVEL%
    pause
)
endlocal
