@echo off
setlocal
cd /d "%~dp0"
echo Starting Material Master Manager V13...

where uv >nul 2>&1
if %errorlevel% equ 0 (
    echo Using 'uv' to run the application...
    uv run src\Material-Master-Manager-V13.py
) else if exist .venv\Scripts\python.exe (
    echo Using local virtual environment...
    .venv\Scripts\python.exe src\Material-Master-Manager-V13.py
) else (
    echo Using system Python...
    python src\Material-Master-Manager-V13.py
)

if %ERRORLEVEL% neq 0 (
    echo.
    echo Application failed with error code %ERRORLEVEL%
    pause
)
endlocal
