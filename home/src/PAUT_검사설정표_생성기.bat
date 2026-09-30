@echo off
setlocal
cd /d "%~dp0"
set "UV_CACHE_DIR=%~dp0..\..\.uv-cache"
if exist "%~dp0..\..\.venv\Scripts\pythonw.exe" (
    start "PAUT 검사설정표 생성기" "%~dp0..\..\.venv\Scripts\pythonw.exe" "%~dp0PAUT_검사설정표_생성기.py"
    exit /b 0
)
where uv >nul 2>nul
if %errorlevel%==0 (
    start "PAUT 검사설정표 생성기" uv run --project "%~dp0..\.." pythonw "%~dp0PAUT_검사설정표_생성기.py"
) else (
    start "PAUT 검사설정표 생성기" python "%~dp0PAUT_검사설정표_생성기.py"
)
endlocal
