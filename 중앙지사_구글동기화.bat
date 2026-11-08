@echo off
chcp 65001 > nul
setlocal
cd /d "%~dp0"
echo ===================================================
echo   🔄 중앙지사 구글 클라우드 모바일 일보 동기화
echo ===================================================

if exist .venv\Scripts\python.exe (
    set "PYTHON_EXE=.venv\Scripts\python.exe"
) else (
    set "PYTHON_EXE=python"
)

%PYTHON_EXE% home\src\site_apps\central\src\sync_google_sheet.py

echo.
echo 동기화 작업이 완료되었습니다.
pause
endlocal

