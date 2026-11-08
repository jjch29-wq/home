@echo off
setlocal
if exist "%~dp0PMI\home\src" (
    cd /d "%~dp0PMI"
) else if exist "C:\Users\jjch2\Desktop\PMI\home\src" (
    cd /d "C:\Users\jjch2\Desktop\PMI"
) else (
    cd /d "%~dp0"
)

if exist .venv\Scripts\python.exe (
    set "PYTHON_EXE=.venv\Scripts\python.exe"
) else (
    set "PYTHON_EXE=python"
)

echo [실행] 중앙지사 구글 클라우드 모바일 일보 동기화...
%PYTHON_EXE% home\src\site_apps\central\src\sync_google_sheet.py

echo.
echo 동기화 작업이 완료되었습니다.
pause
endlocal
