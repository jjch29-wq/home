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

echo [실행] 중앙지사 현장 모바일 입력 서버 (포트 8000)...
%PYTHON_EXE% home\src\site_apps\central\run_mobile.py

pause
endlocal
