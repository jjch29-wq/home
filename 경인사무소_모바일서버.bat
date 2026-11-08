@echo off
chcp 65001 > nul
setlocal
cd /d "%~dp0"
echo ===================================================
echo   🚀 경인사무소 현장 모바일 검사 입력 서버
echo   포트: 8001 / Cloudflare 전용 외부 터널 자동 생성
echo ===================================================

if exist .venv\Scripts\python.exe (
    set "PYTHON_EXE=.venv\Scripts\python.exe"
) else (
    set "PYTHON_EXE=python"
)

%PYTHON_EXE% home\src\site_apps\gyeongin\run_mobile.py

pause
endlocal

