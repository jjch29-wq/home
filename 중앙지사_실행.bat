@echo off
chcp 65001 > nul
setlocal
cd /d "%~dp0"
echo ===================================================
echo   SITCO 한국지역난방공사 중앙지사 자재관리 시스템
echo ===================================================

if exist .venv\Scripts\python.exe (
    set "PYTHON_EXE=.venv\Scripts\python.exe"
) else (
    set "PYTHON_EXE=python"
)

echo [실행] 중앙지사 시스템을 시작합니다...
%PYTHON_EXE% home\src\site_apps\central\main.py

if %ERRORLEVEL% neq 0 (
    echo.
    echo [오류] 프로그램 실행 중 오류가 발생했습니다. (코드: %ERRORLEVEL%)
    pause
)
endlocal

