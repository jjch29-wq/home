@echo off
setlocal
cd /d "C:\Users\jjch2\Desktop\PMI"
if not exist "home\src" cd /d "%~dp0"

REM [0] 바탕화면에 바로가기가 없으면 자동 생성
if not exist "%USERPROFILE%\Desktop\중앙지사_자재관리.bat" (
    copy /y "%~dp0중앙지사_실행.bat" "%USERPROFILE%\Desktop\중앙지사_자재관리.bat" >nul 2>&1
)

if exist ".venv\Scripts\python.exe" (
  set "PYTHON_EXE=.venv\Scripts\python.exe"
) else (
  set "PYTHON_EXE=python"
)

REM [1] 모바일 서버 자동 실행 (포트 8000 미실행 시 백그라운드 숨김 창으로 자동 시작)
powershell.exe -NoProfile -Command "if (-not (Get-NetTCPConnection -LocalPort 8000 -ErrorAction SilentlyContinue)) { Start-Process '%PYTHON_EXE%' -ArgumentList 'home\src\site_apps\central\run_mobile.py' -WindowStyle Hidden }"

REM [2] 중앙지사 자재관리 메인 시스템 실행 (시작 시 구글/모바일 자동 동기화)
"%PYTHON_EXE%" "home\src\site_apps\central\main.py"
endlocal
