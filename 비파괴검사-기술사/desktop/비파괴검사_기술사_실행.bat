@echo off
cd /d "%~dp0\..\.."
if exist ".venv\Scripts\pythonw.exe" (
  start "NDT Study" ".venv\Scripts\pythonw.exe" "비파괴검사-기술사\desktop\app.py"
  exit /b 0
)

where pythonw.exe >nul 2>nul
if not errorlevel 1 (
  start "NDT Study" pythonw.exe "비파괴검사-기술사\desktop\app.py"
  exit /b 0
)

where python.exe >nul 2>nul
if not errorlevel 1 (
  python.exe "비파괴검사-기술사\desktop\app.py"
  exit /b %errorlevel%
)

echo Python을 찾을 수 없습니다.
pause
exit /b 1
