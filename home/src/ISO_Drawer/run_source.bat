@echo off
setlocal
if not exist "%~dp0.venv\Scripts\python.exe" (
  echo The project virtual environment is missing.
  echo Run: py -m pip install -r requirements.txt
  pause
  exit /b 1
)
start "ISO Drawer Console" "%~dp0.venv\Scripts\python.exe" "%~dp0desktop_app.py"
