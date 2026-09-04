@echo off
setlocal
if exist "%~dp0dist\ISO_Drawer\ISO_Drawer.exe" (
  start "" "%~dp0dist\ISO_Drawer\ISO_Drawer.exe"
  exit /b
)
if exist "%~dp0.venv\Scripts\pythonw.exe" (
  start "" "%~dp0.venv\Scripts\pythonw.exe" "%~dp0desktop_app.py"
  exit /b
)
echo ISO Drawer executable was not found.
pause
