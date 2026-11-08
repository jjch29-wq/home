@echo off
setlocal
cd /d "C:\Users\jjch2\Desktop\PMI"
if not exist "home\src" cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
  set "PYTHON_EXE=.venv\Scripts\python.exe"
) else (
  set "PYTHON_EXE=python"
)

"%PYTHON_EXE%" "home\src\site_apps\central\main.py"
endlocal
