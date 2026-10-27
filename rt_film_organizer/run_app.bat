@echo off
setlocal
cd /d "%~dp0"
set PYTHONUTF8=1
"..\.venv\Scripts\python.exe" desktop_app.py
if errorlevel 1 pause

