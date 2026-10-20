@echo off
cd /d "%~dp0"
set "PYTHONPATH=%~dp0home\src"
"%~dp0.venv\Scripts\python.exe" -m site_apps.central.src.attribute_export
if errorlevel 1 pause
