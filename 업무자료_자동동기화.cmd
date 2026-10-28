@echo off
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
  ".venv\Scripts\python.exe" tools\work_data_sync.py --watch
) else (
  py tools\work_data_sync.py --watch
)
pause
