@echo off
set "PYTHONHOME=%LocalAppData%\Programs\Python\Python313"
"%PYTHONHOME%\python.exe" "%~dp0main.py"
if errorlevel 1 pause
