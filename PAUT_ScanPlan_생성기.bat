@echo off
setlocal
cd /d "%~dp0"
if exist "%LOCALAPPDATA%\Programs\Python\Python314\pythonw.exe" (
  start "" "%LOCALAPPDATA%\Programs\Python\Python314\pythonw.exe" "home\src\paut_scanplan_generator.py"
  exit /b
)
if exist "%USERPROFILE%\.local\bin\pythonw.exe" (
  start "" "%USERPROFILE%\.local\bin\pythonw.exe" "home\src\paut_scanplan_generator.py"
  exit /b
)
pythonw "home\src\paut_scanplan_generator.py"
if errorlevel 1 python "home\src\paut_scanplan_generator.py"
