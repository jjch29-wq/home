@echo off
setlocal
cd /d "%~dp0"
set "APP_PYTHON=%~dp0.venv\Scripts\python.exe"
call :check_tk
if not errorlevel 1 goto run
set "APP_PYTHON=%LocalAppData%\Programs\Python\Python313\python.exe"
call :check_tk
if not errorlevel 1 goto run
echo [ERROR] No working Python with Tcl/Tk was found.
echo Install Python with the Tcl/Tk option enabled.
pause
exit /b 1

:run
cd /d "%~dp0home\src"
"%APP_PYTHON%" "%~dp0home\src\paut_rename.py"
if errorlevel 1 (
    echo [ERROR] PAUT Rename failed. See the error above.
    pause
    exit /b 1
)
endlocal
exit /b 0

:check_tk
if not exist "%APP_PYTHON%" exit /b 1
"%APP_PYTHON%" -c "import tkinter as tk; r=tk.Tk(); r.withdraw(); r.destroy()" >nul 2>&1
exit /b %errorlevel%
