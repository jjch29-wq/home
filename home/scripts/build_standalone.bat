@echo off
echo ========================================================
echo   PMI Report Generator (V2 Unified) - Standalone Builder
echo ========================================================

REM Check if PyInstaller is installed
python -m PyInstaller --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [INFO] PyInstaller is not installed. Installing...
    pip install pyinstaller pandas openpyxl pillow
)

REM Define source files
set SCRIPT_NAME=JJCHSITPMI-V2-Unified.py
set EXE_NAME=PMI_Report_V2_Unified

echo.
echo [1/2] Packaging the application...
echo This will include the script and any image files in this folder.

REM Run PyInstaller
REM --onefile: Creates a single executable file.
REM --windowed: Hides the console window (GUI only).
REM --add-data "*.png;." --add-data "*.jpg;." --add-data "*.jp*g;.": Bundles image files into the executable for standalone use.
REM --icon=app.ico: (Optional) Adds an icon if present. You can remove or change this later.

python -m PyInstaller --noconfirm --log-level=WARN ^
    --onefile --windowed ^
    --name "%EXE_NAME%" ^
    --add-data "*.png;." ^
    --add-data "logo_settings_unified.json;." ^
    "%SCRIPT_NAME%"

echo.
echo [2/2] Build process finished.
if exist "dist\%EXE_NAME%.exe" (
    echo.
    echo ✅ SUCCESS: The executable is ready!
    echo 📂 You can find it in the "dist" folder.
    echo 💡 You can move "dist\%EXE_NAME%.exe" ANYWHERE and it will run independently.
) else (
    echo ❌ ERROR: The build seems to have failed. Check the logs above.
)

echo.
pause
