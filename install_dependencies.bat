@echo off
setlocal
echo --------------------------------------------------
echo [1/2] Checking Python Installation...
py --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] 'py' launcher not found. Trying 'python'...
    python --version >nul 2>&1
    if %errorlevel% neq 0 (
        echo [CRITICAL ERROR] Python is not installed or not in PATH.
        pause
        exit /b
    )
    set PY_CMD=python
) else (
    set PY_CMD=py
)

echo.
echo [2/2] Installing OCR Dependencies using %PY_CMD%...
%PY_CMD% -m pip install --upgrade pip
%PY_CMD% -m pip uninstall -y opencv-python opencv-python-headless
%PY_CMD% -m pip install easyocr pillow numpy pyperclip opencv-python google-genai pymupdf
echo.
echo --------------------------------------------------
echo Installation Final Check...
%PY_CMD% -c "import easyocr; import google.genai; import fitz; print('SUCCESS: All modules ready.')"
if %errorlevel% neq 0 (
    echo [WARNING] Some modules might not have installed correctly.
    echo Please check the error messages above.
) else (
    echo [CONGRATULATIONS] All dependencies are successfully installed!
)
echo --------------------------------------------------
echo.
echo Please RESTART the application.
pause
