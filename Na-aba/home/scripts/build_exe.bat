@echo off
chcp 65001 >nul
echo MaterialManager-1 EXE 파일 생성을 시작합니다...
cd /d "C:\Users\jjch2\Desktop\보고서\Project PROVIDENCE\Request\PMI"

REM PyInstaller 설치 확인
python -c "import PyInstaller; print('PyInstaller 설치됨')" 2>nul
if errorlevel 1 (
    echo PyInstaller를 설치합니다...
    python -m pip install pyinstaller
)

REM EXE 파일 생성
echo 빌드를 시작합니다...
python -m PyInstaller --onefile --windowed --name=MaterialManager-1 --clean MaterialManager-1.py

if exist "dist\MaterialManager-1.exe" (
    echo.
    echo === 빌드 성공 ===
    echo EXE 파일: dist\MaterialManager-1.exe
    echo 바탕화면에 복사합니다...
    copy "dist\MaterialManager-1.exe" "MaterialManager-1.exe"
    echo 바탕화면 바로가기 생성 완료: MaterialManager-1.exe
) else (
    echo.
    echo === 빌드 실패 ===
)

pause
