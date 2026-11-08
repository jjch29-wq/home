@echo off
setlocal
if exist "%~dp0PMI\home\src" (
    cd /d "%~dp0PMI"
) else if exist "C:\Users\jjch2\Desktop\PMI\home\src" (
    cd /d "C:\Users\jjch2\Desktop\PMI"
) else (
    cd /d "%~dp0"
)

echo [실행] 중앙지사 시스템 자동 검증 및 배포...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File home\배포_중앙.ps1 %*

pause
endlocal
