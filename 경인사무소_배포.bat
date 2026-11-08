@echo off
chcp 65001 > nul
setlocal
cd /d "%~dp0"
echo ===================================================
echo   🚀 경인사무소 시스템 자동 검증 및 배포
echo ===================================================

powershell.exe -NoProfile -ExecutionPolicy Bypass -File home\배포_경인.ps1 %*

pause
endlocal

