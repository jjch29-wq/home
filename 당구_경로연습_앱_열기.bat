@echo off
setlocal
set "APP_FILE=%~dp0billiards-path-trainer\index.html"

if not exist "%APP_FILE%" (
  echo Carom Lab index file was not found.
  echo %APP_FILE%
  pause
  exit /b 1
)

start "" "%APP_FILE%"
if errorlevel 1 (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%APP_FILE%'"
)

endlocal
