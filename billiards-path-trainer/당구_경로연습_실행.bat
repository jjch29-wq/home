@echo off
setlocal
set "APP_FILE=%~dp0index.html"

if not exist "%APP_FILE%" (
  echo index.html file was not found.
  echo %APP_FILE%
  pause
  exit /b 1
)

start "" "%APP_FILE%"
if errorlevel 1 (
  powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%APP_FILE%'"
)

endlocal
