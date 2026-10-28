@echo off
setlocal
set /p "PMI_WORKSPACE=PMI repository folder: "
if not exist "%PMI_WORKSPACE%\.git" (
  echo Repository not found.
  pause
  exit /b 1
)
if exist "%PMI_WORKSPACE%\.venv\Scripts\python.exe" (
  "%PMI_WORKSPACE%\.venv\Scripts\python.exe" "%~dp0work_data_sync.py" --workspace "%PMI_WORKSPACE%" --cloud "%~dp0.." %*
) else (
  py "%~dp0work_data_sync.py" --workspace "%PMI_WORKSPACE%" --cloud "%~dp0.." %*
)
pause
