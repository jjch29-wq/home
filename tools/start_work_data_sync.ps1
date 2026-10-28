$ErrorActionPreference = 'Stop'
$workspace = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $workspace
$logDir = Join-Path $workspace '.work-sync'
New-Item -ItemType Directory -Path $logDir -Force | Out-Null
$env:PYTHONIOENCODING = 'utf-8'
$syncTool = Join-Path $logDir 'tools\work_data_sync.py'
if (-not (Test-Path -LiteralPath $syncTool)) {
    $syncTool = Join-Path $PSScriptRoot 'work_data_sync.py'
}
& (Join-Path $workspace '.venv\Scripts\python.exe') -X utf8 -u $syncTool --workspace $workspace --watch >> (Join-Path $logDir 'autostart.log') 2>&1
exit $LASTEXITCODE
