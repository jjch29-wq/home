$ErrorActionPreference = 'Stop'

$workspace = Split-Path -Parent $PSScriptRoot
$pythonCandidates = @(
    (Join-Path $workspace '.venv-paut\Scripts\python.exe'),
    (Join-Path $workspace '.venv\Scripts\python.exe')
)
$python = $pythonCandidates | Where-Object {
    if (-not (Test-Path -LiteralPath $_)) { return $false }
    & $_ --version *> $null
    return $LASTEXITCODE -eq 0
} | Select-Object -First 1
$syncScript = Join-Path $PSScriptRoot 'work_data_sync.py'
$logDirectory = Join-Path $workspace '.work-sync'
$logFile = Join-Path $logDirectory 'autostart.log'

New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null

if (-not $python) {
    Add-Content -LiteralPath $logFile -Encoding UTF8 -Value "$(Get-Date -Format s) 동기화 실패: 사용 가능한 Python 실행 파일 없음"
    exit 1
}

if (-not (Test-Path -LiteralPath $syncScript)) {
    Add-Content -LiteralPath $logFile -Encoding UTF8 -Value "$(Get-Date -Format s) 동기화 실패: 동기화 스크립트 없음: $syncScript"
    exit 1
}

Set-Location -LiteralPath $workspace
& $python $syncScript --watch *>> $logFile
