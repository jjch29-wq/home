# ================================================================
#  SITCO Material Manager — 빌드 + 스마트 Git 커밋 스크립트
#  사용법: powershell -ExecutionPolicy Bypass -File 배포.ps1
#  옵션:  -NoExe   : exe 빌드 없이 소스만 커밋
#         -NoPush  : commit만 하고 push 안 함
# ================================================================
param(
    [switch]$NoExe,
    [switch]$NoPush
)

Set-Location $PSScriptRoot
$ErrorActionPreference = 'Stop'

$Today    = (Get-Date).ToString('yyyyMMdd')
$ExeName  = "SITCO-Material-Manager-$Today.exe"
$DistPath = Join-Path $PSScriptRoot "dist"
$ExePath  = Join-Path $DistPath $ExeName

# ── 1. exe 빌드 ────────────────────────────────────────────────
if (-not $NoExe) {
    Write-Host "`n[1/3] exe 빌드 시작 → $ExeName" -ForegroundColor Cyan

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName         = ".venv\Scripts\pyinstaller.exe"
    $psi.Arguments        = "material_manager.spec --distpath dist --workpath build\MaterialManager_V13 --noconfirm --log-level WARN"
    $psi.WorkingDirectory = $PSScriptRoot
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.UseShellExecute  = $false

    $p = [System.Diagnostics.Process]::Start($psi)
    $p.WaitForExit(300000) | Out-Null
    $stderr = $p.StandardError.ReadToEnd()

    if ($p.ExitCode -ne 0) {
        Write-Host "[오류] 빌드 실패:" -ForegroundColor Red
        Write-Host $stderr
        exit 1
    }

    # spec에서 만들어진 exe 확인 (날짜 이름)
    if (Test-Path $ExePath) {
        $mb = [math]::Round((Get-Item $ExePath).Length / 1MB, 1)
        Write-Host "[완료] $ExeName ($mb MB)" -ForegroundColor Green
    } else {
        Write-Host "[경고] exe 파일을 찾지 못했습니다: $ExePath" -ForegroundColor Yellow
    }
} else {
    Write-Host "`n[1/3] exe 빌드 생략 (-NoExe)" -ForegroundColor DarkGray
}

# ── 2. 변경된 파일만 스마트 커밋 ──────────────────────────────
Write-Host "`n[2/3] Git 변경 사항 확인..." -ForegroundColor Cyan

# 소스·spec 스테이징
git add "src/Material-Master-Manager-V13.py" "material_manager.spec" 2>$null

# 오늘 날짜 exe 있으면 추가 (없으면 생략)
if (Test-Path $ExePath) {
    $relExe = "dist/$ExeName"
    git add -f $relExe 2>$null
}

# 변경된 것이 있는지 확인
$staged = git diff --cached --name-only
if (-not $staged) {
    Write-Host "[알림] 변경사항 없음 — 커밋 생략" -ForegroundColor Yellow
} else {
    Write-Host "스테이징된 파일:" -ForegroundColor Gray
    $staged | ForEach-Object { Write-Host "  · $_" -ForegroundColor Gray }

    $msg = "deploy: $Today 빌드 및 소스 업데이트"
    git commit -m $msg
    Write-Host "[완료] 커밋: $msg" -ForegroundColor Green
}

# ── 3. Push ────────────────────────────────────────────────────
if (-not $NoPush) {
    Write-Host "`n[3/3] Push..." -ForegroundColor Cyan

    # 커밋할 게 없어도 push는 해도 무방 (already up to date)
    $pushResult = git push 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "[완료] Push 성공" -ForegroundColor Green
    } else {
        Write-Host "[경고] Push 결과:`n$pushResult" -ForegroundColor Yellow
    }
} else {
    Write-Host "`n[3/3] Push 생략 (-NoPush)" -ForegroundColor DarkGray
}

Write-Host "`n✅ 배포 완료`n" -ForegroundColor Cyan
