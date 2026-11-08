# ================================================================
#  경인사무소 전용 스마트 검증 및 배포 스크립트 (배포_경인.ps1)
#  - 브랜치: 경인사무소-개발
#  - 사업장 격리(Central 등 타 사업장 코드 불변) 자동 검증
#  - 모바일 기능 및 구글 클라우드 동기화 테스트 실행
#  - 스마트 Git 커밋 및 GitHub 원격 푸시
# ================================================================
param(
    [string]$Message = "",
    [switch]$NoTest,
    [switch]$NoPush
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ProjectRoot = Split-Path -Parent $ScriptDir
Set-Location $ProjectRoot
$env:GIT_REDIRECT_STDERR = '2>&1'
$ErrorActionPreference = 'Continue'

$Today = (Get-Date).ToString('yyyy-MM-dd HH:mm')
$PythonExe = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $PythonExe)) {
    $PythonExe = "python"
}

Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "   🚀 경인사무소 자재관리 및 모바일 시스템 배포 프로세스" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan

# ── 1. 브랜치 확인 ──────────────────────────────────────────────
Write-Host "`n[1/4] Git 활성 브랜치 검증..." -ForegroundColor Cyan
$currentBranch = (git rev-parse --abbrev-ref HEAD).Trim()
Write-Host "  · 현재 브랜치: $currentBranch" -ForegroundColor Gray

if ($currentBranch -like "*중앙*" -or $currentBranch -eq "master" -or $currentBranch -eq "main") {
    Write-Host "[경고] 타 사업장 브랜치에서는 경인사무소 배포를 실행할 수 없습니다!" -ForegroundColor Red
    exit 1
}
Write-Host "  => 브랜치 검증 통과 ($currentBranch)" -ForegroundColor Green

# ── 2. 자동화 테스트 및 사업장 격리 검증 ────────────────────────
if (-not $NoTest) {
    Write-Host "`n[2/4] 경인사무소 격리 및 무결성 자동 테스트 실행..." -ForegroundColor Cyan

    Write-Host "  (1) 사업장 격리 테스트 (Central 보호 검증)..." -ForegroundColor Gray
    & $PythonExe -m unittest home/tests/test_site_isolation.py
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[오류] 사업장 격리 테스트 실패! 중앙지사 코드가 변경되었을 수 있습니다." -ForegroundColor Red
        exit 1
    }

    Write-Host "  (2) 모바일 계산/단가/비배관 기능 테스트..." -ForegroundColor Gray
    & $PythonExe home/tests/test_gyeongin_mobile.py
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[오류] 경인 모바일 기능 테스트 실패!" -ForegroundColor Red
        exit 1
    }

    Write-Host "  (3) 구글 클라우드 동기화 테스트..." -ForegroundColor Gray
    & $PythonExe -m unittest home/tests/test_gyeongin_google_sync.py
    if ($LASTEXITCODE -ne 0) {
        Write-Host "[오류] 구글 클라우드 동기화 테스트 실패!" -ForegroundColor Red
        exit 1
    }

    Write-Host "  => 모든 경인사무소 검증 테스트 100% 통과!" -ForegroundColor Green
} else {
    Write-Host "`n[2/4] 테스트 실행 생략 (-NoTest)" -ForegroundColor DarkGray
}

# ── 3. 경인사무소 변경사항 스마트 스테이징 & 커밋 ───────────────
Write-Host "`n[3/4] 경인사무소 변경사항 스테이징 및 커밋..." -ForegroundColor Cyan

git add home/src/site_apps/gyeongin/
git add home/src/경인사무소.py
git add home/tests/test_gyeongin*
git add home/배포_경인.ps1
git add 경인사무소_*.bat 2>$null

$staged = git diff --cached --name-only
if (-not $staged) {
    Write-Host "  [알림] 스테이징된 변경사항 없음 (이미 최신 커밋 상태)" -ForegroundColor Yellow
} else {
    Write-Host "  스테이징된 파일 목록:" -ForegroundColor Gray
    $staged | ForEach-Object { Write-Host "    · $_" -ForegroundColor Gray }

    $commitMsg = if ($Message) { $Message } else { "deploy(gyeongin): $Today 경인사무소 시스템 배포" }
    git commit -m $commitMsg
    Write-Host "  => 커밋 완료: $commitMsg" -ForegroundColor Green
}

# ── 4. 원격 저장소 푸시 ─────────────────────────────────────────
if (-not $NoPush) {
    Write-Host "`n[4/4] GitHub 원격 저장소 푸시 (origin 경인사무소-개발)..." -ForegroundColor Cyan
    $pushResult = git push origin 경인사무소-개발 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Host "  => Push 성공!" -ForegroundColor Green
    } else {
        Write-Host "  [경고] Push 결과:`n$pushResult" -ForegroundColor Yellow
    }
} else {
    Write-Host "`n[4/4] 푸시 생략 (-NoPush)" -ForegroundColor DarkGray
}

Write-Host "`n================================================================" -ForegroundColor Green
Write-Host "   🎉 경인사무소 시스템 배포가 성공적으로 완료되었습니다!" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Green
