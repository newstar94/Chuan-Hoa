param(
    [switch]$OpenAdmin,
    [switch]$OpenApi
)

$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$start = Join-Path $root 'tools\development\Start-ChuanHoaDevelopment.ps1'
$healthUrl = 'http://127.0.0.1:5206/health'
$adminUrl = 'http://127.0.0.1:5206/development/admin'

if (-not (Get-NetTCPConnection -LocalPort 5206 -State Listen -ErrorAction SilentlyContinue)) {
    Write-Host 'Đang khởi động Chuẩn Hóa API tại http://127.0.0.1:5206 ...'
    # Invoke the startup script in this PowerShell process so build/startup
    # errors are visible immediately instead of leaving this wrapper polling
    # a child process that may already have exited.
    & $start -ApiOnly
    if ($LASTEXITCODE -ne 0) {
        throw "Không thể khởi động Chuẩn Hóa API (exit code $LASTEXITCODE)."
    }
}

$ready = $false
for ($attempt = 0; $attempt -lt 60; $attempt++) {
    Start-Sleep -Milliseconds 500
    try {
        $health = Invoke-RestMethod -Uri $healthUrl -TimeoutSec 2
        if ($health.status -eq 'ok') { $ready = $true; break }
    } catch { }
}

if (-not $ready) {
    throw 'Chuẩn Hóa chưa sẵn sàng. Kiểm tra cửa sổ API và file .dev-secrets\api.stderr.log.'
}

Write-Host 'Chuẩn Hóa đã sẵn sàng: http://127.0.0.1:5206'
if ($OpenAdmin) { Start-Process $adminUrl }
if ($OpenApi) { Start-Process 'http://127.0.0.1:5206/openapi/v1.json' }
