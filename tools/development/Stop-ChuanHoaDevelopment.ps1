$ErrorActionPreference = 'Stop'
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$pidPath = Join-Path $projectRoot '.dev-secrets\api.pid'
if (!(Test-Path -LiteralPath $pidPath)) {
    Write-Output 'Không có API Development do script quản lý đang được ghi nhận.'
    exit 0
}

$processId = 0
if (![int]::TryParse((Get-Content -LiteralPath $pidPath -Raw).Trim(), [ref]$processId)) {
    throw 'File PID Development không hợp lệ.'
}
$process = Get-Process -Id $processId -ErrorAction SilentlyContinue
if ($null -ne $process) {
    $path = $process.Path
    $expected = @(
        (Join-Path $projectRoot 'src\ChuanHoa.Api\bin\Release\net10.0\ChuanHoa.Api.exe'),
        'D:\ChuanHoaDevelopment\src\ChuanHoa.Api\bin\Release\net10.0\ChuanHoa.Api.exe'
    ) | ForEach-Object { [IO.Path]::GetFullPath($_) }
    if ($expected -notcontains [IO.Path]::GetFullPath($path)) {
        throw 'PID hiện thuộc tiến trình khác; không dừng tự động.'
    }
    Stop-Process -Id $processId
}
Remove-Item -LiteralPath $pidPath -Force
Write-Output 'Đã dừng API Development của Chuẩn hóa.'
