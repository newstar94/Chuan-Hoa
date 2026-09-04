$ErrorActionPreference = 'Stop'
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$pidPath = Join-Path $projectRoot '.dev-secrets\api.pid'
if (!(Test-Path -LiteralPath $pidPath)) {
    Write-Output 'No managed Development API process is recorded.'
    exit 0
}

$processId = 0
if (![int]::TryParse((Get-Content -LiteralPath $pidPath -Raw).Trim(), [ref]$processId)) {
    throw 'The Development PID file is invalid.'
}
$process = Get-Process -Id $processId -ErrorAction SilentlyContinue
if ($null -ne $process) {
    $path = $process.Path
    $expected = @(
        (Join-Path $projectRoot 'src\ChuanHoa.Api\bin\Release\net10.0\ChuanHoa.Api.exe'),
        'D:\ChuanHoaDevelopment\src\ChuanHoa.Api\bin\Release\net10.0\ChuanHoa.Api.exe'
    ) | ForEach-Object { [IO.Path]::GetFullPath($_) }
    if ($expected -notcontains [IO.Path]::GetFullPath($path)) {
        throw 'The recorded PID belongs to another process; refusing to stop it.'
    }
    Stop-Process -Id $processId
}
Remove-Item -LiteralPath $pidPath -Force
Write-Output 'The Chuan Hoa Development API has stopped.'
