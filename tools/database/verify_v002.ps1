[CmdletBinding()]
param([int]$Port = 55447)
$ErrorActionPreference = 'Stop'
$projectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path
$executionRoot = $projectRoot; $substDrive = $null
if ($projectRoot -match '[^\u0000-\u007F]') {
    foreach ($candidate in @('T:', 'U:', 'V:', 'W:', 'X:', 'Y:', 'Z:')) {
        if (-not (Test-Path -LiteralPath "$candidate\")) { & subst.exe $candidate $projectRoot; if ($LASTEXITCODE -eq 0) { $substDrive = $candidate; $executionRoot = "$candidate\"; break } }
    }
    if ($null -eq $substDrive) { throw 'No ASCII drive available.' }
}
$bin = Join-Path $executionRoot '.tools\postgresql\pgsql\bin'; $data = Join-Path $executionRoot '.tools\postgresql-v002-verify-data'; $log = Join-Path $executionRoot '.tools\postgresql-v002-verify.log'; $db = 'chuanhoa_v002_verify'; $started = $false; $created = $false
function Invoke-Checked([string]$exe, [string[]]$arguments, [string]$label) { & $exe @arguments; if ($LASTEXITCODE -ne 0) { throw "$label failed: $LASTEXITCODE" } }
try {
    if (Test-Path (Join-Path $data 'PG_VERSION')) { Remove-Item -LiteralPath $data -Recurse -Force }
    Invoke-Checked (Join-Path $bin 'initdb.exe') @('-D',$data,'-A','trust','-U','postgres','--encoding=UTF8','--no-locale') 'initdb'
    Invoke-Checked (Join-Path $bin 'pg_ctl.exe') @('-D',$data,'-l',$log,'-o',"-p $Port -h 127.0.0.1",'-w','start') 'start postgres'; $started = $true
    Invoke-Checked (Join-Path $bin 'createdb.exe') @('-h','127.0.0.1','-p',$Port,'-U','postgres',$db) 'createdb'; $created = $true
    $base = @('-X','-v','ON_ERROR_STOP=1','-h','127.0.0.1','-p',$Port,'-U','postgres','-d',$db)
    Invoke-Checked (Join-Path $bin 'psql.exe') ($base + @('-f',(Join-Path $executionRoot 'database\migrations\V001__identity_trial_commercial_foundation.sql'))) 'V001'
    Invoke-Checked (Join-Path $bin 'psql.exe') ($base + @('-f',(Join-Path $executionRoot 'database\migrations\V002__admin_integration_commands.sql'))) 'V002 up'
    Invoke-Checked (Join-Path $bin 'psql.exe') ($base + @('-f',(Join-Path $executionRoot 'tools\database\verify_v002_assertions.sql'))) 'V002 assertions'
    Invoke-Checked (Join-Path $bin 'psql.exe') ($base + @('-f',(Join-Path $executionRoot 'database\migrations\V002_down.sql'))) 'V002 down'
    Invoke-Checked (Join-Path $bin 'psql.exe') ($base + @('-f',(Join-Path $executionRoot 'tools\database\verify_v002_down_assertions.sql'))) 'V002 down assertions'
    'DB-MIGRATION-V002-001=PASS'
} finally {
    if ($created) { & (Join-Path $bin 'dropdb.exe') '--if-exists' '-h' '127.0.0.1' '-p' $Port '-U' 'postgres' $db | Out-Null }
    if ($started) { & (Join-Path $bin 'pg_ctl.exe') '-D' $data '-m' 'fast' '-w' 'stop' | Out-Null }
    if (Test-Path $data) { Remove-Item -LiteralPath $data -Recurse -Force -ErrorAction SilentlyContinue }
    if (Test-Path $log) { Remove-Item -LiteralPath $log -Force -ErrorAction SilentlyContinue }
    if ($null -ne $substDrive) { & subst.exe $substDrive '/d' }
}
