[CmdletBinding()]
param(
    [int]$Port = 55439
)

$ErrorActionPreference = 'Stop'
$projectRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path
$executionRoot = $projectRoot
$substDrive = $null

if ($projectRoot -match '[^\u0000-\u007F]') {
    foreach ($candidate in @('T:', 'U:', 'V:', 'W:', 'X:', 'Y:', 'Z:')) {
        if (-not (Test-Path -LiteralPath "$candidate\")) {
            & subst.exe $candidate $projectRoot
            if ($LASTEXITCODE -eq 0) {
                $substDrive = $candidate
                $executionRoot = "$candidate\"
                break
            }
        }
    }

    if ($null -eq $substDrive) {
        throw 'PostgreSQL portable requires an ASCII path, and no drive letter was available for a temporary project mapping.'
    }
}

$postgresRoot = Join-Path $executionRoot '.tools\postgresql\pgsql'
$postgresBin = Join-Path $postgresRoot 'bin'
$dataDirectory = Join-Path $executionRoot '.tools\postgresql-v001-data'
$logPath = Join-Path $executionRoot '.tools\postgresql-v001.log'
$databaseName = 'chuanhoa_migration_v001_test'
$migrationPath = Join-Path $executionRoot 'database\migrations\V001__identity_trial_commercial_foundation.sql'
$downPath = Join-Path $executionRoot 'database\migrations\V001_down.sql'
$assertionPath = Join-Path $executionRoot 'tools\database\verify_v001_assertions.sql'
$downAssertionPath = Join-Path $executionRoot 'tools\database\verify_v001_down_assertions.sql'
$evidencePath = Join-Path $executionRoot 'shared\docs\implementation\evidence\migration_v001.json'

$initdb = Join-Path $postgresBin 'initdb.exe'
$pgCtl = Join-Path $postgresBin 'pg_ctl.exe'
$createdb = Join-Path $postgresBin 'createdb.exe'
$dropdb = Join-Path $postgresBin 'dropdb.exe'
$psql = Join-Path $postgresBin 'psql.exe'

$serverStarted = $false
$databaseCreated = $false
$startedAt = [DateTimeOffset]::UtcNow
$commands = [System.Collections.Generic.List[string]]::new()

function Invoke-PostgresCommand {
    param(
        [Parameter(Mandatory)]
        [string]$Executable,
        [Parameter(Mandatory)]
        [string[]]$Arguments,
        [Parameter(Mandatory)]
        [string]$EvidenceLabel
    )

    $script:commands.Add($EvidenceLabel)
    & $Executable @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$EvidenceLabel failed with exit code $LASTEXITCODE."
    }
}

try {
    foreach ($requiredPath in @($initdb, $pgCtl, $createdb, $dropdb, $psql, $migrationPath, $downPath, $assertionPath, $downAssertionPath)) {
        if (-not (Test-Path -LiteralPath $requiredPath)) {
            throw "Required file not found: $requiredPath"
        }
    }

    $portInUse = Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue
    if ($null -ne $portInUse) {
        throw "TCP port $Port is already in use. Choose another port."
    }

    if (-not (Test-Path -LiteralPath (Join-Path $dataDirectory 'PG_VERSION'))) {
        if (Test-Path -LiteralPath $dataDirectory) {
            $resolvedDataPath = [System.IO.Path]::GetFullPath($dataDirectory)
            $resolvedToolsPath = [System.IO.Path]::GetFullPath((Join-Path $executionRoot '.tools'))
            if (-not $resolvedDataPath.StartsWith($resolvedToolsPath, [System.StringComparison]::OrdinalIgnoreCase)) {
                throw "Refusing to remove data directory outside the project tool directory: $resolvedDataPath"
            }

            Remove-Item -LiteralPath $resolvedDataPath -Recurse -Force
        }

        Invoke-PostgresCommand -Executable $initdb -Arguments @(
            '-D', $dataDirectory,
            '-A', 'trust',
            '-U', 'postgres',
            '--encoding=UTF8',
            '--no-locale'
        ) -EvidenceLabel 'initdb isolated UTF-8 cluster'
    }

    if (Test-Path -LiteralPath $logPath) {
        Remove-Item -LiteralPath $logPath -Force
    }

    Invoke-PostgresCommand -Executable $pgCtl -Arguments @(
        '-D', $dataDirectory,
        '-l', $logPath,
        '-o', "-p $Port -h 127.0.0.1",
        '-w',
        'start'
    ) -EvidenceLabel 'start isolated PostgreSQL server'
    $serverStarted = $true

    & $dropdb '--if-exists' '-h' '127.0.0.1' '-p' $Port '-U' 'postgres' $databaseName
    if ($LASTEXITCODE -ne 0) {
        throw "drop stale test database failed with exit code $LASTEXITCODE."
    }

    Invoke-PostgresCommand -Executable $createdb -Arguments @(
        '-h', '127.0.0.1',
        '-p', $Port,
        '-U', 'postgres',
        $databaseName
    ) -EvidenceLabel 'create migration test database'
    $databaseCreated = $true

    Invoke-PostgresCommand -Executable $psql -Arguments @(
        '-X',
        '-v', 'ON_ERROR_STOP=1',
        '-h', '127.0.0.1',
        '-p', $Port,
        '-U', 'postgres',
        '-d', $databaseName,
        '-f', $migrationPath
    ) -EvidenceLabel 'apply V001 migration'

    Invoke-PostgresCommand -Executable $psql -Arguments @(
        '-X',
        '-v', 'ON_ERROR_STOP=1',
        '-h', '127.0.0.1',
        '-p', $Port,
        '-U', 'postgres',
        '-d', $databaseName,
        '-f', $assertionPath
    ) -EvidenceLabel 'verify V001 constraints and replay protection'

    Invoke-PostgresCommand -Executable $psql -Arguments @(
        '-X',
        '-v', 'ON_ERROR_STOP=1',
        '-h', '127.0.0.1',
        '-p', $Port,
        '-U', 'postgres',
        '-d', $databaseName,
        '-f', $downPath
    ) -EvidenceLabel 'apply V001 down migration'

    Invoke-PostgresCommand -Executable $psql -Arguments @(
        '-X',
        '-v', 'ON_ERROR_STOP=1',
        '-h', '127.0.0.1',
        '-p', $Port,
        '-U', 'postgres',
        '-d', $databaseName,
        '-f', $downAssertionPath
    ) -EvidenceLabel 'verify V001 rollback removed owned objects'

    $completedAt = [DateTimeOffset]::UtcNow
    $evidence = [ordered]@{
        testId = 'DB-MIGRATION-V001-001'
        status = 'PASS'
        postgresVersion = (& $psql '-X' '-At' '-h' '127.0.0.1' '-p' $Port '-U' 'postgres' '-d' $databaseName '-c' 'SHOW server_version;').Trim()
        startedAtUtc = $startedAt.ToString('O')
        completedAtUtc = $completedAt.ToString('O')
        durationMilliseconds = [int64]($completedAt - $startedAt).TotalMilliseconds
        isolation = [ordered]@{
            host = '127.0.0.1'
            port = $Port
            database = $databaseName
            authentication = 'trust on isolated loopback-only test cluster'
            dataDirectory = '.tools/postgresql-v001-data'
        }
        assertions = @(
            'V001 applies transactionally',
            'required tables exist',
            'launch trial campaign ranges do not overlap',
            'published offer ranges do not overlap for the same dimensions',
            'trial grant history is unique per user and product',
            'execution grant nonce and jti are replay-safe',
            'payment provider event ids are idempotent',
            'scoped request idempotency keys reject duplicate records',
            'transactional outbox idempotency keys reject duplicate messages',
            'audit events reject update and delete',
            'payment events reject update and delete',
            'V001 down migration removes all owned tables, enum types, and trigger function'
        )
        commands = $commands
    }

    $evidence | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $evidencePath -Encoding utf8NoBOM
    Write-Output ($evidence | ConvertTo-Json -Depth 8)
}
finally {
    if ($databaseCreated -and $serverStarted) {
        & $dropdb '--if-exists' '-h' '127.0.0.1' '-p' $Port '-U' 'postgres' $databaseName
    }

    if ($serverStarted) {
        & $pgCtl '-D' $dataDirectory '-m' 'fast' '-w' 'stop'
    }

    if ($null -ne $substDrive) {
        & subst.exe $substDrive '/d'
    }
}
