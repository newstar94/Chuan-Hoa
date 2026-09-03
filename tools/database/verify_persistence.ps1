[CmdletBinding()]
param(
    [int]$Port = 55441
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

$postgresBin = Join-Path $executionRoot '.tools\postgresql\pgsql\bin'
$dataDirectory = Join-Path $executionRoot '.tools\postgresql-persistence-data'
$logPath = Join-Path $executionRoot '.tools\postgresql-persistence.log'
$databaseName = 'chuanhoa_persistence_test'
$migrationPath = Join-Path $executionRoot 'database\migrations\V001__identity_trial_commercial_foundation.sql'
$testProject = Join-Path $executionRoot 'tests\ChuanHoa.Infrastructure.IntegrationTests\ChuanHoa.Infrastructure.IntegrationTests.csproj'
$dotnet = Join-Path $executionRoot '.tools\dotnet\dotnet.exe'
$evidencePath = Join-Path $executionRoot 'shared\docs\implementation\evidence\persistence_integration.json'
$initdb = Join-Path $postgresBin 'initdb.exe'
$pgCtl = Join-Path $postgresBin 'pg_ctl.exe'
$createdb = Join-Path $postgresBin 'createdb.exe'
$dropdb = Join-Path $postgresBin 'dropdb.exe'
$psql = Join-Path $postgresBin 'psql.exe'
$serverStarted = $false
$databaseCreated = $false
$startedAt = [DateTimeOffset]::UtcNow

function Invoke-Checked {
    param(
        [Parameter(Mandatory)]
        [string]$Executable,
        [Parameter(Mandatory)]
        [string[]]$Arguments,
        [Parameter(Mandatory)]
        [string]$Label
    )

    & $Executable @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "$Label failed with exit code $LASTEXITCODE."
    }
}

try {
    foreach ($requiredPath in @($initdb, $pgCtl, $createdb, $dropdb, $psql, $migrationPath, $testProject, $dotnet)) {
        if (-not (Test-Path -LiteralPath $requiredPath)) {
            throw "Required file not found: $requiredPath"
        }
    }

    if ($null -ne (Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue)) {
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

        Invoke-Checked $initdb @(
            '-D', $dataDirectory,
            '-A', 'trust',
            '-U', 'postgres',
            '--encoding=UTF8',
            '--no-locale'
        ) 'initialize isolated PostgreSQL cluster'
    }

    if (Test-Path -LiteralPath $logPath) {
        Remove-Item -LiteralPath $logPath -Force
    }

    Invoke-Checked $pgCtl @(
        '-D', $dataDirectory,
        '-l', $logPath,
        '-o', "-p $Port -h 127.0.0.1",
        '-w',
        'start'
    ) 'start isolated PostgreSQL server'
    $serverStarted = $true

    & $dropdb '--if-exists' '-h' '127.0.0.1' '-p' $Port '-U' 'postgres' $databaseName
    if ($LASTEXITCODE -ne 0) {
        throw "drop stale test database failed with exit code $LASTEXITCODE."
    }

    Invoke-Checked $createdb @(
        '-h', '127.0.0.1',
        '-p', $Port,
        '-U', 'postgres',
        $databaseName
    ) 'create persistence test database'
    $databaseCreated = $true

    Invoke-Checked $psql @(
        '-X',
        '-v', 'ON_ERROR_STOP=1',
        '-h', '127.0.0.1',
        '-p', $Port,
        '-U', 'postgres',
        '-d', $databaseName,
        '-f', $migrationPath
    ) 'apply V001 migration'

    $previousConnectionString = $env:CHUANHOA_TEST_CONNECTION_STRING
    try {
        $env:CHUANHOA_TEST_CONNECTION_STRING = "Host=127.0.0.1;Port=$Port;Database=$databaseName;Username=postgres;Pooling=false;Timeout=5;Command Timeout=10"
        Invoke-Checked $dotnet @(
            'test',
            $testProject,
            '-c', 'Release'
        ) 'run PostgreSQL persistence integration tests'
    }
    finally {
        $env:CHUANHOA_TEST_CONNECTION_STRING = $previousConnectionString
    }

    $completedAt = [DateTimeOffset]::UtcNow
    $evidence = [ordered]@{
        testId = 'DB-PERSISTENCE-001'
        status = 'PASS'
        postgresVersion = (& $psql '-X' '-At' '-h' '127.0.0.1' '-p' $Port '-U' 'postgres' '-d' $databaseName '-c' 'SHOW server_version;').Trim()
        startedAtUtc = $startedAt.ToString('O')
        completedAtUtc = $completedAt.ToString('O')
        durationMilliseconds = [int64]($completedAt - $startedAt).TotalMilliseconds
        testCount = 4
        assertions = @(
            'same key and same request reports in progress while owned',
            'same key and different request hash reports conflict',
            'completed response replays status, content type, headers, and bytes',
            'retryable failure can be reacquired only with a new owner token',
            'outbox insert rolls back with a failed transaction',
            'outbox insert commits with a successful transaction'
        )
    }
    $evidence | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $evidencePath -Encoding utf8NoBOM
    Write-Output ($evidence | ConvertTo-Json -Depth 6)
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
