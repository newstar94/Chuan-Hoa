param(
    [switch]$SkipOnlineNuGetAudit
)

$ErrorActionPreference = 'Stop'
$repositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$localDotnet = Join-Path $repositoryRoot '.tools\dotnet\dotnet.exe'
$dotnet = if (Test-Path -LiteralPath $localDotnet -PathType Leaf) {
    $localDotnet
} else {
    (Get-Command dotnet -ErrorAction Stop).Source
}
$python = (Get-Command python -ErrorAction Stop).Source
$solution = Join-Path $repositoryRoot 'ChuanHoa.slnx'
$integrationProject = Join-Path $repositoryRoot 'tests\ChuanHoa.Infrastructure.IntegrationTests\ChuanHoa.Infrastructure.IntegrationTests.csproj'

if (!(Test-Path -LiteralPath $dotnet -PathType Leaf)) {
    throw '.NET SDK executable was not found.'
}

function Assert-LastExitCode {
    param(
        [Parameter(Mandatory = $true)][string]$Label,
        [Parameter(Mandatory = $true)][int]$ExitCode
    )
    if ($ExitCode -ne 0) {
        throw "$Label failed with exit code $ExitCode."
    }
}

Push-Location $repositoryRoot
try {
    $previousCi = $env:CI
    try {
        $env:CI = 'true'
        & $dotnet restore $solution --locked-mode --no-cache
        Assert-LastExitCode -Label 'Solution locked restore' -ExitCode $LASTEXITCODE
        & $dotnet restore $integrationProject --locked-mode --no-cache
        Assert-LastExitCode -Label 'Integration locked restore' -ExitCode $LASTEXITCODE
    }
    finally {
        $env:CI = $previousCi
    }

    & $python -m py_compile `
        tools/validation/validate_solution_projects.py `
        tools/validation/validate_rule_only_product.py `
        tools/validation/validate_development_packaging.py `
        tools/vsto/validate_vsto_source.py `
        tools/validation/validate_document_privacy.py `
        tools/validation/generate_dotnet_sbom.py `
        tools/validation/validate_dotnet_supply_chain.py `
        tools/validation/audit_nuget_vulnerabilities.py `
        tools/validation/validate_repository_secrets.py `
        tools/validation/validate_product_decision_consistency.py
    Assert-LastExitCode -Label 'Python validator compile' -ExitCode $LASTEXITCODE

    & $python tools/validation/validate_solution_projects.py
    Assert-LastExitCode -Label 'Solution project validator' -ExitCode $LASTEXITCODE
    & $python tools/validation/validate_rule_only_product.py
    Assert-LastExitCode -Label 'Rule-only product validator' -ExitCode $LASTEXITCODE
    & $python tools/validation/validate_development_packaging.py
    Assert-LastExitCode -Label 'Development packaging validator' -ExitCode $LASTEXITCODE
    & $python tools/vsto/validate_vsto_source.py
    Assert-LastExitCode -Label 'VSTO source validator' -ExitCode $LASTEXITCODE
    & $python tools/validation/validate_product_decision_consistency.py --self-test `
        --write-evidence shared/docs/implementation/evidence/product_decision_consistency.json
    Assert-LastExitCode -Label 'Product decision consistency validator' -ExitCode $LASTEXITCODE
    & $python tools/validation/validate_document_privacy.py --self-test `
        --write-evidence shared/docs/implementation/evidence/document_privacy_regression.json
    Assert-LastExitCode -Label 'Document privacy validator' -ExitCode $LASTEXITCODE
    & $python tools/validation/validate_dotnet_supply_chain.py `
        --write-evidence shared/docs/implementation/evidence/dotnet_supply_chain.json
    Assert-LastExitCode -Label 'NuGet lock and SBOM validator' -ExitCode $LASTEXITCODE
    & $python tools/validation/generate_dotnet_sbom.py
    Assert-LastExitCode -Label 'CycloneDX SBOM generator' -ExitCode $LASTEXITCODE

    if (!$SkipOnlineNuGetAudit) {
        & $python tools/validation/audit_nuget_vulnerabilities.py `
            --write-evidence shared/docs/implementation/evidence/nuget_vulnerability_audit.json
        Assert-LastExitCode -Label 'Online NuGet vulnerability audit' -ExitCode $LASTEXITCODE
    }

    & $python tools/validation/validate_repository_secrets.py --self-test --history `
        --write-evidence shared/docs/implementation/evidence/repository_secret_regression.json
    Assert-LastExitCode -Label 'Repository secret regression validator' -ExitCode $LASTEXITCODE

    $temporaryOutput = Join-Path ([IO.Path]::GetTempPath()) `
        ('ChuanHoa-SourceQuality-' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $temporaryOutput | Out-Null
    try {
        # Keep the Development API alive: compile/test into an isolated Temp output
        # instead of replacing the Release DLLs that its process may have loaded.
        $testOutput = @(& $dotnet test $solution -c Release --no-restore `
            "/p:BaseOutputPath=$temporaryOutput\bin\" 2>&1)
        $testExitCode = $LASTEXITCODE
        $testOutput | Write-Output
        Assert-LastExitCode -Label 'Full solution test' -ExitCode $testExitCode
        $passed = 0
        $failed = 0
        foreach ($line in $testOutput) {
            $match = [regex]::Match([string]$line, 'Failed:\s+(\d+),\s+Passed:\s+(\d+)')
            if ($match.Success) {
                $failed += [int]$match.Groups[1].Value
                $passed += [int]$match.Groups[2].Value
            }
        }
        if ($failed -ne 0 -or $passed -lt 272) {
            throw "Test discovery floor failed. Passed=$passed Failed=$failed Minimum=272"
        }
    }
    finally {
        $normalizedTemp = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
        $normalizedTarget = [IO.Path]::GetFullPath($temporaryOutput).TrimEnd('\') + '\'
        if ($normalizedTarget.StartsWith($normalizedTemp, [StringComparison]::OrdinalIgnoreCase) -and
            (Split-Path -Leaf $temporaryOutput).StartsWith('ChuanHoa-SourceQuality-', [StringComparison]::Ordinal)) {
            if (Test-Path -LiteralPath $temporaryOutput) {
                Remove-Item -LiteralPath $temporaryOutput -Recurse -Force
            }
        }
    }

    if ($SkipOnlineNuGetAudit) {
        Write-Output "SOURCE_QUALITY_GATES: PASS_WITH_NOT_RUN_NUGET_AUDIT TESTS=$passed"
    }
    else {
        Write-Output "SOURCE_QUALITY_GATES: PASS TESTS=$passed"
    }
}
finally {
    Pop-Location
}
