param(
    [string]$ApplicationVersion = '',
    [Parameter(Mandatory = $true)][string]$TrustedPublicKeyPath,
    [Parameter(Mandatory = $true)][string]$TrustedPublicKeySha256,
    [Parameter(Mandatory = $true)][string]$SigningCertificateSha256,
    [Parameter(Mandatory = $true)][string]$SigningRootCertificateSha256
)
$ErrorActionPreference = 'Stop'
& (Join-Path $PSScriptRoot 'build_development_test_exe.ps1') @PSBoundParameters -SelfSignedRelease
if ($LASTEXITCODE -ne 0) { throw "Self-signed Release build failed: $LASTEXITCODE" }
