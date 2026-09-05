param(
    [Parameter(Mandatory = $false)]
    [ValidatePattern('^$|^\d+\.\d+\.\d+\.\d+$')]
    [string]$ApplicationVersion = '',

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TrustedPublicKeyPath,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-fA-F -]{64,95}$')]
    [string]$TrustedPublicKeySha256,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-fA-F -]{64,95}$')]
    [string]$SigningCertificateSha256,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-fA-F -]{64,95}$')]
    [string]$SigningRootCertificateSha256
)

$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
. (Join-Path $PSScriptRoot 'BuildContract.ps1')
$ApplicationVersion = Resolve-ChuanHoaApplicationVersion `
    -RepositoryRoot $root -RequestedVersion $ApplicationVersion
$trustedKey = Assert-ChuanHoaTrustedPublicKey `
    -Path $TrustedPublicKeyPath -ExpectedSha256 $TrustedPublicKeySha256
$SigningCertificateSha256 = Normalize-ChuanHoaSha256 `
    -Value $SigningCertificateSha256 -Label 'SigningCertificateSha256'
$certificate = Get-ChuanHoaSigningCertificate -ExpectedSha256 $SigningCertificateSha256
$SigningRootCertificateSha256 = Normalize-ChuanHoaSha256 `
    -Value $SigningRootCertificateSha256 -Label 'SigningRootCertificateSha256'
$rootCertificate = Get-ChuanHoaSigningCertificate `
    -ExpectedSha256 $SigningRootCertificateSha256
if (![string]::Equals($rootCertificate.Subject,
        'CN=Chuan Hoa Local Development Root',
        [System.StringComparison]::Ordinal) -or
    ![string]::Equals($certificate.Issuer, $rootCertificate.Subject,
        [System.StringComparison]::Ordinal)) {
    throw 'The pinned Development signing certificate chain is invalid.'
}
$publishDirectory = Join-Path $root "artifacts\vsto-development-test\$ApplicationVersion"
$publishScript = Join-Path $root 'tools\vsto\publish_development_test.ps1'
$runtimeDirectory = Join-Path $root 'src\ChuanHoa.AddIn.Vsto\bin\Development'
$bootstrapperProject = Join-Path $root 'tools\vsto\DevelopmentTestBootstrapper\DevelopmentTestBootstrapper.csproj'
$bootstrapperSource = Join-Path $root 'tools\vsto\DevelopmentTestBootstrapper\Program.cs'
$accessSmokeProject = Join-Path $root 'tools\vsto\development-access-smoke\ChuanHoa.DevelopmentAccessSmoke.csproj'
$accessSmokeExecutable = Join-Path $root 'tools\vsto\development-access-smoke\bin\Development\ChuanHoa.DevelopmentAccessSmoke.exe'
$outputDirectory = Join-Path $root 'artifacts\installers\development'
$outputName = "ChuanHoa_Development_Test_Setup_$ApplicationVersion.exe"
$outputPath = Join-Path $outputDirectory $outputName
$msbuild = Get-ChuanHoaMsBuild

if (Test-Path -LiteralPath $publishDirectory) {
    # A payload audit can fail after Publish but before the single-file installer
    # exists. That directory was never released, so it is safe to discard and
    # rebuild the same source-of-truth version. A completed installer remains
    # immutable and still requires a new ProductVersion.
    if (Test-Path -LiteralPath $outputPath) {
        throw "The Development version already has build output. Use a newer ProductVersion instead of reusing binaries: $publishDirectory"
    }
    $resolvedPublish = [System.IO.Path]::GetFullPath($publishDirectory)
    $resolvedPublishRoot = [System.IO.Path]::GetFullPath(
        (Join-Path $root 'artifacts\vsto-development-test')).TrimEnd('\') + '\'
    if (!$resolvedPublish.StartsWith($resolvedPublishRoot,
            [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to clean an incomplete publish outside the artifact root: $resolvedPublish"
    }
    [System.IO.Directory]::Delete($resolvedPublish, $true)
}
$previousSigningPin = $env:CHUANHOA_DEVELOPMENT_SIGNING_CERT_SHA256
$env:CHUANHOA_DEVELOPMENT_SIGNING_CERT_SHA256 = $SigningCertificateSha256
try {
    & $publishScript
    $publishExitCode = $LASTEXITCODE
}
finally {
    $env:CHUANHOA_DEVELOPMENT_SIGNING_CERT_SHA256 = $previousSigningPin
}
if ($publishExitCode -ne 0 -or !(Test-Path -LiteralPath $publishDirectory)) {
    throw "Fresh VSTO Development publish failed for $ApplicationVersion."
}
$runtimeManifest = Join-Path $runtimeDirectory 'ChuanHoa.AddIn.Vsto.vsto'
if (!(Test-Path -LiteralPath $runtimeManifest) -or
    !(Select-String -LiteralPath $runtimeManifest -SimpleMatch "version=`"$ApplicationVersion`"" -Quiet)) {
    throw "The directly loadable Development runtime does not match version $ApplicationVersion."
}
$runtimeAssembly = Join-Path $runtimeDirectory 'ChuanHoa.AddIn.Vsto.dll'
if (!(Test-Path -LiteralPath $runtimeAssembly)) {
    throw 'The directly loadable Development VSTO assembly was not found.'
}
$loadedRuntimeAssembly = [System.Reflection.Assembly]::LoadFile($runtimeAssembly)
$embeddedResources = @($loadedRuntimeAssembly.GetManifestResourceNames())
if ($embeddedResources -notcontains 'ChuanHoa.AddIn.Vsto.Ribbon.ChuanHoaRibbon.xml') {
    throw 'The directly loadable Development VSTO assembly is missing the embedded Ribbon XML resource.'
}
Assert-ChuanHoaManagedAssemblyVersion -Path $runtimeAssembly -ExpectedVersion $ApplicationVersion
Assert-ChuanHoaManifestVersion -Path $runtimeManifest `
    -IdentityName 'ChuanHoa.AddIn.Vsto.vsto' -ExpectedVersion $ApplicationVersion
$bootstrapperCode = Get-Content -LiteralPath $bootstrapperSource -Raw
if ($bootstrapperCode -match 'FileName\s*=\s*setupPath' -or $bootstrapperCode -match 'VSTO setup returned') {
    throw 'Development bootstrapper must not invoke ClickOnce/VSTO setup for installation.'
}
if ($bootstrapperCode -notmatch 'ClearWordRibbonValidationCache\(\)' -or
    $bootstrapperCode -notmatch 'ChuanHoa\.AddIn\.Vsto\.Microsoft\.Word\.Document') {
    throw 'Development bootstrapper must clear only the Chuẩn hóa Word Ribbon validation cache during upgrade.'
}
if (Test-Path -LiteralPath $outputPath) {
    throw "The installer already exists. Use a newer version: $outputPath"
}

New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
$temporaryRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('ChuanHoaInstallerBuild-' + [Guid]::NewGuid().ToString('N'))
$payloadDirectory = Join-Path $temporaryRoot 'payload'
$buildDirectory = Join-Path $temporaryRoot 'build'
$payloadZip = Join-Path $temporaryRoot 'payload.zip'
$versionFile = Join-Path $temporaryRoot 'version.txt'

try {
    New-Item -ItemType Directory -Path $payloadDirectory -Force | Out-Null
    New-Item -ItemType Directory -Path $buildDirectory -Force | Out-Null
    foreach ($runtimeFile in @(
        'ChuanHoa.AddIn.Vsto.vsto',
        'ChuanHoa.AddIn.Vsto.dll.manifest',
        'ChuanHoa.AddIn.Vsto.dll',
        'ChuanHoa.Client.Core.dll',
        'Microsoft.Office.Tools.Common.v4.0.Utilities.dll'
    )) {
        $runtimePath = Join-Path $runtimeDirectory $runtimeFile
        if (!(Test-Path -LiteralPath $runtimePath)) {
            throw "The directly loadable Development runtime is incomplete: $runtimeFile"
        }
        Copy-Item -LiteralPath $runtimePath -Destination $payloadDirectory
    }
    & $msbuild $accessSmokeProject /t:Build /p:Configuration=Development /m /nologo /v:minimal
    if ($LASTEXITCODE -ne 0 -or !(Test-Path -LiteralPath $accessSmokeExecutable)) {
        throw 'The Development access verifier build failed.'
    }
    $accessSmokeSignature = Invoke-ChuanHoaAuthenticodeSign `
        -Path $accessSmokeExecutable -Certificate $certificate `
        -ExpectedCertificateSha256 $SigningCertificateSha256
    Copy-Item -LiteralPath $accessSmokeExecutable -Destination $payloadDirectory
    Copy-Item -LiteralPath (Join-Path $publishDirectory 'ChuanHoa.LocalDevelopment.Public.cer') -Destination $payloadDirectory
    $rootCertificateDestination = Join-Path $payloadDirectory `
        'ChuanHoa.LocalDevelopment.Root.cer'
    Export-Certificate -Cert $rootCertificate -FilePath $rootCertificateDestination `
        -Force | Out-Null
    $supportDirectory = Join-Path $payloadDirectory 'DevelopmentSupport'
    New-Item -ItemType Directory -Path $supportDirectory -Force | Out-Null
    Copy-Item -LiteralPath $trustedKey.Path -Destination (Join-Path $supportDirectory 'trusted-key.xml')

    $certificateDestination = Join-Path $payloadDirectory 'ChuanHoa.LocalDevelopment.Public.cer'
    $publishedCertificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
        $certificateDestination)
    $publishedCertificatePin = Get-ChuanHoaCertificateSha256 -Certificate $publishedCertificate
    if (![string]::Equals($publishedCertificatePin, $SigningCertificateSha256,
            [System.StringComparison]::Ordinal)) {
        throw "Published certificate SHA-256 does not match the build contract. Expected=$SigningCertificateSha256 Actual=$publishedCertificatePin"
    }
    $publishedRootCertificate = `
        [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
            $rootCertificateDestination)
    try {
        $publishedRootCertificatePin = Get-ChuanHoaCertificateSha256 `
            -Certificate $publishedRootCertificate
    }
    finally { $publishedRootCertificate.Dispose() }
    if (![string]::Equals($publishedRootCertificatePin,
            $SigningRootCertificateSha256,
            [System.StringComparison]::Ordinal)) {
        throw "Published root certificate SHA-256 does not match the build contract. Expected=$SigningRootCertificateSha256 Actual=$publishedRootCertificatePin"
    }

    Set-Content -LiteralPath $versionFile -Value $ApplicationVersion -Encoding ASCII -NoNewline
    Compress-Archive -Path (Join-Path $payloadDirectory '*') -DestinationPath $payloadZip -CompressionLevel Optimal

    $forbiddenPayloadPatterns = @(
        'VietnameseEngine',
        'LOCAL-TYPO-AI',
        '.onnx',
        '.onnx.data',
        'onnxruntime',
        'model_manifest',
        'training',
        'QRCoder',
        'btnChenQrCode',
        'QrCodeInputDialog',
        'InsertQrCode'
    )
    $payloadEntries = Get-ChildItem -LiteralPath $payloadDirectory -Recurse -File
    foreach ($entry in $payloadEntries) {
        $relativePath = $entry.FullName.Substring($payloadDirectory.Length).TrimStart('\')
        foreach ($pattern in $forbiddenPayloadPatterns) {
            if ($relativePath.IndexOf($pattern, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                throw "Forbidden retired/AI artifact in Development installer payload: $relativePath"
            }
            $payloadBytes = [System.IO.File]::ReadAllBytes($entry.FullName)
            $asciiPayload = [System.Text.Encoding]::ASCII.GetString($payloadBytes)
            $unicodePayload = [System.Text.Encoding]::Unicode.GetString($payloadBytes)
            if ($asciiPayload.IndexOf($pattern, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                $unicodePayload.IndexOf($pattern, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                throw "Forbidden retired/AI marker in Development installer payload file: $relativePath"
            }
        }
    }

    & $msbuild $bootstrapperProject /t:Build /p:Configuration=Release /p:Platform=AnyCPU `
        /p:PayloadZip="$payloadZip" /p:VersionFile="$versionFile" `
        /p:TrustedPublicKeySha256="$($trustedKey.Sha256)" `
        /p:SigningCertificateSha256="$SigningCertificateSha256" `
        /p:SigningRootCertificateSha256="$SigningRootCertificateSha256" `
        /p:OutputPath="$buildDirectory\" `
        /m /nologo /v:minimal
    if ($LASTEXITCODE -ne 0) {
        throw "Development test bootstrapper build failed with exit code $LASTEXITCODE."
    }

    $builtPath = Join-Path $buildDirectory 'ChuanHoa.DevelopmentTestBootstrapper.exe'
    if (!(Test-Path -LiteralPath $builtPath)) {
        throw 'The Development test bootstrapper executable was not produced.'
    }
    Assert-ChuanHoaManagedAssemblyVersion -Path $builtPath -ExpectedVersion $ApplicationVersion
    Copy-Item -LiteralPath $builtPath -Destination $outputPath

    $signTool = Get-ChildItem -LiteralPath 'C:\Program Files (x86)\Windows Kits\10\bin' -Filter signtool.exe -Recurse -ErrorAction SilentlyContinue |
        Where-Object FullName -match '\\x64\\signtool\.exe$' |
        Sort-Object FullName -Descending |
        Select-Object -First 1
    if ($null -ne $signTool) {
        & $signTool.FullName sign /sha1 $certificate.Thumbprint /fd SHA256 /v $outputPath
        if ($LASTEXITCODE -ne 0) {
            throw "Signing the Development test installer failed with exit code $LASTEXITCODE."
        }
    }
    else {
        $signed = Set-AuthenticodeSignature -LiteralPath $outputPath -Certificate $certificate -HashAlgorithm SHA256
        if ($signed.Status -eq 'NotSigned' -or $signed.Status -eq 'HashMismatch') {
            throw "Signing the Development test installer failed: $($signed.StatusMessage)"
        }
    }

    $signature = Get-AuthenticodeSignature -LiteralPath $outputPath
    if ($signature.Status -eq 'NotSigned' -or $signature.Status -eq 'HashMismatch' -or
        $null -eq $signature.SignerCertificate) {
        throw "The Development test installer signature is invalid: $($signature.Status)"
    }
    $outerSignerPin = Get-ChuanHoaCertificateSha256 -Certificate $signature.SignerCertificate
    if (![string]::Equals($outerSignerPin, $SigningCertificateSha256,
            [System.StringComparison]::Ordinal)) {
        throw "Outer installer signer is not the pinned Development certificate. Expected=$SigningCertificateSha256 Actual=$outerSignerPin"
    }
    Assert-ChuanHoaManagedAssemblyVersion -Path $outputPath -ExpectedVersion $ApplicationVersion

    [pscustomobject]@{
        Version = $ApplicationVersion
        InstallerPath = $outputPath
        SizeBytes = (Get-Item -LiteralPath $outputPath).Length
        Sha256 = (Get-FileHash -LiteralPath $outputPath -Algorithm SHA256).Hash
        SignatureStatus = $signature.Status
        SignerThumbprint = $signature.SignerCertificate.Thumbprint
        SignerSha256 = $outerSignerPin
        TrustedPublicKeySha256 = $trustedKey.Sha256
        AccessSmokeSignature = $accessSmokeSignature
        ProductionReady = $false
    }
}
finally {
    $normalizedTemp = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    $normalizedTarget = [System.IO.Path]::GetFullPath($temporaryRoot).TrimEnd('\') + '\'
    if ($normalizedTarget.StartsWith($normalizedTemp, [System.StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path -Leaf $temporaryRoot).StartsWith('ChuanHoaInstallerBuild-', [System.StringComparison]::Ordinal)) {
        if (Test-Path -LiteralPath $temporaryRoot) {
            Remove-Item -LiteralPath $temporaryRoot -Recurse -Force
        }
    }
}
