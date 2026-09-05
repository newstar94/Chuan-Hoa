param(
    [Parameter(Mandatory = $false)]
    [ValidatePattern('^$|^\d+\.\d+\.\d+\.\d+$')]
    [string]$ApplicationVersion = ''
)

$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
. (Join-Path $PSScriptRoot 'BuildContract.ps1')
$ApplicationVersion = Resolve-ChuanHoaApplicationVersion `
    -RepositoryRoot $root -RequestedVersion $ApplicationVersion
$project = Join-Path $root 'src\ChuanHoa.AddIn.Vsto\ChuanHoa.AddIn.Vsto.csproj'
$publishRoot = Join-Path $root 'artifacts\vsto-development-test'
$publishDirectory = Join-Path $publishRoot $ApplicationVersion
$publishAlias = 'D:\ChuanHoaDevelopmentTestPublish'
$msbuild = Get-ChuanHoaMsBuild

if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    throw 'Close Microsoft Word completely before building the Development test installer.'
}
if (Test-Path -LiteralPath $publishDirectory) {
    throw "The output directory already exists. Use a newer version: $publishDirectory"
}

New-Item -ItemType Directory -Path $publishDirectory -Force | Out-Null
if (Test-Path -LiteralPath $publishAlias) {
    $aliasItem = Get-Item -LiteralPath $publishAlias -Force
    if ($aliasItem.LinkType -ne 'Junction') {
        throw "The publish alias exists and is not a junction: $publishAlias"
    }
    $actualTarget = [System.IO.Path]::GetFullPath([string]$aliasItem.Target)
    $expectedRoot = [System.IO.Path]::GetFullPath($publishRoot).TrimEnd('\') + '\'
    if (!$actualTarget.StartsWith($expectedRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "The publish alias points outside the expected artifact root: $publishAlias"
    }
    # Windows PowerShell 5.1 can throw a NullReferenceException when Remove-Item
    # removes a junction whose target path contains Vietnamese characters.
    [System.IO.Directory]::Delete($publishAlias)
}
New-Item -ItemType Junction -Path $publishAlias -Target $publishDirectory | Out-Null

$signingCertificatePin = $env:CHUANHOA_DEVELOPMENT_SIGNING_CERT_SHA256
if ([string]::IsNullOrWhiteSpace($signingCertificatePin)) {
    throw 'Set CHUANHOA_DEVELOPMENT_SIGNING_CERT_SHA256 to the approved Development certificate SHA-256 pin.'
}
$certificate = Get-ChuanHoaSigningCertificate -ExpectedSha256 $signingCertificatePin

$buildArguments = @(
    $project,
    '/p:Configuration=Development',
    '/p:Platform=AnyCPU',
    '/p:SignManifests=true',
    "/p:ManifestCertificateThumbprint=$($certificate.Thumbprint)",
    "/p:TrustedSigningCertificateSha256=$signingCertificatePin",
    "/p:PublishDir=$publishAlias\",
    "/p:PublishUrl=$publishAlias\",
    '/p:Install=true',
    '/p:UpdateEnabled=false',
    "/p:ApplicationVersion=$ApplicationVersion",
    '/p:IsWebBootstrapper=false',
    '/m',
    '/nologo',
    '/v:minimal'
)

# BuildSharedClientAssemblies restores and builds Client.Core through MSBuild's
# dependency graph. The VSTO project restore itself is still explicit so the
# publish starts from clean inputs without deleting generated files by hand.
& $msbuild $project /t:Restore /p:RestoreForce=true /p:Configuration=Development /p:Platform=AnyCPU /m /nologo /v:minimal
if ($LASTEXITCODE -ne 0) {
    throw "VSTO Development restore failed with exit code $LASTEXITCODE."
}

# Always rebuild before publishing. An incremental Publish can otherwise reuse a
# stale VSTO assembly when a non-resx embedded Ribbon resource changes.
& $msbuild @buildArguments /t:Rebuild
if ($LASTEXITCODE -ne 0) {
    throw "VSTO Development rebuild failed with exit code $LASTEXITCODE."
}

# Sign owned PE files before Publish so the VSTO application manifest hashes
# the final signed bytes. The manifest is then signed by the Office build target.
$ownedRuntimePeFiles = @(
    (Join-Path $root 'src\ChuanHoa.AddIn.Vsto\bin\Development\ChuanHoa.AddIn.Vsto.dll'),
    (Join-Path $root 'src\ChuanHoa.AddIn.Vsto\bin\Development\ChuanHoa.Client.Core.dll'),
    (Join-Path $root 'src\ChuanHoa.Client.Core\bin\Development\netstandard2.0\ChuanHoa.Client.Core.dll')
)
$ownedRuntimeSignatures = @(
    foreach ($path in $ownedRuntimePeFiles) {
        Invoke-ChuanHoaAuthenticodeSign -Path $path -Certificate $certificate `
            -ExpectedCertificateSha256 $signingCertificatePin
    }
)

& $msbuild @buildArguments /t:PublishOnly /p:BuildProjectReferences=false
if ($LASTEXITCODE -ne 0) {
    throw "VSTO Development publish-only failed with exit code $LASTEXITCODE."
}

$setupPath = Join-Path $publishDirectory 'setup.exe'
$manifestPath = Join-Path $publishDirectory 'ChuanHoa.AddIn.Vsto.vsto'
$certificatePath = Join-Path $publishDirectory 'ChuanHoa.LocalDevelopment.Public.cer'
if (!(Test-Path -LiteralPath $setupPath) -or !(Test-Path -LiteralPath $manifestPath)) {
    throw 'Publish completed without setup.exe or the deployment manifest.'
}
$publishedAssembly = Get-ChildItem -LiteralPath $publishDirectory -Recurse -File -Filter 'ChuanHoa.AddIn.Vsto.dll.deploy' |
    Select-Object -First 1
if ($null -eq $publishedAssembly) {
    throw 'Publish completed without the deployed VSTO assembly.'
}
$loadedAssembly = [System.Reflection.Assembly]::LoadFile($publishedAssembly.FullName)
$embeddedResources = @($loadedAssembly.GetManifestResourceNames())
if ($embeddedResources -notcontains 'ChuanHoa.AddIn.Vsto.Ribbon.ChuanHoaRibbon.xml') {
    throw 'Published VSTO assembly is missing the embedded Ribbon XML resource.'
}
Assert-ChuanHoaManagedAssemblyVersion -Path $publishedAssembly.FullName -ExpectedVersion $ApplicationVersion
Assert-ChuanHoaManifestVersion -Path $manifestPath `
    -IdentityName 'ChuanHoa.AddIn.Vsto.vsto' -ExpectedVersion $ApplicationVersion
$applicationManifest = Get-ChildItem -LiteralPath $publishDirectory -Recurse -File `
    -Filter 'ChuanHoa.AddIn.Vsto.dll.manifest' | Select-Object -First 1
if ($null -eq $applicationManifest) {
    throw 'Publish completed without the VSTO application manifest.'
}
Assert-ChuanHoaManifestVersion -Path $applicationManifest.FullName `
    -IdentityName 'ChuanHoa.AddIn.Vsto.dll' -ExpectedVersion $ApplicationVersion
Export-Certificate -Cert $certificate -FilePath $certificatePath -Force | Out-Null

[pscustomobject]@{
    Version = $ApplicationVersion
    Configuration = 'Development'
    SetupPath = $setupPath
    SetupSha256 = (Get-FileHash -LiteralPath $setupPath -Algorithm SHA256).Hash
    ManifestPath = $manifestPath
    ManifestSha256 = (Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash
    CertificatePath = $certificatePath
    SigningCertificateThumbprint = $certificate.Thumbprint
    SigningCertificateSha256 = Get-ChuanHoaCertificateSha256 -Certificate $certificate
    InnerPeSignatures = $ownedRuntimeSignatures
    ProductionReady = $false
}
