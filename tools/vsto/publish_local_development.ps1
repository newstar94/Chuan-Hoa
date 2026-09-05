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
$publishDirectory = Join-Path $root 'artifacts\vsto-dev-publish-local'
$publishAlias = 'D:\ChuanHoaPublishLocal'
$expectedTarget = [System.IO.Path]::GetFullPath($publishDirectory).TrimEnd('\')

New-Item -ItemType Directory -Path $publishDirectory -Force | Out-Null
if (Test-Path -LiteralPath $publishAlias) {
    $aliasItem = Get-Item -LiteralPath $publishAlias -Force
    $actualTarget = [System.IO.Path]::GetFullPath([string]$aliasItem.Target).TrimEnd('\')
    if ($aliasItem.LinkType -ne 'Junction' -or
        ![string]::Equals($actualTarget, $expectedTarget, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Publish alias exists but does not target the expected project directory: $publishAlias"
    }
}
else {
    New-Item -ItemType Junction -Path $publishAlias -Target $publishDirectory | Out-Null
}

$signingCertificatePin = $env:CHUANHOA_DEVELOPMENT_SIGNING_CERT_SHA256
if ([string]::IsNullOrWhiteSpace($signingCertificatePin)) {
    throw 'Set CHUANHOA_DEVELOPMENT_SIGNING_CERT_SHA256 to the approved Development certificate SHA-256 pin.'
}
$certificate = Get-ChuanHoaSigningCertificate -ExpectedSha256 $signingCertificatePin

$msbuild = Get-ChuanHoaMsBuild

$buildArguments = @(
    $project,
    '/p:Configuration=Release',
    '/p:SignManifests=true',
    "/p:ManifestCertificateThumbprint=$($certificate.Thumbprint)",
    "/p:TrustedSigningCertificateSha256=$signingCertificatePin",
    "/p:PublishDir=$publishAlias\",
    '/p:PublishUrl=file:///D:/ChuanHoaPublishLocal/',
    '/p:InstallUrl=file:///D:/ChuanHoaPublishLocal/',
    '/p:Install=true',
    '/p:UpdateEnabled=true',
    '/p:UpdateMode=Foreground',
    '/p:UpdateInterval=0',
    '/p:UpdateIntervalUnits=Days',
    "/p:ApplicationVersion=$ApplicationVersion",
    '/p:IsWebBootstrapper=false',
    '/m',
    '/nologo',
    '/v:minimal'
)

& $msbuild $project /t:Restore /p:RestoreForce=true /p:Configuration=Development /p:Platform=AnyCPU /m /nologo /v:minimal
if ($LASTEXITCODE -ne 0) {
    throw "VSTO Development restore failed with exit code $LASTEXITCODE."
}

& $msbuild @buildArguments /t:Rebuild
if ($LASTEXITCODE -ne 0) {
    throw "VSTO rebuild failed with exit code $LASTEXITCODE."
}

$ownedRuntimePeFiles = @(
    (Join-Path $root 'src\ChuanHoa.AddIn.Vsto\bin\Release\ChuanHoa.AddIn.Vsto.dll'),
    (Join-Path $root 'src\ChuanHoa.AddIn.Vsto\bin\Release\ChuanHoa.Client.Core.dll'),
    (Join-Path $root 'src\ChuanHoa.Client.Core\bin\Release\netstandard2.0\ChuanHoa.Client.Core.dll')
)
$ownedRuntimeSignatures = @(
    foreach ($path in $ownedRuntimePeFiles) {
        Invoke-ChuanHoaAuthenticodeSign -Path $path -Certificate $certificate `
            -ExpectedCertificateSha256 $signingCertificatePin
    }
)

& $msbuild @buildArguments /t:PublishOnly /p:BuildProjectReferences=false
if ($LASTEXITCODE -ne 0) {
    throw "VSTO Publish-only failed with exit code $LASTEXITCODE."
}

$setupPath = Join-Path $publishDirectory 'setup.exe'
$manifestPath = Join-Path $publishDirectory 'ChuanHoa.AddIn.Vsto.vsto'
if (!(Test-Path -LiteralPath $setupPath) -or !(Test-Path -LiteralPath $manifestPath)) {
    throw 'Publish completed without the required setup.exe or deployment manifest.'
}
$publishedAssembly = Get-ChildItem -LiteralPath $publishDirectory -Recurse -File -Filter 'ChuanHoa.AddIn.Vsto.dll.deploy' |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
if ($null -eq $publishedAssembly) {
    throw 'Publish completed without the deployed VSTO assembly.'
}
$loadedAssembly = [System.Reflection.Assembly]::LoadFile($publishedAssembly.FullName)
if (@($loadedAssembly.GetManifestResourceNames()) -notcontains 'ChuanHoa.AddIn.Vsto.Ribbon.ChuanHoaRibbon.xml') {
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

[pscustomobject]@{
    Version = $ApplicationVersion
    SetupPath = $setupPath
    SetupSha256 = (Get-FileHash -LiteralPath $setupPath -Algorithm SHA256).Hash
    ManifestPath = $manifestPath
    ManifestSha256 = (Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash
    SigningCertificateThumbprint = $certificate.Thumbprint
    SigningCertificateSha256 = Get-ChuanHoaCertificateSha256 -Certificate $certificate
    InnerPeSignatures = $ownedRuntimeSignatures
    ProductionReady = $false
}
