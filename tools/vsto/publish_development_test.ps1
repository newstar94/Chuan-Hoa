param(
    [Parameter(Mandatory = $false)]
    [ValidatePattern('^$|^\d+\.\d+\.\d+\.\d+$')]
    [string]$ApplicationVersion = '',
    [switch]$SelfSignedRelease
)

$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
. (Join-Path $PSScriptRoot 'BuildContract.ps1')
$ApplicationVersion = Resolve-ChuanHoaApplicationVersion `
    -RepositoryRoot $root -RequestedVersion $ApplicationVersion
$project = Join-Path $root 'src\ChuanHoa.AddIn.Vsto\ChuanHoa.AddIn.Vsto.csproj'
$publishRoot = Join-Path $root 'artifacts\vsto-development-test'
$buildConfiguration = if ($SelfSignedRelease) { 'Release' } else { 'Development' }
if ($SelfSignedRelease) { $publishRoot = Join-Path $root 'artifacts\vsto-self-signed-release' }
$publishDirectory = Join-Path $publishRoot $ApplicationVersion
$publishAlias = Join-Path ([System.IO.Path]::GetTempPath()) `
    ('ChuanHoaDevelopmentPublish-' + [Guid]::NewGuid().ToString('N'))
$msbuild = Get-ChuanHoaMsBuild

if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    throw 'Close Microsoft Word completely before building the Development test installer.'
}
if (Test-Path -LiteralPath $publishDirectory) {
    throw "The output directory already exists. Use a newer version: $publishDirectory"
}

New-Item -ItemType Directory -Path $publishDirectory -Force | Out-Null
# Use a physical ASCII staging directory. Windows can reject traversal through
# a junction as an untrusted mount point during the Office publish target.
New-Item -ItemType Directory -Path $publishAlias | Out-Null

$signingCertificatePin = $env:CHUANHOA_DEVELOPMENT_SIGNING_CERT_SHA256
if ([string]::IsNullOrWhiteSpace($signingCertificatePin)) {
    throw 'Set CHUANHOA_DEVELOPMENT_SIGNING_CERT_SHA256 to the approved Development certificate SHA-256 pin.'
}
$certificate = Get-ChuanHoaSigningCertificate -ExpectedSha256 $signingCertificatePin

$buildArguments = @(
    $project,
    "/p:Configuration=$buildConfiguration",
    '/p:Platform=AnyCPU',
    '/p:SignManifests=true',
    "/p:SelfSignedRelease=$SelfSignedRelease",
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
& $msbuild $project /t:Restore /p:RestoreForce=true /p:Configuration=$buildConfiguration /p:Platform=AnyCPU /m /nologo /v:minimal
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
    (Join-Path $root "src\ChuanHoa.AddIn.Vsto\bin\$buildConfiguration\ChuanHoa.AddIn.Vsto.dll"),
    (Join-Path $root "src\ChuanHoa.AddIn.Vsto\bin\$buildConfiguration\ChuanHoa.Client.Core.dll"),
    (Join-Path $root "src\ChuanHoa.Client.Core\bin\$buildConfiguration\netstandard2.0\ChuanHoa.Client.Core.dll")
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

Get-ChildItem -LiteralPath $publishAlias -Force | ForEach-Object {
    Copy-Item -LiteralPath $_.FullName -Destination $publishDirectory -Recurse -Force
}
$normalizedTemp = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath()).TrimEnd('\') + '\'
$normalizedStage = [System.IO.Path]::GetFullPath($publishAlias).TrimEnd('\') + '\'
if ($normalizedStage.StartsWith($normalizedTemp, [System.StringComparison]::OrdinalIgnoreCase) -and
    (Split-Path -Leaf $publishAlias).StartsWith('ChuanHoaDevelopmentPublish-', [System.StringComparison]::Ordinal)) {
    Remove-Item -LiteralPath $publishAlias -Recurse -Force
}

$publishedApplicationManifest = Get-ChildItem -LiteralPath $publishDirectory `
    -Recurse -File -Filter 'ChuanHoa.AddIn.Vsto.dll.manifest' |
    Select-Object -First 1
$publishedAddIn = Get-ChildItem -LiteralPath $publishDirectory -Recurse -File `
    -Filter 'ChuanHoa.AddIn.Vsto.dll.deploy' | Select-Object -First 1
$publishedClientCore = Get-ChildItem -LiteralPath $publishDirectory -Recurse -File `
    -Filter 'ChuanHoa.Client.Core.dll.deploy' | Select-Object -First 1
$publishedDeploymentManifest = Join-Path $publishDirectory 'ChuanHoa.AddIn.Vsto.vsto'
if ($null -eq $publishedApplicationManifest -or $null -eq $publishedAddIn -or
    $null -eq $publishedClientCore -or
    !(Test-Path -LiteralPath $publishedDeploymentManifest)) {
    throw 'Published VSTO payload is incomplete before manifest finalization.'
}
Update-ChuanHoaSignedVstoManifests `
    -ApplicationManifestPath $publishedApplicationManifest.FullName `
    -DeploymentManifestPath $publishedDeploymentManifest `
    -ApplicationPayloadPaths @{
        'ChuanHoa.AddIn.Vsto.dll' = $publishedAddIn.FullName
        'ChuanHoa.Client.Core.dll' = $publishedClientCore.FullName
    } `
    -Certificate $certificate

# The one-file Development bootstrapper stages the directly loadable runtime
# from bin\Development. Keep its manifests byte-identical to the finalized
# published manifests whose hashes cover the signed DLL bytes.
Copy-Item -LiteralPath $publishedApplicationManifest.FullName -Destination `
    (Join-Path $root "src\ChuanHoa.AddIn.Vsto\bin\$buildConfiguration\ChuanHoa.AddIn.Vsto.dll.manifest") -Force
Copy-Item -LiteralPath $publishedDeploymentManifest -Destination `
    (Join-Path $root "src\ChuanHoa.AddIn.Vsto\bin\$buildConfiguration\ChuanHoa.AddIn.Vsto.vsto") -Force
$runtimeApplicationManifest = Join-Path $root `
    "src\ChuanHoa.AddIn.Vsto\bin\$buildConfiguration\ChuanHoa.AddIn.Vsto.dll.manifest"
$runtimeDeploymentManifest = Join-Path $root `
    "src\ChuanHoa.AddIn.Vsto\bin\$buildConfiguration\ChuanHoa.AddIn.Vsto.vsto"
Update-ChuanHoaSignedVstoManifests `
    -ApplicationManifestPath $runtimeApplicationManifest `
    -DeploymentManifestPath $runtimeDeploymentManifest `
    -ApplicationPayloadPaths @{
        'ChuanHoa.AddIn.Vsto.dll' = (Join-Path $root "src\ChuanHoa.AddIn.Vsto\bin\$buildConfiguration\ChuanHoa.AddIn.Vsto.dll")
        'ChuanHoa.Client.Core.dll' = (Join-Path $root "src\ChuanHoa.AddIn.Vsto\bin\$buildConfiguration\ChuanHoa.Client.Core.dll")
    } `
    -DeploymentApplicationCodebase 'ChuanHoa.AddIn.Vsto.dll.manifest' `
    -DisableDeploymentFileExtensionMapping `
    -Certificate $certificate

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
    Configuration = $buildConfiguration
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
