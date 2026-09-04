param(
    [Parameter(Mandatory = $false)]
    [ValidatePattern('^$|^\d+\.\d+\.\d+\.\d+$')]
    [string]$ApplicationVersion = ''
)

$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
if ([string]::IsNullOrWhiteSpace($ApplicationVersion)) {
    [xml]$versionProps = Get-Content -LiteralPath (Join-Path $root 'Directory.Build.props')
    $ApplicationVersion = [string]$versionProps.Project.PropertyGroup.ProductVersion
}
$project = Join-Path $root 'src\ChuanHoa.AddIn.Vsto\ChuanHoa.AddIn.Vsto.csproj'
$publishRoot = Join-Path $root 'artifacts\vsto-development-test'
$publishDirectory = Join-Path $publishRoot $ApplicationVersion
$publishAlias = 'D:\ChuanHoaDevelopmentTestPublish'
$msbuild = 'C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe'

if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    throw 'Close Microsoft Word completely before building the Development test installer.'
}
if (!(Test-Path -LiteralPath $msbuild)) {
    throw 'Visual Studio Build Tools 2022 MSBuild was not found.'
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

$certificate = Get-ChildItem Cert:\CurrentUser\My |
    Where-Object Subject -eq 'CN=Chuan Hoa Local Development' |
    Sort-Object NotAfter -Descending |
    Select-Object -First 1
if ($null -eq $certificate) {
    throw 'The local Development signing certificate was not found.'
}

$buildArguments = @(
    $project,
    '/p:Configuration=Development',
    '/p:Platform=AnyCPU',
    '/p:SignManifests=true',
    "/p:ManifestCertificateThumbprint=$($certificate.Thumbprint)",
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

# A clean checkout has no VSTO project.assets.json because the SDK solution does
# not include this legacy Office project. Restore it explicitly before build.
# Remove NuGet's generated dependency snapshot first. When the final
# PackageReference is removed, NuGet reports "Nothing to do" and otherwise
# leaves the old dependency in the VSTO deployment manifest.
foreach ($generatedNuGetFile in @(
    'obj\project.assets.json',
    'obj\ChuanHoa.AddIn.Vsto.csproj.nuget.dgspec.json',
    'obj\ChuanHoa.AddIn.Vsto.csproj.nuget.g.props',
    'obj\ChuanHoa.AddIn.Vsto.csproj.nuget.g.targets'
)) {
    $generatedNuGetPath = Join-Path (Split-Path -Parent $project) $generatedNuGetFile
    if (Test-Path -LiteralPath $generatedNuGetPath) {
        Remove-Item -LiteralPath $generatedNuGetPath -Force
    }
}
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

& $msbuild @buildArguments /t:Publish
if ($LASTEXITCODE -ne 0) {
    throw "VSTO Development publish failed with exit code $LASTEXITCODE."
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
Export-Certificate -Cert $certificate -FilePath $certificatePath -Force | Out-Null

[pscustomobject]@{
    Version = $ApplicationVersion
    Configuration = 'Development'
    SetupPath = $setupPath
    SetupSha256 = (Get-FileHash -LiteralPath $setupPath -Algorithm SHA256).Hash
    ManifestPath = $manifestPath
    ManifestSha256 = (Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash
    CertificatePath = $certificatePath
    SigningCertificate = $certificate.Subject
    ProductionReady = $false
}
