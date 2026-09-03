param(
    [Parameter(Mandatory = $false)]
    [ValidatePattern('^\d+\.\d+\.\d+\.\d+$')]
    [string]$ApplicationVersion = '1.0.0.14'
)

$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
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

$certificate = Get-ChildItem Cert:\CurrentUser\My |
    Where-Object Subject -eq 'CN=Chuan Hoa Local Development' |
    Sort-Object NotAfter -Descending |
    Select-Object -First 1
if ($null -eq $certificate) {
    throw 'Local development signing certificate was not found.'
}

$msbuild = 'C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe'
if (!(Test-Path -LiteralPath $msbuild)) {
    throw 'Visual Studio Build Tools 2022 MSBuild was not found.'
}

$buildArguments = @(
    $project,
    '/p:Configuration=Release',
    '/p:SignManifests=true',
    "/p:ManifestCertificateThumbprint=$($certificate.Thumbprint)",
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

& $msbuild @buildArguments /t:Rebuild
if ($LASTEXITCODE -ne 0) {
    throw "VSTO rebuild failed with exit code $LASTEXITCODE."
}

& $msbuild @buildArguments /t:Publish
if ($LASTEXITCODE -ne 0) {
    throw "VSTO Publish failed with exit code $LASTEXITCODE."
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

[pscustomobject]@{
    Version = $ApplicationVersion
    SetupPath = $setupPath
    SetupSha256 = (Get-FileHash -LiteralPath $setupPath -Algorithm SHA256).Hash
    ManifestPath = $manifestPath
    ManifestSha256 = (Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash
    SigningCertificate = $certificate.Subject
    ProductionReady = $false
}
