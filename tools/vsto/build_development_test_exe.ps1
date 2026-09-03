param(
    [Parameter(Mandatory = $false)]
    [ValidatePattern('^\d+\.\d+\.\d+\.\d+$')]
    [string]$ApplicationVersion = '1.0.0.16'
)

$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$publishDirectory = Join-Path $root "artifacts\vsto-development-test\$ApplicationVersion"
$runtimeDirectory = Join-Path $root 'src\ChuanHoa.AddIn.Vsto\bin\Development'
$bootstrapperProject = Join-Path $root 'tools\vsto\DevelopmentTestBootstrapper\DevelopmentTestBootstrapper.csproj'
$bootstrapperSource = Join-Path $root 'tools\vsto\DevelopmentTestBootstrapper\Program.cs'
$trustedKeyPath = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'ChuanHoa\Development\trusted-key.xml'
$outputDirectory = Join-Path $root 'artifacts\installers\development'
$outputName = "ChuanHoa_Development_Test_Setup_$ApplicationVersion.exe"
$outputPath = Join-Path $outputDirectory $outputName
$msbuild = 'C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe'

if (!(Test-Path -LiteralPath $publishDirectory)) {
    throw "The published Development package was not found: $publishDirectory"
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
if (!(Test-Path -LiteralPath $trustedKeyPath)) {
    throw 'The Development public trust key was not found. Run Start-ChuanHoaDevelopment.ps1 first.'
}
if (!(Test-Path -LiteralPath $msbuild)) {
    throw 'Visual Studio Build Tools 2022 MSBuild was not found.'
}
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
        'Microsoft.Office.Tools.Common.v4.0.Utilities.dll',
        'QRCoder.dll'
    )) {
        $runtimePath = Join-Path $runtimeDirectory $runtimeFile
        if (!(Test-Path -LiteralPath $runtimePath)) {
            throw "The directly loadable Development runtime is incomplete: $runtimeFile"
        }
        Copy-Item -LiteralPath $runtimePath -Destination $payloadDirectory
    }
    Copy-Item -LiteralPath (Join-Path $publishDirectory 'ChuanHoa.LocalDevelopment.Public.cer') -Destination $payloadDirectory
    $supportDirectory = Join-Path $payloadDirectory 'DevelopmentSupport'
    New-Item -ItemType Directory -Path $supportDirectory -Force | Out-Null
    Copy-Item -LiteralPath $trustedKeyPath -Destination (Join-Path $supportDirectory 'trusted-key.xml')
    $cacheDir = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'ChuanHoa\Cache'
    foreach ($cacheFile in @('lease.xml', 'rules.xml', 'server-time.txt')) {
        $sourceCache = Join-Path $cacheDir $cacheFile
        if (Test-Path -LiteralPath $sourceCache) {
            Copy-Item -LiteralPath $sourceCache -Destination (Join-Path $supportDirectory $cacheFile)
        }
    }
    Set-Content -LiteralPath $versionFile -Value $ApplicationVersion -Encoding ASCII -NoNewline
    Compress-Archive -Path (Join-Path $payloadDirectory '*') -DestinationPath $payloadZip -CompressionLevel Optimal

    & $msbuild $bootstrapperProject /t:Build /p:Configuration=Release /p:Platform=AnyCPU `
        /p:PayloadZip="$payloadZip" /p:VersionFile="$versionFile" /p:OutputPath="$buildDirectory\" `
        /m /nologo /v:minimal
    if ($LASTEXITCODE -ne 0) {
        throw "Development test bootstrapper build failed with exit code $LASTEXITCODE."
    }

    $builtPath = Join-Path $buildDirectory 'ChuanHoa.DevelopmentTestBootstrapper.exe'
    if (!(Test-Path -LiteralPath $builtPath)) {
        throw 'The Development test bootstrapper executable was not produced.'
    }
    Copy-Item -LiteralPath $builtPath -Destination $outputPath

    $certificate = Get-ChildItem Cert:\CurrentUser\My |
        Where-Object Subject -eq 'CN=Chuan Hoa Local Development' |
        Sort-Object NotAfter -Descending |
        Select-Object -First 1
    if ($null -eq $certificate) {
        throw 'The local Development signing certificate was not found.'
    }
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
    if ($signature.Status -eq 'NotSigned' -or $signature.Status -eq 'HashMismatch') {
        throw "The Development test installer signature is invalid: $($signature.Status)"
    }

    [pscustomobject]@{
        Version = $ApplicationVersion
        InstallerPath = $outputPath
        SizeBytes = (Get-Item -LiteralPath $outputPath).Length
        Sha256 = (Get-FileHash -LiteralPath $outputPath -Algorithm SHA256).Hash
        SignatureStatus = $signature.Status
        Signer = $signature.SignerCertificate.Subject
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
