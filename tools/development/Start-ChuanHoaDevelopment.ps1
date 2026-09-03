param(
    [switch]$LaunchWord,
    [switch]$LaunchAdmin
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Xml.Linq
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$dotnet = Join-Path $projectRoot '.tools\dotnet\dotnet.exe'
$apiProject = Join-Path $projectRoot 'src\ChuanHoa.Api\ChuanHoa.Api.csproj'
$vstoProject = Join-Path $projectRoot 'src\ChuanHoa.AddIn.Vsto\ChuanHoa.AddIn.Vsto.csproj'
$secretDirectory = Join-Path $projectRoot '.dev-secrets'
$privateKeyPath = Join-Path $secretDirectory 'development-signing-key.xml'
$pidPath = Join-Path $secretDirectory 'api.pid'
$apiStdout = Join-Path $secretDirectory 'api.stdout.log'
$apiStderr = Join-Path $secretDirectory 'api.stderr.log'
$developmentHome = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'ChuanHoa\Development'
$trustedKeyPath = Join-Path $developmentHome 'trusted-key.xml'
$cacheDirectory = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'ChuanHoa\Cache'
$apiUrl = 'http://127.0.0.1:5206'
$keyId = 'CHUANHOA-LOCAL-DEVELOPMENT-1'
$developmentAlias = 'D:\ChuanHoaDevelopment'
$msbuild = 'C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe'

if (!(Test-Path -LiteralPath $dotnet)) { throw 'Local .NET SDK was not found.' }
if (!(Test-Path -LiteralPath $msbuild)) { throw 'Visual Studio Build Tools/VSTO was not found.' }
if ($LaunchWord -and (Get-Process WINWORD -ErrorAction SilentlyContinue)) {
    throw 'Close Microsoft Word completely, then run this script again.'
}

if (Test-Path -LiteralPath $pidPath) {
    $recordedId = 0
    if ([int]::TryParse((Get-Content -LiteralPath $pidPath -Raw).Trim(), [ref]$recordedId)) {
        $recordedProcess = Get-CimInstance Win32_Process -Filter "ProcessId = $recordedId" -ErrorAction SilentlyContinue
        $expectedApiExecutables = @(
            (Join-Path $projectRoot 'src\ChuanHoa.Api\bin\Release\net10.0\ChuanHoa.Api.exe'),
            (Join-Path $developmentAlias 'src\ChuanHoa.Api\bin\Release\net10.0\ChuanHoa.Api.exe')
        ) | ForEach-Object { [IO.Path]::GetFullPath($_) }
        if ($null -ne $recordedProcess -and
            $expectedApiExecutables -contains [IO.Path]::GetFullPath($recordedProcess.ExecutablePath)) {
            Stop-Process -Id $recordedId
            Start-Sleep -Milliseconds 300
        }
    }
    Remove-Item -LiteralPath $pidPath -Force
}

New-Item -ItemType Directory -Path $secretDirectory -Force | Out-Null
New-Item -ItemType Directory -Path $developmentHome -Force | Out-Null
if (Test-Path -LiteralPath $developmentAlias) {
    $aliasItem = Get-Item -LiteralPath $developmentAlias -Force
    $actualTarget = [IO.Path]::GetFullPath([string]$aliasItem.Target).TrimEnd('\')
    $expectedTarget = [IO.Path]::GetFullPath($projectRoot).TrimEnd('\')
    if ($aliasItem.LinkType -ne 'Junction' -or
        ![string]::Equals($actualTarget, $expectedTarget, [StringComparison]::OrdinalIgnoreCase)) {
        throw "The Development alias points to an unexpected location: $developmentAlias"
    }
}
else {
    New-Item -ItemType Junction -Path $developmentAlias -Target $projectRoot | Out-Null
}

if (!(Test-Path -LiteralPath $privateKeyPath)) {
    $rsa = [System.Security.Cryptography.RSA]::Create(3072)
    try {
        $key = $rsa.ExportParameters($true)
        $privateXml = [System.Xml.Linq.XElement]::new('RSAKeyValue',
            [System.Xml.Linq.XElement]::new('Modulus', [Convert]::ToBase64String($key.Modulus)),
            [System.Xml.Linq.XElement]::new('Exponent', [Convert]::ToBase64String($key.Exponent)),
            [System.Xml.Linq.XElement]::new('P', [Convert]::ToBase64String($key.P)),
            [System.Xml.Linq.XElement]::new('Q', [Convert]::ToBase64String($key.Q)),
            [System.Xml.Linq.XElement]::new('DP', [Convert]::ToBase64String($key.DP)),
            [System.Xml.Linq.XElement]::new('DQ', [Convert]::ToBase64String($key.DQ)),
            [System.Xml.Linq.XElement]::new('InverseQ', [Convert]::ToBase64String($key.InverseQ)),
            [System.Xml.Linq.XElement]::new('D', [Convert]::ToBase64String($key.D)))
        [System.IO.File]::WriteAllText($privateKeyPath, $privateXml.ToString([System.Xml.Linq.SaveOptions]::DisableFormatting), [Text.UTF8Encoding]::new($false))
    }
    finally { $rsa.Dispose() }
}

$privateRoot = [System.Xml.Linq.XElement]::Load($privateKeyPath)
$trustedXml = [System.Xml.Linq.XElement]::new('trustedDevelopmentKey',
    [System.Xml.Linq.XAttribute]::new('keyId', $keyId),
    [System.Xml.Linq.XElement]::new('RSAKeyValue',
        [System.Xml.Linq.XElement]::new('Modulus', $privateRoot.Element('Modulus').Value),
        [System.Xml.Linq.XElement]::new('Exponent', $privateRoot.Element('Exponent').Value)))
[System.IO.File]::WriteAllText($trustedKeyPath, $trustedXml.ToString([System.Xml.Linq.SaveOptions]::DisableFormatting), [Text.UTF8Encoding]::new($false))

& $dotnet build $apiProject --configuration Release --nologo
if ($LASTEXITCODE -ne 0) { throw 'The Development API build failed.' }

$certificate = Get-ChildItem Cert:\CurrentUser\My |
    Where-Object Subject -eq 'CN=Chuan Hoa Local Development' |
    Sort-Object NotAfter -Descending |
    Select-Object -First 1
if ($null -eq $certificate) { throw 'The Chuan Hoa Local Development certificate was not found.' }

& $msbuild $vstoProject /t:Build /p:Configuration=Development /p:Platform=AnyCPU `
    /p:SignManifests=true /p:ManifestCertificateThumbprint=$($certificate.Thumbprint) /m /nologo /v:minimal
if ($LASTEXITCODE -ne 0) { throw 'The Development VSTO build failed.' }

$apiReady = $false

if (!$apiReady) {
    $previousEnvironment = $env:ASPNETCORE_ENVIRONMENT
    $previousUrls = $env:ASPNETCORE_URLS
    $previousKeyPath = $env:ChuanHoa__DevelopmentSigningKeyPath
    $previousDictionaryPath = $env:ChuanHoa__DevelopmentRuleDictionaryPath
    $previousLexiconDirectory = $env:ChuanHoa__DevelopmentRuleLexiconDirectory
    $previousBootstrapFlag = $env:ChuanHoa__EnableDevelopmentBootstrap
    $env:ASPNETCORE_ENVIRONMENT = 'Development'
    $env:ASPNETCORE_URLS = $apiUrl
    $env:ChuanHoa__DevelopmentSigningKeyPath = Join-Path $developmentAlias '.dev-secrets\development-signing-key.xml'
    $env:ChuanHoa__DevelopmentRuleDictionaryPath = Join-Path $developmentAlias 'shared\dictionaries\typo_dictionary.json'
    $env:ChuanHoa__DevelopmentRuleLexiconDirectory = Join-Path $developmentAlias 'shared\dictionaries\hunspell-vi'
    $env:ChuanHoa__EnableDevelopmentBootstrap = 'true'
    try {
        $apiExecutable = Join-Path $developmentAlias 'src\ChuanHoa.Api\bin\Release\net10.0\ChuanHoa.Api.exe'
        $process = Start-Process -FilePath $apiExecutable `
            -WorkingDirectory $developmentAlias -WindowStyle Hidden -PassThru -RedirectStandardOutput $apiStdout -RedirectStandardError $apiStderr
        [System.IO.File]::WriteAllText($pidPath, $process.Id.ToString([Globalization.CultureInfo]::InvariantCulture))
    }
    finally {
        $env:ASPNETCORE_ENVIRONMENT = $previousEnvironment
        $env:ASPNETCORE_URLS = $previousUrls
        $env:ChuanHoa__DevelopmentSigningKeyPath = $previousKeyPath
        $env:ChuanHoa__DevelopmentRuleDictionaryPath = $previousDictionaryPath
        $env:ChuanHoa__DevelopmentRuleLexiconDirectory = $previousLexiconDirectory
        $env:ChuanHoa__EnableDevelopmentBootstrap = $previousBootstrapFlag
    }
    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        Start-Sleep -Milliseconds 500
        try {
            $health = Invoke-RestMethod -Uri "$apiUrl/health" -TimeoutSec 2
            if ($health.status -eq 'ok') { $apiReady = $true; break }
        }
        catch { }
    }
}
if (!$apiReady) { throw "The Development API did not start. See $apiStderr" }

$manifest = Join-Path $projectRoot 'src\ChuanHoa.AddIn.Vsto\bin\Development\ChuanHoa.AddIn.Vsto.vsto'
$addinRegistry = 'HKCU:\Software\Microsoft\Office\Word\Addins\ChuanHoa.AddIn.Vsto'
$developmentFriendlyName = [string]::Concat('Chu', [char]0x1EA9, 'n h', [char]0x00F3, 'a (Development)')
New-Item -Path $addinRegistry -Force | Out-Null
New-ItemProperty -Path $addinRegistry -Name Description -Value $developmentFriendlyName -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinRegistry -Name FriendlyName -Value $developmentFriendlyName -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addinRegistry -Name LoadBehavior -Value 3 -PropertyType DWord -Force | Out-Null
New-ItemProperty -Path $addinRegistry -Name Manifest -Value (([Uri]$manifest).AbsoluteUri + '|vstolocal') -PropertyType String -Force | Out-Null

if (Test-Path -LiteralPath $cacheDirectory) {
    foreach ($name in @('lease.xml', 'rules.xml', 'server-time.txt')) {
        $target = Join-Path $cacheDirectory $name
        if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Force }
    }
}

$smokeProject = Join-Path $projectRoot 'tools\vsto\development-access-smoke\ChuanHoa.DevelopmentAccessSmoke.csproj'
& $msbuild $smokeProject /t:Build /p:Configuration=Development /m /nologo /v:minimal
if ($LASTEXITCODE -ne 0) { throw 'The Development license smoke-test build failed.' }
$smokeExe = Join-Path $projectRoot 'tools\vsto\development-access-smoke\bin\Development\ChuanHoa.DevelopmentAccessSmoke.exe'
& $smokeExe
if ($LASTEXITCODE -ne 0) { throw 'The Development API, lease, and rule pack failed the end-to-end smoke test.' }

if ($LaunchWord) {
    $word = (Get-Command WINWORD.EXE -ErrorAction SilentlyContinue).Source
    if ([string]::IsNullOrWhiteSpace($word)) {
        $word = 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE'
    }
    if (!(Test-Path -LiteralPath $word)) { throw 'Microsoft Word was not found.' }
    Start-Process -FilePath $word
}
if ($LaunchAdmin) {
    Start-Process "$apiUrl/development/admin"
}

[pscustomobject]@{
    Mode = 'Development'
    Api = $apiUrl
    VstoManifest = $manifest
    OfflineLeaseDays = 7
    TrustedKey = $trustedKeyPath
    WordLaunched = [bool]$LaunchWord
    Admin = "$apiUrl/development/admin"
    AdminLaunched = [bool]$LaunchAdmin
    ProductionReady = $false
}
