param(
    [Parameter(Mandatory = $false)]
    [string]$InstallerPath = '',

    [Parameter(Mandatory = $false)]
    [string]$EvidencePath = ''
)

$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$props = [xml](Get-Content -LiteralPath (Join-Path $root 'Directory.Build.props') -Raw)
$version = [string]$props.Project.PropertyGroup.ProductVersion
if ($version -notmatch '^\d+\.\d+\.\d+\.\d+$') {
    throw 'Directory.Build.props does not contain one valid ProductVersion.'
}
if ([string]::IsNullOrWhiteSpace($InstallerPath)) {
    $InstallerPath = Join-Path $root (
        'artifacts\installers\development\ChuanHoa_Development_Test_Setup_' + $version + '.exe')
}
$installer = (Resolve-Path -LiteralPath $InstallerPath).Path
if ((Get-Process WINWORD -ErrorAction SilentlyContinue).Count -ne 0) {
    throw 'Microsoft Word is open. Save or recover the document and close every WINWORD process before lifecycle verification.'
}

$baseDirectory = Join-Path $env:LOCALAPPDATA 'ChuanHoa\DevelopmentInstaller'
$currentDirectory = Join-Path $baseDirectory 'Current'
$dictionaryDirectory = Join-Path $env:LOCALAPPDATA 'ChuanHoa\Dictionaries'
$trustedKeyPath = Join-Path $env:LOCALAPPDATA 'ChuanHoa\Development\trusted-key.xml'
$accessCachePaths = @(
    (Join-Path $env:LOCALAPPDATA 'ChuanHoa\Cache\lease.xml'),
    (Join-Path $env:LOCALAPPDATA 'ChuanHoa\Cache\rules.xml'),
    (Join-Path $env:LOCALAPPDATA 'ChuanHoa\Cache\server-time.txt')
)
$addInRegistryPath = 'HKCU:\Software\Microsoft\Office\Word\Addins\ChuanHoa.AddIn.Vsto'
$installerRegistryPath = 'HKCU:\Software\ChuanHoa\DevelopmentInstaller'
$appsFeaturesRegistryPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\ChuanHoa.DevelopmentTest'
$installerCacheRoot = Join-Path $env:LOCALAPPDATA 'ChuanHoa\InstallerCache\Development'
$installerCacheVersionDirectory = Join-Path $installerCacheRoot $version
$cachedInstallerPath = Join-Path $installerCacheVersionDirectory 'ChuanHoa_Development_Test_Setup.exe'
$faultPoints = @(
    'after-trusted-certificate',
    'after-trusted-key',
    'after-access-smoke',
    'after-current-switch',
    'after-registry',
    'after-verification'
)

function Get-FileInventory([string]$path) {
    if (!(Test-Path -LiteralPath $path)) { return @() }
    return @(
        Get-ChildItem -LiteralPath $path -File -Recurse -ErrorAction Stop |
            Sort-Object FullName |
            ForEach-Object {
                [ordered]@{
                    Path = $_.FullName.Substring($path.Length).Replace('\', '/')
                    Length = $_.Length
                    Sha256 = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
                }
            }
    )
}

function Get-RegistryInventory([string]$path) {
    if (!(Test-Path -LiteralPath $path)) { return $null }
    $item = Get-ItemProperty -LiteralPath $path
    $values = [ordered]@{}
    foreach ($property in $item.PSObject.Properties |
        Where-Object { $_.Name -notlike 'PS*' } | Sort-Object Name) {
        $values[$property.Name] = [string]$property.Value
    }
    return $values
}

function Get-FileSnapshot([string]$path) {
    if (!(Test-Path -LiteralPath $path -PathType Leaf)) {
        return [ordered]@{ Path = $path; Exists = $false }
    }
    $item = Get-Item -LiteralPath $path
    return [ordered]@{
        Path = $path
        Exists = $true
        Length = $item.Length
        Sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    }
}

function Get-OwnedCertificateInventory {
    $rows = @()
    foreach ($storeName in @('Root', 'TrustedPublisher')) {
        $store = New-Object System.Security.Cryptography.X509Certificates.X509Store(
            $storeName,
            [System.Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
        try {
            $store.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
            foreach ($certificate in $store.Certificates |
                Where-Object { $_.Subject -eq 'CN=Chuan Hoa Local Development' } |
                Sort-Object Thumbprint) {
                $temporary = Join-Path $env:TEMP ('ChuanHoaCertificate-' + [Guid]::NewGuid().ToString('N') + '.cer')
                try {
                    [IO.File]::WriteAllBytes($temporary, $certificate.RawData)
                    $sha256 = (Get-FileHash -LiteralPath $temporary -Algorithm SHA256).Hash
                }
                finally {
                    if (Test-Path -LiteralPath $temporary) { [IO.File]::Delete($temporary) }
                }
                $rows += [ordered]@{ Store = $storeName; Thumbprint = $certificate.Thumbprint; Sha256 = $sha256 }
            }
        }
        finally { $store.Close() }
    }
    return @($rows)
}

function Get-ManagedState {
    return [ordered]@{
        Files = Get-FileInventory $baseDirectory
        TrustedKey = Get-FileSnapshot $trustedKeyPath
        AccessCache = @($accessCachePaths | ForEach-Object { Get-FileSnapshot $_ })
        PersonalDictionary = Get-FileInventory $dictionaryDirectory
        AddInRegistry = Get-RegistryInventory $addInRegistryPath
        InstallerRegistry = Get-RegistryInventory $installerRegistryPath
        AppsFeaturesRegistry = Get-RegistryInventory $appsFeaturesRegistryPath
        InstallerCache = Get-FileInventory $installerCacheRoot
        Certificates = Get-OwnedCertificateInventory
    }
}

function Get-StateHash($state) {
    $json = $state | ConvertTo-Json -Depth 10 -Compress
    $bytes = [Text.Encoding]::UTF8.GetBytes($json)
    $sha256 = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha256.ComputeHash($bytes))).Replace('-', '') }
    finally { $sha256.Dispose() }
}

function Invoke-Installer([string[]]$arguments, [string]$faultPoint = '') {
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = $installer
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $start.Arguments = [string]::Join(' ', $arguments)
    if (![string]::IsNullOrWhiteSpace($faultPoint)) {
        $start.EnvironmentVariables['CHUANHOA_INSTALLER_ENABLE_FAULT_INJECTION'] = '1'
        $start.EnvironmentVariables['CHUANHOA_INSTALLER_FAULT_POINT'] = $faultPoint
    }
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $start
    try {
        if (!$process.Start()) { throw 'Could not start the Development installer.' }
        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        return [ordered]@{
            ExitCode = $process.ExitCode
            Stdout = $stdout.Trim()
            Stderr = $stderr.Trim()
        }
    }
    finally { $process.Dispose() }
}

function Invoke-RegisteredUninstall {
    $entry = Get-ItemProperty -LiteralPath $appsFeaturesRegistryPath
    $command = [string]$entry.QuietUninstallString
    if ($command -notmatch '^"([^"]+)"\s+(.+)$') {
        throw 'QuietUninstallString is not a safely quoted executable command.'
    }
    $executable = $Matches[1]
    $arguments = $Matches[2]
    if (![string]::Equals($executable, $cachedInstallerPath,
            [StringComparison]::OrdinalIgnoreCase)) {
        throw 'QuietUninstallString does not use the expected versioned installer cache.'
    }
    if (![string]::Equals($arguments, '/uninstall /quiet', [StringComparison]::Ordinal)) {
        throw 'QuietUninstallString has unexpected arguments.'
    }
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = $executable
    $start.Arguments = $arguments
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $start
    try {
        if (!$process.Start()) { throw 'Could not start registered Development uninstaller.' }
        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        return [ordered]@{
            ExitCode = $process.ExitCode
            Stdout = $stdout.Trim()
            Stderr = $stderr.Trim()
            Command = $command
        }
    }
    finally { $process.Dispose() }
}

function Assert-InstalledCurrent {
    $dll = Join-Path $currentDirectory 'ChuanHoa.AddIn.Vsto.dll'
    $manifest = Join-Path $currentDirectory 'ChuanHoa.AddIn.Vsto.vsto'
    if (!(Test-Path -LiteralPath $dll) -or !(Test-Path -LiteralPath $manifest)) {
        throw 'Installed Current payload is incomplete.'
    }
    $fileVersion = (Get-Item -LiteralPath $dll).VersionInfo.FileVersion
    if (![string]::Equals($fileVersion, $version, [StringComparison]::Ordinal)) {
        throw "Installed DLL version mismatch. Expected=$version Actual=$fileVersion"
    }
    if (!(Select-String -LiteralPath $manifest -SimpleMatch "version=`"$version`"" -Quiet)) {
        throw 'Installed VSTO manifest does not match ProductVersion.'
    }
    $installerState = Get-ItemProperty -LiteralPath $installerRegistryPath
    $expectedSignerPin = [string]$installerState.SigningCertificateSha256
    foreach ($ownedPe in @(
        $dll,
        (Join-Path $currentDirectory 'ChuanHoa.Client.Core.dll'),
        (Join-Path $currentDirectory 'ChuanHoa.DevelopmentAccessSmoke.exe')
    )) {
        $signature = Get-AuthenticodeSignature -LiteralPath $ownedPe
        if ($signature.Status -ne 'Valid' -or $null -eq $signature.SignerCertificate) {
            throw "Installed owned PE signature is not valid: $ownedPe ($($signature.Status))"
        }
        $sha256 = [Security.Cryptography.SHA256]::Create()
        try {
            $signerPin = ([BitConverter]::ToString(
                $sha256.ComputeHash($signature.SignerCertificate.RawData))).Replace('-', '')
        }
        finally { $sha256.Dispose() }
        if (![string]::Equals($signerPin, $expectedSignerPin, [StringComparison]::Ordinal)) {
            throw "Installed owned PE signer pin mismatch: $ownedPe"
        }
    }
    $registration = Get-ItemProperty -LiteralPath $addInRegistryPath
    if ([int]$registration.LoadBehavior -ne 3 -or
        [string]$registration.Manifest -notlike '*DevelopmentInstaller/Current/ChuanHoa.AddIn.Vsto.vsto|vstolocal') {
        throw 'Installed Word registration is not active or does not target Current.'
    }
    if (!(Test-Path -LiteralPath $cachedInstallerPath -PathType Leaf)) {
        throw 'Cached signed installer is missing.'
    }
    $sourceHash = (Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash
    $cachedHash = (Get-FileHash -LiteralPath $cachedInstallerPath -Algorithm SHA256).Hash
    if (![string]::Equals($sourceHash, $cachedHash, [StringComparison]::Ordinal)) {
        throw 'Cached installer differs from the released installer.'
    }
    $entry = Get-ItemProperty -LiteralPath $appsFeaturesRegistryPath
    if (![string]::Equals([string]$entry.DisplayVersion, $version, [StringComparison]::Ordinal) -or
        ![string]::Equals([string]$entry.InstallLocation, $currentDirectory, [StringComparison]::OrdinalIgnoreCase) -or
        ![string]::Equals([string]$entry.QuietUninstallString,
            ('"' + $cachedInstallerPath + '" /uninstall /quiet'), [StringComparison]::Ordinal) -or
        ![string]::Equals([string]$entry.ModifyPath,
            ('"' + $cachedInstallerPath + '" /repair'), [StringComparison]::Ordinal)) {
        throw 'Apps & Features registration is incomplete or points outside the signed cache.'
    }
}

function Assert-Uninstalled {
    if (Test-Path -LiteralPath $addInRegistryPath) { throw 'Word add-in registration remains after uninstall.' }
    if (Test-Path -LiteralPath $currentDirectory) { throw 'Current payload remains after uninstall.' }
    if (Test-Path -LiteralPath $appsFeaturesRegistryPath) { throw 'Apps & Features entry remains after uninstall.' }
    if (Test-Path -LiteralPath $installerCacheVersionDirectory) { throw 'Versioned installer cache remains after uninstall.' }
}

function Assert-DictionaryPreserved([string]$expectedHash) {
    $actualHash = Get-StateHash (Get-FileInventory $dictionaryDirectory)
    if (![string]::Equals($actualHash, $expectedHash, [StringComparison]::Ordinal)) {
        throw 'Personal dictionary changed during installer lifecycle verification.'
    }
}

$dictionaryHash = Get-StateHash (Get-FileInventory $dictionaryDirectory)
Assert-InstalledCurrent

$repair = Invoke-Installer @('/repair', '/quiet')
if ($repair.ExitCode -ne 0) { throw "Repair failed with exit code $($repair.ExitCode): $($repair.Stderr)" }
Assert-InstalledCurrent
Assert-DictionaryPreserved $dictionaryHash

$rollbackResults = @()
foreach ($faultPoint in $faultPoints) {
    $beforeHash = Get-StateHash (Get-ManagedState)
    $result = Invoke-Installer @('/repair', '/quiet') $faultPoint
    $afterHash = Get-StateHash (Get-ManagedState)
    if ($result.ExitCode -ne 10) {
        throw "Fault point $faultPoint returned $($result.ExitCode), expected 10. $($result.Stderr)"
    }
    if (![string]::Equals($beforeHash, $afterHash, [StringComparison]::Ordinal)) {
        throw "Installer rollback did not restore exact managed state at $faultPoint."
    }
    Assert-DictionaryPreserved $dictionaryHash
    $rollbackResults += [ordered]@{ FaultPoint = $faultPoint; ExitCode = $result.ExitCode; StateRestored = $true }
}

$uninstall = Invoke-RegisteredUninstall
if ($uninstall.ExitCode -ne 0) { throw "Uninstall failed with exit code $($uninstall.ExitCode): $($uninstall.Stderr)" }
$cleanupDeadline = [DateTime]::UtcNow.AddSeconds(15)
while ((Test-Path -LiteralPath $installerCacheVersionDirectory) -and
    [DateTime]::UtcNow -lt $cleanupDeadline) {
    Start-Sleep -Milliseconds 100
}
Assert-Uninstalled
Assert-DictionaryPreserved $dictionaryHash

$reinstall = Invoke-Installer @('/quiet')
if ($reinstall.ExitCode -ne 0) { throw "Fresh reinstall failed with exit code $($reinstall.ExitCode): $($reinstall.Stderr)" }
Assert-InstalledCurrent
Assert-DictionaryPreserved $dictionaryHash

$accessSmokePath = Join-Path $currentDirectory 'ChuanHoa.DevelopmentAccessSmoke.exe'
$accessSmoke = New-Object Diagnostics.ProcessStartInfo
$accessSmoke.FileName = $accessSmokePath
$accessSmoke.UseShellExecute = $false
$accessSmoke.CreateNoWindow = $true
$accessSmoke.RedirectStandardOutput = $true
$accessSmoke.RedirectStandardError = $true
$accessProcess = New-Object Diagnostics.Process
$accessProcess.StartInfo = $accessSmoke
try {
    if (!$accessProcess.Start()) { throw 'Could not start installed Development access smoke.' }
    $accessOutput = $accessProcess.StandardOutput.ReadToEnd().Trim()
    $accessError = $accessProcess.StandardError.ReadToEnd().Trim()
    $accessProcess.WaitForExit()
    if ($accessProcess.ExitCode -ne 0) {
        throw "Installed access smoke failed with exit code $($accessProcess.ExitCode): $accessError"
    }
}
finally { $accessProcess.Dispose() }

$evidence = [ordered]@{
    SchemaVersion = 2
    Status = 'PASS_DEVELOPMENT_LIFECYCLE'
    CapturedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
    ProductVersion = $version
    InstallerPath = $installer
    InstallerSizeBytes = (Get-Item -LiteralPath $installer).Length
    InstallerSha256 = (Get-FileHash -LiteralPath $installer -Algorithm SHA256).Hash
    RepairExitCode = $repair.ExitCode
    RollbackFaultPoints = $rollbackResults
    UninstallExitCode = $uninstall.ExitCode
    UninstallCommand = $uninstall.Command
    AppsFeaturesRegistered = $true
    CachedInstallerSha256 = (Get-FileHash -LiteralPath $cachedInstallerPath -Algorithm SHA256).Hash
    ReinstallExitCode = $reinstall.ExitCode
    PersonalDictionaryPreserved = $true
    InstalledVersion = (Get-Item -LiteralPath (Join-Path $currentDirectory 'ChuanHoa.AddIn.Vsto.dll')).VersionInfo.FileVersion
    LoadBehavior = [int](Get-ItemProperty -LiteralPath $addInRegistryPath).LoadBehavior
    AccessSmokeExitCode = 0
    AccessSmokeOutput = $accessOutput
}

$json = $evidence | ConvertTo-Json -Depth 10
if (![string]::IsNullOrWhiteSpace($EvidencePath)) {
    $resolvedEvidence = [IO.Path]::GetFullPath((Join-Path $root $EvidencePath))
    $evidenceDirectory = [IO.Path]::GetDirectoryName($resolvedEvidence)
    if (![string]::IsNullOrWhiteSpace($evidenceDirectory)) {
        [IO.Directory]::CreateDirectory($evidenceDirectory) | Out-Null
    }
    [IO.File]::WriteAllText($resolvedEvidence, $json + [Environment]::NewLine, [Text.UTF8Encoding]::new($false))
}
$json
