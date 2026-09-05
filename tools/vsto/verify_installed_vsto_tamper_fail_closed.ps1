param(
    [Parameter(Mandatory = $false)]
    [string]$InstallerPath = '',

    [Parameter(Mandatory = $false)]
    [string]$EvidencePath = ''
)

$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$props = [xml](Get-Content -LiteralPath `
    (Join-Path $root 'Directory.Build.props') -Raw)
$version = [string]$props.Project.PropertyGroup.ProductVersion
if ($version -notmatch '^\d+\.\d+\.\d+\.\d+$') {
    throw 'Directory.Build.props does not contain a valid ProductVersion.'
}
if ([string]::IsNullOrWhiteSpace($InstallerPath)) {
    $InstallerPath = Join-Path $root (
        'artifacts\installers\development\ChuanHoa_Development_Test_Setup_' +
        $version + '.exe')
}
$installer = (Resolve-Path -LiteralPath $InstallerPath).Path
$currentDirectory = Join-Path $env:LOCALAPPDATA `
    'ChuanHoa\DevelopmentInstaller\Current'
$addInPath = Join-Path $currentDirectory 'ChuanHoa.AddIn.Vsto.dll'
$registryPath = 'HKCU:\Software\Microsoft\Office\Word\Addins\ChuanHoa.AddIn.Vsto'
$passiveSmoke = Join-Path $root `
    'tools\vsto\passive-startup-smoke\bin\Development\ChuanHoa.PassiveStartupSmoke.exe'
$ribbonSmoke = Join-Path $root `
    'tools\vsto\ribbon-capability-smoke\bin\Development\ChuanHoa.RibbonCapabilitySmoke.exe'

if ((Get-Process WINWORD -ErrorAction SilentlyContinue).Count -ne 0) {
    throw 'Microsoft Word must be closed before installed tamper verification.'
}
foreach ($path in @($installer, $addInPath, $passiveSmoke, $ribbonSmoke)) {
    if (!(Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Required file is missing: $path"
    }
}
$installed = Get-ItemProperty -LiteralPath $registryPath
if ([int]$installed.LoadBehavior -ne 3 -or
    [string]$installed.Manifest -notlike '*DevelopmentInstaller/Current/ChuanHoa.AddIn.Vsto.vsto|vstolocal') {
    throw 'Installed Word add-in registration is not the expected Development channel.'
}

function Get-Inventory([string]$directory) {
    return @(
        Get-ChildItem -LiteralPath $directory -File -Recurse |
            Sort-Object FullName |
            ForEach-Object {
                [ordered]@{
                    RelativePath = ($_.FullName.Substring($directory.Length)).TrimStart('\').Replace('\', '/')
                    Length = $_.Length
                    Sha256 = (Get-FileHash -LiteralPath $_.FullName `
                        -Algorithm SHA256).Hash
                }
            }
    )
}

function Invoke-ProcessWithTimeout {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $false)][string]$Arguments = '',
        [Parameter(Mandatory = $true)][int]$TimeoutMilliseconds
    )
    $start = [Diagnostics.ProcessStartInfo]::new()
    $start.FileName = $Path
    $start.Arguments = $Arguments
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.WindowStyle = [Diagnostics.ProcessWindowStyle]::Hidden
    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $start
    try {
        if (!$process.Start()) { throw "Could not start: $Path" }
        if (!$process.WaitForExit($TimeoutMilliseconds)) {
            try { $process.Kill() } catch { }
            throw "Process timed out: $Path"
        }
        return $process.ExitCode
    }
    finally { $process.Dispose() }
}

function Stop-TestWordProcesses {
    Get-Process WINWORD -ErrorAction SilentlyContinue |
        ForEach-Object {
            try { $_.Kill(); $_.WaitForExit(5000) } catch { }
            finally { $_.Dispose() }
        }
}

$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) `
    ('ChuanHoa-Installed-Tamper-' + [Guid]::NewGuid().ToString('N'))
$backupDirectory = Join-Path $temporaryRoot 'Current'
$baselineInventory = Get-Inventory $currentDirectory
$baselineManifest = [string]$installed.Manifest
$baselineLoadBehavior = [int]$installed.LoadBehavior
$tamperedSignatureStatus = $null
$rejectionExitCode = $null
$repairExitCode = $null
$postRepairPassiveExitCode = $null
$postRepairRibbonExitCode = $null
$repairSucceeded = $false

try {
    New-Item -ItemType Directory -Path $temporaryRoot | Out-Null
    Copy-Item -LiteralPath $currentDirectory -Destination $backupDirectory `
        -Recurse -Force

    $bytes = [IO.File]::ReadAllBytes($addInPath)
    if ($bytes.Length -lt 8192) { throw 'Installed add-in DLL is unexpectedly small.' }
    $bytes[4096] = $bytes[4096] -bxor 0x01
    [IO.File]::WriteAllBytes($addInPath, $bytes)
    $tamperedSignatureStatus = [string](
        Get-AuthenticodeSignature -LiteralPath $addInPath).Status
    if ($tamperedSignatureStatus -ne 'HashMismatch') {
        throw "Tampered DLL did not produce HashMismatch: $tamperedSignatureStatus"
    }

    $rejectionExitCode = Invoke-ProcessWithTimeout `
        -Path $passiveSmoke -TimeoutMilliseconds 45000
    Stop-TestWordProcesses
    if ($rejectionExitCode -eq 0) {
        throw 'Word still connected the add-in after its signed DLL was modified.'
    }
}
finally {
    Stop-TestWordProcesses
    try {
        $repairExitCode = Invoke-ProcessWithTimeout `
            -Path $installer -Arguments '/repair /quiet' `
            -TimeoutMilliseconds 60000
        $repairSucceeded = $repairExitCode -eq 0
    }
    catch {
        $repairSucceeded = $false
    }
    if (!$repairSucceeded) {
        if (!(Test-Path -LiteralPath $backupDirectory -PathType Container)) {
            throw 'Repair failed and the verified Temp backup is unavailable.'
        }
        Get-ChildItem -LiteralPath $backupDirectory -File -Recurse |
            ForEach-Object {
                $relative = ($_.FullName.Substring($backupDirectory.Length)).TrimStart('\')
                $target = Join-Path $currentDirectory $relative
                $parent = Split-Path -Parent $target
                New-Item -ItemType Directory -Path $parent -Force | Out-Null
                Copy-Item -LiteralPath $_.FullName -Destination $target -Force
            }
        Set-ItemProperty -LiteralPath $registryPath -Name Manifest `
            -Value $baselineManifest
        Set-ItemProperty -LiteralPath $registryPath -Name LoadBehavior `
            -Value $baselineLoadBehavior -Type DWord
        throw 'Installer repair failed; the original installed payload was restored from Temp.'
    }
}

$postInventory = Get-Inventory $currentDirectory
if (($baselineInventory | ConvertTo-Json -Depth 4 -Compress) -ne
    ($postInventory | ConvertTo-Json -Depth 4 -Compress)) {
    throw 'Installed payload hashes did not return to the exact baseline after repair.'
}
$postRegistration = Get-ItemProperty -LiteralPath $registryPath
if ([int]$postRegistration.LoadBehavior -ne 3 -or
    ![string]::Equals([string]$postRegistration.Manifest, $baselineManifest,
        [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Word add-in registration was not restored after repair.'
}
$postRepairPassiveExitCode = Invoke-ProcessWithTimeout `
    -Path $passiveSmoke -TimeoutMilliseconds 45000
if ($postRepairPassiveExitCode -ne 0) {
    throw 'Passive startup smoke did not recover after installer repair.'
}
$postRepairRibbonExitCode = Invoke-ProcessWithTimeout `
    -Path $ribbonSmoke -TimeoutMilliseconds 45000
if ($postRepairRibbonExitCode -ne 0) {
    throw 'Ribbon capability smoke did not recover after installer repair.'
}

$evidence = [ordered]@{
    SchemaVersion = 1
    Status = 'PASS_INSTALLED_DLL_TAMPER_FAIL_CLOSED_AND_REPAIRED'
    ProductVersion = $version
    InstallerFile = [IO.Path]::GetFileName($installer)
    InstallerSha256 = (Get-FileHash -LiteralPath $installer `
        -Algorithm SHA256).Hash
    TamperedTarget = 'ChuanHoa.AddIn.Vsto.dll'
    TamperedAuthenticodeStatus = $tamperedSignatureStatus
    WordRejectedTamperedAddIn = $true
    RejectionSmokeExitCode = $rejectionExitCode
    RepairExitCode = $repairExitCode
    ExactPayloadHashRestored = $true
    LoadBehavior = [int]$postRegistration.LoadBehavior
    PostRepairPassiveStartupExitCode = $postRepairPassiveExitCode
    PostRepairRibbonCapabilityExitCode = $postRepairRibbonExitCode
    UserDocumentsTouched = $false
}
if (![string]::IsNullOrWhiteSpace($EvidencePath)) {
    $resolvedEvidence = if ([IO.Path]::IsPathRooted($EvidencePath)) {
        $EvidencePath
    } else {
        Join-Path $root $EvidencePath
    }
    New-Item -ItemType Directory -Path (Split-Path -Parent $resolvedEvidence) `
        -Force | Out-Null
    $evidence | ConvertTo-Json -Depth 5 |
        Set-Content -LiteralPath $resolvedEvidence -Encoding UTF8
}
$evidence | ConvertTo-Json -Depth 5

$normalizedTemp = ([IO.Path]::GetFullPath([IO.Path]::GetTempPath())).TrimEnd('\') + '\'
$normalizedTarget = ([IO.Path]::GetFullPath($temporaryRoot)).TrimEnd('\') + '\'
if ($normalizedTarget.StartsWith($normalizedTemp,
        [StringComparison]::OrdinalIgnoreCase) -and
    (Split-Path -Leaf $temporaryRoot).StartsWith(
        'ChuanHoa-Installed-Tamper-', [StringComparison]::Ordinal) -and
    (Test-Path -LiteralPath $temporaryRoot)) {
    Remove-Item -LiteralPath $temporaryRoot -Recurse -Force
}
