param(
    [Parameter(Mandatory = $true)]
    [string]$InstallerPath,

    [Parameter(Mandatory = $false)]
    [string]$EvidencePath = ''
)

$ErrorActionPreference = 'Stop'
$resolvedInstaller = (Resolve-Path -LiteralPath $InstallerPath).Path
$signature = Get-AuthenticodeSignature -LiteralPath $resolvedInstaller
if ($signature.Status -ne 'Valid') {
    throw "Baseline installer Authenticode must be Valid, actual=$($signature.Status)."
}

Add-Type -AssemblyName System.IO.Compression
$installerBytes = [IO.File]::ReadAllBytes($resolvedInstaller)
$assembly = [Reflection.Assembly]::Load($installerBytes)
$programType = $assembly.GetType(
    'ChuanHoa.DevelopmentTestBootstrapper.Program', $true)
$bindingFlags = [Reflection.BindingFlags]::Static -bor `
    [Reflection.BindingFlags]::NonPublic
$verifyPayload = $programType.GetMethod('VerifyStagedPayload', $bindingFlags)
$verifySigner = $programType.GetMethod('VerifyInstallerSigner', $bindingFlags)
if ($null -eq $verifyPayload -or $null -eq $verifySigner) {
    throw 'Installer does not expose the expected private integrity gates.'
}

function Read-TextResource {
    param([Parameter(Mandatory = $true)][string]$Name)
    $stream = $assembly.GetManifestResourceStream($Name)
    if ($null -eq $stream) { throw "Missing installer resource: $Name" }
    try {
        $reader = [IO.StreamReader]::new($stream)
        try { return $reader.ReadToEnd().Trim() }
        finally { $reader.Dispose() }
    }
    finally { $stream.Dispose() }
}

function Invoke-PrivateGate {
    param(
        [Parameter(Mandatory = $true)]
        [Reflection.MethodInfo]$Method,
        [Parameter(Mandatory = $true)]
        [object[]]$Arguments
    )
    try {
        [object[]]$rawArguments = [object[]]::new($Arguments.Length)
        for ($index = 0; $index -lt $Arguments.Length; $index++) {
            $rawArguments[$index] = $Arguments[$index].PSObject.BaseObject
        }
        $null = $Method.Invoke($null, $rawArguments)
        return $null
    }
    catch {
        $exception = $_.Exception
        while ($null -ne $exception.InnerException) {
            $exception = $exception.InnerException
        }
        return $exception.Message
    }
}

function Flip-Byte {
    param([Parameter(Mandatory = $true)][string]$Path)
    $bytes = [IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -lt 512) { throw "File is too small for tamper test: $Path" }
    $offset = [Math]::Min(4096, $bytes.Length - 257)
    $bytes[$offset] = $bytes[$offset] -bxor 0x01
    [IO.File]::WriteAllBytes($Path, $bytes)
}

function Tamper-ManifestVersion {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Version
    )
    $text = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
    $needle = 'version="' + $Version + '"'
    $offset = $text.IndexOf($needle, [StringComparison]::Ordinal)
    if ($offset -lt 0) { throw "Manifest version marker was not found: $Path" }
    $replacementVersion = $Version.Substring(0, $Version.Length - 1) +
        $(if ($Version.EndsWith('9')) { '8' } else { '9' })
    $replacement = 'version="' + $replacementVersion + '"'
    $text = $text.Remove($offset, $needle.Length).Insert($offset, $replacement)
    [IO.File]::WriteAllText($Path, $text, [Text.UTF8Encoding]::new($false))
}

$version = Read-TextResource `
    -Name 'ChuanHoa.DevelopmentInstaller.Version.txt'
$trustedKeyPin = Read-TextResource `
    -Name 'ChuanHoa.DevelopmentInstaller.TrustedPublicKey.sha256'
$signingPin = Read-TextResource `
    -Name 'ChuanHoa.DevelopmentInstaller.SigningCertificate.sha256'
$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) `
    ('ChuanHoa-Tamper-Rejection-' + [Guid]::NewGuid().ToString('N'))
$payloadDirectory = Join-Path $temporaryRoot 'Payload'
$outerCopy = Join-Path $temporaryRoot 'Installer.exe'
$results = @()

try {
    New-Item -ItemType Directory -Path $payloadDirectory -Force | Out-Null
    [IO.File]::WriteAllBytes($outerCopy, $installerBytes)
    $payloadStream = $assembly.GetManifestResourceStream(
        'ChuanHoa.DevelopmentInstaller.Payload.zip')
    if ($null -eq $payloadStream) { throw 'Missing embedded installer payload.' }
    try {
        $archive = [IO.Compression.ZipArchive]::new(
            $payloadStream, [IO.Compression.ZipArchiveMode]::Read, $true)
        try {
            $normalizedPayload = ([IO.Path]::GetFullPath($payloadDirectory)).TrimEnd('\') + '\'
            foreach ($entry in $archive.Entries) {
                if ([string]::IsNullOrEmpty($entry.Name)) { continue }
                $destination = [IO.Path]::GetFullPath(
                    (Join-Path $payloadDirectory $entry.FullName))
                if (!$destination.StartsWith($normalizedPayload,
                        [StringComparison]::OrdinalIgnoreCase)) {
                    throw "Unsafe embedded path: $($entry.FullName)"
                }
                $parent = Split-Path -Parent $destination
                New-Item -ItemType Directory -Path $parent -Force | Out-Null
                $input = $entry.Open()
                try {
                    $output = [IO.File]::Create($destination)
                    try { $input.CopyTo($output) }
                    finally { $output.Dispose() }
                }
                finally { $input.Dispose() }
            }
        }
        finally { $archive.Dispose() }
    }
    finally { $payloadStream.Dispose() }

    [object[]]$payloadArguments = @(
        $payloadDirectory, $version, $trustedKeyPin, $signingPin)
    [object[]]$signerArguments = @($resolvedInstaller, $signingPin)
    $baselinePayloadError = Invoke-PrivateGate `
        -Method $verifyPayload -Arguments $payloadArguments
    $baselineSignerError = Invoke-PrivateGate `
        -Method $verifySigner -Arguments $signerArguments
    if ($null -ne $baselinePayloadError -or $null -ne $baselineSignerError) {
        throw "Baseline integrity gate failed. Payload=$baselinePayloadError Signer=$baselineSignerError"
    }

    $cases = @(
        [ordered]@{
            target = 'outer-installer'
            path = $outerCopy
            gate = $verifySigner
            expectedMessage = '0x80096010'
        },
        [ordered]@{
            target = 'ChuanHoa.AddIn.Vsto.dll'
            path = (Join-Path $payloadDirectory 'ChuanHoa.AddIn.Vsto.dll')
            gate = $verifyPayload
            expectedMessage = '0x80096010'
        },
        [ordered]@{
            target = 'ChuanHoa.Client.Core.dll'
            path = (Join-Path $payloadDirectory 'ChuanHoa.Client.Core.dll')
            gate = $verifyPayload
            expectedMessage = '0x80096010'
        },
        [ordered]@{
            target = 'ChuanHoa.DevelopmentAccessSmoke.exe'
            path = (Join-Path $payloadDirectory 'ChuanHoa.DevelopmentAccessSmoke.exe')
            gate = $verifyPayload
            expectedMessage = '0x80096010'
        },
        [ordered]@{
            target = 'Microsoft.Office.Tools.Common.v4.0.Utilities.dll'
            path = (Join-Path $payloadDirectory 'Microsoft.Office.Tools.Common.v4.0.Utilities.dll')
            gate = $verifyPayload
            expectedMessage = 'manifest'
        },
        [ordered]@{
            target = 'ChuanHoa.AddIn.Vsto.vsto'
            path = (Join-Path $payloadDirectory 'ChuanHoa.AddIn.Vsto.vsto')
            gate = $verifyPayload
            expectedMessage = 'manifest'
        },
        [ordered]@{
            target = 'ChuanHoa.AddIn.Vsto.dll.manifest'
            path = (Join-Path $payloadDirectory 'ChuanHoa.AddIn.Vsto.dll.manifest')
            gate = $verifyPayload
            expectedMessage = 'manifest'
        }
    )

    foreach ($case in $cases) {
        $original = [IO.File]::ReadAllBytes($case.path)
        try {
            if ($case.target -eq 'ChuanHoa.AddIn.Vsto.vsto' -or
                $case.target -eq 'ChuanHoa.AddIn.Vsto.dll.manifest') {
                Tamper-ManifestVersion -Path $case.path -Version $version
            }
            else {
                Flip-Byte -Path $case.path
            }
            [object[]]$arguments = if ($case.target -eq 'outer-installer') {
                @($case.path, $signingPin)
            } else {
                $payloadArguments
            }
            $message = Invoke-PrivateGate -Method $case.gate -Arguments $arguments
            if ([string]::IsNullOrWhiteSpace($message) -or
                $message.IndexOf($case.expectedMessage,
                    [StringComparison]::OrdinalIgnoreCase) -lt 0) {
                throw "Tamper was not rejected at the expected gate: $($case.target). Actual=$message"
            }
            $results += [ordered]@{
                target = $case.target
                status = 'PASS_REJECTED'
                rejection = $message
            }
        }
        finally {
            [IO.File]::WriteAllBytes($case.path, $original)
        }

        if ($case.target -ne 'outer-installer') {
            $restoredError = Invoke-PrivateGate `
                -Method $verifyPayload -Arguments $payloadArguments
            if ($null -ne $restoredError) {
                throw "Restored payload did not return to baseline: $restoredError"
            }
        }
    }

    $evidence = [ordered]@{
        schemaVersion = 1
        status = 'PASS_TAMPER_REJECTED_ON_TEMP_COPY'
        productVersion = $version
        installerFile = [IO.Path]::GetFileName($resolvedInstaller)
        installerSha256 = (Get-FileHash -LiteralPath $resolvedInstaller `
            -Algorithm SHA256).Hash
        baseline = 'PASS'
        cases = $results
        installedStateTouched = $false
    }
    if (![string]::IsNullOrWhiteSpace($EvidencePath)) {
        $resolvedEvidence = if ([IO.Path]::IsPathRooted($EvidencePath)) {
            $EvidencePath
        } else {
            Join-Path (Get-Location).Path $EvidencePath
        }
        $evidenceParent = Split-Path -Parent $resolvedEvidence
        New-Item -ItemType Directory -Path $evidenceParent -Force | Out-Null
        $evidence | ConvertTo-Json -Depth 6 |
            Set-Content -LiteralPath $resolvedEvidence -Encoding UTF8
    }
    $evidence | ConvertTo-Json -Depth 6
}
finally {
    $normalizedTemp = ([IO.Path]::GetFullPath([IO.Path]::GetTempPath())).TrimEnd('\') + '\'
    $normalizedTarget = ([IO.Path]::GetFullPath($temporaryRoot)).TrimEnd('\') + '\'
    if ($normalizedTarget.StartsWith($normalizedTemp,
            [StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path -Leaf $temporaryRoot).StartsWith(
            'ChuanHoa-Tamper-Rejection-', [StringComparison]::Ordinal) -and
        (Test-Path -LiteralPath $temporaryRoot)) {
        Remove-Item -LiteralPath $temporaryRoot -Recurse -Force
    }
}
