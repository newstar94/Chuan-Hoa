param(
    [Parameter(Mandatory = $true)]
    [string]$InstallerPath,

    [Parameter(Mandatory = $false)]
    [ValidatePattern('^$|^\d+\.\d+\.\d+\.\d+$')]
    [string]$ExpectedProductVersion = '',

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-fA-F -]{64,95}$')]
    [string]$TrustedPublicKeySha256,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^[0-9a-fA-F -]{64,95}$')]
    [string]$SigningCertificateSha256
)

$ErrorActionPreference = 'Stop'
$root = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
. (Join-Path $root 'tools\vsto\BuildContract.ps1')
$expectedVersion = Resolve-ChuanHoaApplicationVersion `
    -RepositoryRoot $root -RequestedVersion $ExpectedProductVersion
$trustedKeyPin = Normalize-ChuanHoaSha256 `
    -Value $TrustedPublicKeySha256 -Label 'TrustedPublicKeySha256'
$certificatePin = Normalize-ChuanHoaSha256 `
    -Value $SigningCertificateSha256 -Label 'SigningCertificateSha256'

$resolvedInstaller = (Resolve-Path -LiteralPath $InstallerPath).Path
Assert-ChuanHoaManagedAssemblyVersion `
    -Path $resolvedInstaller -ExpectedVersion $expectedVersion
$signature = Get-AuthenticodeSignature -LiteralPath $resolvedInstaller
if ($signature.Status -eq 'NotSigned' -or $signature.Status -eq 'HashMismatch' -or
    $null -eq $signature.SignerCertificate) {
    throw "Installer signature is invalid: $($signature.Status)"
}
$outerSignerPin = Get-ChuanHoaCertificateSha256 `
    -Certificate $signature.SignerCertificate
if (![string]::Equals($outerSignerPin, $certificatePin,
        [System.StringComparison]::Ordinal)) {
    throw "Installer signer does not match the approved certificate SHA-256. Expected=$certificatePin Actual=$outerSignerPin"
}

Add-Type -AssemblyName System.IO.Compression
# PowerShell 7 can reject a signed .NET Framework executable through LoadFile
# with BadImageFormatException even though the CLR can inspect the exact same
# PE image. Loading an immutable byte copy avoids path/load-context probing and
# does not execute the installer's entry point.
$installerBytes = [IO.File]::ReadAllBytes($resolvedInstaller)
$assembly = [Reflection.Assembly]::Load($installerBytes)
function Read-InstallerTextResource {
    param([Parameter(Mandatory = $true)][string]$Name)
    $stream = $assembly.GetManifestResourceStream($Name)
    if ($null -eq $stream) { throw "Installer resource is missing: $Name" }
    try {
        $reader = New-Object System.IO.StreamReader($stream)
        try { return $reader.ReadToEnd().Trim() }
        finally { $reader.Dispose() }
    }
    finally { $stream.Dispose() }
}

$payloadResource = 'ChuanHoa.DevelopmentInstaller.Payload.zip'
$versionResource = 'ChuanHoa.DevelopmentInstaller.Version.txt'
$trustedKeyPinResource = 'ChuanHoa.DevelopmentInstaller.TrustedPublicKey.sha256'
$certificatePinResource = 'ChuanHoa.DevelopmentInstaller.SigningCertificate.sha256'

$embeddedVersion = Read-InstallerTextResource -Name $versionResource
if (![string]::Equals($embeddedVersion, $expectedVersion,
        [System.StringComparison]::Ordinal)) {
    throw "Installer version resource differs from ProductVersion. ProductVersion=$expectedVersion Embedded=$embeddedVersion"
}
$embeddedTrustedKeyPin = Normalize-ChuanHoaSha256 `
    -Value (Read-InstallerTextResource -Name $trustedKeyPinResource) `
    -Label 'EmbeddedTrustedPublicKeySha256'
$embeddedCertificatePin = Normalize-ChuanHoaSha256 `
    -Value (Read-InstallerTextResource -Name $certificatePinResource) `
    -Label 'EmbeddedSigningCertificateSha256'
if (![string]::Equals($embeddedTrustedKeyPin, $trustedKeyPin,
        [System.StringComparison]::Ordinal) -or
    ![string]::Equals($embeddedCertificatePin, $certificatePin,
        [System.StringComparison]::Ordinal)) {
    throw 'Installer embedded pins do not match the external build/audit contract.'
}

$payloadStream = $assembly.GetManifestResourceStream($payloadResource)
if ($null -eq $payloadStream) {
    throw "Installer resource is missing: $payloadResource"
}

$expectedEntries = @(
    'ChuanHoa.AddIn.Vsto.vsto',
    'ChuanHoa.AddIn.Vsto.dll.manifest',
    'ChuanHoa.AddIn.Vsto.dll',
    'ChuanHoa.Client.Core.dll',
    'ChuanHoa.DevelopmentAccessSmoke.exe',
    'Microsoft.Office.Tools.Common.v4.0.Utilities.dll',
    'ChuanHoa.LocalDevelopment.Public.cer',
    'DevelopmentSupport/trusted-key.xml'
)
$forbiddenPatterns = @(
    'VietnameseEngine',
    'LOCAL-TYPO-AI',
    '.onnx',
    '.onnx.data',
    'onnxruntime',
    'model_manifest',
    'training',
    'QRCoder',
    'btnChenQrCode',
    'QrCodeInputDialog',
    'InsertQrCode',
    '.pfx',
    '.p12',
    'BEGIN PRIVATE KEY',
    'BEGIN RSA PRIVATE KEY'
)
$entryNames = @()
$forbiddenHits = @()
$entryBytes = @{}
$payloadEvidence = @()

try {
    $archive = New-Object System.IO.Compression.ZipArchive(
        $payloadStream,
        [System.IO.Compression.ZipArchiveMode]::Read,
        $true)
    try {
        foreach ($entry in $archive.Entries) {
            if ([string]::IsNullOrEmpty($entry.Name)) { continue }
            $normalizedName = $entry.FullName.Replace('\', '/')
            if ($normalizedName.StartsWith('/') -or $normalizedName.Contains('../') -or
                $normalizedName.Contains(':')) {
                throw "Installer payload contains an unsafe path: $normalizedName"
            }
            if ($entryBytes.ContainsKey($normalizedName)) {
                throw "Installer payload contains a duplicate path: $normalizedName"
            }
            $entryNames += $normalizedName
            $entryStream = $entry.Open()
            try {
                $memory = New-Object System.IO.MemoryStream
                try {
                    $entryStream.CopyTo($memory)
                    $bytes = $memory.ToArray()
                }
                finally { $memory.Dispose() }
            }
            finally { $entryStream.Dispose() }
            $entryBytes[$normalizedName] = $bytes

            $memoryHash = [System.Security.Cryptography.SHA256]::Create()
            try {
                $entryHash = [BitConverter]::ToString(
                    $memoryHash.ComputeHash($bytes)).Replace('-', '')
            }
            finally { $memoryHash.Dispose() }
            $payloadEvidence += [ordered]@{
                path = $normalizedName
                sizeBytes = $bytes.Length
                sha256 = $entryHash
            }

            $ascii = [Text.Encoding]::ASCII.GetString($bytes)
            $unicode = [Text.Encoding]::Unicode.GetString($bytes)
            foreach ($pattern in $forbiddenPatterns) {
                if ($normalizedName.IndexOf($pattern,
                        [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $forbiddenHits += "$normalizedName (name: $pattern)"
                }
                if ($ascii.IndexOf($pattern, [StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                    $unicode.IndexOf($pattern, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $forbiddenHits += "$normalizedName (content: $pattern)"
                }
            }
        }
    }
    finally { $archive.Dispose() }
}
finally { $payloadStream.Dispose() }

$missing = @($expectedEntries | Where-Object { $entryNames -notcontains $_ })
$unexpected = @($entryNames | Where-Object { $expectedEntries -notcontains $_ })
if ($missing.Count -gt 0 -or $unexpected.Count -gt 0) {
    throw "Installer payload does not match allowlist. Missing=[$($missing -join ', ')]; Unexpected=[$($unexpected -join ', ')]"
}
if ($forbiddenHits.Count -gt 0) {
    throw "Installer contains prohibited AI/model/private-key markers: $($forbiddenHits -join '; ')"
}

$trustedKeyBytes = [byte[]]$entryBytes['DevelopmentSupport/trusted-key.xml']
$trustedKeyHash = [System.Security.Cryptography.SHA256]::Create()
try {
    $actualTrustedKeyPin = [BitConverter]::ToString(
        $trustedKeyHash.ComputeHash($trustedKeyBytes)).Replace('-', '')
}
finally { $trustedKeyHash.Dispose() }
if (![string]::Equals($actualTrustedKeyPin, $trustedKeyPin,
        [System.StringComparison]::Ordinal)) {
    throw "Payload trusted public key does not match the approved SHA-256. Expected=$trustedKeyPin Actual=$actualTrustedKeyPin"
}
$trustedKeyText = [Text.Encoding]::UTF8.GetString($trustedKeyBytes).TrimStart([char]0xFEFF)
$trustedKeyDocument = [Xml.XmlDocument]::new()
$trustedKeyDocument.PreserveWhitespace = $true
$trustedKeyDocument.LoadXml($trustedKeyText)
$trustedKeyRoot = $trustedKeyDocument.DocumentElement
$rsaKey = if ($null -eq $trustedKeyRoot) { $null } else {
    $trustedKeyRoot.SelectSingleNode('./RSAKeyValue')
}
$rsaElementNames = if ($null -eq $rsaKey) { @() } else {
    @($rsaKey.ChildNodes | Where-Object NodeType -eq ([Xml.XmlNodeType]::Element) |
        ForEach-Object LocalName)
}
if ($null -eq $trustedKeyRoot -or $trustedKeyRoot.LocalName -ne 'trustedDevelopmentKey' -or
    [string]::IsNullOrWhiteSpace($trustedKeyRoot.GetAttribute('keyId')) -or
    $rsaElementNames.Count -ne 2 -or $rsaElementNames -notcontains 'Modulus' -or
    $rsaElementNames -notcontains 'Exponent') {
    throw 'Payload trusted key is not a public-only trustedDevelopmentKey document.'
}

$payloadCertificate = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new(
    [byte[]]$entryBytes['ChuanHoa.LocalDevelopment.Public.cer'])
try {
    $actualCertificatePin = Get-ChuanHoaCertificateSha256 -Certificate $payloadCertificate
}
finally { $payloadCertificate.Dispose() }
if (![string]::Equals($actualCertificatePin, $certificatePin,
        [System.StringComparison]::Ordinal)) {
    throw "Payload certificate does not match the approved SHA-256. Expected=$certificatePin Actual=$actualCertificatePin"
}

function Assert-ManifestBytesVersion {
    param(
        [Parameter(Mandatory = $true)][byte[]]$Bytes,
        [Parameter(Mandatory = $true)][string]$IdentityName
    )
    $manifestText = [Text.Encoding]::UTF8.GetString($Bytes).TrimStart([char]0xFEFF)
    $manifest = [Xml.XmlDocument]::new()
    $manifest.PreserveWhitespace = $true
    $manifest.LoadXml($manifestText)
    $identities = @($manifest.SelectNodes("//*[local-name()='assemblyIdentity']") |
        Where-Object { [string]::Equals($_.GetAttribute('name'), $IdentityName,
            [StringComparison]::Ordinal) })
    if ($identities.Count -eq 0 -or @($identities | Where-Object {
            ![string]::Equals($_.GetAttribute('version'), $expectedVersion,
                [StringComparison]::Ordinal)
        }).Count -ne 0) {
        throw "Manifest identity '$IdentityName' does not match ProductVersion $expectedVersion."
    }
    if ($manifest.SelectSingleNode("//*[local-name()='Signature']") -eq $null) {
        throw "Manifest identity '$IdentityName' has no XML signature."
    }
}
Assert-ManifestBytesVersion `
    -Bytes ([byte[]]$entryBytes['ChuanHoa.AddIn.Vsto.vsto']) `
    -IdentityName 'ChuanHoa.AddIn.Vsto.vsto'
Assert-ManifestBytesVersion `
    -Bytes ([byte[]]$entryBytes['ChuanHoa.AddIn.Vsto.dll.manifest']) `
    -IdentityName 'ChuanHoa.AddIn.Vsto.dll'

$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) `
    ('ChuanHoaInstallerAudit-' + [Guid]::NewGuid().ToString('N'))
$innerSignatures = @()
try {
    New-Item -ItemType Directory -Path $temporaryRoot -Force | Out-Null
    foreach ($managedEntry in @(
        'ChuanHoa.AddIn.Vsto.dll',
        'ChuanHoa.Client.Core.dll',
        'ChuanHoa.DevelopmentAccessSmoke.exe'
    )) {
        $temporaryPath = Join-Path $temporaryRoot ([IO.Path]::GetFileName($managedEntry))
        [IO.File]::WriteAllBytes($temporaryPath, [byte[]]$entryBytes[$managedEntry])
        Assert-ChuanHoaManagedAssemblyVersion `
            -Path $temporaryPath -ExpectedVersion $expectedVersion
        $innerSignature = Get-AuthenticodeSignature -LiteralPath $temporaryPath
        $innerSignatures += [ordered]@{
            path = $managedEntry
            status = [string]$innerSignature.Status
            signer = if ($null -eq $innerSignature.SignerCertificate) {
                $null
            } else {
                $innerSignature.SignerCertificate.Subject
            }
            signerSha256 = if ($null -eq $innerSignature.SignerCertificate) {
                $null
            } else {
                Get-ChuanHoaCertificateSha256 -Certificate $innerSignature.SignerCertificate
            }
            timestamped = $null -ne $innerSignature.TimeStamperCertificate
        }
    }
}
finally {
    $normalizedTemp = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    $normalizedTarget = [IO.Path]::GetFullPath($temporaryRoot).TrimEnd('\') + '\'
    if ($normalizedTarget.StartsWith($normalizedTemp,
            [StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path -Leaf $temporaryRoot).StartsWith(
            'ChuanHoaInstallerAudit-', [StringComparison]::Ordinal) -and
        (Test-Path -LiteralPath $temporaryRoot)) {
        Remove-Item -LiteralPath $temporaryRoot -Recurse -Force
    }
}

$developmentGaps = @()
if (@($innerSignatures | Where-Object {
        $_.status -ne 'Valid' -or
        ![string]::Equals([string]$_.signerSha256, $certificatePin,
            [StringComparison]::Ordinal)
    }).Count -ne 0) {
    $developmentGaps += 'INNER_PE_AUTHENTICODE_NOT_VALID'
}
$developmentGaps += 'SELF_SIGNED_DEVELOPMENT_CERTIFICATE_NOT_PRODUCTION_TRUST'
$developmentCapabilities = @('APPS_FEATURES_AND_CACHED_UNINSTALL_IMPLEMENTED')
if (@($innerSignatures | Where-Object {
        $_.status -ne 'Valid' -or
        ![string]::Equals([string]$_.signerSha256, $certificatePin,
            [StringComparison]::Ordinal)
    }).Count -eq 0) {
    $developmentCapabilities += 'OWNED_INNER_PE_AUTHENTICODE_VALID'
}

[ordered]@{
    schemaVersion = 2
    status = 'PASS_DEVELOPMENT'
    productVersion = $expectedVersion
    installerPath = $resolvedInstaller
    sizeBytes = (Get-Item -LiteralPath $resolvedInstaller).Length
    sha256 = (Get-FileHash -LiteralPath $resolvedInstaller -Algorithm SHA256).Hash
    signatureStatus = [string]$signature.Status
    signerThumbprint = $signature.SignerCertificate.Thumbprint
    signerSha256 = $outerSignerPin
    trustedPublicKeySha256 = $actualTrustedKeyPin
    payloadFiles = $payloadEvidence
    innerPeSignatures = $innerSignatures
    forbiddenHits = $forbiddenHits.Count
    developmentGaps = $developmentGaps
    developmentCapabilities = $developmentCapabilities
    productionReady = $false
} | ConvertTo-Json -Depth 8
