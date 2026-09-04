param(
    [Parameter(Mandatory = $true)]
    [string]$InstallerPath
)

$ErrorActionPreference = 'Stop'
$resolvedInstaller = (Resolve-Path -LiteralPath $InstallerPath).Path
$signature = Get-AuthenticodeSignature -LiteralPath $resolvedInstaller
if ($signature.Status -eq 'NotSigned' -or $signature.Status -eq 'HashMismatch') {
    throw "Installer signature is invalid: $($signature.Status)"
}

Add-Type -AssemblyName System.IO.Compression
$assembly = [Reflection.Assembly]::LoadFile($resolvedInstaller)
$payloadResource = 'ChuanHoa.DevelopmentInstaller.Payload.zip'
$payloadStream = $assembly.GetManifestResourceStream($payloadResource)
if ($null -eq $payloadStream) {
    throw "Installer resource is missing: $payloadResource"
}

$expectedEntries = @(
    'ChuanHoa.AddIn.Vsto.vsto',
    'ChuanHoa.AddIn.Vsto.dll.manifest',
    'ChuanHoa.AddIn.Vsto.dll',
    'ChuanHoa.Client.Core.dll',
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
    'InsertQrCode'
)
$entryNames = @()
$forbiddenHits = @()

try {
    $archive = New-Object System.IO.Compression.ZipArchive(
        $payloadStream,
        [System.IO.Compression.ZipArchiveMode]::Read,
        $true)
    try {
        foreach ($entry in $archive.Entries) {
            if ([string]::IsNullOrEmpty($entry.Name)) { continue }
            $normalizedName = $entry.FullName.Replace('\', '/')
            $entryNames += $normalizedName
            $entryStream = $entry.Open()
            try {
                $memory = New-Object System.IO.MemoryStream
                try {
                    $entryStream.CopyTo($memory)
                    $bytes = $memory.ToArray()
                }
                finally {
                    $memory.Dispose()
                }
            }
            finally {
                $entryStream.Dispose()
            }

            $ascii = [Text.Encoding]::ASCII.GetString($bytes)
            $unicode = [Text.Encoding]::Unicode.GetString($bytes)
            foreach ($pattern in $forbiddenPatterns) {
                if ($normalizedName.IndexOf($pattern, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $forbiddenHits += "$normalizedName (name: $pattern)"
                }
                if ($ascii.IndexOf($pattern, [StringComparison]::OrdinalIgnoreCase) -ge 0 -or
                    $unicode.IndexOf($pattern, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $forbiddenHits += "$normalizedName (content: $pattern)"
                }
            }
        }
    }
    finally {
        $archive.Dispose()
    }
}
finally {
    $payloadStream.Dispose()
}

$missing = @($expectedEntries | Where-Object { $entryNames -notcontains $_ })
$unexpected = @($entryNames | Where-Object { $expectedEntries -notcontains $_ })
if ($missing.Count -gt 0 -or $unexpected.Count -gt 0) {
    throw "Installer payload does not match allowlist. Missing=[$($missing -join ', ')]; Unexpected=[$($unexpected -join ', ')]"
}
if ($forbiddenHits.Count -gt 0) {
    throw "Installer contains prohibited AI/model markers: $($forbiddenHits -join '; ')"
}

[pscustomobject]@{
    Status = 'PASS'
    InstallerPath = $resolvedInstaller
    SizeBytes = (Get-Item -LiteralPath $resolvedInstaller).Length
    Sha256 = (Get-FileHash -LiteralPath $resolvedInstaller -Algorithm SHA256).Hash
    SignatureStatus = [string]$signature.Status
    Signer = $signature.SignerCertificate.Subject
    PayloadFiles = $entryNames.Count
    ForbiddenHits = $forbiddenHits.Count
} | ConvertTo-Json
