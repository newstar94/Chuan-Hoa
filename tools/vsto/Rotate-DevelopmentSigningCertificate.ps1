# Run with Windows PowerShell 5.1 (PKI module).
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Security
$repository = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$backup = Join-Path $repository ('.dev-secrets\certificate-backup-' + [DateTime]::UtcNow.ToString('yyyyMMddTHHmmssZ'))
New-Item -ItemType Directory -Path $backup | Out-Null
$subjects = @('CN=Chuan Hoa Local Development', 'CN=Chuan Hoa Local Development Root',
    'CN=Chuan Hoa Local Development Leaf Test')
$old = @()
foreach ($store in @('My', 'Root', 'TrustedPublisher')) {
    foreach ($certificate in Get-ChildItem "Cert:\CurrentUser\$store") {
        if ($subjects -notcontains $certificate.Subject) { continue }
        $old += [pscustomobject]@{ Store=$store; Thumbprint=$certificate.Thumbprint; Subject=$certificate.Subject }
        Export-Certificate -Cert $certificate -FilePath (Join-Path $backup "$store-$($certificate.Thumbprint).cer") | Out-Null
        if ($store -eq 'My' -and $certificate.HasPrivateKey) {
            $random = New-Object byte[] 48
            $rng = [Security.Cryptography.RandomNumberGenerator]::Create()
            try { $rng.GetBytes($random) } finally { $rng.Dispose() }
            $password = [Convert]::ToBase64String($random)
            $secure = ConvertTo-SecureString $password -AsPlainText -Force
            Export-PfxCertificate -Cert $certificate -Password $secure -FilePath (Join-Path $backup "$($certificate.Thumbprint).pfx") | Out-Null
            $protected = [Security.Cryptography.ProtectedData]::Protect([Text.Encoding]::UTF8.GetBytes($password), $null,
                [Security.Cryptography.DataProtectionScope]::CurrentUser)
            [IO.File]::WriteAllBytes((Join-Path $backup "$($certificate.Thumbprint).password.dpapi"), $protected)
            $password = $null
        }
    }
}
$old | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $backup 'stores.json') -Encoding UTF8

# First create and validate the replacement. Do not remove usable identities if
# creation fails. Root is a CA; leaf explicitly carries the Code Signing EKU.
$notBefore = (Get-Date).AddMinutes(-10)
$root = New-SelfSignedCertificate -Type Custom -Subject 'CN=Chuan Hoa Local Development Root' `
    -CertStoreLocation Cert:\CurrentUser\My -KeyAlgorithm RSA -KeyLength 3072 -HashAlgorithm SHA256 `
    -KeyExportPolicy Exportable -KeyUsage CertSign,CRLSign -NotBefore $notBefore -NotAfter (Get-Date).AddYears(3) `
    -TextExtension @('2.5.29.19={critical}{text}ca=1&pathlength=0')
$leaf = New-SelfSignedCertificate -Type CodeSigningCert -Subject 'CN=Chuan Hoa Local Development' `
    -Signer $root -CertStoreLocation Cert:\CurrentUser\My -KeyAlgorithm RSA -KeyLength 3072 `
    -HashAlgorithm SHA256 -KeyExportPolicy Exportable -KeyUsage DigitalSignature `
    -NotBefore $notBefore -NotAfter (Get-Date).AddYears(1)
if (!$leaf.HasPrivateKey -or @($leaf.EnhancedKeyUsageList | Where-Object { $_.ObjectId -eq '1.3.6.1.5.5.7.3.3' }).Count -ne 1) {
    throw 'New certificate does not carry Code Signing EKU and private key.'
}
$rootPath = Join-Path $backup 'replacement-root.cer'
$leafPath = Join-Path $backup 'replacement-leaf.cer'
Export-Certificate -Cert $root -FilePath $rootPath | Out-Null
Export-Certificate -Cert $leaf -FilePath $leafPath | Out-Null
Import-Certificate -FilePath $rootPath -CertStoreLocation Cert:\CurrentUser\Root | Out-Null
Import-Certificate -FilePath $leafPath -CertStoreLocation Cert:\CurrentUser\TrustedPublisher | Out-Null
$probe = Join-Path $backup 'signature-probe.exe'
Copy-Item -LiteralPath (Join-Path $repository 'tools\vsto\development-access-smoke\bin\Development\ChuanHoa.DevelopmentAccessSmoke.exe') -Destination $probe
$signature = Set-AuthenticodeSignature -LiteralPath $probe -Certificate $leaf -HashAlgorithm SHA256
if ($signature.Status -ne 'Valid') { throw "Replacement signature validation failed: $($signature.Status)" }
foreach ($item in $old) {
    $path = "Cert:\CurrentUser\$($item.Store)\$($item.Thumbprint)"
    $certificate = Get-Item -LiteralPath $path
    if ($subjects -notcontains $certificate.Subject) { throw 'Certificate target changed during rotation.' }
    Remove-Item -LiteralPath $path
}
function Get-Pin($certificate) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($sha.ComputeHash($certificate.RawData)).Replace('-','') }
    finally { $sha.Dispose() }
}
$result = [pscustomobject]@{ SigningCertificateSha256=(Get-Pin $leaf); SigningRootCertificateSha256=(Get-Pin $root);
    LeafThumbprint=$leaf.Thumbprint; RootThumbprint=$root.Thumbprint; SignatureProbe='Valid'; Removed=$old.Count; Backup=$backup }
$result | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $repository '.dev-secrets\development-signing-pins.json') -Encoding UTF8
$result | ConvertTo-Json
