function Get-ChuanHoaProductVersion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot
    )

    $propsPath = Join-Path $RepositoryRoot 'Directory.Build.props'
    if (!(Test-Path -LiteralPath $propsPath -PathType Leaf)) {
        throw "Directory.Build.props was not found: $propsPath"
    }

    [xml]$props = Get-Content -LiteralPath $propsPath -Raw
    $productVersionNodes = @($props.SelectNodes('/Project/PropertyGroup/ProductVersion'))
    if ($productVersionNodes.Count -ne 1) {
        throw "Directory.Build.props must define ProductVersion exactly once; found $($productVersionNodes.Count)."
    }

    $productVersion = ([string]$productVersionNodes[0].InnerText).Trim()
    if ($productVersion -notmatch '^\d+\.\d+\.\d+\.\d+$') {
        throw "ProductVersion must contain exactly four numeric components: $productVersion"
    }

    foreach ($propertyName in @('Version', 'AssemblyVersion', 'FileVersion', 'InformationalVersion')) {
        $nodes = @($props.SelectNodes("/Project/PropertyGroup/$propertyName"))
        if ($nodes.Count -ne 1 -or ([string]$nodes[0].InnerText).Trim() -ne '$(ProductVersion)') {
            throw "$propertyName must be defined exactly once as `$`(ProductVersion`) in Directory.Build.props."
        }
    }
    $revisionSetting = @($props.SelectNodes(
        '/Project/PropertyGroup/IncludeSourceRevisionInInformationalVersion'))
    if ($revisionSetting.Count -ne 1 -or
        ![string]::Equals(([string]$revisionSetting[0].InnerText).Trim(), 'false',
            [System.StringComparison]::OrdinalIgnoreCase)) {
        throw 'IncludeSourceRevisionInInformationalVersion must be false to keep signed assembly ProductVersion exact.'
    }

    return $productVersion
}

function Get-ChuanHoaMsBuild {
    $candidates = [System.Collections.Generic.List[string]]::new()
    $vsWhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
    if (Test-Path -LiteralPath $vsWhere -PathType Leaf) {
        $found = @(& $vsWhere -products * -requires Microsoft.Component.MSBuild `
            -find 'MSBuild\**\Bin\MSBuild.exe' 2>$null)
        foreach ($path in $found) {
            if (![string]::IsNullOrWhiteSpace($path)) { $candidates.Add($path.Trim()) }
        }
    }
    foreach ($path in @(
        'C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe',
        'C:\Program Files\Microsoft Visual Studio\18\BuildTools\MSBuild\Current\Bin\MSBuild.exe',
        'C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe'
    )) {
        if (!$candidates.Contains($path)) { $candidates.Add($path) }
    }

    $available = @($candidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
        Sort-Object { [System.Diagnostics.FileVersionInfo]::GetVersionInfo($_).FileVersion } -Descending)
    if ($available.Count -eq 0) {
        throw 'MSBuild with the Visual Studio Office Developer Tools was not found.'
    }
    return $available[0]
}

function Resolve-ChuanHoaApplicationVersion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepositoryRoot,

        [Parameter(Mandatory = $false)]
        [string]$RequestedVersion = ''
    )

    $productVersion = Get-ChuanHoaProductVersion -RepositoryRoot $RepositoryRoot
    if (![string]::IsNullOrWhiteSpace($RequestedVersion) -and
        ![string]::Equals($RequestedVersion.Trim(), $productVersion,
            [System.StringComparison]::Ordinal)) {
        throw "ApplicationVersion '$RequestedVersion' does not match ProductVersion '$productVersion'. Directory.Build.props is the only version source."
    }
    return $productVersion
}

function Normalize-ChuanHoaSha256 {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string]$Label
    )

    $normalized = $Value.Replace(' ', '').Replace('-', '').Trim().ToUpperInvariant()
    if ($normalized -notmatch '^[0-9A-F]{64}$') {
        throw "$Label must be a 64-character SHA-256 hexadecimal value."
    }
    return $normalized
}

function Get-ChuanHoaCertificateSha256 {
    param(
        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        return [BitConverter]::ToString($sha256.ComputeHash($Certificate.RawData)).Replace('-', '')
    }
    finally {
        $sha256.Dispose()
    }
}

function Get-ChuanHoaSigningCertificate {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExpectedSha256
    )

    $pin = Normalize-ChuanHoaSha256 -Value $ExpectedSha256 -Label 'SigningCertificateSha256'
    $matches = @(
        Get-ChildItem Cert:\CurrentUser\My |
            Where-Object {
                $_.HasPrivateKey -and
                [string]::Equals(
                    (Get-ChuanHoaCertificateSha256 -Certificate $_),
                    $pin,
                    [System.StringComparison]::Ordinal)
            }
    )
    if ($matches.Count -ne 1) {
        throw "Expected exactly one CurrentUser signing certificate pinned by SHA-256 $pin; found $($matches.Count)."
    }
    if ($matches[0].NotBefore.ToUniversalTime() -gt [DateTime]::UtcNow -or
        $matches[0].NotAfter.ToUniversalTime() -le [DateTime]::UtcNow) {
        throw 'The pinned Development signing certificate is not currently valid.'
    }
    return $matches[0]
}

function Get-ChuanHoaSignTool {
    $roots = @(
        (Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10\bin'),
        (Join-Path $env:ProgramFiles 'Windows Kits\10\bin')
    ) | Where-Object { Test-Path -LiteralPath $_ -PathType Container }
    $candidates = @(
        foreach ($root in $roots) {
            Get-ChildItem -LiteralPath $root -Filter signtool.exe -Recurse -File `
                -ErrorAction SilentlyContinue |
                Where-Object FullName -match '\\x64\\signtool\.exe$'
        }
    )
    $selected = $candidates |
        Sort-Object { [System.Diagnostics.FileVersionInfo]::GetVersionInfo($_.FullName).FileVersion } -Descending |
        Select-Object -First 1
    if ($null -eq $selected) { return $null }
    return $selected.FullName
}

function Invoke-ChuanHoaAuthenticodeSign {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedCertificateSha256
    )

    $resolvedPath = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
    $expectedPin = Normalize-ChuanHoaSha256 `
        -Value $ExpectedCertificateSha256 -Label 'ExpectedCertificateSha256'
    $actualPin = Get-ChuanHoaCertificateSha256 -Certificate $Certificate
    if (!$Certificate.HasPrivateKey -or
        ![string]::Equals($actualPin, $expectedPin, [System.StringComparison]::Ordinal)) {
        throw 'Authenticode signing certificate does not match the pinned private-key identity.'
    }

    $signTool = Get-ChuanHoaSignTool
    if ($null -ne $signTool) {
        & $signTool sign /sha1 $Certificate.Thumbprint /fd SHA256 /v $resolvedPath
        if ($LASTEXITCODE -ne 0) {
            throw "Authenticode signing failed for $resolvedPath with exit code $LASTEXITCODE."
        }
    }
    else {
        $signed = Set-AuthenticodeSignature -LiteralPath $resolvedPath `
            -Certificate $Certificate -HashAlgorithm SHA256
        if ($signed.Status -eq 'NotSigned' -or $signed.Status -eq 'HashMismatch') {
            throw "Authenticode signing failed for ${resolvedPath}: $($signed.StatusMessage)"
        }
    }
    $signature = Get-AuthenticodeSignature -LiteralPath $resolvedPath
    if ($signature.Status -eq 'NotSigned' -or $signature.Status -eq 'HashMismatch' -or
        $null -eq $signature.SignerCertificate) {
        throw "Authenticode signature verification failed for ${resolvedPath}: $($signature.Status)."
    }
    $signerPin = Get-ChuanHoaCertificateSha256 -Certificate $signature.SignerCertificate
    if (![string]::Equals($signerPin, $expectedPin, [System.StringComparison]::Ordinal)) {
        throw "Authenticode signer pin mismatch for $resolvedPath."
    }
    return [pscustomobject]@{
        Path = $resolvedPath
        Status = [string]$signature.Status
        SignerSha256 = $signerPin
        Timestamped = $null -ne $signature.TimeStamperCertificate
    }
}

function Assert-ChuanHoaTrustedPublicKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedSha256
    )

    $resolvedPath = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
    if (!(Test-Path -LiteralPath $resolvedPath -PathType Leaf)) {
        throw "The trusted Development public key is not a file: $resolvedPath"
    }
    $pin = Normalize-ChuanHoaSha256 -Value $ExpectedSha256 -Label 'TrustedPublicKeySha256'
    $actualHash = (Get-FileHash -LiteralPath $resolvedPath -Algorithm SHA256).Hash.ToUpperInvariant()
    if (![string]::Equals($actualHash, $pin, [System.StringComparison]::Ordinal)) {
        throw "The trusted Development public key SHA-256 does not match the build contract. Expected=$pin Actual=$actualHash"
    }

    [xml]$keyDocument = Get-Content -LiteralPath $resolvedPath -Raw
    $root = $keyDocument.DocumentElement
    if ($null -eq $root -or $root.LocalName -ne 'trustedDevelopmentKey') {
        throw 'The trusted Development public key must use the trustedDevelopmentKey root element.'
    }
    $keyId = ([string]$root.GetAttribute('keyId')).Trim()
    if ([string]::IsNullOrWhiteSpace($keyId)) {
        throw 'The trusted Development public key is missing keyId.'
    }
    $rootElements = @($root.ChildNodes | Where-Object NodeType -eq ([System.Xml.XmlNodeType]::Element))
    if ($rootElements.Count -ne 1 -or $rootElements[0].LocalName -ne 'RSAKeyValue') {
        throw 'The trusted Development key must contain exactly one RSAKeyValue element.'
    }
    $rsaElements = @($rootElements[0].ChildNodes | Where-Object NodeType -eq ([System.Xml.XmlNodeType]::Element))
    $elementNames = @($rsaElements | ForEach-Object LocalName)
    if ($rsaElements.Count -ne 2 -or $elementNames -notcontains 'Modulus' -or
        $elementNames -notcontains 'Exponent') {
        throw 'The trusted Development key must contain public Modulus and Exponent only; private RSA material is forbidden.'
    }
    foreach ($element in $rsaElements) {
        try {
            $decoded = [Convert]::FromBase64String(([string]$element.InnerText).Trim())
        }
        catch {
            throw "The trusted Development key contains invalid Base64 in $($element.LocalName)."
        }
        if ($decoded.Length -eq 0) {
            throw "The trusted Development key contains an empty $($element.LocalName)."
        }
    }

    return [pscustomobject]@{
        Path = $resolvedPath
        Sha256 = $actualHash
        KeyId = $keyId
    }
}

function Assert-ChuanHoaManagedAssemblyVersion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedVersion
    )

    $resolvedPath = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
    $assemblyVersion = [System.Reflection.AssemblyName]::GetAssemblyName($resolvedPath).Version.ToString()
    $fileInfo = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($resolvedPath)
    $observed = @($assemblyVersion, $fileInfo.FileVersion, $fileInfo.ProductVersion)
    if (@($observed | Where-Object {
            ![string]::Equals($_, $ExpectedVersion, [System.StringComparison]::Ordinal)
        }).Count -ne 0) {
        throw "Managed assembly version split-brain in $resolvedPath. Expected=$ExpectedVersion Assembly=$assemblyVersion File=$($fileInfo.FileVersion) Product=$($fileInfo.ProductVersion)"
    }
}

function Assert-ChuanHoaManifestVersion {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$IdentityName,

        [Parameter(Mandatory = $true)]
        [string]$ExpectedVersion
    )

    [xml]$manifest = Get-Content -LiteralPath $Path -Raw
    $identities = @($manifest.SelectNodes("//*[local-name()='assemblyIdentity']") |
        Where-Object { [string]::Equals($_.GetAttribute('name'), $IdentityName,
            [System.StringComparison]::Ordinal) })
    if ($identities.Count -eq 0) {
        throw "Manifest does not contain assembly identity '$IdentityName': $Path"
    }
    $mismatches = @($identities | Where-Object {
        ![string]::Equals($_.GetAttribute('version'), $ExpectedVersion,
            [System.StringComparison]::Ordinal)
    })
    if ($mismatches.Count -ne 0) {
        throw "Manifest version split-brain for '$IdentityName' in $Path; expected $ExpectedVersion."
    }
}
