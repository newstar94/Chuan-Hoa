$folder = "d:\chuan-hoa-the-thuc-workspace\client-web-addin"
$shareName = "VietDocAddin"

# 1. Tao SMB Share
try {
    $existing = Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue
    if ($null -eq $existing) {
        New-SmbShare -Name $shareName -Path $folder -FullAccess $env:USERNAME -ReadAccess "Everyone" -ErrorAction SilentlyContinue | Out-Null
        Write-Output "Created SMB Share: $shareName"
    } else {
        Write-Output "SMB Share already exists: $shareName"
    }
} catch {
    Write-Output "SMB Share Note: $($_.Exception.Message)"
}

$comp = $env:COMPUTERNAME
$unc = "\\$comp\$shareName"

# 2. Dang ky vao Registry Office 16.0 TrustedCatalogs
$guid = "{D3B07384-D113-4F44-972A-60589A19D826}"
$regKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$guid"

if (-not (Test-Path $regKey)) {
    New-Item -Path $regKey -Force | Out-Null
}

Set-ItemProperty -Path $regKey -Name "Url" -Value $unc -Force
Set-ItemProperty -Path $regKey -Name "Flags" -Value 1 -Type DWord -Force

Write-Output "UNC Path: $unc"
Write-Output "Trusted Catalog registered successfully!"
