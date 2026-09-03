param(
    [switch]$LaunchWord
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "   SỬA CHỮA & KHÔI PHỤC TAB CHUẨN HÓA     " -ForegroundColor Cyan
Write-Host "==========================================" -ForegroundColor Cyan

# 1. Đóng Microsoft Word nếu đang mở
$runningWord = Get-Process WINWORD -ErrorAction SilentlyContinue
if ($runningWord) {
    Write-Host "Đang đóng các tiến trình Microsoft Word đang chạy..." -ForegroundColor Yellow
    Stop-Process -Name WINWORD -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 1
}

# 2. Xóa Chuẩn Hóa khỏi danh sách Disabled Items (Word Resiliency)
$officeVersions = @('14.0', '15.0', '16.0')
$clearedDisabled = 0
foreach ($ver in $officeVersions) {
    $disabledKey = "HKCU:\Software\Microsoft\Office\$ver\Word\Resiliency\DisabledItems"
    if (Test-Path $disabledKey) {
        $props = Get-ItemProperty $disabledKey
        foreach ($prop in $props.PSObject.Properties) {
            if ($prop.Name -notmatch '^(PS|__)') {
                $decoded = [System.Text.Encoding]::Unicode.GetString($prop.Value)
                if ($decoded -like '*chuanhoa*') {
                    Remove-ItemProperty -Path $disabledKey -Name $prop.Name -Force -ErrorAction SilentlyContinue
                    Write-Host "[OK] Đã xóa mục vô hiệu hóa ($($prop.Name)) trong Word $ver" -ForegroundColor Green
                    $clearedDisabled++
                }
            }
        }
    }
    
    $crashingKey = "HKCU:\Software\Microsoft\Office\$ver\Word\Resiliency\CrashingAddinList"
    if (Test-Path $crashingKey) {
        $props = Get-ItemProperty $crashingKey
        foreach ($prop in $props.PSObject.Properties) {
            if ($prop.Name -like '*chuanhoa*') {
                Remove-ItemProperty -Path $crashingKey -Name $prop.Name -Force -ErrorAction SilentlyContinue
                Write-Host "[OK] Đã xóa khỏi danh sách add-in bị lỗi trong Word $ver" -ForegroundColor Green
            }
        }
    }
}

# 3. Xóa bộ nhớ đệm Custom UI Ribbon Validation Cache
foreach ($ver in $officeVersions) {
    $validationCache = "HKCU:\Software\Microsoft\Office\$ver\Common\CustomUIValidationCache"
    if (Test-Path $validationCache) {
        $props = Get-ItemProperty $validationCache
        foreach ($prop in $props.PSObject.Properties) {
            if ($prop.Name -like '*ChuanHoa*') {
                Remove-ItemProperty -Path $validationCache -Name $prop.Name -Force -ErrorAction SilentlyContinue
                Write-Host "[OK] Đã làm mới Ribbon Cache ($($prop.Name)) trong Office $ver" -ForegroundColor Green
            }
        }
    }
}

# 4. Dọn dẹp SolutionMetadata lỗi của VSTO runtime
$vstoMeta = 'HKCU:\Software\Microsoft\VSTO\SolutionMetadata'
if (Test-Path $vstoMeta) {
    $metaProps = Get-ItemProperty $vstoMeta
    foreach ($prop in $metaProps.PSObject.Properties) {
        if ($prop.Name -like '*ChuanHoa*') {
            $guid = $prop.Value
            Remove-ItemProperty -Path $vstoMeta -Name $prop.Name -Force -ErrorAction SilentlyContinue
            if ($guid -and (Test-Path "$vstoMeta\$guid")) {
                Remove-Item -Path "$vstoMeta\$guid" -Recurse -Force -ErrorAction SilentlyContinue
            }
            Write-Host "[OK] Đã dọn dẹp metadata VSTO lỗi: $guid" -ForegroundColor Green
        }
    }
}

# 5. Đảm bảo LoadBehavior = 3 (Tự động tải khi mở Word)
$addinKey = 'HKCU:\Software\Microsoft\Office\Word\Addins\ChuanHoa.AddIn.Vsto'
if (Test-Path $addinKey) {
    Set-ItemProperty -Path $addinKey -Name 'LoadBehavior' -Value 3 -Type DWord
    Write-Host "[OK] Đã cấu hình LoadBehavior = 3 (Tải khi khởi động Word)" -ForegroundColor Green
} else {
    Write-Host "[CẢNH BÁO] Chưa tìm thấy đăng ký add-in tại $addinKey. Hãy chạy bộ cài hoặc script Start-ChuanHoaDevelopment.ps1." -ForegroundColor Yellow
}

# 6. Kiểm tra chứng chỉ tin cậy
$cert = Get-ChildItem Cert:\CurrentUser\TrustedPublisher -ErrorAction SilentlyContinue |
    Where-Object { $_.Subject -like '*CN=Chuan Hoa Local Development*' } |
    Select-Object -First 1
if ($cert) {
    Write-Host "[OK] Chứng chỉ 'Chuan Hoa Local Development' đã được tin cậy trong TrustedPublisher." -ForegroundColor Green
} else {
    Write-Host "[CẢNH BÁO] Chứng chỉ phát triển chưa có trong TrustedPublisher. Cần cài chứng chỉ để Word cấp quyền chạy." -ForegroundColor Yellow
}

# 7. Kiểm tra và khôi phục bộ nhớ đệm bản quyền/quy tắc (để các nút sáng và hoạt động)
$cacheDir = [System.IO.Path]::Combine($env:LOCALAPPDATA, 'ChuanHoa', 'Cache')
$leaseFile = Join-Path $cacheDir 'lease.xml'
$rulesFile = Join-Path $cacheDir 'rules.xml'
if (!(Test-Path $leaseFile) -or !(Test-Path $rulesFile)) {
    Write-Host "Đang kiểm tra nguồn khôi phục giấy phép offline..." -ForegroundColor Yellow
    $backupDir = [System.IO.Path]::Combine($env:LOCALAPPDATA, 'ChuanHoa', 'DevelopmentInstaller', 'Current', 'DevelopmentSupport')
    if (Test-Path $backupDir) {
        New-Item -ItemType Directory -Path $cacheDir -Force | Out-Null
        foreach ($f in @('lease.xml', 'rules.xml', 'server-time.txt')) {
            $src = Join-Path $backupDir $f
            if (Test-Path $src) {
                Copy-Item -LiteralPath $src -Destination (Join-Path $cacheDir $f) -Force
            }
        }
        Write-Host "[OK] Đã khôi phục thành công giấy phép và quy tắc offline vào Cache!" -ForegroundColor Green
    } else {
        Write-Host "[CẢNH BÁO] Thiếu file giấy phép (lease.xml, rules.xml). Hãy chạy bộ cài đặt mới hoặc Start-ChuanHoaDevelopment.ps1 để kích hoạt lại các nút." -ForegroundColor Yellow
    }
} else {
    Write-Host "[OK] Bộ nhớ đệm bản quyền và gói quy tắc (lease.xml, rules.xml) hợp lệ." -ForegroundColor Green
}

Write-Host "==========================================" -ForegroundColor Cyan
Write-Host "Khôi phục hoàn tất thành công!" -ForegroundColor Green

if ($LaunchWord) {
    Write-Host "Đang khởi động Microsoft Word..." -ForegroundColor Cyan
    Start-Process WINWORD.EXE
}
