# Cài đặt VSTO local development

## Chạy Development một lệnh

```powershell
& 'D:\Chuẩn Hóa\tools\development\Start-ChuanHoaDevelopment.ps1' -LaunchWord -LaunchAdmin
```

Lệnh build API/VSTO, tạo khóa Development ngoài source nếu chưa có, cài public trust key, chạy API loopback, kiểm tra signed lease/rule E2E, đăng ký add-in và tùy chọn mở Word cùng trang quản trị. Admin ở `http://127.0.0.1:5206/development/admin` và không mở nếu thiếu cờ Development/loopback.

Dừng API bằng:

```powershell
& 'D:\Chuẩn Hóa\tools\development\Stop-ChuanHoaDevelopment.ps1'
```

## File nào được dùng

- Không mở `src/ChuanHoa.AddIn.Vsto/bin/Release/ChuanHoa.AddIn.Vsto.vsto`. Đây là manifest chạy/debug tại máy phát triển, registry dùng hậu tố `|vstolocal`; mở trực tiếp có thể hiện lỗi “An add-in could not be found or could not be loaded”.
- Chạy `artifacts/vsto-dev-publish-local/setup.exe` để kiểm tra luồng cài đặt ClickOnce local. `setup.exe` kiểm tra prerequisite .NET Framework 4.8 và VSTO Runtime trước khi gọi deployment manifest.
- Gói này chỉ được ký bằng `CN=Chuan Hoa Local Development`; không phân phối cho người dùng thật và không dùng làm production evidence.

## Tạo lại gói

```powershell
& 'D:\Chuẩn Hóa\tools\vsto\publish_local_development.ps1' -ApplicationVersion 1.0.0.10
```

Script dùng alias ASCII `D:\ChuanHoaPublishLocal` trỏ vào thư mục artefact trong project vì VSTO MSBuild 17 không phân tích đúng deployment URI chứa ký tự tiếng Việt. Alias chỉ dùng lúc build; artefact thực nằm trong `D:\Chuẩn Hóa\artifacts\vsto-dev-publish-local`.

## EXE thử nghiệm một file

```powershell
& 'D:\Chuẩn Hóa\tools\vsto\build_development_test_exe.ps1'
```

Version lấy từ `Directory.Build.props`. Script luôn rebuild source, publish VSTO, audit allowlist và tạo `artifacts\installers\development\ChuanHoa_Development_Test_Setup_<version>.exe`. Không build lại cùng một version; hãy tăng `ProductVersion` khi source thay đổi.

```powershell
# Cài hoặc nâng cấp không hiện hộp thoại
& '.\ChuanHoa_Development_Test_Setup_1.0.0.90.exe' /quiet

# Repair
& '.\ChuanHoa_Development_Test_Setup_1.0.0.90.exe' /repair /quiet

# Gỡ bản Development; giữ từ điển cá nhân và tài liệu
& '.\ChuanHoa_Development_Test_Setup_1.0.0.90.exe' /uninstall /quiet
```

## Giới hạn evidence

- Build Publish hiện đã đạt 0 warning/0 error và có `setup.exe`, top-level `.vsto`, version directory cùng các dependency `.deploy`.
- Fresh install, repair, uninstall và reinstall của EXE Development 1.0.0.90 đã PASS; đây không phải silent enterprise install.
- Word 2010 và Office x86 chưa có VM evidence.
- Production vẫn cần certificate tin cậy có timestamp, HTTPS immutable version directory và CI signing identity.
