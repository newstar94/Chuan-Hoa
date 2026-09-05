# Release Runbook

## Promotion path

`development → internal → pilot → stable → legacy-maintenance`

## Điều kiện trước build phát hành

1. Source revision và dependency lock đã cố định.
2. Unit, contract, rule, security và relevant Word tests có evidence.
3. Không có production secret trong repository hoặc artefact.
4. Artefact dùng cùng một build khi promote giữa rings.
5. DLL/EXE/installer được Authenticode sign và timestamp.
6. Application/deployment manifests được ký và chứa hash mọi runtime file.
7. Release record có version, commit, build ID, hashes, SBOM, signatures, compatibility và test report.

## Dependency/privacy gates

Chạy toàn bộ source-quality gate bằng một lệnh (locked restore, validators, SBOM, online NuGet audit, secret/privacy/decision regression và 403 test trong Temp để không khóa API Development):

```powershell
powershell -ExecutionPolicy Bypass -File .\tools\validation\run_source_quality_gates.ps1
```

Các lệnh thành phần tương đương:

```powershell
$env:CI = 'true'
.\.tools\dotnet\dotnet.exe restore .\ChuanHoa.slnx --locked-mode
.\.tools\dotnet\dotnet.exe restore .\tests\ChuanHoa.Infrastructure.IntegrationTests\ChuanHoa.Infrastructure.IntegrationTests.csproj --locked-mode
Remove-Item Env:CI
python .\tools\validation\validate_dotnet_supply_chain.py --write-evidence .\shared\docs\implementation\evidence\dotnet_supply_chain.json
python .\tools\validation\generate_dotnet_sbom.py
python .\tools\validation\audit_nuget_vulnerabilities.py --write-evidence .\shared\docs\implementation\evidence\nuget_vulnerability_audit.json
python .\tools\validation\validate_repository_secrets.py --self-test --history --write-evidence .\shared\docs\implementation\evidence\repository_secret_regression.json
python .\tools\validation\validate_document_privacy.py --self-test --write-evidence .\shared\docs\implementation\evidence\document_privacy_regression.json
python .\tools\validation\validate_product_decision_consistency.py --self-test --write-evidence .\shared\docs\implementation\evidence\product_decision_consistency.json
```

- `packages.lock.json` phải được cập nhật có chủ ý cùng thay đổi dependency; CI không được tự nới lock.
- SBOM hiện hành là CycloneDX JSON tại `artifacts/sbom/chuanhoa-dotnet.cdx.json`; hash phải được đưa vào release record.
- NuGet advisory audit là kết quả tại thời điểm chạy và phải chạy lại mỗi release. Secret regression quét current indexed/non-ignored tree cùng Git history; nội dung secret Development đã ignore không bị đọc hoặc ghi vào evidence.
- Đây là gate local NuGet/SBOM/secret/privacy source và là cấu hình cho managed CI. Remote run, branch protection, CI provenance, production signing và production network observability vẫn là gate riêng, chưa được suy thành PASS.
- Lệnh hợp nhất hiện quét cả current tree và toàn bộ Git history. Workflow `.github/workflows/source-quality.yml` chạy cùng lệnh trên Windows với checkout full history và action pin bằng SHA. Chỉ ghi remote CI PASS khi có run URL/commit SHA và branch protection thực tế; file workflow local không tự chứng minh trạng thái GitHub.

## ClickOnce consumer

- HTTPS deployment URL ổn định và immutable version directory.
- Ring/cohort dùng deployment manifest URL đã ký riêng; client không tự chọn production ring.
- Update stage khi Word chạy và apply ở lần mở kế tiếp; không thay binary trong process Word.
- Required minimum version được server enforce.
- Người dùng chạy `setup.exe` từ publish root; không mở manifest `bin\Release\*.vsto` dành cho `vstolocal` development.

## MSI/Intune enterprise

- Detect Office bitness và registry view.
- Check/install .NET Framework/VSTO Runtime theo policy.
- Tắt ClickOnce self-update trong enterprise mode.
- Managed compliance chỉ authoritative khi connector khỏe; thiếu connector hiển thị PARTIAL.

## Rollback

- Binary: build/ký last-known-good dưới version mới cao hơn, smoke test và promote nhanh.
- Rule/config: đổi mapping về immutable compatible release trước đó.
- Không sửa hoặc ghi đè stable artefact lịch sử.

## Development Test EXE lifecycle

- Mỗi `ProductVersion` chỉ được build một lần; muốn thay source phải tăng version.
- Cài/upgrade bằng cách chạy EXE; bootstrapper bung allowlist vào `Staging-*`, kiểm tra rồi chuyển nguyên tử sang `Current`, giữ `Previous` để rollback.
- Repair không hiện UI: dùng Modify/Repair trong Apps & Features hoặc chạy EXE cache với `/repair /quiet`; lặp lại đúng luồng staging/verification/activation.
- Gỡ bản Development: dùng Apps & Features; `QuietUninstallString` gọi đúng EXE đã ký được cache theo version. Trình gỡ chỉ xóa registration, `Current`/`Previous`/`Staging-*`, trusted Development key, certificate thuộc installer và cache version của chính nó. Từ điển cá nhân và tài liệu người dùng phải được giữ.
- Các cờ này chỉ là quy trình Development; không thay thế MSI/Intune/ClickOnce production.

## Trạng thái hiện tại

Build/load local không còn bị chặn bởi toolchain: bản Development 1.0.0.107 đã được ký Authenticode cả outer EXE và ba PE do dự án sở hữu; VSTO application/deployment manifest được sinh và ký sau khi PE có bytes cuối. Audit payload, cài và smoke-test startup thụ động/Ribbon trên Word 16 x64 PASS. Repair, uninstall qua Apps & Features command, cleanup signed versioned cache, reinstall và rollback 6/6 fault point trên chính artifact `.107` PASS; state cài đặt, trusted key, cache ký, Apps & Features, installer cache, chữ ký inner PE và từ điển cá nhân được hậu kiểm. Builder bắt buộc pin SHA-256 của khóa công khai và certificate; bootstrapper/audit dùng cùng allowlist 8 file, gồm `Microsoft.Office.Tools.Common.v4.0.Utilities.dll`. Các artifact cũ là lịch sử bất biến. Add/remove certificate trong `/quiet` dùng đúng `certutil.exe` hệ thống, không tạo cửa sổ Root Certificate Store. Build/Rebuild kiểm chứng không được đăng ký, hủy đăng ký hoặc xóa inclusion-list trust; chỉ launcher/bootstrapper được quản lý registration. Release production vẫn bị chặn vì chưa có production code-signing certificate/timestamp/CI signing identity, HTTPS deployment URL bất biến, KMS/HSM cho grant/FixPlan/rule và VM Word 2010/x86. Certificate development tự ký không được dùng để promote production.
