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

## Trạng thái hiện tại

Build/load local không còn bị chặn bởi toolchain: Release artefact đã được ký bằng certificate development và smoke-test trên Word 16 x64. Release production vẫn bị chặn vì chưa có production code-signing certificate/CI signing identity, HTTPS deployment URL bất biến, KMS/HSM cho grant/FixPlan/rule, VM Word 2010/x86 và installer/update/uninstall evidence. Certificate development không được dùng để promote production.
