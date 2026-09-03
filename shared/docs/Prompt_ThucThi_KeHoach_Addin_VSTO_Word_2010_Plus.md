# MASTER EXECUTION PROMPT — TRIỂN KHAI ADD-IN VSTO WORD 2010+

Bạn là Codex chịu trách nhiệm triển khai thực tế sản phẩm Add-in Chuẩn hóa Thể thức trong workspace:

D:\Chuẩn Hóa

Đây là nhiệm vụ xây dựng sản phẩm, không phải nhiệm vụ chỉ phân tích hoặc viết lại kế hoạch. Hãy làm việc liên tục theo từng giai đoạn cho đến khi đạt Definition of Done hoặc gặp một blocker thật sự cần quyết định kinh doanh, quyền truy cập, chứng thư, hạ tầng hay hệ thống bên ngoài.

## 1. Nguồn sự thật bắt buộc

Trước khi sửa bất kỳ code nào, phải đọc đầy đủ:

1. D:\Chuẩn Hóa\shared\docs\KeHoach_TrienKhai_Addin_VSTO_Word_2010_Plus.md
2. Toàn bộ README, tài liệu kiến trúc và tài liệu nghiệp vụ trong workspace.
3. Toàn bộ VBA trong shared\vba_extracted.
4. Ribbon chuẩn trong shared\ChuanHoaTheThuc_Full_Ribbon.dotm và các artefact Ribbon đã trích xuất.
5. Toàn bộ client-vsto-csharp, client-web-addin, backend-api, rules, dictionaries, installer và script liên quan.
6. Hướng dẫn dành cho agent trong AGENTS.md, CONTEXT.md hoặc file tương đương nếu có.
7. Trạng thái thay đổi hiện hữu trong workspace; không ghi đè thay đổi của người dùng.

Kế hoạch chi tiết là hợp đồng triển khai chính. Nếu prompt này và kế hoạch có khác biệt, ưu tiên:

1. Yêu cầu mới nhất của người dùng.
2. Quyết định đã khóa trong kế hoạch.
3. Nguyên tắc an toàn, bảo mật và bảo toàn dữ liệu.
4. Chi tiết triển khai có thể đảo ngược do Codex lựa chọn và ghi lại bằng ADR.

Ngoại lệ đã được rà soát và khóa trong prompt này: `CORRECTION-ADMIN-PHASE-MAPPING-001` sửa riêng nhãn phân kỳ Admin đang không khớp giữa Mục 22 và Bảng 14.22 của kế hoạch. Mapping bắt buộc là Admin B/C/D ở Giai đoạn 7, Admin E ở Giai đoạn 8 và Admin F ở Giai đoạn 9, vì Admin D là commercial/payment, Admin E là release/rollout/incident và Admin F là production hardening. Correction này có chủ đích, được ưu tiên hơn các nhãn Admin B/C, D, E/F cũ tại Mục 22; không được coi đó là lý do dừng triển khai hoặc tự chọn một mapping khác. Mọi nội dung khác của kế hoạch vẫn giữ nguyên thứ tự ưu tiên nêu trên.

Correction kỹ thuật `CORRECTION-RULE-ROUTE-BASELINE-001` được tạo sau phép đo tái lập tại `shared/docs/implementation/evidence/rule_reconciliation.json`: 96 definition VBA, 94 route đã đăng ký, 19 route trỏ tới ba implementation luôn trả `Nothing`, do đó có 75 route có đường logic. Correction này thay thế số 17/77 cũ trong prompt; không thay đổi 52 mã JSON và khoảng 14 mã backend prototype.

Không được tự giảm phạm vi chỉ vì code hiện tại mới là prototype.

## 2. Kết quả cuối cùng phải đạt

Xây dựng một sản phẩm hoàn chỉnh gồm:

1. Một Word VSTO Add-in production-ready cho Word 2010 trở lên trên Windows.
2. Một tab Ribbon duy nhất, giữ đúng contract giao diện VBA đã phê duyệt.
3. Đầy đủ chức năng kiểm tra, chuẩn hóa và AutoFix theo plan.
4. Backend authoritative quản lý danh tính, thiết bị, trial, entitlement, rules, FixPlan, giá, đơn hàng, thanh toán, release và audit.
5. Admin Portal web riêng để chủ sản phẩm quản lý ứng dụng, người dùng và vận hành.
6. Customer Portal web riêng để người dùng và quản trị viên tổ chức tự phục vụ trong đúng phạm vi quyền.
7. Chuỗi build, ký số, cài đặt, cập nhật, rollout, forward-rollback và thu hồi phiên bản.
8. Bộ test, corpus, bằng chứng tương thích, tài liệu vận hành, bảo mật và migration khỏi VBA.

Không phát triển Office Web Add-in như client sản phẩm trong phạm vi này. Thư mục client-web-addin hiện tại chỉ là artefact tham khảo và không được dùng để tuyên bố hoàn thành.

## 3. Các quyết định sản phẩm đã khóa

Không hỏi lại các quyết định sau:

- Client sản phẩm là VSTO/C#.
- Phạm vi là Word 2010 trở lên trên Windows, Office x86 và x64 theo support lane trong plan.
- Không hỗ trợ Word Web, macOS, iPad, Android, Office 2007 hoặc Word 2003 trong phạm vi này.
- Trong Word chỉ có một tab Ribbon; không có task pane thường trực.
- Dialog tạm thời được phép cho login, progress, preview, warning, QR và About.
- Word được đọc và chỉnh sửa tại local bằng Word Object Model.
- Scan premium, compliance engine, canonical rules và signed FixPlan thuộc server.
- Add-in không chứa toàn bộ premium engine/rules để chạy offline vô hạn.
- Admin Portal và Customer Portal là website riêng, không phải Word Web Add-in.
- Trial ra mắt và trial cá nhân không cộng dồn.
- Giá không hardcode trong DLL.
- Offer đã publish, quote, payment, trial grant, entitlement history, release và audit là bất biến theo quy tắc trong plan.
- Binary rollback dùng forward-fix với version cao hơn; không hạ ClickOnce version.
- Rule/config mapping chỉ được rollback trực tiếp khi schema và protocol còn tương thích.
- Client không được coi là nguồn sự thật cho giá, thời gian, entitlement, DLL integrity, payment hoặc trial.
- Không được hứa chống crack tuyệt đối trên máy người dùng có toàn quyền quản trị.

## 4. Các quyết định chưa có dữ liệu production

Không tự bịa các giá trị sau:

- Ngày bắt đầu và kết thúc launch trial.
- Thời lượng personal trial.
- Giới hạn thiết bị theo loại tài khoản/gói.
- Trial lease TTL, paid lease offline grace và chính sách khi server không reachable.
- Giá, currency, subscription term, renewal policy và grandfathering policy.
- Refund/chargeback policy và hạn mức phê duyệt.
- Payment provider production và thông tin merchant.
- Production identity tenant/provider configuration.
- Admin IdP, role holders, approval thresholds và break-glass SLA.
- Code-signing certificate production.
- KMS/HSM keys và secret production.
- Domain production, deployment URL và callback URL.
- Ngưỡng thời hạn/phạm vi manual entitlement cần phê duyệt.
- Ranh giới organization/tenant, SSO/domain verification và quyền customer administrator.
- Support lane, compatibility lane và việc có hỗ trợ Windows 32-bit/legacy theo hợp đồng hay không.
- Chính sách cloud snapshot, retention và trường hợp bắt buộc on-prem.
- Bộ legal/business sources được phê duyệt cho ND30, HD05 và Viettel.
- VM, Office licenses và môi trường dùng làm compatibility evidence.
- Design sign-off, brand assets và accessibility acceptance owner cho Portal.
- Quyết định có xây signed updater riêng để đạt zero-click trên máy cá nhân hay không.
- Mức cam kết client authenticity trên máy cá nhân và cơ chế nào được chọn để phân biệt clone/re-signed client dùng chính tài khoản, thiết bị và entitlement hợp lệ.
- Danh sách khách hàng cần on-prem engine.

Mọi giá trị ghi là “mặc định đề xuất” trong kế hoạch chỉ dùng để ước lượng, thiết kế contract và test cấu hình; không phải quyền tự động chốt production.

Với mỗi mục còn thiếu:

1. Ghi vào Decision Register với mã BLOCKED_DECISION.
2. Tạo interface, configuration contract, validation, test double và runbook cần thiết.
3. Không tạo secret, merchant ID, certificate hoặc ngày giả cho production.
4. Tiếp tục mọi hạng mục độc lập không bị chặn.
5. Chỉ dừng để hỏi người dùng khi blocker đó thực sự ngăn bước tiếp theo và không còn công việc an toàn nào khác.

## 5. Cách làm việc bắt buộc

### 5.1. Không chỉ lập kế hoạch

- Sau khi audit và khóa baseline, phải thực hiện code, migration, test và tài liệu.
- Không kết thúc chỉ với danh sách việc cần làm.
- Không tạo handler chỉ hiện thông báo thành công.
- Không tạo mock được bật trong production.
- Test double chỉ tồn tại trong test hoặc môi trường development được khóa rõ.
- Không dùng localhost URL, debug bypass, test license hoặc giá hardcode trong release.

### 5.2. Làm theo phase và exit gate

Thực hiện theo dependency và exit gate của Mục 22, đồng thời áp dụng bắt buộc `CORRECTION-ADMIN-PHASE-MAPPING-001` để đồng bộ phase với phạm vi chức năng tại Bảng 14.22:

1. Giai đoạn 0 — Phê duyệt sản phẩm và threat model.
2. Giai đoạn 1 — Đóng băng Ribbon và hành vi VBA.
3. Giai đoạn 2 — Canonical rules và corpus.
4. Giai đoạn 3 — Nền VSTO production.
5. Giai đoạn 4 — Nền server, identity, security và Admin A.
6. Giai đoạn 5 — Compliance Engine và signed FixPlan.
7. Giai đoạn 6 — Port chức năng theo các wave.
8. Giai đoạn 7 — Trial, thương mại, Admin B/C/D và Customer Portal.
9. Giai đoạn 8 — Signing, update, installer, Admin E và release operations.
10. Giai đoạn 9 — Pilot, Admin F, migration và ra mắt.

Không đánh dấu phase hoàn thành nếu chưa có bằng chứng cho exit gate. Có thể chạy các workstream độc lập song song nhưng không được đi qua dependency chưa đạt.

### 5.3. Duy trì bằng chứng triển khai

Tạo và cập nhật các tài liệu sau trong shared\docs\implementation:

- Implementation_Status.md
- Decision_Register.md
- Architecture_Decisions.md
- Ribbon_Contract.md và bản machine-readable tương ứng.
- VBA_Migration_Ledger.md hoặc định dạng machine-readable tương đương.
- Rule_Catalog_Reconciliation.md.
- Test_Evidence.md.
- Compatibility_Matrix.md.
- Security_Threat_Model.md.
- Release_Runbook.md.
- Incident_Runbook.md.
- Data_Privacy_Register.md.

Mỗi hạng mục trạng thái phải có:

- Owner hoặc workstream.
- Trạng thái.
- Dependency.
- File/code liên quan.
- Test ID.
- Evidence thực tế.
- Blocker nếu có.
- Ngày cập nhật.

Không ghi “đã test” nếu không có lệnh, kết quả hoặc bằng chứng môi trường thật.

### 5.4. Quản lý thay đổi

- Bảo toàn mọi thay đổi hiện hữu không thuộc nhiệm vụ.
- Không dùng thao tác phá hủy để làm sạch workspace.
- Không xóa prototype VBA/Web/VSTO trước khi đã archive, đối chiếu và qua phase retirement.
- Không thay đổi dữ liệu production hoặc gửi thông báo ra ngoài nếu chưa có ủy quyền cụ thể.
- Mọi migration database phải có up/down hoặc chiến lược forward-only được giải thích, test và backup/restore plan.
- Mọi thay đổi contract phải cập nhật server, generated client, validation, migration và test liên quan.

## 6. Gate baseline trước khi port

Phải tạo bằng chứng xác nhận:

- Ribbon chuẩn có đúng 1 tab, 7 group, 36 button, 4 menu, 2 dropdown và 3 checkbox.
- Tổng cộng 45 control có ID, label, order, icon, size, screentip, supertip, callback, state và test ID.
- Có đúng 68 module VBA trong migration ledger.
- Mỗi module có public entry points, callers, dependency, side effects, đích C#, quyết định PORT/REPLACE/MERGE/RETIRE/REFERENCE_ONLY, fixture, test và sign-off.
- Các số liệu rules phải được tái hiện bằng script/test và ghi rõ provenance: 96 definition VBA; 94 route đã đăng ký; 75 route có đường logic sau khi loại 19 route trỏ tới implementation luôn trả rỗng; tài liệu/metadata từng tuyên bố 82; JSON có 52 mã duy nhất; backend prototype phát khoảng 14 mã. Không gộp các khái niệm này thành một con số marketing.
- Không công bố số lượng rule thương mại trước khi declared, loaded, routed, implemented và reported đã khớp.
- Nút CHUẨN HÓA TOÀN BỘ được nối với workflow AutoFix thật, không gọi nhầm định dạng trang.
- Kiểm tra thể thức và kiểm tra chính tả có contract read-only.

Nếu một kiểm kê khác plan, không âm thầm sửa plan hoặc code. Ghi evidence, xác định nguyên nhân và cập nhật ADR/Decision Register trước khi tiếp tục.

## 7. Yêu cầu VSTO bắt buộc

### 7.1. Nền project

- Tạo project Word VSTO Add-in thật, không tiếp tục giả định prototype hiện tại là buildable.
- Target .NET Framework 4.8.
- Dùng Word 2010 Object Model làm semantic baseline.
- RibbonX dùng schema tương thích Word 2010.
- Embed Interop Types khi phù hợp.
- Có generated ThisAddIn partial, Office build targets, manifests và publish configuration đúng chuẩn.
- Không dùng regsvr32 để đăng ký managed VSTO DLL.
- Không load assembly/plugin thực thi nằm ngoài signed manifest inventory.

Nếu máy hiện tại thiếu Visual Studio Office Developer Tools, VSTO targets hoặc Office VM:

1. Ghi blocker môi trường với lỗi chính xác.
2. Hoàn thiện source, project, CI validation và test không phụ thuộc host trong khả năng an toàn.
3. Tiếp tục backend, contracts, Admin Portal, test và tài liệu.
4. Không tuyên bố VSTO build/load đã qua nếu chưa chạy trên toolchain thật.

### 7.2. Ribbon

- Giữ đúng một giao diện Ribbon trong Word.
- Không thêm task pane thường trực.
- Đủ 38 control target theo Mục 6.3 của plan; giữ baseline VBA 45 control làm provenance.
- Đủ 31 command button có handler thật.
- Không có chức năng, entitlement hoặc callback Lưu thành DOCX.
- Không có chức năng, entitlement hoặc callback Bỏ dấu hay chuẩn hóa i/y.
- Tên tab và tên sản phẩm hiển thị là “Chuẩn hóa”, không dùng “CHUẨN HÓA THỂ THỨC (2026)” hoặc tên kèm năm.
- Xử lý trực tiếp tài liệu đã lưu `.doc` và `.docx`, giữ nguyên định dạng; định dạng khác fail closed.
- Menu chỉ là container khi contract quy định.
- Dropdown và checkbox phải có dynamic callbacks đầy đủ.
- getEnabled, selected item và pressed state phải phản ánh active document/window.
- Ribbon invalidation chạy đúng Office UI thread và có debounce.
- Không dùng global static state cho regime, document type hoặc view.
- Mỗi Document/Window có context tách biệt.
- Chuyển document/window không rò state.

### 7.3. Word safety

- Scan là read-only.
- Scan không đổi text, formatting, selection, comments, Track Changes, custom properties, undo stack hoặc Word global options.
- Mỗi mutation phải được phân loại vào đúng một trong hai lane; không được trộn contract hoặc tạo artefact giả để hợp thức hóa luồng:
  - Compliance/AutoFix do server hoạch định: Snapshot → Finding → signed FixPlan → Preview → Apply → Verify → Undo/Rollback.
  - Local deterministic command không cần Compliance Engine: Snapshot/preflight → signed ExecutionGrant → Preview/Confirm theo risk tier → Apply → Verify → Undo/Rollback. Không tạo Finding hoặc FixPlan giả cho lane này.
- FixPlan của lane Compliance/AutoFix phải ký số, ràng buộc user, device, command, document fingerprint, revision, schema, ruleset, engine version, thời hạn và one-time identity.
- ExecutionGrant của lane local deterministic phải tuân thủ đầy đủ Mục 9.2; command vẫn phải kiểm tra document state, protected/read-only state, allowlist, bounds, risk tier, preview/confirm, fingerprint ngay trước apply, one-time grant, Undo và backup/rollback phù hợp.
- Chụp lại fingerprint ngay trước apply; mismatch phải fail closed.
- Một command dùng một custom Undo record khi Word version hỗ trợ.
- Có backup và rollback policy tương ứng risk tier.
- Risk tier chỉ có bốn semantics:
  - SAFE: thay đổi xác định, giới hạn rõ và có thể áp dụng sau preview theo product policy.
  - CONFIRM: thay đổi text, cấu trúc hoặc bảng có rủi ro; phải hiển thị operation cụ thể và người dùng xác nhận.
  - REPORT_ONLY: chỉ tạo finding, tuyệt đối không sinh operation tự sửa.
  - BLOCKED: không thay đổi tài liệu vì protected/read-only, fingerprint mismatch, plan hết hạn, schema lạ hoặc precondition không đạt.
- Nội dung pháp lý/nghiệp vụ, số tiền, ngày nghiệp vụ, tên cơ quan, chữ ký, con dấu và component không chắc chắn luôn REPORT_ONLY, kể cả khi server rule bị cấu hình sai; client safety policy phải từ chối operation ngoài allowlist/risk contract.
- Khôi phục ScreenUpdating, events, selection, view và trạng thái Word trong mọi error path.
- Không xóa hoặc làm mất comments, Track Changes, fields, bookmarks, content controls, headers, footers, footnotes, endnotes, hyperlinks, hình, bảng, section, chữ ký hoặc con dấu ngoài operation scope đã được ủy quyền bởi signed FixPlan của lane Compliance/AutoFix hoặc bởi command contract kèm signed ExecutionGrant của lane local deterministic. Mọi cấu trúc nằm ngoài operation scope đã authorize luôn phải được bảo toàn.
- Unknown operation hoặc schema không tương thích phải fail closed và không sửa tài liệu.
- Không để lỗi auth, server hoặc integrity làm Word crash hay khóa tài liệu.

### 7.4. Port chức năng

Port đủ các wave trong plan:

- Wave A: đọc dữ liệu trực tiếp từ `.doc`/`.docx`, regime, loại văn bản, kiểm tra thể thức, kiểm tra chính tả và trình bày findings.
- Wave B: page setup, section ngang/dọc, trang thừa, styles 13/14/15, cỡ chữ, Keep with next, page numbers và character spacing.
- Wave C: table headers, chuẩn hóa bảng, căn ô, Excel cleanup, ảnh và QR.
- Wave D: TCVN3/Unicode, whitespace, punctuation, dash, ellipsis và decimal; không triển khai tone hoặc i/y.
- Wave E: ba View toggles, update status, feedback và About.
- AutoFix được triển khai sau khi các operation thành phần và safety kernel đạt gate.

Mỗi command phải có:

- Contract.
- Precondition.
- Scope.
- Entitlement.
- Capability guard.
- Risk tier.
- Handler thật.
- Positive, negative và boundary tests.
- Golden document fixture.
- Undo/rollback evidence nếu có mutation.

## 8. Canonical rules và Compliance Engine

- Không dùng JSON hiện tại như nguồn canonical trước khi tái tạo và validation.
- Tách rõ quy định pháp lý, phép kiểm tra thực thi và thao tác sửa tự động.
- Mỗi rule có stable code, regime, document type, component, source, effective period, severity, checker, fixability, risk tier và tests.
- Mỗi rule được gắn trạng thái IMPLEMENTED, WARN_ONLY, NOT_APPLICABLE hoặc RETIRED.
- NOT_CHECKED không được tính là PASS.
- Không hardcode số rule trong API hoặc UI.
- Dictionaries phải bảo toàn Unicode, NFC, dấu thanh và từ tiếng Việt.
- Rules/dictionaries release bất biến, có schema, signature, engine compatibility, effective date, corpus evidence, pilot và rollback mapping.
- Premium engine và rule logic cốt lõi chạy phía server.
- Rule release không được chứa script, assembly hoặc code tùy ý gửi xuống client.

## 9. Server-authoritative security

### 9.1. Identity và device

- Dùng system browser và OAuth/OIDC Authorization Code với PKCE.
- VSTO là public client, không chứa client secret.
- Access token ngắn hạn.
- Refresh-token rotation và reuse detection.
- Device dùng asymmetric key; ưu tiên TPM-backed CNG non-exportable, fallback DPAPI/software key phải có assurance level thấp hơn.
- HWID, MAC hoặc serial chỉ là risk signal, không phải credential.
- Premium request dùng device proof, nonce, timestamp, jti và replay protection.
- Copy cache/token/lease sang máy khác phải thất bại vì thiếu device private key.
- Không coi TPM/device key là bằng chứng DLL chính thức.

### 9.2. Entitlement và lease

- Account, device, client release, entitlement và trial được kiểm tra tại server.
- Lease được ký bằng asymmetric key trong KMS/HSM và hỗ trợ key rotation bằng kid.
- Lease gắn subject, organization, device key, product/features, source, release, protocol, issued/not-before/expiry và jti.
- Trial lease không vượt trial end.
- Paid offline grace chỉ duy trì identity/quyền cache theo policy; không tạo premium FixPlan mới khi engine không reachable.
- Revocation có hiệu lực tại request online tiếp theo theo SLA; ghi rõ residual risk của lease offline.
- Local command có giá trị nhưng không cần Compliance Engine vẫn phải lấy ExecutionGrant ngắn hạn có chữ ký.
- ExecutionGrant phải gắn command ID, subject, organization, device key, client release, document fingerprint/scope, issued/not-before/expiry, nonce và one-time jti.
- Handler local xác minh chữ ký, key ID, audience, command, user/device/release/document binding và expiry trước khi mutate.
- Server và client có replay cache/idempotency; grant bị copy, sửa, hết hạn, đổi command/document/device hoặc dùng lần hai phải bị từ chối.
- ExecutionGrant chỉ tăng chi phí bypass; không được coi là bằng chứng DLL chính thức nếu implementation đầy đủ nằm local.

### 9.3. Anti-tamper và release trust

- Ký Authenticode DLL, EXE, bootstrapper và installer.
- Ký application manifest và deployment manifest.
- Manifest liệt kê và hash mọi DLL, resource và config thực thi.
- Timestamp mọi chữ ký.
- Production signing key non-exportable và không nằm trong repo/runner lâu dài.
- Tách code-signing, lease-signing, FixPlan-signing, rule-signing và TLS keys.
- Strong name, obfuscation, self-hash và anti-debug chỉ là defense-in-depth.
- Không tin client tự báo DLL hash/version như attestation.
- DLL chính thức bị đổi một byte phải không đi vào VSTO startup qua official deployment.
- Add-in giả hoặc re-signed không có production entitlement phải không nhận premium findings/FixPlan.
- Client dưới minimum protocol/version hoặc release bị revoke phải fail closed cho premium.
- Phải lập ADR CLIENT_AUTHENTICITY_BOUNDARY trước khi giữ cam kết “mọi client giả/re-signed đều không nhận lease/FixPlan”. OAuth public client, device key, TPM và client-reported hash không phân biệt official DLL với clone chạy dưới cùng user/device.
- ADR phải chọn và chứng minh một trong các mức: endpoint managed bằng WDAC/App Control publisher policy; trusted updater/helper hoặc attestation service có threat model độc lập; cơ chế platform attestation được kiểm chứng; hoặc chấp nhận giới hạn consumer và sửa acceptance/marketing claim cho đúng.
- Không được coi obfuscation, self-check hoặc certificate embedded trong DLL là giải pháp cho boundary này.
- Security test bắt buộc clone/re-signed add-in dùng tài khoản, entitlement, token và thiết bị hợp lệ. Chỉ được PASS “client authenticity” khi cơ chế đã chọn thực sự từ chối; nếu không phải ghi BLOCKED_DECISION hoặc ACCEPTED_RISK, không được báo PASS.

### 9.4. FixPlan

- FixPlan chỉ chứa operation enum trong allowlist và dữ liệu khai báo có bounds.
- Không chứa macro, script, assembly, process, shell, path hoặc URL thực thi.
- Server kiểm tra account, device, entitlement, command, release, protocol, quota, nonce và input limits trước khi chạy engine.
- Client xác minh signature, kid, schema, user, device, command, document fingerprint, revision và expiry.
- Replay, sửa operation, đổi document/user/device/command hoặc dùng unknown operation phải bị từ chối.
- Server compromise được coi là high-impact risk; client vẫn validate operation allowlist, bounds, preview và document state.

## 10. Privacy và document processing

- Local đọc Word, tạo snapshot và chỉ apply operation đã được ủy quyền bởi signed FixPlan hoặc command contract kèm signed ExecutionGrant tương ứng.
- Cloud chỉ nhận snapshot tối thiểu cần cho rule đang chạy.
- Không gửi nguyên DOCX nếu không có contract và consent riêng.
- Không gửi mặc định filename, path, comment người dùng, hình, chữ ký, con dấu, macro, external link hoặc QR payload.
- Không log request body chứa nội dung tài liệu.
- Snapshot có schema, size/depth/time limits.
- Worker xử lý input không tin cậy, có sandbox, timeout, regex timeout và tenant isolation.
- Queue/storage tạm mã hóa và có TTL; ưu tiên zero-retention/delete-on-completion.
- Identity/commercial data tách khỏi document-processing worker.
- Có privacy notice, consent, DPA, retention, deletion và incident flow.
- Khách không cho phép cloud cần on-prem engine; ghi rõ on-prem làm giảm mức bảo vệ IP/anti-crack vì admin khách kiểm soát server.

## 11. Trial, tài khoản và entitlement

Triển khai đúng Mục 10–12 của plan:

- Backend là nguồn sự thật duy nhất.
- Server time và UTC là authoritative.
- Launch trial có start/end toàn cục, người vào muộn chỉ dùng thời gian còn lại.
- Personal trial chỉ áp dụng sau launch campaign theo eligibility đã khóa.
- Một user/product chỉ nhận tối đa một trial grant chuẩn.
- Người đã nhận launch trial không nhận personal trial.
- Reinstall, đổi máy, đổi clock, xóa cache hoặc đồng thời claim từ hai device không tạo trial mới.
- Trial grant lịch sử không bị xóa.
- Admin exception tạo entitlement/grant nguồn ADMIN mới, có expiry, reason và audit.
- Paid entitlement, manual grant, launch trial, personal trial và paid-required tuân thủ thứ tự ưu tiên trong plan.
- Account/device/release security denial thắng quyền thương mại.
- Boundary tests chạy tại trước, đúng và sau start/end.

## 12. Giá, quote, payment và subscription

- Product, plan, feature, quota và device policy có schema rõ.
- Offer được version hóa theo audience, channel, currency và thời gian hiệu lực.
- Offer PUBLISHED bất biến.
- Khoảng offer hiệu lực không được chồng lấn.
- Tiền lưu bằng integer minor units.
- Backend tự chọn offer; client không gửi amount authoritative.
- Quote chụp snapshot bất biến của giá, currency, tax, discount, feature, term và expiry.
- Quote có integrity binding, thuộc đúng subject và chỉ dùng một lần.
- Order tạo từ quote còn hiệu lực với idempotency key.
- Payment chỉ VERIFIED_PAID sau webhook/provider verification.
- Webhook xác minh signature, merchant, order reference, amount, currency và state.
- Duplicate và out-of-order webhook không tạo hai payment, subscription hoặc entitlement.
- Mismatch đi vào MANUAL_REVIEW, không tự cấp quyền.
- Refund/chargeback tạo event/record bù, không xóa payment.
- Dùng provider command/outbox, retry, dead-letter, replay an toàn và reconciliation.
- Không lưu card data.
- Payment provider chưa chọn không được thay bằng flow production giả.

## 13. Admin Portal

Triển khai toàn bộ Mục 14 của plan, không chỉ một dashboard minh họa.

### 13.1. Kiến trúc

- Portal web riêng với hostname và OAuth audience riêng.
- Admin login bắt buộc MFA.
- Dùng Admin BFF với Secure, HttpOnly, SameSite cookie.
- Không lưu admin access token trong localStorage.
- Portal không gọi database trực tiếp.
- Authorization thực thi tại backend, không chỉ ẩn nút.
- Mỗi môi trường có scope, banner và permission riêng.
- Portal không đọc hoặc hiển thị nội dung tài liệu Word.
- Production Portal chỉ promote artefact đã ký và CI-attested; không cho upload DLL/rule tùy ý.

### 13.2. Module giao diện

Phải có đầy đủ:

- Dashboard vận hành.
- User list và user detail.
- Organization, contract, member, seat và verified domain.
- Device và session.
- Trial campaign và trial grant.
- Entitlement, lease và resolver explanation.
- Product, plan, feature và offer versions.
- Quote, order, payment, webhook, refund, subscription và reconciliation.
- Client releases, distribution mode, release ring và compatibility lane.
- Rollout, minimum version, revoke và update health.
- Command catalog, feature flags và kill switch.
- Rule/dictionary releases.
- Service health, queue, incident và postmortem.
- Approval requests.
- Audit và security events.
- Admin user, role, permission và scope.
- Retention, privacy và notification configuration.

### 13.3. RBAC và approval

- Implement permission chi tiết và scope theo environment/organization.
- Người đề xuất không tự duyệt.
- Dùng các cặp maker-checker trong plan cho product/trial, entitlement, pricing, finance, rules, releases và security.
- Approval gắn target type, target ID, target version/ETag và payload hash.
- State machine gồm DRAFT, SUBMITTED, APPROVED, REJECTED, EXPIRED, CANCELLED, EXECUTING, EXECUTED và FAILED.
- Payload hoặc target thay đổi phải trả STALE_APPROVAL.
- Approve chỉ quyết định payload đã khóa và đưa command vào outbox; không nhận payload nghiệp vụ mới.
- Break-glass yêu cầu step-up MFA, incident ID, reason, TTL và post-review.
- Bật lại global command/release sau emergency phải qua approval bình thường.

### 13.4. Tenant isolation

- organizationId canonical được lấy từ principal/membership đã xác minh.
- Không suy tenant từ domain email hoặc ID client tự gửi.
- Kiểm tra cross-tenant cho detail, list, search, count, aggregate, export, cache, object path, queue và background job.
- Organization roles chỉ có quyền member/seat/device/billing/deployment trong tenant.
- Không cho Organization Admin truy cập Admin Portal nội bộ hoặc tenant khác.

### 13.5. Release và update UI

- Tách distributionMode, releaseRing và compatibilityLane.
- ClickOnce cohort dùng signed manifest/ring riêng hoặc signed updater đã phê duyệt.
- Client không tự chọn production ring.
- Client heartbeat/update attempt chỉ là last-observed telemetry, không phải DLL attestation.
- MSI/Intune chỉ hiển thị compliance authoritative khi connector quản trị có health tốt.
- Thiếu connector phải hiển thị coverage PARTIAL.
- Rollout có cohort, percentage, schedule, health threshold, pause và approval.
- Binary rollback là forward-fix version cao hơn.

### 13.6. UX

- Operations dashboard trung tính, chuyên nghiệp, data-dense nhưng dễ quét.
- Không dùng phong cách landing page, typography phóng đại, gradient trang trí hoặc animation gây nhiễu.
- Dùng semantic design tokens, một bộ SVG icon và nhịp spacing 4/8 px.
- Mọi status có text/icon, không chỉ màu.
- Phân biệt zero, no-data, partial, stale và unavailable.
- Bảng dùng server pagination/filter/sort; list lớn dùng virtualization phù hợp.
- Filter nằm trong URL và hỗ trợ deep link.
- Mutation có loading, idempotency, ETag conflict, impact preview, reason và audit reference.
- Error nêu nguyên nhân, cách phục hồi và correlation ID.
- Đạt WCAG 2.2 AA.
- Test keyboard, focus, screen reader, contrast, zoom 200%, reduced motion và touch target.
- Test responsive tại các breakpoint trong plan.

## 14. Customer Portal

- Subject chỉ xem và sửa các trường hồ sơ được phép của chính mình.
- Xem trial, entitlement, subscription, order, invoice và update status của mình.
- Revoke thiết bị của mình.
- Nhận installer/link phát hành chính thức theo policy.
- Mở quote/order từ offer do backend chọn.
- Gửi support/privacy request không tự đính kèm tài liệu.
- Organization Member/Admin/Billing/Deployment chỉ dùng API đúng tenant role.
- Không cho sửa trial date, amount, entitlement, risk state, release production control hoặc payment status.

## 15. API, database và contracts

- Backend mới dùng ASP.NET Core LTS phù hợp tại thời điểm triển khai.
- PostgreSQL là transactional database.
- OpenAPI và JSON Schema là nguồn contract.
- Sinh C# và TypeScript clients từ contract khi phù hợp.
- Mọi mutation có idempotency key.
- Mutation trên state hiện hữu có ETag/version và optimistic concurrency.
- Mọi response có correlation ID và versioned error contract.
- Mọi state transition quan trọng có append-only event/audit.
- Dùng transaction/outbox để không có trạng thái PAID thiếu entitlement hoặc ngược lại.
- Enforce unique/exclusion/foreign-key constraints tại database, không chỉ service.
- Không có admin API set-paid, set-trial-end, edit-published-offer hoặc upload-production-dll.
- Implement đầy đủ public/client, account, organization và admin API trong Mục 17 của plan.
- Implement đầy đủ data model trong Mục 18 của plan.
- Tenant scope được kiểm tra ở query, cache, export và background worker.

## 16. Cài đặt, cập nhật và phát hành

### 16.1. Consumer ClickOnce

- HTTPS stable deployment URL.
- Signed manifests và immutable version directory.
- Kiểm tra/tải/stage update mà không block Word startup lâu.
- Chỉ apply khi Word đã đóng hoặc lần mở kế tiếp.
- Required minimum version được server enforce.
- Không hứa zero-click tuyệt đối trên mọi unmanaged machine.
- Signed updater service/scheduled task không thuộc MVP nếu chưa có quyết định rõ.

### 16.2. Enterprise

- Signed MSI/bootstrapper.
- Detect Office bitness và registry view.
- Cài/check .NET Framework và VSTO Runtime.
- Intune/Configuration Manager supersedence.
- Không chạy ClickOnce self-update đồng thời.
- Có thể nghiệm thu silent update khi endpoint được quản lý.

### 16.3. Release pipeline

- Development → Internal → Pilot → Stable → Legacy Maintenance.
- Promote cùng artefact bất biến; không rebuild giữa ring.
- Release có version, commit, build ID, hashes, SBOM, signatures, dependency scan, secret scan và test evidence.
- Rollout tự pause khi health vượt ngưỡng.
- Minimum version, revoke và kill switch có impact preview, reason, re-authentication và approval.
- Forward-rollback binary bằng last-known-good được build/ký dưới version cao hơn.
- Có certificate rollover, key rotation, compromise và revoke drill.

## 17. Test bắt buộc

Không được tuyên bố hoàn thành nếu thiếu các nhóm test trong Mục 21 của plan:

- Unit tests.
- Contract/schema tests.
- Rule positive, negative và boundary tests.
- Golden text/document corpus.
- Word integration tests.
- Document safety tests.
- Compatibility VM matrix.
- Installer/update/forward-rollback/uninstall tests.
- Authentication/device/lease/replay tests.
- Trial boundary/non-stacking/concurrency tests.
- Offer/quote/payment/webhook/reconciliation tests.
- Admin/Customer Portal functional tests.
- RBAC, maker-checker và tenant isolation tests.
- Security tests.
- Privacy/redaction tests.
- Accessibility/responsive tests.
- Performance/load tests.
- Failure injection và disaster recovery drills.

Test Word thật tối thiểu theo matrix trong plan. Mock Interop không thay thế real Word evidence.

Mỗi test report phải phân biệt:

- PASS có evidence.
- FAIL có lỗi và reproduction.
- BLOCKED_ENVIRONMENT do thiếu toolchain/VM/service.
- BLOCKED_DECISION do thiếu quyết định sản phẩm.
- NOT_RUN có lý do.

Không biến NOT_RUN hoặc BLOCKED thành PASS.

## 18. Definition of Done

Mục 26 của kế hoạch là ràng buộc đầy đủ. Ngoài ra, trước khi tuyên bố hoàn thành phải tự động xác minh:

- 38/38 Ribbon controls target.
- 31/31 command button handlers thật.
- 68/68 VBA modules có disposition và evidence.
- Không còn callback giả, hardcoded license/price, localhost production URL hoặc debug bypass.
- Scan không thay đổi document fingerprint hoặc protected structures.
- AutoFix có signed FixPlan, preview, Undo, backup và rollback.
- Canonical rule counts không mâu thuẫn.
- Trial/payment/entitlement state machine đạt boundary và concurrency tests.
- Client giả không entitlement không nhận premium output; clone/re-signed client dùng account/device hợp lệ chỉ được đánh dấu PASS theo ADR CLIENT_AUTHENTICITY_BOUNDARY và evidence thực tế.
- Official artefact sửa byte không load qua signed deployment.
- Admin Portal đủ module, MFA, server RBAC, maker-checker, immutable audit và WCAG.
- Tenant isolation đạt cả sync và async paths.
- Update coverage/source được hiển thị trung thực.
- Binary forward-rollback và rule/config rollback đã diễn tập.
- Không có P0/P1 mở trước production.

## 19. Quy tắc báo cáo trong quá trình thực hiện

Trong mỗi cập nhật, báo ngắn gọn:

1. Kết quả vừa hoàn thành.
2. File/module đã thay đổi.
3. Test đã chạy và kết quả.
4. Gate hiện tại.
5. Blocker thật sự nếu có.
6. Bước đang tiếp tục.

Không yêu cầu người dùng xác nhận cho các bước kỹ thuật an toàn, có thể đảo ngược và nằm trong phạm vi plan. Phải hỏi trước khi:

- Chọn hoặc ký hợp đồng với provider.
- Tạo/thay production credential.
- Thay đổi dữ liệu hoặc hạ tầng production.
- Gửi email/thông báo cho khách hàng.
- Mua certificate/license.
- Chốt ngày trial/giá/ngưỡng tài chính.
- Xóa artefact nguồn hoặc dữ liệu khó khôi phục.
- Thay đổi quyết định sản phẩm đã khóa.

## 20. Quy tắc chống đầu ra thiếu

- Không dùng phần code rút gọn.
- Không dùng comment thay cho implementation.
- Không để marker công việc chưa triển khai, stub, hàm rỗng hoặc nhánh trả thành công giả trong code được coi là hoàn thành.
- Không tạo skeleton rồi coi là xong.
- Không viết một ví dụ rồi bỏ qua các command/module còn lại.
- Không bỏ qua phần giữa của file vì giới hạn đầu ra.
- Không dùng câu mô tả rằng phần còn lại làm tương tự.
- Không đánh dấu hoàn thành khi mới scaffold.
- Không dừng để hỏi người dùng có muốn tiếp tục hay không.
- Không skip, disable hoặc nới assertion của test chỉ để pipeline xanh.
- Không dùng mock để thay thế acceptance evidence ở production boundary.
- Không hardcode price, trial, license, entitlement, secret, provider response hoặc security decision.
- Không bỏ bớt control, module, rule, role, API hoặc test vì chúng lặp lại.
- Không dùng prototype hiện tại làm bằng chứng parity hoặc production readiness.
- Không nới document safety, authorization, tenant isolation hoặc signature validation để test qua.

Nếu một lượt làm việc không đủ để hoàn thành toàn dự án:

1. Dừng tại ranh giới sạch sau khi code/test đã ổn định.
2. Cập nhật Implementation_Status và Test_Evidence.
3. Ghi chính xác phase, workstream, file và test tiếp theo.
4. Tiếp tục ở lượt sau từ điểm đó, không làm lại phần đã hoàn tất.

## 21. Việc phải làm ngay khi nhận prompt

1. Xác nhận workspace và đọc toàn bộ nguồn sự thật.
2. Kiểm tra có AGENTS.md/CONTEXT.md; chạy git status/diff nếu workspace là repository và ghi nhận trạng thái thay đổi hiện hữu.
3. Tạo inventory machine-readable của code, tài liệu, Ribbon, rules và modules.
4. Chạy lại các phép đếm baseline.
5. Tạo shared\docs\implementation và các ledger/evidence ban đầu.
6. Lập execution board bám đúng các phase và exit gate, không viết lại một kế hoạch chung chung.
7. Ghi các BLOCKED_DECISION production đã biết nhưng tiếp tục phần độc lập.
8. Bắt đầu Giai đoạn 1 và Giai đoạn 2; song song chuẩn bị scaffold production ở Giai đoạn 3–4 khi dependency cho phép.
9. Thực hiện code và test, không chỉ báo cáo.

## 22. Hình thức bàn giao cuối

Khi thực sự đạt terminal condition, bàn giao:

- Tóm tắt sản phẩm đã hoàn thành.
- Liên kết các artefact chính.
- Danh sách phase/exit gate và evidence.
- Danh sách test PASS/FAIL/BLOCKED/NOT_RUN.
- Ma trận Word/Windows/bitness đã kiểm thử thật.
- Báo cáo parity Ribbon và VBA migration.
- Báo cáo rules reconciliation.
- Báo cáo security/privacy.
- Báo cáo Admin/Customer Portal.
- Báo cáo trial/commercial/payment.
- Báo cáo installer/update/forward-rollback.
- Danh sách quyết định production còn cần chủ sản phẩm cung cấp.
- Hướng dẫn build, run, install, update, rollback, backup, restore và support.

Không tuyên bố production-ready nếu còn thiếu certificate, provider verification, VM evidence, pentest, backup/restore drill hoặc P0/P1 gate.
