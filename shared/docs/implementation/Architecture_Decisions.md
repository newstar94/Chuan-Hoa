# Architecture Decisions

## ADR-001 — VSTO-only Word client

- Trạng thái: ACCEPTED.
- Quyết định: client sản phẩm là Word VSTO Add-in C# cho Word 2010+ trên Windows.
- Hệ quả: Web Add-in hiện tại chỉ là reference artefact; không được ship hoặc dùng làm parity evidence.
- Kiểm chứng bắt buộc: project VSTO thật, signed manifests, Word VM matrix x86/x64.

## ADR-002 — Signed rules, local-authoritative document processing

- Trạng thái: SUPERSEDED ngày 2026-09-02 bởi yêu cầu sản phẩm.
- Quyết định hiện hành: toàn bộ đọc snapshot, detector, kiểm tra thể thức/chính tả, lập finding và mutation tài liệu chạy trong Word/VSTO tại local. Server không nhận nội dung, filename, path, binary DOC/DOCX, ảnh, comments, chữ ký hoặc con dấu.
- Server xác minh identity/device/entitlement/release và phát hành rule pack ký RS256 cùng offline lease ký RS256 tối đa 7 ngày. Client kiểm tra signature, key id, loại artefact, thời gian, thiết bị, phiên bản và feature trước khi dùng.
- Hệ quả: khi mất mạng, cache hợp lệ tiếp tục chạy đến hết hạn lease; sau đó feature có phí fail closed nhưng Word vẫn mở. Startup không chờ mạng; refresh chạy nền và Ribbon được invalidate trên UI thread.

## ADR-003 — Một Ribbon, không persistent task pane

- Trạng thái: ACCEPTED.
- Quyết định: đúng một tab Ribbon theo artefact VBA; chỉ dialog tạm thời cho login, progress, preview, warning và About.
- Hệ quả: `TaskPaneControl` của prototype bị RETIRE sau khi đối chiếu, không được đưa vào project VSTO mới.

## ADR-004 — CORRECTION-RULE-ROUTE-BASELINE-001

- Trạng thái: ACCEPTED_TECHNICAL.
- Evidence: `evidence/rule_reconciliation.json` và script `tools/baseline/generate_inventory.py`.
- Quan sát: 96 definitions; 94 registered routes; hai definition chưa route là `ND30-PL1-M1-K4-ENC` và `ND30-PL1-M1-K4-NFC`.
- Ba implementation `CheckComponentUnderline`, `CheckComponentNeverDetected` và `CheckCapitalizationNotDetectable` luôn trả `Nothing`; chúng nhận lần lượt 3, 5 và 11 routes, tổng 19.
- Kết luận: baseline có 75 routes với đường logic, không phải 77. Prompt và kế hoạch trong project mới đã được sửa công khai; nguồn workspace cũ không bị thay đổi.

## ADR-005 — CLIENT_AUTHENTICITY_BOUNDARY

- Trạng thái: BLOCKED_DECISION.
- Ranh giới: OAuth public client, device key, TPM và client-reported hash không tự phân biệt official DLL với clone chạy dưới cùng user/device/entitlement.
- Phương án cần chọn: managed endpoint với WDAC/App Control publisher policy; trusted updater/helper hoặc attestation service có threat model riêng; platform attestation đã kiểm chứng; hoặc chấp nhận giới hạn consumer và sửa claim.
- Cấm dùng làm bằng chứng: obfuscation, self-hash, strong name hoặc certificate nhúng trong DLL.
- Acceptance: clone/re-signed client dùng account, token, device và entitlement hợp lệ phải được test; chưa có boundary được chọn thì kết quả là BLOCKED_DECISION, không phải PASS.

## ADR-006 — ASP.NET Core trên .NET 10 LTS

- Trạng thái: ACCEPTED_TECHNICAL.
- Căn cứ: metadata release chính thức của Microsoft truy vấn ngày 2026-09-01 cho channel 10.0: release type LTS, SDK 10.0.400, EOL 2028-11-14.
- Quyết định: backend mới target `net10.0`; SDK được đặt cục bộ trong `.tools` để không thay đổi toolchain toàn máy.
- Hệ quả: source và CI khóa SDK; production runtime/container phải dùng patch được hỗ trợ và qua dependency scan.

## ADR-007 — PostgreSQL idempotency và transactional outbox

- Trạng thái: ACCEPTED_TECHNICAL.
- Quyết định: mutation dùng scoped idempotency record với SHA-256 key/request binding; business state và outbox message phải ghi trong cùng PostgreSQL transaction.
- Replay: cùng key/cùng request trả đúng response đã lưu; cùng key/khác request trả conflict; request đang sở hữu record không được chạy đồng thời.
- Hệ quả: không endpoint nghiệp vụ nào được mở trước khi auth scope và idempotency replay được nối hoàn chỉnh; chỉ kiểm tra header không được coi là exactly-once.
- Evidence: `DB-MIGRATION-V001-001` và `DB-PERSISTENCE-001` trên PostgreSQL 17.11.

## ADR-008 — VSTO source foundation fail closed

- Trạng thái: ACCEPTED_TECHNICAL.
- Quyết định: project production mới ở `src/ChuanHoa.AddIn.Vsto` dùng đúng một Ribbon 39 control, không task pane và không giữ `Word.Document` tĩnh.
- Command chưa port hoặc chưa vượt exit gate bị disabled và không có handler mô phỏng. About, ba tùy chọn hiển thị, hai lane scan và các tiện ích Word local đã có handler thật; không còn nút `Đọc dữ liệu`, mỗi command tự chuẩn bị đúng lane cần dùng. Mỗi mutation local yêu cầu signed `DOCUMENT_TOOLS` và preflight. Chỉ `Chuẩn hóa toàn bộ` cùng hai lệnh đổi cách đặt dấu tạo recovery copy; thao tác cục bộ dùng Word Undo và không clone/save cưỡng bức.
- AutoFix dùng callback riêng `OnAutoFixAll2026`; không còn nối nhầm vào định dạng trang giấy.
- Evidence: `VSTO-SOURCE-001` PASS; `VSTO-BUILD-001` PASS_LOCAL_DEVELOPMENT; `VSTO-WORD16-X64-001` PASS_LOCAL_SMOKE. Production signing, Word 2010/x86, Word COM safety adapter và command parity vẫn chưa đạt.

## ADR-010 — Bỏ Save-as-DOCX, xử lý trực tiếp DOC và DOCX

- Trạng thái: ACCEPTED_PRODUCT ngày 2026-09-01.
- Quyết định: loại bỏ `btnLuuDocx`, entitlement `SAVE_DOCX` và callback tương ứng khỏi target VSTO; giữ nguyên VBA extracted và `ribbon_actual.json` làm provenance.
- Định dạng xử lý được phép là tài liệu đã lưu `.doc` với Word binary format, hoặc `.docx` với Word Open XML transitional/strict format. So khớp không phân biệt hoa thường nhưng phải khớp cả extension và `Document.SaveFormat`.
- Không tự chuyển `.doc` sang `.docx`, không thay đổi định dạng lưu hiện tại. `.docm`, template, RTF, tài liệu chưa lưu hoặc extension/SaveFormat không khớp đều fail closed.
- Compatibility Mode không phải blocker. Mọi mutation vẫn phải vượt authorization, backup, fingerprint, Undo/rollback và exit gate hiện hữu.

## ADR-011 — Bỏ i/y, giữ hai lệnh đồng nhất vị trí dấu thanh

- Trạng thái: ACCEPTED_PRODUCT ngày 2026-09-01.
- Quyết định cập nhật theo yêu cầu sản phẩm: loại bỏ `mnuIY`, `btnKieuI` và `btnKieuY`; giữ `mnuBoDau`, `btnKieuOaUy` và `btnKieuOaUy2` để người dùng chủ động đồng nhất `oà/uý` hoặc `òa/úy`.
- Hệ quả: không phát hành handler cho `IY_NORMALIZE`; lệnh đặt dấu là mutation toàn tài liệu nên có recovery copy trong Windows Temp.
- VBA extracted và `ribbon_actual.json` giữ nguyên làm provenance, không còn là target sản phẩm.

## ADR-012 — Tên sản phẩm hiển thị là “Chuẩn hóa”

- Trạng thái: ACCEPTED_PRODUCT ngày 2026-09-01.
- Quyết định: tab Ribbon và tiêu đề sản phẩm trong VSTO dùng đúng tên “Chuẩn hóa”; không dùng “CHUẨN HÓA THỂ THỨC (2026)” hoặc biến thể kèm năm.
- Nhãn chức năng như “CHUẨN HÓA TOÀN BỘ” vẫn là tên command, không phải tên sản phẩm.
- Baseline VBA và evidence lịch sử giữ nguyên tên cũ làm provenance.

## ADR-013 — Comment có neo chính xác và tô chữ đỏ có sở hữu

- Trạng thái: ACCEPTED_PRODUCT ngày 2026-09-01.
- Quyết định: mỗi finding cục bộ phải có neo machine-readable gồm story, paragraph, offset, length và exact expected text. Comment ghi mã quy tắc, mức độ, hiện trạng, yêu cầu đúng và căn cứ; chữ trong phạm vi sai được tô đỏ.
- Lỗi paragraph-level mới tô cả paragraph. Lỗi section/page setup hoặc toàn tài liệu chỉ neo comment an toàn, không tô nội dung giả định.
- Client phải so khớp document fingerprint, revision và exact expected text; không fallback tìm occurrence khác. Neo mơ hồ, stale, protected hoặc không hỗ trợ trả `ANCHOR_UNRESOLVED` và không annotation.
- Comment dùng marker theo lane/finding; red marker có bookmark và document-variable state để khôi phục màu gốc. Clear chỉ xóa marker của lane do add-in sở hữu; không xóa comment/màu của người dùng hoặc lane khác.
- Scan vẫn read-only đối với nội dung/format nghiệp vụ; annotation là presentation mutation riêng theo yêu cầu người dùng. `btnKiemTra` chỉ được mở sau khi backend findings, entitlement, rollback và real-Word test đạt exit gate.
- Evidence: `ANNOTATION-PLANNER-001` PASS_UNIT; `ANNOTATION-VSTO-ADAPTER-001` PASS_WORD16_X64_LOCAL_SMOKE cho apply/rerun/clear và bảo toàn comment/màu người dùng. Injected-failure rollback, golden DOC/DOCX và end-to-end scan còn NOT_RUN.

## ADR-014 — Development tách biệt Release

- Trạng thái: ACCEPTED_TECHNICAL ngày 2026-09-02.
- Development bootstrap và Admin chỉ bật khi environment là `Development`, cờ `ChuanHoa:EnableDevelopmentBootstrap=true` và request đến từ loopback.
- Private key Development nằm ngoài source trong `.dev-secrets`; VSTO chỉ giữ public trust key ở profile local. Release không biên dịch `CHUANHOA_DEVELOPMENT`, không chứa trust key Development và fail closed cho đến khi có production endpoint/key.
- Admin Development quản lý dữ liệu thử local; không phải production IdP/RBAC/payment evidence.

## ADR-009 — Canonical rule publication fail closed

- Trạng thái: ACCEPTED_TECHNICAL.
- Quyết định: `shared/rules/rules_*.json` và backend prototype không phải canonical. Canonical v1 dùng closed JSON Schema, strict .NET parser và publication validator tại `src/ChuanHoa.Rules`.
- Baseline logic path chỉ là provenance observation, không đồng nghĩa implementation đã verified.
- Release `Published` yêu cầu từng rule active, legal traceability approved, detector/engine verified, đủ positive/negative/boundary fixtures approved và fix policy không blocked.
- Evidence: `RULE-CANONICAL-001` PASS; baseline 96 rule là `PASS_DRAFT_ONLY` và `publishable=false`.

## ADR-015 — Loại bỏ hoàn toàn tính năng QR

- Trạng thái: ACCEPTED_PRODUCT ngày 2026-09-04.
- Quyết định: QR không còn thuộc sản phẩm hiện tại. Loại `btnChenQrCode`, callback, dialog, renderer, QRCoder, test thao tác QR và binary QR khỏi installer.
- Theo quyết định sản phẩm ngày 2026-09-05, nhóm Hiển thị và ba checkbox bị loại; target Ribbon có 34 button, 3 menu, 2 dropdown, không checkbox và 39 control tương tác.
- VBA extracted, `ribbon_actual.json` và migration ledger được giữ làm provenance, không được dùng để tự sinh lại QR vào target.
