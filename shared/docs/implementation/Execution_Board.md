# Execution Board

| Phase | Workstream | Trạng thái | Exit gate/evidence |
| ---: | --- | --- | --- |
| 0 | Product decisions và threat model | IN_PROGRESS | ADR nền đã ghi; production decisions còn BLOCKED_DECISION |
| 1 | Ribbon contract và VBA baseline | IN_PROGRESS | Baseline VBA 45 control được giữ làm provenance; target hiện hành 42 control (34 button, 3 menu, 2 dropdown, 3 checkbox), bỏ Save-as-DOCX, `Đọc dữ liệu`, i/y và QR; gộp dọn khoảng trắng/dấu câu/dấu ngoặc vào `Sửa nhanh chính tả`; 68-module ledger có đủ; cần golden parity, visual bản hiện tại, Word 2010/x86 và sign-off |
| 2 | Canonical rules và corpus | IN_PROGRESS | Closed schema/parser và draft 96 rule đã có; legal/detector/golden corpus chưa đạt |
| 3 | VSTO production foundation | IN_PROGRESS | DOC/DOCX, signed local lease/rules, background refresh, exact-anchor annotation, 34/34 button có handler và local scan E2E đã có; Word 16 x64 COM load và smoke command PASS. Cần injected-failure rollback, production key/signing, visual bản hiện tại và Word 2010/x86 |
| 4 | Server, identity, security, Admin A | IN_PROGRESS | .NET 10/API/PostgreSQL nền; Development Admin local có user/trial/versioned price. Production identity/MFA/RBAC/audit/API đầy đủ còn thiếu |
| 5 | Local compliance engine và signed rules | COMPLETE_SOURCE_IN_PROGRESS_APPROVAL | 73/73 route baseline còn trong sản phẩm đã port local, cộng 4 detector Line Shape có nguồn trực tiếp và exact anchors; 2 route tone/i-y đã loại; rule pack và lease RS256. Chờ legal review, golden corpus và Word 2010/x86 evidence |
| 6 | Command waves A–E | COMPLETE_DEVELOPMENT_WORD16_X64 | Chuyển Unicode, ba bộ style, ba cỡ chữ, tiện ích bảng/shape, `Sửa nhanh chính tả` và 1-Click đã nối local; QR retired; DOC/DOCX command smoke PASS. Word 2010/x86 và golden parity vẫn mở |
| 7 | Trial, commercial, Admin B/C/D, Customer Portal | IN_PROGRESS_FOUNDATION | Domain trial/offer và Development Admin có; production IdP/DB/RBAC/provider/date/price vẫn blocked |
| 8 | Signing, update, Admin E, installer | COMPLETE_DEVELOPMENT_BLOCKED_PRODUCTION | EXE Development rule-only `.90` đã build/cài/repair/gỡ/cài lại; payload allowlist 7 file, không AI/ONNX/QR. Production certificate, HTTPS immutable release và KMS/HSM vẫn blocked |
| 9 | Pilot, Admin F, migration, launch | PENDING_DEPENDENCY | Chờ toàn bộ exit gates và VM/pentest evidence |

Board chỉ chuyển phase khi exit gate có file, test ID và evidence thực tế.
