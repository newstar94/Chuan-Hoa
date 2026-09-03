# Execution Board

| Phase | Workstream | Trạng thái | Exit gate/evidence |
| ---: | --- | --- | --- |
| 0 | Product decisions và threat model | IN_PROGRESS | ADR nền đã ghi; production decisions còn BLOCKED_DECISION |
| 1 | Ribbon contract và VBA baseline | IN_PROGRESS | Baseline VBA 45 control được giữ; target 38 control đã bỏ Save-as-DOCX, Bỏ dấu và i/y theo ADR-010/011; 68-module ledger có đủ; cần làm mới visual Word, golden parity, Word 2010/x86 và sign-off |
| 2 | Canonical rules và corpus | IN_PROGRESS | Closed schema/parser và draft 96 rule đã có; legal/detector/golden corpus chưa đạt |
| 3 | VSTO production foundation | IN_PROGRESS | DOC/DOCX, signed local lease/rules, background refresh, exact-anchor annotation và local scan E2E có nền; 16 utility command đã nối local. Cần injected-failure rollback, production key/signing, Word 2010/x86 và command parity còn lại |
| 4 | Server, identity, security, Admin A | IN_PROGRESS | .NET 10/API/PostgreSQL nền; Development Admin local có user/trial/versioned price. Production identity/MFA/RBAC/audit/API đầy đủ còn thiếu |
| 5 | Local compliance engine và signed rules | COMPLETE_SOURCE_IN_PROGRESS_APPROVAL | 73/73 route baseline còn trong sản phẩm đã port local, cộng 4 detector Line Shape có nguồn trực tiếp và exact anchors; 2 route tone/i-y đã loại; rule pack và lease RS256. Chờ legal review, golden corpus và Word 2010/x86 evidence |
| 6 | Command waves A–E | IN_PROGRESS | 16 utility command local có signed feature gate, recovery copy và Undo; Unicode/style/QR/AutoFix còn khóa |
| 7 | Trial, commercial, Admin B/C/D, Customer Portal | IN_PROGRESS_FOUNDATION | Domain trial/offer và Development Admin có; production IdP/DB/RBAC/provider/date/price vẫn blocked |
| 8 | Signing, update, Admin E, installer | PENDING_DEPENDENCY | Chờ artefact buildable và production certificate |
| 9 | Pilot, Admin F, migration, launch | PENDING_DEPENDENCY | Chờ toàn bộ exit gates và VM/pentest evidence |

Board chỉ chuyển phase khi exit gate có file, test ID và evidence thực tế.
