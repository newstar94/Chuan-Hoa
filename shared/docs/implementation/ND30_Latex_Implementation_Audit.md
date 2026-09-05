# Audit triển khai NĐ30/HD05 và khuyến nghị LaTeX/Typst

## Baseline xác minh ngày 05/09/2026

- Repository: `D:\Chuẩn Hóa`
- Branch: `main`
- `HEAD = origin/main = 73f9d96f8eb04035b08a6cbf831a4576c6ed84e1`
- Remote đã được kiểm tra bằng `git fetch --prune origin`.
- ProductVersion trước triển khai: `1.0.0.94`.
- ProductVersion đã reserve cho triển khai: `1.0.0.95`.
- Worktree trước triển khai chỉ có file prompt chưa track.
- Solution baseline: 272/272 test PASS (Core 227, Contracts 4, Domain 14, Rules 10, Application 6, API 11).
- Validator baseline: `SOLUTION_PROJECTS`, `RULE_ONLY_PRODUCT`, `VSTO-SOURCE-001` PASS.
- Ribbon source: 1 tab, 6 group, 34 button, 3 menu, 2 dropdown, 0 checkbox, tổng 39 control.
- VSTO build/load của source commit mới: `NOT_RUN`; artifact `.94` không được dùng làm bằng chứng cho commit này.

## Bảo toàn tài liệu nguồn

| Tệp | Size | LastWriteTimeUtc | SHA-256 trước triển khai |
|---|---:|---|---|
| `Nghị định 30.doc` | 530432 | `2026-08-31T17:46:02.1451433Z` | `0C875C3C69D081892A41E5B3C1A5B731A1728AF141AAADFB013978C2EF67B793` |
| `Phụ lục Nghị định 30.doc` | 1625088 | `2026-06-21T05:36:11.0000000Z` | `DCA5247B53E65492F67061C52CEF3135104E44A267143568DD2F8FF5348ADB00` |
| `Hướng dẫn 05.docx` | 122999 | `2026-08-31T17:46:40.0660962Z` | `DFD22896C32C8F2EC273176F3E92F9ED724019103011EF646C779DBB572BB3C2` |

Hash phải được kiểm tra lại sau toàn bộ Word smoke. Không được Save/Save As đè ba tệp này.

## Artifact `.94` bất biến

| Artifact | Size | SHA-256 | Trạng thái chữ ký trên host audit |
|---|---:|---|---|
| `artifacts/installers/development/ChuanHoa_Development_Test_Setup_1.0.0.94.exe` | 263976 | `299B543F542D8CFC25222A2A9C4C47ADFF37FC627FEF3B654FB071276356E98D` | `UnknownError`; không coi là production trust |

`.94` có lỗi bootstrapper allowlist thiếu `Microsoft.Office.Tools.Common.v4.0.Utilities.dll`, nên có thể dừng với thông báo “Payload không đúng allowlist”. Artifact này được giữ bất biến để bảo toàn lịch sử; lỗi đã được sửa từ `.95` và không tái xuất hiện trong `.97`.

## Ma trận audit code trước triển khai

| Trạng thái | Phát hiện | Bằng chứng chính | Hướng xử lý |
|---|---|---|---|
| Đã có | Scan/mutation tài liệu chạy local; lifecycle không chủ động quét nội dung | `ThisAddIn.cs`, `RibbonRuntime.cs`, `WordDocumentReadRuntime.cs` | Giữ invariant và bổ sung regression gate |
| Đã có | Comment hiện có hợp đồng hai dòng và annotation ownership | `AnnotationPlanner.cs`, `WordFindingAnnotationAdapter.cs` | Giữ hai dòng; tách visual theo severity/source |
| Đã có | Backup bulk mutation ở `%TEMP%\ChuanHoa\Backups` với retention hữu hạn | `WordRecoveryCopyManager.cs` | Giữ policy; không tạo backup cho sửa cục bộ |
| Một phần | `HeadingDetector` có decimal/Roman/article/unnumbered | `HeadingDetector.cs` | Xây cây theo parent/block/scheme và chống false positive |
| Nguy hiểm | Dòng khoản/list ngắn có thể bị coi là heading; `AlphabetPattern` không được dùng | `HeadingDetector.cs` | Không coi a/b/c là heading; yêu cầu evidence kết hợp |
| Nguy hiểm | Role Word snapshot là `Unknown`; scanner advisory không nhận role/block canonical | `WordDocumentSnapshotBuilder.cs`, `CanonicalRuleScanner.cs` | Tách raw snapshot và derived analysis context |
| Thiếu | Continuity không khóa parent prefix/block, thiếu backward/missing-parent/level-jump | `HeadingDetector.cs` | Viết Heading Tree và full negative/boundary corpus |
| Nguy hiểm | Scanner LaTeX chạy vô điều kiện và policy hard-code trong DLL | `CanonicalRuleScanner.cs`, `LocalRulePack.cs` | Signed `AcademicTypography` profile, mặc định tắt |
| Thiếu | Snapshot bảng chỉ có boolean viền dọc; caption Hình chưa có association | `WordDocumentSnapshotBuilder.cs`, `LatexTypographicScanner.cs` | Tri-state border/style/weight/range và raw object facts |
| Nguy hiểm | Suggestion LaTeX vẫn có thể bị tô đỏ như lỗi pháp lý | `AnnotationPlanner.cs` | Presentation policy theo severity/family |
| Nguy hiểm | Finding LaTeX không hỗ trợ có thể rơi vào generic paragraph auto-fix | `WordOneClickRuntime.cs` | Explicit allowlist, typed not-supported, post-scan |
| Nguy hiểm | Sửa cục bộ xóa cache/comment trước khi hậu kiểm thực tế | `WordOneClickRuntime.cs` | Re-snapshot/re-run chính rule trước khi xóa annotation |
| Nguy hiểm | Prefix `CHUANHOA2_` có thể bị nhận là legacy; Shape người dùng có thể bị nhận ownership | `WordOneClickRuntime.cs` | Ownership/version predicate chính xác và idempotency test |
| Nguy hiểm | Body formatter xóa custom TabStops quá rộng | `WordOneClickRuntime.cs` | Chỉ sửa tab ở role/policy chắc chắn |
| Thiếu | Fingerprint chưa chứa Keep/Widow/style/border raw state | `WordDocumentSnapshotBuilder.cs` | Tách document fingerprint và scan/cache key |
| Thiếu | VSTO build graph dùng HintPath + `--no-restore` dễ phụ thuộc cache máy | `ChuanHoa.AddIn.Vsto.csproj` | Restore/build dependency rõ và clean-copy gate |
| Nguy hiểm | `ApplicationVersion` có thể khác `ProductVersion`; activation rollback chưa đủ evidence | publish/builder/bootstrapper scripts | Một nguồn version, pin public key và rollback test cô lập |

## Quyết định safe-fix của release `1.0.0.95`

- `LATEX-SEC-STYLE`: comment-only.
- `LATEX-SEC-CONTINUITY`: comment-only.
- `LATEX-PAGINATION-KEEP`: chỉ auto-fix khi signed allowlist bật và confidence đủ.
- `LATEX-PAGINATION-WIDOW`: chỉ auto-fix khi signed allowlist bật và Body classifier đủ.
- `LATEX-TABLE-BOOKTABS`: comment-only.
- `LATEX-CAPTION-POS`: comment-only.
- `LATEX-MATH-SYNTAX`: comment-only.

## Kết quả sau triển khai

| Phát hiện ban đầu | Trạng thái 1.0.0.95 | Bằng chứng |
|---|---|---|
| Scanner advisory chạy vô điều kiện | RESOLVED | Rule-pack v2 typed `AcademicTypography`; signed OFF/ON Word smoke PASS |
| Heading/list/căn cứ và continuity nhầm block/parent | RESOLVED | Positive/negative/boundary corpus; Roman canonical; 349 Core test PASS |
| Snapshot thiếu raw state/table/caption/protected range | RESOLVED | Snapshot schema v3 và fingerprint mutation smoke trên Word PASS |
| Suggestion bị tô đỏ/nhầm lỗi pháp lý | RESOLVED | Typed family/severity; LATEX Suggestion không có visual red marker |
| Generic auto-fix/xóa comment trước hậu kiểm | RESOLVED | Explicit safe-fix matrix; selected-fix và 1-Click recapture/post-scan theo tọa độ |
| Line Shape ownership và idempotency | RESOLVED | Exact released prefix; dashed/sai geometry vẫn báo; Shape người dùng không bị chiếm/xóa |
| Xóa custom TabStops | RESOLVED | Không còn `TabStops.ClearAll()`; chỉ thêm canonical tab khi thiếu |
| Build graph/version split-brain | RESOLVED | VSTO `ProjectReference`; shared build contract; assembly/file/product/manifest/lease cùng `1.0.0.95` |
| Installer bootstrapper/audit allowlist lệch nhau | RESOLVED | Thêm VSTO Utilities DLL và validator so sánh hai allowlist; install/repair/uninstall/reinstall PASS |

Full solution build 0 warning/0 error và 402/402 test PASS. Các validator `SOLUTION_PROJECTS`, `RULE_ONLY_PRODUCT`, `DEVELOPMENT_PACKAGING`, `VSTO-SOURCE-001` PASS. Word 16 x64 runtime smoke PASS cho 39 control, DOC/DOCX/Document1, scan lặp, annotation, 1-Click, tài liệu lớn, ba tài liệu nguồn và signed profile OFF/ON.

Hash sau Word smoke của ba tài liệu nguồn trùng tuyệt đối bảng ban đầu. Artifact `.94` vẫn 263.976 byte và SHA-256 `299B543F542D8CFC25222A2A9C4C47ADFF37FC627FEF3B654FB071276356E98D`.

Bộ cài Development `.95` là 317.224 byte, SHA-256 `0A4D94650179DE16E708A3237FC0AA9B20173F3C10E42DBD4273C7C04291C541`; audit 8 file/0 forbidden hit, upgrade `.94→.95`, repair, uninstall/reinstall, access smoke, COM load và rollback 6/6 fault point PASS. Word 2010/x86 và visual screenshot Ribbon/icon `.95` là `NOT_RUN`; production signing/timestamp, inner PE Authenticode, Apps & Features, HTTPS immutable release và KMS/HSM vẫn là production blocker.

## Bằng chứng hiện hành release `1.0.0.97`

- Full solution restore/build/test vẫn đạt 0 warning, 0 error và 402/402 test; bốn validator cùng VSTO Development clean-copy build PASS.
- Annotation fault injection PASS: lỗi sau visual đầu tiên phục hồi nguyên trạng comment/màu và registry annotation trước đó.
- Benchmark format+spelling Word 16 x64, warmup 1 và đo 10 vòng: median 6.170,5 ms, p95 7.340 ms; working set 362.500.096→364.519.424 byte, private bytes 257.605.632→256.462.848 byte, handle 1.873→1.794, GDI 169→169, USER 161→154; WINWORD thoát sạch.
- Bộ cài `.97` là 314.152 byte, SHA-256 `FCAEB75A529CB6653F4478E50B3A238EA35C4745DEDB197074A1B7B13F4A1486`, Authenticode `Valid`, payload 8 file và 0 marker AI/ONNX/QR. Fresh install, repair, silent uninstall/reinstall, rollback 6/6 và COM load online/offline PASS. Certificate lifecycle dùng `certutil.exe` hidden, pin subject/SHA-256 và hậu kiểm store; không còn cửa sổ Root Certificate Store trong `/quiet`.
- `.94`, `.95` và `.96` được giữ bất biến. Visual screenshot Ribbon/icon của chính `.97` và Word 2010/x86 vẫn `NOT_RUN`; production signing/timestamp, inner PE Authenticode, Apps & Features, HTTPS immutable release và KMS/HSM vẫn là production blocker.

## Bằng chứng hiện hành release `1.0.0.101`

- Startup hoàn toàn thụ động: không tạo runtime, không đăng ký document event, không đọc `ActiveDocument`/`Selection`, không tải license và không chạy scanner trước thao tác Ribbon chủ động đầu tiên. Passive startup smoke trên bản cài thật xác nhận `Connect=true`, 5 giây idle không đổi cache/backup, không background save/print và Word thoát sạch.
- `Sửa lỗi đang chọn` dùng finding/snapshot/role cache từ lần kiểm tra chủ động; không `Prepare`, không dựng snapshot, không detect/scan toàn tài liệu, bắt đầu trực tiếp ở trạng thái `Mutating` và chỉ xóa annotation đúng `FindingId` sau hậu kiểm mutation. Annotation fault rollback PASS.
- 1-Click có policy cho đủ 92 mã finding hiện hành: 85 mã literal do scanner phát và bảy mã `LATEX-*` typed. Validator chỉ chấp nhận mã nằm trong đúng method dispatch/report-only, không còn coi việc mã xuất hiện ngẫu nhiên trong helper/comment là bằng chứng xử lý. Lỗi có cách sửa xác định được dispatch vào handler style/section/Line Shape/text/dấu câu/chính tả; lỗi thiếu dữ kiện pháp lý hoặc cần quyết định/tái cấu trúc vẫn report-only và giữ annotation. DOC/DOCX/Document1/multiple-document/dashed-line/quick-spelling smoke PASS.
- Full solution 403/403 test PASS. Clean-copy `.101` không `.git`, `.tools`, `artifacts`, `bin`, `obj`, cache hoặc secret đạt restore/build 0 warning/0 error, 403/403 test, bốn validator và VSTO Development rebuild; installed registration không đổi.
- EXE `.101` là 319.272 byte, SHA-256 `A351A2764760444DFDBD17DBC5831C4588673670690343C507D82D4B0623D460`, Authenticode `Valid`, audit 8 payload/0 marker AI-ONNX-QR; cài và access/Word smoke PASS. Repair/uninstall-reinstall/rollback 6/6 fault point của đúng artifact `.101` PASS; từ điển cá nhân và state ký được giữ/khôi phục. Người dùng đã xác nhận visual tab/icon/con trỏ idle trên `.101`; Word 2010/x86 vẫn chưa có. Production blockers giữ nguyên.
