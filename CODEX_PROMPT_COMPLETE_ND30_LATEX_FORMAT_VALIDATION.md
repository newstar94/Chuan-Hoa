# PROMPT THỰC THI: HOÀN THIỆN KIỂM TRA THỂ THỨC NĐ30/HD05 VÀ KHUYẾN NGHỊ LATEX/TYPST

## 1. Vai trò và kết quả phải bàn giao

Bạn là Codex chịu trách nhiệm trực tiếp rà soát, sửa code, bổ sung kiểm thử, kiểm chứng trên Microsoft Word, cập nhật tài liệu kỹ thuật và đóng gói bản Development Test của dự án Word VSTO Add-in **Chuẩn hóa** tại:

```text
D:\Chuẩn Hóa
```

Remote chính:

```text
https://github.com/newstar94/Chuan-Hoa.git
```

Đây là nhiệm vụ triển khai thực tế, không phải nhiệm vụ chỉ viết kế hoạch. Phải làm việc trên code mới nhất, bảo toàn các sửa đổi hợp lệ hiện có và chỉ tuyên bố hoàn thành khi có bằng chứng đạt các gate và Definition of Done trong prompt này.

Mục tiêu của đợt triển khai:

1. Hoàn thiện nút **Kiểm tra thể thức** để giữ nguyên tính đúng đắn của Nghị định 30/2020/NĐ-CP và Hướng dẫn 05-HD/VPTW, đồng thời hỗ trợ một nhóm khuyến nghị biên tập theo LaTeX/Typst trong cùng pipeline local.
2. Khắc phục các false positive, false negative và lỗi kiến trúc đang có trong `HeadingDetector`, `LatexTypographicScanner`, snapshot Word, annotation, cache và luồng sửa lỗi.
3. Phân biệt tuyệt đối giữa **quy tắc thể thức có thẩm quyền** và **khuyến nghị thẩm mỹ LaTeX/Typst**; khuyến nghị không được xuất hiện như lỗi pháp lý.
4. Giữ Word ổn định với `.doc`, `.docx`, tài liệu chưa lưu, file nhiều văn bản, file lớn và file nhiều bảng/Shape.
5. Chỉ tự sửa những finding có phương án an toàn, xác định và hậu kiểm được; không làm đẹp hàng loạt theo suy đoán.
6. Giữ sản phẩm rule-only, không AI/ONNX/QR, không gửi nội dung tài liệu lên máy chủ và không nới lỏng license/gói quy tắc có chữ ký.
7. Tạo một bộ cài Development Test mới bằng một file EXE phân phối, dùng phiên bản chưa từng được build và không ghi đè bản `1.0.0.94`.

## 2. Baseline đã xác minh ngày 05/09/2026

Trước khi sửa, phải kiểm tra lại baseline. Nếu code đã thay đổi sau mốc này, lập diff và tiếp tục từ code mới hơn; không reset hoặc quay lại snapshot cũ.

Baseline tại thời điểm lập prompt:

```text
Repository: D:\Chuẩn Hóa
Branch: main
HEAD: 73f9d96f8eb04035b08a6cbf831a4576c6ed84e1
origin/main: 73f9d96f8eb04035b08a6cbf831a4576c6ed84e1
Commit: feat: implement LaTeX typographic validation scanner and heading structure detection engine
ProductVersion: 1.0.0.94
Worktree: sạch
```

Evidence baseline đã lưu và phải tái lập trên máy/host thực thi trước khi sửa:

```text
ChuanHoa.Client.Core.Tests: 227 PASS
ChuanHoa.Contracts.Tests: 4 PASS
ChuanHoa.Domain.Tests: 14 PASS
ChuanHoa.Rules.Tests: 10 PASS
ChuanHoa.Application.Tests: 6 PASS
ChuanHoa.Api.Tests: 11 PASS
Tổng solution: 272/272 PASS
```

Baseline VSTO tĩnh:

```text
Target framework: .NET Framework 4.8
Minimum Office version: 14.0 / Word 2010
Ribbon: 1 tab, 6 group, 34 button, 3 menu, 2 dropdown, 0 checkbox
Tổng interactive control: 39
Callback methods: 51
validate_vsto_source.py: PASS về source
VSTO build gate cho source mới nhất: NOT_RUN
Word load/smoke gate cho source mới nhất: NOT_RUN
```

Artifact sau đã tồn tại và phải coi là bất biến:

```text
D:\Chuẩn Hóa\artifacts\installers\development\ChuanHoa_Development_Test_Setup_1.0.0.94.exe
D:\Chuẩn Hóa\artifacts\vsto-development-test\1.0.0.94\
```

Artifact `.94` không được ghi đè, chỉnh sửa hoặc dùng làm bằng chứng build/load cho code sau commit mới nếu hash/evidence không khớp. Ngay sau preflight, trước khi tạo bất kỳ binary, signed pack, lease hoặc Word smoke artifact nào từ source đã sửa, phải reserve và cập nhật một phiên bản lớn hơn mọi version đã dùng, tối thiểu `1.0.0.95`. Trước khi chọn số, quét `Directory.Build.props`, toàn bộ thư mục artifact, manifest, tag và tài liệu phát hành; chọn số chưa từng xuất hiện. Không tạo binary mới mang version `.94`.

Không chỉ truyền `-ApplicationVersion` khác với `ProductVersion`. DLL hiện lấy Assembly/File/ProductVersion từ `Directory.Build.props`; manifest và DLL khác phiên bản sẽ gây `LEASE_RELEASE_MISMATCH`. Phải dùng một nguồn version thống nhất và thêm gate chặn split-brain.

## 3. Nguồn tham chiếu và thứ tự thẩm quyền

Phải đọc đầy đủ trước khi sửa:

```text
D:\Chuẩn Hóa\shared\docs\KeHoach_KiemTra_TheThuc_ND30_va_LaTeX.md
C:\Users\newst\.codex\attachments\f31711c3-b39a-4670-9589-763cf9550886\pasted-text.txt
D:\Chuẩn Hóa\README.md
D:\Chuẩn Hóa\shared\docs\implementation\Implementation_Status.md
D:\Chuẩn Hóa\shared\docs\implementation\Test_Evidence.md
D:\Chuẩn Hóa\shared\docs\implementation\Release_Runbook.md
D:\Chuẩn Hóa\shared\docs\implementation\Rule_Catalog_Reconciliation.md
D:\Chuẩn Hóa\shared\docs\implementation\Local_Rule_Port_Ledger.md
D:\Chuẩn Hóa\shared\contracts\ribbon\ribbon-contract.v1.json
```

Ba tài liệu Word dưới đây là nguồn nghiệp vụ, chỉ được mở/convert/render theo chế độ chỉ đọc để đối chiếu; không được sửa, Save As đè, đổi metadata hoặc commit bản chuyển đổi:

```text
D:\Chuẩn Hóa\Nghị định 30.doc
D:\Chuẩn Hóa\Phụ lục Nghị định 30.doc
D:\Chuẩn Hóa\Hướng dẫn 05.docx
```

Hai tài liệu kế hoạch là nguồn ý định, không phải đặc tả đã được kiểm chứng. Khi mâu thuẫn, áp dụng thứ tự:

1. Yêu cầu trong prompt này và quyết định sản phẩm mới nhất.
2. Nghị định 30 cho văn bản hành chính và Hướng dẫn 05 cho văn bản Đảng.
3. Rule catalog/canonical contracts đã có traceability và test được duyệt.
4. Code production mới nhất đã có kiểm thử.
5. Khuyến nghị LaTeX/Typst chỉ bù đắp nhu cầu cấu trúc và thẩm mỹ, không được ghi đè các nguồn phía trên.

Không sửa prototype cũ trong các thư mục sau, trừ khi chỉ cập nhật ghi chú provenance có căn cứ rõ ràng:

```text
D:\Chuẩn Hóa\client-vsto-csharp
D:\Chuẩn Hóa\client-web-addin
D:\Chuẩn Hóa\backend-api
```

## 4. Sự thật về rule registry hiện tại

Không dùng tuyên bố “82 quy tắc NĐ30” làm baseline thực thi. `rules_compliance_checks.json` có con số 82 của catalog cũ, nhưng registry production tại commit hiện tại có:

```text
ND30-*: 70 mã
HD05-*: 1 mã
LOCAL-*: 6 mã
Registry hiện hành trước LaTeX: 77 mã
LATEX-*: 7 mã
AllRegisteredRuleCodes: 84 mã duy nhất
```

Giữ đúng bảy mã canonical hiện có:

```text
LATEX-SEC-STYLE
LATEX-SEC-CONTINUITY
LATEX-PAGINATION-KEEP
LATEX-PAGINATION-WIDOW
LATEX-TABLE-BOOKTABS
LATEX-CAPTION-POS
LATEX-MATH-SYNTAX
```

Không tạo biến thể `LATEX-SEC-DETECT` hoặc `LATEX-MATH-DELIMITER`. Nếu cập nhật catalog cũ, phải ghi rõ trạng thái legacy/provenance hoặc tạo mapping có kiểm chứng; không đổi số đếm để làm tài liệu trông nhất quán.

## 5. Audit hiện trạng phải được dùng làm điểm xuất phát

Commit `73f9d96` đã có nền móng sau; không viết lại từ đầu:

- `HeadingDetector.cs` và unit test cơ bản.
- `LatexTypographicScanner.cs` và bảy mã `LATEX-*`.
- Snapshot đã bắt đầu đọc `KeepWithNext`, `WidowControl`, `StyleName` và `HasVerticalBorders`.
- `CanonicalRuleScanner.ScanFormat()` đã gọi scanner LaTeX.
- Popup đã bắt đầu tách số finding LaTeX.
- Tổng test tăng từ 259 lên 272 và đang PASS.

Tuy nhiên phần mới chưa đủ điều kiện phát hành. Phải xác nhận và xử lý các lỗi sau:

1. Scanner LaTeX đang chạy vô điều kiện cho mọi regime và policy bị hard-code trong DLL, nằm ngoài signed advisory profile.
2. Regex heading quá rộng; dòng `1. Nội dung khoản` ngắn có thể bị coi là heading. `AlphabetPattern` được khai báo nhưng chưa dùng.
3. `Điều`, Khoản, Điểm và danh sách có thể bị nhận nhầm thành heading học thuật.
4. `WordDocumentSnapshotBuilder` đang đặt `Role="Unknown"`; role/block canonical do `DocumentRoleDetector` tạo ra chưa được truyền vào `HeadingDetector`/LaTeX scanner. Vấn đề là thiếu propagation/vocabulary thống nhất, không phải so sánh casing vì collection hiện dùng `OrdinalIgnoreCase`.
5. Continuity chỉ nhớ số cuối theo level, không khóa theo parent prefix hoặc logical document block; có thể so nhầm nhánh `1.1` với `2.1`. Chưa xử lý đầy đủ missing parent, nhảy cấp, số lùi, duplicate theo nhánh, first-gap và reset scheme.
6. Roman parser chấp nhận biểu diễn không canonical như `IIII` hoặc `IC`.
7. Caption mới suy luận caption Bảng từ paragraph lân cận; `FigureCaptionRegex` chưa được dùng và snapshot chưa liên kết hình với caption.
8. `ReadTableCoordinates()` chưa cung cấp `TableIndex` thật cho paragraph Word; unit test snapshot dựng tay không chứng minh integration trên Word thật.
9. Booktabs mới chỉ biết boolean viền dọc, chưa biết top/bottom/inside-horizontal/header separator, line style, weight hoặc trạng thái unknown.
10. Math scanner chưa loại đủ tiền tệ, escaped dollar, unmatched/nested delimiter, Field, OMath, code và protected range.
11. Widow dùng độ dài 120 ký tự làm proxy cho Body/multiline, dễ báo sai header, footer, chú thích, bảng và đoạn hẹp/rộng khác nhau.
12. `Style.NameLocal.StartsWith("Heading")` phụ thuộc ngôn ngữ Office; `Style` RCW và từng `Border` RCW mới lấy chưa được quản lý đầy đủ.
13. Lỗi đọc COM có thể bị biến thành `false`, tạo false negative thay vì trạng thái `Unknown`.
14. Fingerprint/cache chưa hash các thuộc tính mới ảnh hưởng finding: `KeepWithNext`, `WidowControl`, style identity và border/header state. Snapshot schema là 2 nhưng hash domain vẫn còn dấu vết tên version cũ.
15. `AnnotationPlanner` tô đỏ mọi paragraph/text finding mà không xét severity; `Suggestion` LaTeX đang có nguy cơ hiển thị như lỗi pháp lý.
16. Popup đang suy ra lỗi bắt buộc theo phép trừ tổng với LATEX, dễ gắn nhãn sai cho family khác và chưa có thống kê heading đáng tin cậy.
17. `WordOneClickRuntime` chưa có implementation xác định cho các rule LaTeX. Finding không hỗ trợ có thể rơi vào generic paragraph formatter, trả thành công và làm comment biến mất dù lỗi chưa được sửa.
18. Sửa cục bộ đang xóa finding khỏi cache thay vì hậu kiểm rule trên snapshot mới.
19. Line Shape thuộc sở hữu add-in có nguy cơ bị xóa/tạo lại mỗi lần vì predicate legacy `StartsWith("CHUANHOA")` cũng khớp prefix mới `CHUANHOA2_`.
20. Heuristic Shape hiện có nguy cơ nhận ownership hoặc xóa Shape của người dùng khi association chưa chắc chắn.
21. Xóa toàn bộ custom TabStops ở body có thể phá tab chủ ý trong danh sách, biểu mẫu hoặc đoạn bố cục đặc biệt.
22. VSTO dùng `HintPath` và custom build Client.Core với `--no-restore`; clean agent có thể thiếu `project.assets.json` nếu build graph không restore đúng thứ tự.
23. Script publish cho phép `ApplicationVersion` tách khỏi `ProductVersion`; bootstrapper chưa chứng minh rollback transaction ở mọi lỗi sau khi đổi `Current`.

## 6. Nguyên tắc bất biến

1. Văn bản hành chính dùng NĐ30; văn bản Đảng dùng HD05. Không dùng LaTeX để thay font, cỡ chữ, lề, quốc hiệu, tiêu ngữ, tên cơ quan, số/ký hiệu, trích yếu, căn cứ, Điều/Khoản/Điểm, chữ ký, nơi nhận hoặc mẫu phụ lục.
2. LaTeX/Typst là **khuyến nghị thẩm mỹ/biên tập**, không phải quy định pháp lý và luôn có `Severity=Suggestion`.
3. Dùng nút **Kiểm tra thể thức** hiện có. Không thêm Ribbon button, checkbox, tab hoặc task pane cho nhóm LaTeX.
4. Nội dung, đường dẫn và binary tài liệu không được upload. Capture, scan, annotation và mutation chạy local. Mạng chỉ được dùng ngoài command Word để refresh danh tính, lease 7 ngày và gói quy tắc có chữ ký.
5. Không gọi mạng đồng bộ trong Ribbon callback, scan hoặc mutation.
6. Không tự đọc/quét tài liệu ở startup, DocumentOpen, WindowActivate, Alt-Tab, Save, `getEnabled` hoặc khi chỉ chuyển focus Modern Comments. Chỉ lệnh do người dùng bấm mới được chuẩn bị snapshot cần thiết.
7. Word COM chỉ được truy cập trên UI/STA thread. Sau khi capture xong immutable DTO và không còn RCW, pure-.NET analysis có thể chạy trên worker có cancellation; trước annotate/mutate phải quay lại UI/STA và xác minh document identity, revision và fingerprint. Không dùng worker để gọi COM; không dùng `Application.DoEvents`, sleep, retry mù hoặc tăng timeout để che treo.
8. Hỗ trợ `.doc`, `.docx` và `Document1` chưa lưu. Không ép Save As và không đổi `.doc` thành `.docx` chỉ để chạy chức năng.
9. Không dùng `Marshal.FinalReleaseComObject`. Chỉ release RCW do code sở hữu, đúng một lần, theo thứ tự con trước cha.
10. Không xóa comment, màu, highlight, Style, Shape, TabStop hoặc border do người dùng sở hữu khi chưa có ownership/association chắc chắn.
11. Không nới license, signature, rule-pack expiry, device/client/feature binding, replay protection hoặc fail-closed security để làm nút sáng.
12. Không thêm lại AI/ML/ONNX/VietnameseEngine, IPC AI, training artifact hoặc QR.
13. Không thay đổi ba tài liệu Word nguồn.
14. Không sửa generated Ribbon output mà bỏ qua contract/generator canonical.
15. Không xóa hoặc làm yếu test hiện có để đạt PASS.

## 7. Trình tự triển khai bắt buộc

Thực hiện theo thứ tự sau; không bật auto-fix trước khi detector và snapshot đạt gate:

1. Preflight và tái lập baseline.
2. Chốt taxonomy, applicability, severity và signed advisory policy.
3. Reserve/cập nhật ProductVersion mới chưa từng dùng trước mọi release-bound build, pack, lease hoặc smoke artifact.
4. Mở rộng snapshot/fingerprint với trạng thái tri-state và COM ownership an toàn.
5. Sửa role/block propagation và Heading Tree.
6. Hoàn thiện bảy detector ở chế độ scan/comment-only trước.
7. Sửa annotation, popup, cache invalidation và hậu kiểm finding.
8. Chỉ mở các auto-fix được phép trong Safe-Fix Matrix.
9. Chạy unit/integration/source gates.
10. Chạy Word smoke trên source mới.
11. Cập nhật tài liệu/evidence, build/audit/install/repair/upgrade bộ cài mới.

Sau mỗi phase, chạy test tập trung. Không gom toàn bộ thay đổi rồi mới kiểm tra.

## 8. Phase 0 — Preflight và báo cáo ma trận hiện trạng

1. Đọc toàn bộ file policy/agent instruction trong repo nếu có.
2. Chạy `git fetch --prune origin`, sau đó chạy `git status --short --branch`, xác minh HEAD và origin. Nếu fetch bị chặn bởi network/authentication, ghi `REMOTE_NOT_VERIFIED` và không gọi local HEAD là “mới nhất từ GitHub”. Không tự `reset`, `checkout --` hoặc xóa thay đổi người dùng nếu worktree không còn sạch.
3. Lập dependency map bằng `rg` cho:
   - `HeadingDetector`;
   - `LatexTypographicScanner`;
   - `LocalScanSnapshot` và Word snapshot DTO;
   - `DocumentRoleDetector`/`LogicalDocumentBlock`;
   - `AnnotationFinding`, planner và Word adapter;
   - `WordOneClickRuntime`/selected fix;
   - `LocalRulePack`, parser, signature và development issuer;
   - fingerprint/cache identity;
   - installer/publish/bootstrapper.
4. Chạy lại 272 test, ba validator chính và static VSTO validator trước khi sửa.
5. Ghi một bảng `Đã có / Một phần / Thiếu / Nguy hiểm / Bằng chứng file:line` vào `shared/docs/implementation/ND30_Latex_Implementation_Audit.md`. Không ghi kết luận chỉ dựa trên tên class hoặc unit test snapshot dựng tay.
6. Kiểm tra artifact `.94` và evidence; đánh dấu rõ `STALE_FOR_HEAD` hoặc `NOT_RUN` nếu không khớp source hiện tại.
7. Ghi SHA-256, size và mtime trước khi sửa của ba tài liệu Word nguồn; đối chiếu lại SHA-256 ở cuối nhiệm vụ và bắt buộc không đổi.
8. Sau audit, reserve ProductVersion mới theo mục 19.1 trước khi sinh binary/gói ký/lease gắn release. Build tạm trước mốc này chỉ dùng chẩn đoán và không được coi là release evidence.

## 9. Phase 1 — Taxonomy và signed advisory policy

### 9.1. Phân loại nguồn và mức độ

Mỗi finding phải có family/source được xác định trực tiếp, không suy ra bằng phép trừ tổng:

```text
ND30: thể thức văn bản hành chính
HD05: thể thức văn bản Đảng
LOCAL: chính tả/typography deterministic của sản phẩm
LATEX: khuyến nghị thẩm mỹ LaTeX/Typst
```

Không mass-convert mọi `ND30-*` thành Error nếu catalog chưa có căn cứ. Giữ hoặc chuẩn hóa severity của NĐ30/HD05 theo rule catalog có traceability. `LATEX-*` luôn là `Suggestion`, không được nâng thành Error/Warning bằng dữ liệu không tin cậy.

### 9.2. Applicability phải nằm trong gói quy tắc có chữ ký

Không để `_latexScanner.Scan()` chạy vô điều kiện. Tạo schema mới `chuanhoa.local-rule-pack.v2` với DTO typed `AdvisoryProfile`; parser v2 closed-by-default. Nếu cần đọc v1 để nâng cấp cache, v1 được hiểu là không có advisory profile và không được bật `LATEX-*`. Profile v2 có contract:

```text
Code = AcademicTypography
Enabled = true/false
EnabledRuleCodes = tập con của đúng bảy mã LATEX-* canonical
DetectorPolicyVersion = version thuật toán/policy được client hỗ trợ
Thresholds = các giá trị typed, bounded và được parser validate
AutoFixRuleCodes = tập con chỉ gồm LATEX-PAGINATION-KEEP và LATEX-PAGINATION-WIDOW trong release này
```

XML v2 Development mặc định dùng shape canonical sau; không tạo tên field khác ở nhiều lớp:

```xml
<advisoryProfiles>
  <profile code="AcademicTypography" enabled="true" detectorPolicyVersion="1">
    <enabledRules>
      <rule code="LATEX-SEC-STYLE" />
      <rule code="LATEX-SEC-CONTINUITY" />
      <rule code="LATEX-PAGINATION-KEEP" />
      <rule code="LATEX-PAGINATION-WIDOW" />
      <rule code="LATEX-TABLE-BOOKTABS" />
      <rule code="LATEX-CAPTION-POS" />
      <rule code="LATEX-MATH-SYNTAX" />
    </enabledRules>
    <autoFixRules>
      <rule code="LATEX-PAGINATION-KEEP" />
      <rule code="LATEX-PAGINATION-WIDOW" />
    </autoFixRules>
    <thresholds headingConfidenceMinimum="0.90"
                bodyConfidenceMinimum="0.95"
                captionMaxBlankParagraphs="1"
                mathMinimumSignalCount="1" />
  </profile>
</advisoryProfiles>
```

Parser phải giới hạn confidence trong `[0,1]`, `captionMaxBlankParagraphs` trong `[0,2]`, `mathMinimumSignalCount` trong `[1,10]`, reject duplicate/unknown advisory rule code và reject auto-fix code không nằm trong `enabledRules`. `advisoryProfiles` phải là một nested payload có ranh giới parse riêng: lỗi field/rule/threshold bên trong làm vô hiệu toàn bộ advisory profile với diagnostic typed, nhưng không làm hợp thức hóa dữ liệu advisory một phần. Các giá trị Development trên là ngưỡng bảo thủ ban đầu; chỉ được thay sau khi có golden-corpus evidence và không được hạ để làm test dễ PASS.

Yêu cầu hành vi:

1. Khi signed rule pack không có advisory profile hoặc profile bị tắt, scan NĐ30/HD05 không sinh bất kỳ finding `LATEX-*` nào.
2. Development issuer/Admin setting là nơi chủ động phát hành một gói ký hợp lệ bật `AcademicTypography` để kiểm thử cùng nút **Kiểm tra thể thức**. Production mặc định tắt và không được âm thầm opt-in thay người dùng/tenant.
3. DLL chứa thuật toán detector; signed pack quyết định bật/tắt, version policy, ngưỡng và quyền auto-fix. Không cho payload chưa ký mở feature.
4. Sai outer signed envelope, chữ ký, core schema, thời hạn, core legal-rule payload hoặc release compatibility phải từ chối toàn bộ pack, hoặc dùng last-known-good pack còn hạn và hợp lệ; không được “cứu” riêng phần NĐ30 từ outer/core payload malformed. Sau khi outer/core pack đã verify và parse hợp lệ, profile vắng/tắt hoặc nested `advisoryProfiles` lỗi mới được cô lập bằng cách vô hiệu toàn bộ advisory profile; NĐ30/HD05 từ core pack hợp lệ tiếp tục chạy và LATEX không chạy. Diagnostic phải phân biệt `RULE_PACK_REJECTED` với `ADVISORY_PROFILE_DISABLED`.
5. Lease `clientReleaseId`, assembly và feature grant phải khớp release hiện hành; rule pack giữ compatibility contract qua `minimumClientReleaseId`, không bị ép có `version` bằng app. Offline lease tối đa 7 ngày vẫn giữ nguyên.
6. Parser, issuer Development, serializer, signature tests và cache migration phải được cập nhật đồng bộ.
7. Không copy rule/lease cache cá nhân từ máy build vào installer dùng chung.

## 10. Phase 2 — Snapshot Word, fingerprint và COM safety

Thiết kế DTO thuần .NET trước; scanner không được giữ hoặc gọi COM.

### 10.1. Raw paragraph snapshot và derived analysis

Tách rõ:

```text
RawWordSnapshot: dữ liệu capture trực tiếp từ Word, không chứa kết luận role/heading.
DerivedAnalysisContext: LogicalDocumentBlock, role map, heading tree và confidence sinh từ raw snapshot + detector policy.
```

Bổ sung hoặc chuẩn hóa các trường raw ảnh hưởng rule:

- absolute `Start`/`End`, story, section và paragraph index;
- `KeepWithNext` tri-state;
- `WidowControl` tri-state;
- locale-independent built-in Style identity, outline level và tên style chỉ dùng chẩn đoán;
- table index, row, cell và nested state thật nếu paragraph nằm trong bảng;
- protected-span intersection;
- các thuộc tính hình học cần thiết nhưng không buộc layout pagination cho mọi paragraph.

Không dựa vào `Style.NameLocal.StartsWith("Heading")`. Ưu tiên built-in style ID/outline metadata không phụ thuộc Office en-US/vi-VN. Nếu cần đọc `Style`, release RCW trong `finally` và không giữ nó trong cache.

### 10.2. Table snapshot

Mỗi table phải có:

- stable index trong story và absolute range `Start`/`End`;
- section, nesting depth, row/column count;
- merged/split raw facts, header row indexes và các dấu hiệu Word thô dùng để phân loại;
- border snapshot riêng cho top, bottom, left, right, inside-horizontal, inside-vertical và header separator;
- mỗi border có trạng thái `Unknown/None/Present`, line style và line weight;
- thông tin đủ để phân biệt bảng dữ liệu với bảng bố cục quốc hiệu, biểu mẫu pháp lý, bảng phụ lục mẫu, nested table và bảng lưới có ý nghĩa.

Không biến lỗi đọc border thành `false`. Mọi `Word.Borders`, từng `Word.Border`, `Rows`, `Row`, `Cells`, `Cell`, `Range` được lấy riêng phải có ownership và release rõ ràng.

### 10.3. Hình và caption

Snapshot cần mô tả `InlineShape` và Shape nổi đủ dùng để liên kết caption:

- story/section và anchor absolute range;
- loại đối tượng;
- thứ tự trước/sau;
- Word Caption/SEQ field nếu có;
- protected state;
- dữ kiện vị trí/range thô để tầng derived analysis tính association.

`DerivedAnalysisContext` giữ layout/data-table classification, caption association và confidence tương ứng. Không dùng khoảng cách paragraph đơn thuần khi không có object/range thật.

### 10.4. Schema, document fingerprint và scan/cache key

Tăng schema snapshot khi thay contract và đổi hash domain tương ứng. Document fingerprint chỉ hash dữ liệu thể hiện trạng thái Word, tối thiểu:

- raw text và Word-owned format/state;
- `KeepWithNext`;
- `WidowControl`;
- built-in style identity/outline;
- table range, raw merge/nesting/header-row facts;
- đầy đủ border state/style/weight;
- raw object/anchor/field/range facts của hình và caption;
- relevant protected spans.

Không hash role, logical block, heading tree, table classification, caption association/confidence, RulePackId, advisory policy version hoặc lane vào **document fingerprint**. Chúng là dữ liệu dẫn xuất/cấu hình, không phải identity trạng thái Word. Tạo **scan/cache key** riêng gồm tối thiểu:

```text
DocumentFingerprint + document revision + lane + RulePackId + rule-pack version + DetectorPolicyVersion
```

Thêm test chứng minh thay một thuộc tính format làm cache invalid đúng, còn focus/Alt-Tab không làm snapshot tự chạy.

### 10.5. Hiệu năng và hủy

1. Capture theo batch hữu hạn với checkpoint/cancellation giữa các batch.
2. Không gọi pagination/layout API đắt tiền trên toàn bộ tài liệu nếu rule không cần.
3. Không giữ RCW trong static/singleton/context lâu dài.
4. Pure-.NET scan trên immutable DTO có thể chạy nền với cancellation; không truyền RCW vào worker và phải stale-check trước annotate/mutate.
5. Tài liệu lớn phải có đường đi giảm chi phí nhưng không đổi `Unknown` thành “đúng”.
6. Nếu không thu đủ dữ liệu, trả trạng thái typed `NotEvaluated/Incomplete`, thống kê rule/component chưa đánh giá trong popup và diagnostic aggregate không chứa nội dung tài liệu. Advisory có thể bỏ finding cục bộ; rule pháp lý không được hiển thị “0 lỗi/đạt” khi chưa đánh giá đủ.

## 11. Phase 3 — Role propagation, logical block và Heading Tree

### 11.1. Role contract

1. Dùng một vocabulary canonical camelCase từ `DocumentRoleDetector` end-to-end.
2. Không duy trì danh sách role PascalCase riêng trong `HeadingDetector`.
3. `CanonicalRuleScanner` đã có `DetectBlocks(snapshot)` và `MergeRoles(blocks)`; truyền chính các block/role này vào scanner advisory thay vì chạy detector rời hoặc để `Role="Unknown"`.
4. File chứa nhiều văn bản phải có `LogicalDocumentBlock` riêng; mọi state heading/continuity reset tại ranh giới block.
5. Nếu role/block confidence không đủ, không phát sinh finding phụ thuộc role.

### 11.2. Nhận diện heading

Tách rõ:

- thành phần hành chính;
- Điều/Khoản/Điểm pháp lý;
- danh sách/bullet;
- heading học thuật;
- heading đặc biệt không số.

Các đoạn `Điều 1`, `1. Nội dung khoản`, `a) Nội dung điểm`, dòng căn cứ, số/ký hiệu, ngày tháng, nơi nhận, phụ lục mẫu và nội dung trong bảng không được phát sinh `LATEX-SEC-STYLE` hoặc `LATEX-PAGINATION-KEEP` chỉ vì khớp regex.

Heading học thuật phải đạt evidence kết hợp, không chỉ “ngắn hơn 200 ký tự”:

- numbering scheme hợp lệ hoặc built-in outline/style;
- title shape hợp lý;
- bold/style/outline và ngữ cảnh lân cận;
- role không nằm trong exclusion;
- không phải list item, legal structure, citation, table cell hoặc protected range;
- confidence đạt threshold từ signed advisory policy.

`AlphabetPattern` phải được xử lý dứt điểm: hoặc triển khai trong một numbering scheme học thuật có context/confidence và test đầy đủ, hoặc xóa dead code. Không coi `a)`, `b)`, `c)` mặc định là heading.

Roman numeral phải canonical: parse xong phải round-trip về đúng biểu diễn chuẩn; từ chối `IIII`, `IC` và chuỗi chữ cái tình cờ.

### 11.3. Cây phân cấp và continuity

Xây dựng cây theo từng block và numbering scheme:

1. Decimal node giữ toàn bộ path, ví dụ `[1,2,3]`, không chỉ số cuối.
2. Sibling continuity chỉ so trong cùng parent prefix.
3. Phát hiện duplicate sibling, gap, backward numbering, missing parent và nhảy cấp.
4. `1.2` sang `2.1` là chuyển parent hợp lệ, không được báo do state level 2 cũ.
5. Quy tắc first-gap chỉ chạy khi policy của scheme yêu cầu bắt đầu từ 1; không tự áp vào trích đoạn tài liệu.
6. Roman, decimal và unnumbered scheme có state riêng; reset hợp lý khi đổi scheme hoặc block.
7. Không tự đổi số đề mục. Continuity luôn comment-only.
8. Thống kê H1/H2/H3 lấy từ cây đã lọc, không lấy từ regex match thô.

## 12. Phase 4 — Hoàn thiện bảy rule LaTeX/Typst

Mỗi rule phải có applicability, required snapshot fields, severity, message, safe-fix policy, positive/negative/boundary/unknown tests và production caller.

### 12.1. `LATEX-SEC-STYLE`

- Chỉ áp dụng cho heading học thuật có confidence cao trong profile `AcademicTypography`.
- So sánh level với built-in Heading identity/outline locale-independent.
- Không báo cho Điều/Khoản/Điểm, list, component NĐ30/HD05, table, header/footer, caption hoặc protected range.
- Comment-only tuyệt đối trong release này. Không tự gán Heading Style vì Style definition có thể đổi font, cỡ, màu, spacing và làm hỏng format pháp lý.

### 12.2. `LATEX-SEC-CONTINUITY`

- Dùng Heading Tree theo parent/block/scheme.
- Message phải nêu số hiện tại và chuỗi mong đợi cụ thể.
- Không báo continuity giữa hai logical document block.
- Không tự renumber, không xóa comment nếu người dùng chưa sửa thật.

### 12.3. `LATEX-PAGINATION-KEEP`

- Chỉ áp dụng cho heading học thuật confidence cao.
- `KeepWithNext=false` mới báo; `Unknown` không báo.
- Có thể auto-fix thành `true` khi signed allowlist bật, đoạn không phải legal/list/table/protected, và có đoạn nội dung non-empty kế tiếp trong cùng logical block; không tạo chuỗi `KeepWithNext` kéo nhiều đoạn/trang ngoài policy.
- Hậu kiểm chính thuộc tính trên snapshot mới trước khi xóa finding.

### 12.4. `LATEX-PAGINATION-WIDOW`

- Chỉ áp dụng khi `DocumentRoleDetector` không gán component/legal/list/table/caption role, paragraph ở main story, ngoài bảng/vùng bảo vệ, có prose-token ratio và punctuation shape của thân bài đạt threshold ký; không dùng `Text.Length >= 120` làm điều kiện chính. Nếu classifier chưa đạt confidence thì `NotEvaluated`, không báo.
- Loại header/footer/footnote/endnote/caption/table/list/legal basis/component/protected range.
- `WidowControl=false` mới báo; `Unknown` không báo.
- Chỉ auto-fix khi Body confidence cao và signed allowlist bật.

### 12.5. `LATEX-TABLE-BOOKTABS`

- Chỉ áp dụng cho bảng dữ liệu đơn giản có classification confidence cao trong profile AcademicTypography.
- Không áp dụng cho bảng bố cục quốc hiệu/tiêu ngữ, biểu mẫu pháp lý, bảng phụ lục theo mẫu, bảng lồng, merged/split phức tạp, bảng không có header rõ, hoặc bảng mà lưới thể hiện ý nghĩa nghiệp vụ.
- Detector phải đánh giá đủ top, bottom, header separator và vertical border; không kết luận “thiếu ba đường ngang” từ một boolean viền dọc.
- `Unknown` ở bất kỳ border bắt buộc nào làm rule không đủ bằng chứng.
- Comment-only tuyệt đối trong release này; không có nhánh auto-booktabs. Chỉ thu detector/classifier evidence để quyết định một giai đoạn nâng cấp sau.

### 12.6. `LATEX-CAPTION-POS`

- Caption Bảng được khuyến nghị ở trên bảng; caption Hình được khuyến nghị ở dưới hình.
- Nhận diện cả text pattern đã duyệt và Word Caption/SEQ field.
- Liên kết bằng absolute range/object association; hỗ trợ vị trí trên/dưới và khoảng trắng được policy cho phép.
- Không báo nếu không liên kết chắc chắn, object nằm trong vùng bảo vệ hoặc layout pháp lý đặc biệt.
- Comment-only; không tự di chuyển caption.

### 12.7. `LATEX-MATH-SYNTAX`

- Phát hiện raw inline/display LaTeX math có delimiter cân bằng và nội dung có tín hiệu toán rõ.
- Không báo `US$`, `$100`, giá tiền, escaped `\$`, delimiter unmatched, Field, OMath đã tồn tại, hyperlink, code, content control/protected span hoặc text metadata.
- Nested/ambiguous delimiter không được auto-fix; có thể bỏ qua hoặc comment rõ là cần kiểm tra nếu policy ký cho phép.
- Không tự chuyển sang OMath trong đợt này.

## 13. Phase 5 — Annotation và trải nghiệm người dùng

### 13.1. Hợp đồng comment

Comment tiếp tục chỉ có đúng hai dòng:

```text
Hiện tại: <nguồn và trạng thái thực tế>
Yêu cầu đúng: <nguồn và cách sửa cụ thể>
```

Không thêm dòng citation, metadata, confidence hoặc technical detail vào comment. Ownership/citation kỹ thuật lưu ngoài nội dung comment.

Ví dụ advisory:

```text
Hiện tại: [Khuyến nghị LaTeX/Typst] Đề mục “1.2. Phạm vi” chưa bật Keep with next.
Yêu cầu đúng: [Khuyến nghị LaTeX/Typst] Bật Keep with next để giữ đề mục cùng đoạn nội dung ngay sau; đây là khuyến nghị trình bày, không phải lỗi NĐ30.
```

NĐ30/HD05 vẫn phải nêu cách sửa cụ thể, không dùng câu chung chung như “định dạng lại theo quy định”.

### 13.2. Visual policy

1. Finding NĐ30/HD05 đủ bằng chứng tiếp tục comment và tô đỏ đúng range theo chính sách hiện hành.
2. Finding `LATEX-*` dạng Suggestion không được đổi màu chữ sang đỏ và không được trình bày như vi phạm pháp lý.
3. Mặc định advisory dùng comment có nhãn nguồn, không mutation text formatting.
4. Nếu thiết kế một visual khác cho advisory, phải có ownership, lưu trạng thái gốc, restore chính xác, idempotency và test giữ màu/highlight của người dùng; nếu chưa đủ thì không thêm visual.
5. `Unknown` snapshot không sinh annotation.
6. Rerun/clear chỉ xóa annotation do add-in sở hữu và đúng scan/rule; giữ comment và formatting của người dùng.

### 13.3. Popup kết quả

Không dùng `total - latexCount`. Aggregate trực tiếp theo family và severity. Khi advisory profile bật, hiển thị tương đương:

```text
Đã kiểm tra thể thức hoàn toàn tại máy.
- Vi phạm/cảnh báo thể thức NĐ30 hoặc HD05: X.
- Khuyến nghị thẩm mỹ LaTeX/Typst: Y.
- Đề mục học thuật nhận diện: Z (H1: A, H2: B, H3: C).
```

Khi profile không bật, không hiển thị số 0 gây hiểu nhầm đã quét; ghi rõ khuyến nghị LaTeX/Typst chưa được bật trong gói quy tắc hiện tại hoặc ẩn dòng theo UI contract đã test.

## 14. Phase 6 — Sửa lỗi đang chọn, 1-Click và hậu kiểm

### 14.1. Dispatch explicit, không generic fallback

1. Mỗi rule auto-fix phải có handler explicit trong allowlist.
2. Finding `LATEX-*` chưa hỗ trợ không được rơi vào generic paragraph formatter và không được trả `true`.
3. Handler không hỗ trợ trả kết quả typed `NotSupported/AdvisoryOnly`, giữ nguyên comment và Ribbon vẫn hoạt động.
4. Capture document/selection/finding trước khi Modern Comments đổi focus.
5. Không xóa finding khỏi cache như bằng chứng đã sửa.

### 14.2. Hậu kiểm bắt buộc

Sau mutation:

1. Chụp lại snapshot tối thiểu đủ cho rule; rule có context toàn cục như continuity/caption có thể phải chụp lại block liên quan.
2. Chạy lại chính rule với cùng signed policy.
3. Chỉ xóa comment/visual khi cùng rule và lỗi tương ứng thực sự không còn.
4. Nếu anchor dịch chuyển, dùng stable identity/range remap có kiểm chứng; không xóa theo paragraph index cũ.
5. Nếu hậu kiểm lỗi hoặc trả `Unknown`, rollback khi có thể và giữ annotation.
6. Invalidate cache theo fingerprint/revision/lane/policy, không làm mất context document.

### 14.3. Safe-Fix Matrix

| Rule | Mặc định | Điều kiện được auto-fix |
|---|---|---|
| `LATEX-SEC-STYLE` | Comment-only | Không auto-fix trong release này |
| `LATEX-SEC-CONTINUITY` | Comment-only | Không bao giờ tự renumber |
| `LATEX-PAGINATION-KEEP` | Có thể auto-fix | Heading học thuật chắc chắn, giá trị hiện tại false, signed allowlist bật, hậu kiểm PASS |
| `LATEX-PAGINATION-WIDOW` | Có thể auto-fix | Body role chắc chắn, không nằm exclusion, signed allowlist bật, hậu kiểm PASS |
| `LATEX-TABLE-BOOKTABS` | Comment-only | Không auto-fix trong release này; chỉ thu evidence cho một giai đoạn sau |
| `LATEX-CAPTION-POS` | Comment-only | Không tự di chuyển caption |
| `LATEX-MATH-SYNTAX` | Comment-only | Không tự chuyển OMath |

NĐ30/HD05 giữ các auto-fix đã có. Mọi handler được chạm/sửa trong nhiệm vụ này bắt buộc hậu kiểm trước khi xóa comment; lập ledger riêng cho các handler cũ chưa được migrate thay vì refactor mù toàn bộ 77 route trong cùng đợt.

### 14.4. Backup, undo và rollback

1. 1-Click tạo recovery copy trước bulk mutation tại `%TEMP%\ChuanHoa\Backups` theo retention hiện hành 7 ngày/tối đa 20 bản.
2. Hai lệnh đồng nhất cách đặt dấu được coi là bulk mutation và tiếp tục có backup.
3. Sửa lỗi đang chọn, KeepWithNext tại selection, co/giãn chữ và mutation cục bộ không tạo backup toàn file.
4. Tài liệu `.doc`/`.docx` đã lưu giữ nguyên định dạng; tài liệu chưa lưu dùng clone tạm `.docx` mà không ép người dùng Save As.
5. Dùng Word UndoRecord khi version hỗ trợ; luôn restore ScreenUpdating, DisplayAlerts, operation gate và selection state trong `finally`.
6. Chạy 1-Click lần hai không tạo thay đổi thừa.

## 15. Phase 7 — Khóa hồi quy Line Shape, TabStop và format pháp lý

Phần LaTeX không được làm hồi quy các lỗi NĐ30/HD05 đã sửa trước đây.

### 15.1. Line Shape

1. Phân biệt chính xác legacy prefix và current owned prefix; `CHUANHOA2_` không được tự coi là legacy chỉ vì bắt đầu bằng `CHUANHOA`.
2. Ownership không được dựa duy nhất vào paragraph index vì index thay đổi sau chỉnh sửa.
3. Match Shape theo ownership marker, semantic role, anchor story/range, page/cell, line style, weight, chiều dài và tọa độ.
4. Phát hiện line owned bị lệch, quá ngắn/dài, nét đứt hoặc sai weight.
5. Line dưới Tiêu ngữ căn theo text Tiêu ngữ; line dưới tên cơ quan/trích yếu theo đúng component tương ứng.
6. Không chiếm ownership, đổi tên hoặc xóa Shape người dùng khi association chưa chắc chắn.
7. Chuẩn hóa hai lần không nhân đôi, xóa rồi tạo lại hoặc làm line trôi vị trí.

### 15.2. TabStop

1. Không `ClearAll()` custom TabStops vô điều kiện cho mọi paragraph body.
2. Chỉ loại TabStop xung đột khi paragraph được nhận diện chắc chắn thuộc policy cần chuẩn hóa.
3. Giữ tab chủ ý trong bảng, biểu mẫu, list, signature layout, component chuyên biệt và vùng không đủ confidence.
4. Với danh sách NĐ30 thực sự, áp hanging indent/marker policy đã được duyệt và test; không biến khoản/điểm thành heading.

### 15.3. Golden legal behavior

Phải giữ các hợp đồng:

- không đổi `VIỆT NAM` thành `Việt Nam`;
- không làm `NAM` xuống dòng;
- không bôi đậm toàn bộ nội dung Điều 1;
- không tạo line dưới Điều 1;
- không báo line hợp lệ là thiếu;
- căn cứ cuối kết thúc bằng dấu chấm, các căn cứ trước dùng dấu chấm phẩy theo block;
- ngày short hợp lệ trong ngữ cảnh đã chốt không bị format scanner báo sai;
- phụ lục theo Mẫu 2.1/2.2, số phụ lục không bắt buộc La Mã;
- không comment hướng giấy ngang/dọc do người dùng chủ động chọn;
- không phá bảng layout hoặc bảng mẫu pháp lý bằng booktabs.

## 16. Phase 8 — Code quality và build graph

Không thực hiện đại tu không kiểm soát, nhưng khi sửa các class quá lớn phải tách seam có trách nhiệm rõ:

- heading parsing và Heading Tree;
- advisory applicability/policy;
- table border capture/classification;
- figure/caption association;
- annotation presentation policy;
- selected-fix dispatch và post-fix verifier.

Yêu cầu:

1. Core scanner không phụ thuộc Word COM.
2. Không duplicate role detection hoặc rule code string rải rác.
3. Regex được compile/culture policy rõ và có negative corpus.
4. Tránh quét phrase/từ điển theo thuật toán lặp toàn bộ danh sách cho mỗi token nếu profiling cho thấy bottleneck; mọi tối ưu phải giữ kết quả deterministic và có benchmark trước/sau.
5. Nếu baseline API test thực sự fail vì test host không có quyền Windows Event Log, cấu hình logger test phù hợp hoặc làm EventLog non-fatal; không mở rộng scope nếu baseline tái lập vẫn PASS.
6. Sửa VSTO build graph để clean clone restore/build đáng tin cậy. Ưu tiên ProjectReference hoặc dependency target có restore rõ; không dựa vào binary stale hay `project.assets.json` có sẵn trên máy.
7. TreatWarningsAsErrors tiếp tục bật; không suppress warning hàng loạt.

## 17. Ma trận kiểm thử bắt buộc

### 17.1. Unit và core integration

Mỗi rule mới phải có positive, negative, boundary, ambiguity và `Unknown` test.

Heading/continuity:

- nhận diện đúng `1. TỔNG QUAN`, `1.1. Mục tiêu`, `1.1.1. Phạm vi`, Roman canonical và heading không số đã duyệt;
- không nhận nhầm `Điều 1`, khoản `1.`, điểm `a)`, list có dấu chấm phẩy, dòng căn cứ, số/ký hiệu, ngày tháng, tên cơ quan, nơi nhận, phụ lục mẫu;
- `1.1` sang `1.3` cùng parent báo gap;
- `1.2` sang `2.1` không so nhầm state cấp hai;
- duplicate, backward, missing parent, level jump, first-gap policy;
- reset đúng khi một file chứa nhiều logical document block;
- Roman invalid `IIII` và `IC` không được nhận.

Snapshot integration:

- snapshot được tạo từ Word-like adapter/production mapping, không chỉ constructor dựng tay;
- paragraph trong table có `TableIndex` thật;
- style identity hoạt động với tên hiển thị Office en-US và vi-VN;
- border top/bottom/inside-horizontal/inside-vertical/header có style, weight và unknown;
- Style/Border/Shape/Range RCW ownership test hoặc smoke chứng minh không leak;
- đổi Keep/Widow/Style/border làm fingerprint/cache invalid.

Table/caption/math:

- data table booktabs đúng và sai;
- bảng quốc hiệu, layout table, biểu mẫu pháp lý, phụ lục mẫu, nested table, merged phức tạp và semantic grid không bị đề xuất/sửa sai;
- caption Bảng ở trên/dưới và caption Hình ở trên/dưới;
- Word Caption/SEQ, empty paragraph policy, object không liên kết;
- `$E=mc^2$` và display math hợp lệ được phát hiện;
- `US$`, `$100`, escaped dollar, unmatched/nested ambiguous, Field, OMath, hyperlink, code và protected range không bị báo sai.

Annotation/fix:

- legal finding tô đỏ đúng ownership;
- advisory Suggestion không tô đỏ;
- comment đúng hai dòng và có source label;
- rerun/clear/idempotency giữ comment, màu và highlight của người dùng;
- rule không có handler không được generic-format hoặc xóa comment;
- rule có handler chỉ xóa comment sau post-scan xác nhận;
- hậu kiểm `Unknown` giữ comment;
- selected fix vẫn hoạt động sau khi focus ở Modern Comments.

Signed policy:

- profile vắng/tắt không sinh `LATEX-*`;
- signed AcademicTypography profile bật đúng allowlist;
- payload bị sửa, hết hạn, sai release/device/feature bị từ chối;
- advisory fail không làm mất rule NĐ30/HD05 hợp lệ;
- không bypass lease offline 7 ngày.

### 17.2. False-positive corpus bắt buộc

Tạo golden fixtures cho:

- Điều/Khoản/Điểm và danh sách a/b/c;
- văn bản có nhiều loại văn bản trong một file;
- số/ký hiệu và ngày tháng;
- tên cơ quan, quốc hiệu, tiêu ngữ, căn cứ, trích yếu, chữ ký, nơi nhận;
- tiền tệ có `$`;
- bảng header quốc hiệu và bảng mẫu phụ lục;
- header/footer/footnote/endnote;
- table lồng/merged;
- field, content control, hyperlink, Modern Comments và protected range.

### 17.3. Word smoke thực tế

Chạy trên Word 16 x64 hiện có, với cả `.doc`, `.docx` và `Document1` chưa lưu:

1. Add-in load, tab **Chuẩn hóa** xuất hiện khi người dùng tự mở Word.
2. Cả 39 control đúng trạng thái với lease/rule pack hợp lệ; không cần nút Đọc dữ liệu.
3. Kiểm tra thể thức hai lần liên tiếp trước và sau sửa.
4. Kiểm tra chính tả rồi kiểm tra thể thức rồi 1-Click; không mất ActiveDocument hoặc làm Ribbon mờ.
5. Save/Save As và chạy tiếp; tài liệu chưa lưu không bị ép Save As.
6. Alt-Tab/WindowActivate không tự capture/scan và không làm Word Not Responding.
7. Modern Comments + Sửa lỗi đang chọn.
8. File nhiều văn bản, nhiều section, nhiều bảng, Shape, caption, OMath và Field.
9. Tài liệu 10/50/100 trang hoặc corpus lớn tương đương; progress/cancel ở checkpoint an toàn.
10. Chạy 1-Click hai lần; kiểm tra line, style, border, tab, comment và backup idempotent.
11. Chạy tối thiểu 10 vòng sau một vòng warm-up cho scan thể thức và chuỗi command hỗn hợp; ghi median/p95 duration, working set, private bytes, GDI/USER handle và RCW-related symptom. Không có tăng đơn điệu qua các vòng; p95 không hồi quy quá 25% so với baseline cùng máy/corpus trừ khi có phân tích được duyệt. Đóng Word không còn WINWORD treo.
12. Advisory profile tắt và bật bằng hai signed pack riêng; cùng một nút cho kết quả đúng policy.

Word 2010 x86 và các version/bitness khác là compatibility gate trước production. Nếu không có VM/máy phù hợp, ghi `NOT_RUN` hoặc `BLOCKED_ENVIRONMENT`; tuyệt đối không suy ra PASS từ Word 16 x64 hoặc static target Office 14.0.

## 18. Lệnh gate tối thiểu

Dùng local SDK của repo khi tồn tại. Từ `D:\Chuẩn Hóa`, chạy tối thiểu:

```powershell
Set-Location 'D:\Chuẩn Hóa'

& '.\.tools\dotnet\dotnet.exe' restore '.\ChuanHoa.slnx'
& '.\.tools\dotnet\dotnet.exe' build '.\ChuanHoa.slnx' -c Release --no-restore
& '.\.tools\dotnet\dotnet.exe' test '.\ChuanHoa.slnx' -c Release --no-build

python '.\tools\validation\validate_solution_projects.py'
python '.\tools\validation\validate_rule_only_product.py'
python '.\tools\vsto\generate_vsto_ribbon.py'
python '.\tools\vsto\validate_vsto_source.py'

git diff --check
```

Sau khi generator chạy, mọi diff generated phải đúng với contract canonical. Nếu không có thay đổi Ribbon, generator không được tạo drift.

Build VSTO Development bằng MSBuild của Visual Studio Build Tools sau khi solution restore thành công. Ghi lại command, exit code, 0 warning/0 error và hash source/artifact; không dùng VSTO binary cũ.

Không tự push. Được phép tạo local verification commit sau khi mọi diff đã được review và không chứa artifact/secret; nếu policy môi trường không cho phép commit, tạo temp copy từ clean clone rồi áp đúng patch được hash để kiểm thử. Báo rõ phương thức đã dùng. Restore/build/test/validators/VSTO Development build trong môi trường sạch; không trỏ HintPath về workspace gốc và không dùng `bin/obj` copy sẵn.

Mọi test được discover hiện có phải tiếp tục PASS, không test nào bị xóa/skip; 272 là floor cảnh báo của baseline, còn số cuối lấy từ discovery thực tế và dự kiến tăng do có test mới. Không hard-code output giả hoặc sửa expected count xuống thấp hơn.

## 19. Versioning, đóng gói và kiểm chứng cài đặt

### 19.1. Chọn version mới

1. Quét mọi version đã dùng.
2. Chọn version lớn hơn `1.0.0.94` và chưa có artifact/tag/manifest.
3. Cập nhật `Directory.Build.props` làm source-of-truth.
4. Sửa script để `-ApplicationVersion` bị từ chối nếu khác `ProductVersion`, hoặc loại đường gây split-brain.
5. Assert equality giữa ProductVersion, AssemblyVersion, FileVersion, InformationalVersion, ClickOnce application version, VSTO manifest, installer version và lease `clientReleaseId`. Rule pack giữ contract tương thích `minimumClientReleaseId` hiện hành; không ép `rulePack.version == ProductVersion` trừ khi có ADR/schema migration riêng được duyệt.
6. Sau khi reserve/bump version, phát hành lại signed Development pack/lease cho release mới và chạy lại toàn bộ access/binding tests; không dùng artifact đã ký cho `.94`.

### 19.2. Tạo một EXE Development Test

Sau khi tăng `ProductVersion`, đóng hoàn toàn Word và chạy builder không truyền version rời:

```powershell
Set-Location 'D:\Chuẩn Hóa'
& '.\tools\vsto\build_development_test_exe.ps1'
```

Bộ cài dự kiến nằm tại `artifacts\installers\development` với tên lấy từ ProductVersion mới. Một EXE là **một file phân phối**; khi cài, VSTO vẫn phải giải nén DLL, `.vsto`, manifest và đăng ký Office. Không mô tả Development EXE hiện tại là auto-update vì `UpdateEnabled=false` và production updater chưa hoàn thành.

Audit bộ cài bằng script hiện hành, sử dụng path được tính từ `Directory.Build.props`, rồi ghi path, size, SHA-256, Authenticode status và payload entries. Development trusted **public** key là build input được phép nếu được pin bằng SHA-256/thumbprint đã duyệt; private key và cache lease/rule cá nhân vẫn bị cấm. Sửa builder để nhận public key qua input rõ như `-TrustedPublicKeyPath` cùng expected SHA-256 từ build contract, rồi verify trước khi copy; không âm thầm lấy một `trusted-key.xml` tùy ý từ `%LOCALAPPDATA%`. Bộ cài Development phụ thuộc Development provisioning/API hợp lệ và không được mô tả là portable offline installer cho người dùng production. Payload phải dùng allowlist và không có:

```text
VietnameseEngine
LOCAL-TYPO-AI
.onnx
.onnx.data
onnxruntime
model_manifest
training
QRCoder
btnChenQrCode
QrCodeInputDialog
InsertQrCode
private key
PFX
machine-specific cache
document content
```

Ký PE DLL/EXE bên trong trước, sau đó ký VSTO application/deployment manifest, cuối cùng ký outer installer EXE; audit từng lớp. Nếu pipeline Development hiện chưa Authenticode-sign inner DLL hoặc chưa đăng ký Apps & Features/cached uninstall, ghi rõ `DEVELOPMENT_GAP/NOT_PRODUCTION_READY` hoặc triển khai gate tương ứng; không che giấu bằng chữ ký outer EXE.

### 19.3. Install/upgrade/repair/rollback

1. Upgrade từ Development `.94` sang version mới.
2. Repair cùng version.
3. Uninstall/reinstall khi test plan yêu cầu và giữ từ điển cá nhân.
4. Kiểm tra registry `LoadBehavior=3`, manifest trỏ đúng `DevelopmentInstaller\Current`, assembly version đúng và `COMAddIn.Connect=True`.
5. Mở Word theo cách người dùng bình thường, kiểm tra tab trực quan trên tài liệu đã lưu và chưa lưu.
6. Fault-injection ở mọi điểm sau staging/activation: verification, registry, access smoke và switch Current. Chỉ chạy trong registry/profile/installation root cô lập hoặc VM có snapshot phục hồi, không phá bản cài người dùng hiện tại. Nếu lỗi, phải khôi phục `Previous`, registry và trạng thái chạy được; không chỉ trả exit code 10 rồi để installation nửa vời.
7. Bootstrapper Development không được tin certificate chỉ bằng `Subject=CN=...`; phải pin SHA-256 public key/certificate thumbprint từ build contract. Production tuyệt đối không tự thêm arbitrary self-signed certificate vào CurrentUser Root; dùng publisher CA/enterprise-provisioned trust phù hợp.
8. Development certificate tự ký chỉ là bằng chứng Development; production vẫn blocked nếu chưa có publisher certificate tin cậy/timestamp, HTTPS immutable release, KMS/HSM và VM compatibility.

## 20. Cập nhật tài liệu và evidence

Sửa drift theo kết quả thực tế, tối thiểu:

1. `README.md` đang ghi Ribbon 42 control; baseline đúng là 39.
2. Tài liệu Ribbon cũ còn ghi 7 group/3 checkbox; contract/validator hiện là 6 group/0 checkbox/39 control.
3. `Implementation_Status.md` và `Test_Evidence.md` còn lẫn `.91`, `.94` và 259 test. Không thay lịch sử đã xảy ra, nhưng thêm trạng thái hiện hành rõ ràng và không dùng evidence cũ cho source mới.
4. Ghi snapshot schema, advisory profile, family/severity, safe-fix matrix, privacy và compatibility state.
5. Thêm test IDs/evidence cho heading tree, snapshot-border-caption, advisory presentation, post-fix verification, signed profile, cache invalidation, Word heavy/multi-block và installer rollback.
6. Mọi PASS có command, thời điểm, environment, exit code và artifact/hash. Lane chưa chạy phải là `NOT_RUN`, không được ghi COMPLETE.

Tạo/cập nhật evidence máy đọc được tại các path ổn định:

```text
shared/docs/implementation/evidence/nd30_latex_validation.json
shared/docs/implementation/evidence/nd30_latex_word_smoke.json
shared/docs/implementation/evidence/nd30_latex_installer.json
```

JSON phải có schema/version, commit, ProductVersion, environment, command, exit code, discovered/passed/failed/skipped counts, artifact/source hashes, timestamp và trạng thái `PASS/FAIL/NOT_RUN/BLOCKED_ENVIRONMENT`. Không chứa nội dung tài liệu, đường dẫn tài liệu người dùng ngoài tên ba fixture nguồn đã công khai trong repo, token hoặc secret.

## 21. Điều cấm

- Không triển khai nguyên xi hai tài liệu kế hoạch khi code/evidence cho thấy rủi ro.
- Không cho LaTeX/booktabs chạy mặc định ngoài signed advisory profile.
- Không gọi khuyến nghị LaTeX là lỗi NĐ30/HD05.
- Không tô đỏ `Suggestion` LaTeX.
- Không tự renumber heading, tự di chuyển caption hoặc tự chuyển raw math sang OMath.
- Không áp booktabs cho mọi bảng.
- Không gán Heading Style cho Điều/Khoản/Điểm hoặc list item.
- Không dùng độ dài text làm bằng chứng duy nhất cho heading/Body/Widow.
- Không coi dữ liệu COM đọc lỗi là `false` hoặc “đúng”.
- Không xóa comment trước post-scan.
- Không dùng generic fallback để báo auto-fix thành công.
- Không giữ Word RCW trong cache hoặc gọi COM từ background thread.
- Không dùng `FinalReleaseComObject`, `DoEvents`, sleep, retry mù, timeout tăng tùy tiện hoặc test skip để che hang.
- Không tự scan ở lifecycle/getEnabled.
- Không upload nội dung, path hoặc binary tài liệu.
- Không bypass lease/rule signature hoặc dùng development private key trong Release.
- Không thêm AI/ONNX/QR.
- Không ghi đè artifact `.94` hoặc build lại cùng ProductVersion.
- Không dùng `-ApplicationVersion` để tạo manifest khác version DLL.
- Không commit `bin`, `obj`, `artifacts`, secret, private key, PFX, lease/rule cache cá nhân hoặc file tạm chuyển đổi tài liệu.
- Không push remote nếu người dùng không yêu cầu rõ.
- Không xóa test/gate, hạ severity hoặc hard-code test output để đạt PASS.
- Không sửa ba tài liệu Word nguồn.
- Không tuyên bố Word 2010/x86 PASS khi chưa chạy trên môi trường đó.
- Không tuyên bố production-ready bằng certificate Development.

## 22. Definition of Done

Chỉ hoàn thành khi đồng thời đạt:

1. Baseline và diff được audit; không mất thay đổi mới hơn của người dùng.
2. 77 mã hiện hành không hồi quy; bảy mã `LATEX-*` duy nhất, đúng tên và có traceability.
3. Advisory bị tắt mặc định khi signed profile vắng/tắt; profile AcademicTypography ký hợp lệ bật đúng rule allowlist.
4. NĐ30/HD05 giữ thẩm quyền; `Điều 1`, khoản `1.`, điểm `a)`, list, căn cứ, thành phần đầu/cuối và bảng mẫu không bị nhận nhầm là heading/booktabs.
5. Heading Tree xử lý parent, block, scheme, duplicate, gap, backward, missing parent và level jump đúng.
6. Snapshot có locale-independent style identity, tri-state Keep/Widow/border, table range/index và figure/caption association đủ dùng.
7. Document fingerprint bao phủ raw Word state; scan/cache key bao phủ rule pack, policy version và lane; cả hai invalid đúng sau mutation/policy change.
8. Required snapshot state `Unknown` trả `NotEvaluated/Incomplete`, không sinh false finding và không được hiển thị như tài liệu đã đạt.
9. Caption Bảng/Hình và math có positive/negative corpus, không dựa vào regex lân cận đơn giản.
10. Advisory Suggestion không tô đỏ; comment vẫn đúng hai dòng, có nhãn nguồn và cách sửa cụ thể.
11. Selected fix/1-Click chỉ xóa comment sau hậu kiểm; rule không hỗ trợ giữ comment và không báo thành công giả.
12. Chỉ KeepWithNext/Widow đủ confidence mới có thể auto-fix theo signed allowlist; SEC-STYLE/continuity/booktabs/caption/math đều comment-only trong release này.
13. Line Shape owned idempotent, line dashed/lệch được phát hiện, Shape người dùng không bị chiếm/xóa.
14. TabStop chủ ý không bị xóa hàng loạt.
15. Không auto-scan ở startup/open/activate/getEnabled; không Word COM trên worker thread; không còn WINWORD treo trong smoke lặp.
16. `.doc`, `.docx`, `Document1`, file nhiều văn bản và tài liệu lớn hoạt động; không ép Save As.
17. Backup/undo/rollback đúng policy và 1-Click chạy lần hai không tạo thay đổi thừa.
18. Toàn bộ test được discover PASS, không xóa/skip test; số lượng không thấp hơn floor 272; validators PASS; VSTO Development clean build 0 warning/0 error.
19. Word 16 x64 load/smoke có bằng chứng cho source mới. Word 2010/x86 ghi đúng PASS hoặc NOT_RUN/BLOCKED theo môi trường thật.
20. Docs/evidence khớp 39 control, test count mới, source hash và version mới.
21. ProductVersion mới chưa từng dùng; `.94` nguyên vẹn; manifest/DLL/lease/rule/installer không split-brain.
22. Development EXE mới build, ký, audit payload, upgrade/repair/rollback và Word load thành công.
23. Không AI/ONNX/QR/secret/document content trong source runtime hoặc installer.
24. Production blocker còn lại được báo trung thực, không dùng Development evidence thay thế.

## 23. Báo cáo cuối bắt buộc

Khi hoàn tất, trả báo cáo có bằng chứng:

1. Branch, commit đầu/cuối và `git status`.
2. Ma trận audit ban đầu và root cause của từng lỗi đã sửa.
3. File/contract/migration đã thay đổi và lý do.
4. Rule registry trước/sau, signed advisory profile và safe-fix policy.
5. Kết quả false-positive corpus và bảy rule.
6. Snapshot schema/fingerprint/COM ownership thay đổi ra sao.
7. Annotation, popup, selected fix và post-fix verification.
8. Kết quả Line Shape/TabStop/golden NĐ30-HD05 regression.
9. Ma trận unit/integration/validator với command, exit code và số test.
10. Ma trận Word smoke theo file type, Word version và bitness.
11. Version cũ/mới và bằng chứng không ghi đè `.94`.
12. Installer path, size, SHA-256, signature, payload audit, upgrade/repair/rollback và registry/COM load.
13. Danh sách lane `NOT_RUN/BLOCKED` và điều kiện để đóng blocker.
14. Xác nhận không sửa ba tài liệu Word nguồn, không upload tài liệu và không thêm AI/ONNX/QR.

Không kết thúc bằng một kế hoạch mới. Hãy triển khai, kiểm chứng và chỉ tuyên bố hoàn thành khi Definition of Done thực sự đạt; nếu thiếu môi trường production/Word 2010 thì hoàn thành phần có thể kiểm chứng và báo blocker chính xác thay vì giả PASS.
