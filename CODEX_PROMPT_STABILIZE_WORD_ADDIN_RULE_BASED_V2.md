# PROMPT THỰC THI V2: ỔN ĐỊNH CHUẨN HÓA, HOÀN THIỆN RULE-BASED VÀ TỪ ĐIỂN CÁ NHÂN

## 1. Vai trò và yêu cầu kết quả

Bạn là Codex chịu trách nhiệm rà soát, sửa code, kiểm thử và đóng gói dự án Word VSTO Add-in `Chuẩn hóa` tại:

```text
D:\Chuẩn Hóa
```

Remote chính:

```text
https://github.com/newstar94/Chuan-Hoa.git
```

Không chỉ viết kế hoạch. Hãy triển khai đầy đủ các thay đổi trong source, tests, tài liệu và pipeline đóng gói. Chỉ tuyên bố hoàn thành khi có bằng chứng build/test và đạt Definition of Done trong prompt này.

Mục tiêu đợt này:

1. Ổn định Word Add-in trên Word 2010 trở lên, Office x86 và x64.
2. Xóa hoàn toàn AI/ML/ONNX/VietnameseEngine khỏi phiên bản hiện tại.
3. Hoàn thiện kiểm tra chính tả thuần rule-based chạy local.
4. Hoàn thiện kiểm tra và chuẩn hóa thể thức theo Nghị định 30, Hướng dẫn 05 và các quyết định nghiệp vụ đã chốt.
5. Khắc phục Word Not Responding, lỗi ActiveDocument, Ribbon bị mờ/mất và lỗi khi chạy nhiều chức năng liên tiếp.
6. Làm repository build/test được từ clean clone.
7. Hoàn thiện tính năng `Từ điển cá nhân`, tích hợp icon riêng đã cung cấp vào Ribbon.
8. Tạo bộ cài Development Test một file EXE, không chứa engine/model AI.

## 2. Baseline đã kiểm tra trên máy ngày 04/09/2026

Không được giả định baseline GitHub cũ còn đúng. Trước khi sửa, hãy xác minh lại nhưng dùng các phát hiện sau làm điểm xuất phát:

```text
Branch: main
HEAD: 2dcb9f29f9dc920524d09d4c15daf945bc3de42f
origin/main: 2dcb9f29f9dc920524d09d4c15daf945bc3de42f
Worktree: sạch tại thời điểm audit
```

Các commit sau `19c5996` đã bổ sung đáng kể Ribbon, callback, annotation, `WordOneClickRuntime`, typography cleaner và dialog Từ điển cá nhân. Không quay lại code cũ và không ghi đè các sửa lỗi mới hơn.

### 2.1. Lỗi clean-clone đã xác định chính xác

Trên máy có thư mục:

```text
src\ChuanHoa.Application\
```

và source project tồn tại. Tuy nhiên project không được Git theo dõi. Nguyên nhân đã xác minh:

```text
.gitignore:39: *.application
```

Trên Windows, pattern này cũng bỏ qua thư mục `ChuanHoa.Application` do filesystem không phân biệt hoa/thường. `git check-ignore` xác nhận toàn bộ project bị ignore. Vì vậy máy đang làm việc có thể dùng source vật lý và binary cũ, nhưng clean clone vẫn thiếu:

```text
src\ChuanHoa.Application\ChuanHoa.Application.csproj
```

và `dotnet test ChuanHoa.slnx -c Release` thất bại với `MSB3202`.

Đây là P0 và phải sửa đúng nguyên nhân:

1. Sửa `.gitignore` để vẫn ignore deployment manifest `*.application` nhưng không ignore thư mục/source `src/ChuanHoa.Application`.
2. Đưa toàn bộ source cần thiết của `ChuanHoa.Application` vào Git.
3. Không đưa `bin/` và `obj/` vào Git.
4. Tạo một clean clone mới từ commit sau sửa và chứng minh project vẫn hiện diện.
5. Chạy restore/build/test trong clone mới; không dùng source hoặc binary sót lại ngoài Git.

### 2.2. Baseline Ribbon hiện tại

Source validator hiện báo PASS về cấu trúc với:

```text
34 buttons
3 menus
2 dropDowns
3 checkBoxes
42 interactive controls
58 callback methods
Target: net48
Minimum Office version: 14.0
```

Đây chỉ là static source validation, không chứng minh VSTO build/load hoặc Word runtime PASS. Phải chạy lại validator sau mọi thay đổi contract/Ribbon/callback.

### 2.3. AI vẫn còn trong source và installer

Các thành phần sau vẫn được Git theo dõi hoặc còn nằm trong build/runtime:

```text
tools/VietnameseEngine/
training/
src/ChuanHoa.Client.Core/Scanning/VietnameseEngineIpcClient.cs
tests/ChuanHoa.Client.Core.Tests/VietnameseEngineEndToEndIpcTests.cs
NHIỆM VỤ_ Tích hợp bộ sửa lỗi chính tả tiếng Việt AI chạy local-offline vào Add-in hiện có.md
TASK_ Research, Train, Distill, Quantize and Export a Tiny Vietnamese ONNX Context Model.md
VIETNAMESE_SPELL_ARCHITECTURE.md
VIETNAMESE_SPELL_PERFORMANCE.md
VIETNAMESE_SPELL_TEST_REPORT.md
```

`LocalDocumentScanner` vẫn gọi IPC theo từng paragraph. Development installer vẫn build/copy `VietnameseEngine.exe` và `artifacts/models`. Bootstrapper vẫn dừng process theo tên `VietnameseEngine`.

### 2.4. Từ điển cá nhân đã có nhưng chưa hoàn chỉnh

Hiện có:

```text
src/ChuanHoa.AddIn.Vsto/Runtime/CustomDictionaryDialog.cs
src/ChuanHoa.Client.Core/Lexicon/PersonalDictionaryManager.cs
btnTuDienCaNhan
OnTuDienCaNhan
```

Các lỗi đã xác minh:

1. Nút đang dùng icon Office mặc định `imageMso="SpellingOptions"`, chưa dùng icon riêng.
2. Dialog gọi `ClearDocumentIgnores(string.Empty)`, nhưng manager trả về ngay khi `documentId` rỗng; nút báo thành công nhưng thực tế không xóa gì.
3. `CanonicalRuleScanner` gọi `IsKnownOrIgnored(actual)` không truyền document fingerprint/id, nên danh sách bỏ qua theo tài liệu không được production scanner sử dụng.
4. File `user_custom_dictionary.txt` được ghi trực tiếp, không atomic; Word hoặc máy dừng giữa lúc ghi có thể làm hỏng dữ liệu.
5. Load/save đang nuốt mọi exception, khiến UI có thể báo thành công giả.
6. Singleton double-check hiện không dùng `volatile` hoặc `Lazy<T>`.
7. Dialog mở với owner `null`, có nguy cơ nằm sau Word.
8. Nút Ribbon phụ thuộc `GetEnabledHasDocument`, trong khi việc quản lý từ điển người dùng không cần tài liệu đang mở.
9. Chưa có test đầy đủ cho duplicate, Unicode NFC, giới hạn đầu vào, persistence atomic, lỗi I/O, clear current/all document ignores và scanner production sử dụng ignore scope.

## 3. Nguyên tắc thực thi bắt buộc

1. Đọc `AGENTS.md`, `CONTEXT.md`, README, ADR, rule catalog, migration ledger, test evidence và release runbook nếu tồn tại.
2. Kiểm tra `git status --short` trước mỗi nhóm thay đổi; bảo toàn mọi thay đổi của người dùng.
3. Dùng `rg` lập dependency map trước khi xóa AI hoặc sửa shared code.
4. Không sửa hàng loạt trước khi xác định nguyên nhân.
5. Không dùng sleep, retry mù, tăng timeout hoặc skip test để che deadlock/hang.
6. Không gọi Word COM từ `Task.Run`, thread pool hoặc background thread.
7. Không thực hiện network I/O trong Ribbon callback, scan hoặc mutation Word.
8. Không tự động đọc/quét tài liệu khi Word khởi động, mở tài liệu, WindowActivate, chuyển focus hoặc `getEnabled`.
9. Chỉ dùng `apply_patch` cho chỉnh sửa source; không ghi đè file bằng script ad-hoc.
10. Không sửa các tài liệu Word nguồn. Bản PDF/render dùng để nghiên cứu phải là bản tạm chỉ đọc.
11. Mọi báo cáo PASS phải kèm lệnh, exit code và output thực tế.

## 4. Kiến trúc phiên bản hiện tại

### 4.1. Phạm vi được giữ

- VSTO Add-in dành cho Word 2010+.
- .NET Framework 4.8, Office 14.0+.
- Hỗ trợ trực tiếp `.doc` và `.docx`.
- Một tab Ribbon duy nhất tên `Chuẩn hóa`.
- Kiểm tra thể thức và chính tả chạy local.
- Rule pack/lease có thể tải và cache trước, nhưng command Word không gọi mạng đồng bộ.
- Rule-based là nguồn thực thi duy nhất của spelling checker.

### 4.2. Xóa hoàn toàn AI ở giai đoạn hiện tại

Không được giữ trong runtime, build graph, tests hoặc installer hiện tại:

- `VietnameseEngine.exe`;
- Named Pipe dành cho engine;
- `VietnameseEngineIpcClient`;
- `CheckWithAiEngine`;
- `LOCAL-TYPO-AI`;
- ONNX Runtime hoặc `InferenceSession`;
- `.onnx`, `.onnx.data`, model checkpoint, AI tokenizer hoặc model manifest;
- training, distillation, quantization và export code;
- AI benchmark/report không tái lập;
- AI feature flag, AI stub, AI mock hoặc code dự phòng không có caller production.

Không đổi tên AI thành `Advanced`, `Smart` hoặc `Contextual` để giữ nguyên code. AI chỉ xuất hiện trong một tài liệu roadmap tương lai ngắn gọn, không có binary, dependency hoặc code thực thi.

## 5. Workstream A — Làm repository tái lập

1. Sửa xung đột `.gitignore` với `ChuanHoa.Application`.
2. Track các source sau nếu đúng dependency hiện tại:
   - `ChuanHoa.Application.csproj`;
   - `Clock.cs`;
   - `ServiceCollectionExtensions.cs`;
   - `Persistence/*`;
   - `Scanning/*`.
3. Kiểm tra không có secret hoặc machine-specific path trước khi add.
4. Bổ sung test/CI thực hiện clean checkout và kiểm tra project paths trong solution đều tồn tại.
5. Không xóa reference đến Application khỏi solution/API/Infrastructure để né lỗi.
6. Nếu NuGet bị lỗi SSL cục bộ, phân biệt rõ lỗi môi trường với lỗi source; CI vẫn phải có restore chính thức.

## 6. Workstream B — Xóa AI không làm mất rule-based

1. Xóa `tools/VietnameseEngine` khỏi tree.
2. Xóa `training` và các tài liệu nhiệm vụ/báo cáo AI đã lỗi thời khỏi nhánh sản phẩm hiện tại; có thể thay bằng một file `docs/roadmap/AI_CONTEXTUAL_SPELLING_FUTURE.md` chỉ mô tả phạm vi tương lai.
3. Xóa `VietnameseEngineIpcClient` và IPC E2E test.
4. Sửa `LocalDocumentScanner`:
   - bỏ engine client field/constructor;
   - bỏ `CheckWithAiEngine`;
   - `ScanSpelling` chỉ gọi các module deterministic;
   - không còn label/source/rule code AI.
5. Sửa `WordDocumentReadRuntime` để tạo scanner rule-only.
6. Xóa start/stop VietnameseEngine khỏi bootstrapper.
7. Xóa build/copy Engine và `artifacts/models` khỏi installer script.
8. Chuyển installer sang payload allowlist; cấm wildcard copy toàn output.
9. Thêm gate fail nếu payload có `VietnameseEngine`, `.onnx`, `.onnx.data`, `onnxruntime`, `model_manifest` hoặc thư mục `training`.
10. Chỉ giữ tokenizer/candidate generator/confusion set nếu có caller rule-based thật; đổi mô tả về đúng bản chất và bổ sung test caller.

## 7. Workstream C — Ổn định Word lifecycle và hiệu năng

### 7.1. Ngăn Word Not Responding

1. Rà soát `RibbonRuntime`, `WordDocumentReadRuntime`, `WordDocumentSnapshotBuilder`, `WordLocalScanRuntime`, `WordOneClickRuntime` và mọi COM adapter.
2. Bỏ IPC sẽ loại một nguyên nhân treo, nhưng vẫn phải xử lý snapshot tài liệu lớn.
3. Word COM chỉ được truy cập trên STA, theo vòng lặp có giới hạn.
4. Tách dữ liệu thành DTO thuần .NET càng sớm càng tốt; rule scan chỉ nhận DTO/text.
5. Thay các `Application.DoEvents()` trong scan/snapshot bằng cơ chế tiến độ không cho phép reentrancy vào mutation hoặc callback khác.
6. Có operation state rõ ràng: Idle, Capturing, Scanning, Annotating, Mutating, Cancelling, FailedRecoverable.
7. `finally` luôn giải phóng operation gate, ScreenUpdating, DisplayAlerts, undo record và COM scope thuộc quyền sở hữu.
8. Tài liệu lớn xử lý theo batch; có progress và cancellation; không dùng một thao tác COM không giới hạn.
9. Trước annotate/mutate, xác nhận document identity, revision/fingerprint và range còn hợp lệ.
10. Không giữ `Word.Document`, `Range`, `Table`, `Shape` hoặc RCW trong static/singleton cache.

### 7.2. ActiveDocument và Ribbon

Khắc phục và test:

- đang mở tài liệu nhưng báo `Hãy mở một tài liệu Word`;
- dùng chức năng thứ hai sau chức năng thứ nhất thì mất document;
- sau exception toàn bộ Ribbon bị mờ;
- bấm kiểm tra lần hai bị treo;
- chuyển sang ứng dụng khác rồi quay lại Word thì add-in tự quét;
- Save/Save As làm context bị invalid;
- `Sửa lỗi đang chọn` làm Modern Comments lấy focus và mất document;
- Codex mở Word thì có tab nhưng người dùng tự mở Word lại mất tab.

Yêu cầu thiết kế:

1. Capture document đúng một lần ở đầu callback.
2. Mỗi document có context riêng.
3. Không gọi lại `_application.ActiveDocument` giữa command nếu đã có document ổn định.
4. `getEnabled` chỉ dùng metadata/cache nhẹ; không đọc content, Selection, table, shape hoặc server.
5. Exception không được biến trạng thái transient thành `ACTIVE_DOCUMENT_REQUIRED` vĩnh viễn.
6. Sau lỗi phải invalidate Ribbon an toàn và trở lại trạng thái dùng được.
7. Không yêu cầu nút `Đọc dữ liệu` như prerequisite.
8. Mỗi command tự chuẩn bị scope dữ liệu tối thiểu và chỉ tái sử dụng cache khi chắc chắn hợp lệ.

## 8. Workstream D — Hoàn thiện spelling rule-based

Pipeline bắt buộc:

1. Unicode NFC normalization có mapping offset về text gốc.
2. Xác định story, paragraph, table cell và protected spans.
3. Tokenize tiếng Việt và punctuation.
4. Áp dụng từ điển, personal dictionary, acronym/proper-name protection.
5. Kiểm tra spacing, punctuation và casing theo role/ngữ cảnh xác định.
6. Kiểm tra cấu trúc âm tiết và lexicon.
7. Sinh candidate từ Telex, lỗi bàn phím, dấu thanh, phụ âm đầu/cuối và edit distance giới hạn.
8. Confusion pair/phrase rule chỉ chạy khi điều kiện deterministic khớp.
9. Xếp hạng candidate bằng điểm giải thích được; không gọi đây là semantic AI.
10. Chỉ auto-fix khi có một candidate rõ ràng, rule cho phép và hậu điều kiện kiểm chứng được.
11. Trường hợp mơ hồ chỉ comment hoặc không báo, ưu tiên tránh false positive.

Các test bắt buộc:

- `Quyết định xố` → `Quyết định số`;
- `ự án` → `dự án`;
- `bàn dao hồ sơ` → `bàn giao hồ sơ` chỉ khi phrase rule rõ;
- giữ `CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM` viết hoa;
- không báo chỉ mục `a)`, `b)`, `1.`, `1.1`, số La Mã;
- không báo tên riêng, tên cơ quan, địa danh, viết tắt, số hiệu, mã hồ sơ, URL, email;
- không coi `HOÀ/HÒA` là lỗi cần comment;
- giữ lệnh chủ động đồng nhất `oà/òa`, `uý/úy`;
- ngày dạng short trong căn cứ được phép về thể thức; rule thiếu từ `ngày` chỉ chạy đúng ngữ cảnh đã chốt;
- không xử lý field code, hyperlink metadata, content-control metadata hoặc marker bảng như văn bản thường.

## 9. Workstream E — Từ điển cá nhân

### 9.1. Hợp đồng chức năng

Từ điển cá nhân là rule-based, chạy local và không yêu cầu AI/server. Dữ liệu mặc định tại:

```text
%LocalAppData%\ChuanHoa\Dictionaries\user_custom_dictionary.txt
```

Người dùng phải có thể:

1. Mở dialog ngay cả khi chưa mở document.
2. Xem danh sách đã sắp xếp.
3. Tìm kiếm không phân biệt hoa/thường và tương thích Unicode tiếng Việt.
4. Thêm từ hoặc cụm được policy cho phép.
5. Xóa từ với xác nhận phù hợp.
6. Thêm từ đang được comment/chọn vào từ điển từ workflow kiểm tra chính tả.
7. Bỏ qua một từ trong document hiện tại theo fingerprint ổn định.
8. Xóa danh sách bỏ qua của document hiện tại.
9. Xóa tất cả danh sách bỏ qua trong session bằng API riêng, không truyền chuỗi rỗng để giả lập.
10. Thấy lỗi lưu/đọc bằng thông báo rõ ràng; không báo thành công giả.

### 9.2. Sửa manager

1. Thay singleton double-check bằng `Lazy<PersonalDictionaryManager>` hoặc cơ chế thread-safe đơn giản.
2. Không dùng chung static lock cho mọi instance test; tách instance lock khỏi singleton initialization.
3. Chuẩn hóa NFC và Trim trước khi lưu/so sánh.
4. Đặt giới hạn chiều dài, số dòng, ký tự điều khiển và input rỗng.
5. Định nghĩa rõ từ đơn, cụm từ và cách scanner áp dụng từng loại.
6. Ghi file atomic: ghi temp cùng volume, flush, replace/move; giữ last-known-good hoặc backup nhỏ.
7. Không nuốt lỗi I/O. Trả result/error typed cho UI.
8. `GetUserWords` trả snapshot sorted, không lộ collection mutable.
9. Cung cấp `ClearDocumentIgnores(documentId)` và `ClearAllDocumentIgnores()` riêng biệt.
10. Truyền document fingerprint/id vào scanner để `IgnoreWordForDocument` thực sự có hiệu lực.
11. Dọn document ignore cache khi document đóng; không để tăng bộ nhớ vô hạn.
12. Không ghi nội dung tài liệu hoặc personal words vào log thông thường.

### 9.3. Sửa dialog

1. Dialog phải có owner là cửa sổ Word hiện tại và không nằm sau Word.
2. Không giữ COM object trong dialog.
3. Hỗ trợ DPI 100%, 125%, 150%, 200%; không cắt chữ tiếng Việt.
4. Có trạng thái rỗng, lỗi đọc, lỗi lưu và duplicate rõ ràng.
5. Disable `Xóa từ` khi chưa chọn.
6. Enter thêm từ, Escape đóng; focus và tab order đúng.
7. Không hiện message success thừa sau thao tác thông thường.
8. `Xóa bỏ qua` phải nói rõ xóa của document hiện tại hay toàn session và thực hiện đúng lựa chọn.
9. Sau khi thay đổi từ điển, invalidate spelling cache của các document liên quan để lần kiểm tra tiếp theo dùng dữ liệu mới.

## 10. Workstream F — Icon riêng cho Từ điển cá nhân

Các asset đã được thiết kế và cung cấp tại workspace:

```text
D:\chuan-hoa-the-thuc-workspace\assets\icons\personal-dictionary.svg
D:\chuan-hoa-the-thuc-workspace\assets\icons\personal-dictionary-16.png
D:\chuan-hoa-the-thuc-workspace\assets\icons\personal-dictionary-32.png
D:\chuan-hoa-the-thuc-workspace\assets\icons\personal-dictionary-48.png
D:\chuan-hoa-the-thuc-workspace\assets\icons\personal-dictionary-256.png
```

Ý nghĩa hình ảnh: quyển từ điển mở màu xanh kết hợp huy hiệu người dùng màu xanh lục; nền trong suốt. Không tự tạo một icon khác nếu các asset này render đúng.

### 10.1. Tích hợp bắt buộc

1. Copy asset cần thiết vào thư mục source có tên rõ ràng, ví dụ:

```text
src\ChuanHoa.AddIn.Vsto\Resources\Icons\personal-dictionary-16.png
src\ChuanHoa.AddIn.Vsto\Resources\Icons\personal-dictionary-32.png
```

2. Nhúng PNG vào VSTO assembly bằng `EmbeddedResource`; không dùng absolute path runtime.
3. Với `btnTuDienCaNhan`, bỏ `imageMso="SpellingOptions"` và dùng callback `getImage="GetImageTuDienCaNhan"` trong ribbon contract canonical.
4. Không sửa trực tiếp generated XML/callback mà bỏ qua contract/generator. Cập nhật theo đúng source-of-truth:
   - `shared/contracts/ribbon/ribbon-contract.v1.json`;
   - `tools/vsto/generate_vsto_ribbon.py` nếu cần;
   - generated Ribbon XML/callback;
   - source validator và evidence.
5. Generator hiện đã biết sinh `getImage` bằng `Runtime.GetImage(controlId)`, nhưng `IChuanHoaRibbonRuntime` và `UnavailableRibbonRuntime` chưa có `GetImage`. Bổ sung contract runtime đầy đủ.
6. Loader ảnh phải:
   - dùng manifest resource name ổn định;
   - cache ảnh/COM picture theo lifecycle phù hợp;
   - không dispose stream trước khi bitmap đã được clone hoàn chỉnh;
   - trả loại tương thích Office 2010 Ribbon XML (`IPictureDisp`/object được Office chấp nhận);
   - không làm rò GDI handle;
   - không đọc file từ disk mỗi lần Office gọi `getImage`.
7. Nếu resource không tồn tại hoặc load thất bại, callback không được ném exception ra Office. Dùng fallback an toàn để tab Ribbon vẫn hiển thị và ghi diagnostic không chứa dữ liệu tài liệu.
8. Test icon ở 16 px và 32 px, Office light/dark ribbon nếu phiên bản hỗ trợ, DPI 100–200%.
9. Kiểm tra PNG có alpha transparency, đúng kích thước và hash ổn định.
10. Không nhúng SVG trực tiếp vào Office 2010 Ribbon; SVG chỉ là master thiết kế.

## 11. Workstream G — Thể thức NĐ30/HD05

Đối chiếu đầy đủ ba tài liệu nguồn và rule catalog. Cập nhật traceability cho mỗi rule: nguồn, loại văn bản, điều kiện, finding, cách sửa, auto-fix, positive test, negative test và caller production.

Yêu cầu trọng yếu:

1. Tự nhận diện loại văn bản, không bắt người dùng chọn.
2. Không đổi `VIỆT NAM` thành `Việt Nam`.
3. Không để `NAM` xuống dòng sau chuẩn hóa.
4. Header dùng một tier cỡ chữ nhất quán:
   - Quốc hiệu 13, Tiêu ngữ 14, địa danh/ngày 14; hoặc
   - Quốc hiệu 12, Tiêu ngữ 13, địa danh/ngày 13.
5. Nhận diện đúng trích yếu và trạng thái bold thực tế.
6. Không bôi đậm toàn bộ nội dung Điều 1 và không tạo line dưới Điều 1.
7. Line dưới Tiêu ngữ, tên cơ quan và Trích yếu là Line Shape, không phải underline.
8. Line matching dùng anchor, tọa độ trang/cell, chiều dài, kiểu line và semantic role.
9. Không báo thiếu line hợp lệ; nếu line sai thì auto-fix được.
10. Auto-fix line idempotent: lần hai không tạo thêm hoặc làm lệch.
11. Line Tiêu ngữ căn theo text Tiêu ngữ, không theo page/cell sai.
12. Phụ lục theo Mẫu 2.1/2.2; số phụ lục không bắt buộc La Mã.
13. Không comment orientation ngang/dọc do người dùng chủ động chọn.
14. Kéo bảng lên cùng tiêu đề phụ lục khi khoảng trống là do paragraph/section break dư nhưng không phá orientation.
15. Lặp tiêu đề bảng cho toàn bộ bảng, xử lý merged rows an toàn.
16. Xóa trang thừa có vòng lặp hữu hạn, giữ end-of-cell marker và section break cần thiết.
17. Khi 1-Click chuẩn hóa đoạn nội dung, phải xóa custom TabStops và đặt lại Left/Right/FirstLineIndent. Danh sách bắt đầu bằng dấu gạch/bullet dùng hanging indent ổn định: marker 10 mm, nội dung và dòng xuống hàng 15 mm; không áp quy tắc này cho bảng hoặc component có role chuyên biệt.

## 12. Workstream H — Comment, sửa lỗi và 1-Click

Comment chỉ có:

```text
Hiện tại: <lỗi thực tế>
Yêu cầu đúng: <cách sửa cụ thể>
```

Yêu cầu:

1. Không tạo `Missing content`.
2. Neo comment đúng range/story và tô đỏ đúng vùng.
3. `Sửa lỗi đang chọn` capture selection/range trước khi đổi focus khỏi Modern Comments.
4. Sửa xong phải verify cục bộ rồi mới xóa comment.
5. Không xóa comment nếu correction thất bại hoặc lỗi vẫn còn.
6. Thành công không hiện modal gây phiền.
7. 1-Click sửa cả format và spelling rule-based đủ điều kiện.
8. Sau mutation chỉ quét lại scope cần thiết nếu có thể, không mặc định đọc lại toàn tài liệu.
9. Co/giãn chữ và các lệnh cục bộ không tạo backup toàn file.
10. Chỉ lệnh chỉnh sửa hàng loạt như Chuẩn hóa toàn bộ/Thay dấu tạo recovery copy trong Windows Temp và có cleanup an toàn.

## 13. Workstream I — Ribbon contract regression audit

Quyết định sản phẩm mới nhất loại bỏ hoàn toàn QR, nhưng vẫn giữ ba checkbox nhóm Hiển thị. Quyết định này ghi đè yêu cầu khôi phục QR trong bản prompt cũ.

1. Đối chiếu yêu cầu đã chốt, ribbon contract, migration ledger và VBA gốc.
2. Xóa `btnChenQrCode`, callback, handler, dialog, renderer, dependency QRCoder, test thao tác QR và file QR trong payload installer. Không xóa VBA extracted/evidence baseline; đánh dấu chúng là provenance của tính năng đã retired.
3. Khôi phục/giữ đầy đủ ba checkbox Hiển thị theo contract; không xóa tính năng ngoài phạm vi.
4. Mọi control trong contract phải có handler production; unregistered command fail closed.
5. Không để generator làm rơi control chỉ vì source DOTM và contract không đồng bộ.
6. Sau reconcile, target phải có 34 button, 3 menu, 2 dropdown, 3 checkbox và 42 interactive control; generator và validator phải fail nếu QR bị đưa trở lại.

## 14. Workstream J — License và security

1. Không bỏ/bypass license để làm Ribbon sáng.
2. Development và Release dùng key/config riêng.
3. Release phải fail closed nếu chưa có endpoint/public key production.
4. Lease/rule verification local không đọc nội dung tài liệu.
5. Không gọi server cho từng paragraph hoặc mỗi mutation.
6. Không hard-code private key/token production.
7. Không tuyên bố chống crack tuyệt đối.
8. Không để certificate Development được coi là production evidence.

## 15. Workstream K — Một EXE Development Test rule-only

Sửa pipeline để:

1. Build lại source mỗi lần; không dùng binary cũ vì file đã tồn tại.
2. Payload dùng allowlist.
3. Không chứa VietnameseEngine, Engine directory, ONNX, model hoặc training artifact.
4. Không chép lease/rule cache cá nhân của máy build vào installer dùng chung.
5. Version có một source-of-truth cho ApplicationVersion, AssemblyVersion, FileVersion và InformationalVersion.
6. Installer, VSTO manifests, DLL và executable nội bộ được ký đúng policy môi trường.
7. Cài qua staging, verify rồi atomic switch `Current`; giữ rollback.
8. Có repair/uninstall/upgrade test.
9. Người dùng vẫn chỉ cần chạy một EXE; sau cài VSTO được phép bung DLL/manifest đúng chuẩn.

## 16. Kiểm thử bắt buộc

### 16.1. Core/static tests

- solution project-existence test từ clean clone;
- AI-removal dependency/payload test;
- spelling positive/negative/protected-span tests;
- Unicode offset tests;
- document classifier tests;
- Line Shape geometry/idempotency tests;
- comment schema và selected-fix tests;
- appendix/table/section tests;
- personal dictionary manager/dialog integration tests;
- icon resource/callback/fallback tests;
- recovery copy retention/safe cleanup tests.

### 16.2. Word smoke tests

1. `.doc` và `.docx`.
2. Chạy hai chức năng liên tiếp.
3. Kiểm tra chính tả hai lần sau sửa.
4. Save/Save As rồi chạy tiếp.
5. Chuyển ứng dụng rồi quay lại Word, xác nhận không auto-scan.
6. Modern Comments + Sửa lỗi đang chọn.
7. Tài liệu nhiều bảng/shape/section.
8. Tài liệu 10/50/100 trang với progress/cancel.
9. Xóa trang thừa trên tài liệu kết thúc bằng bảng.
10. Chuẩn hóa hai lần, line không nhân đôi/lệch.
11. Từ điển cá nhân mở không cần document, nằm trước Word, thêm/xóa/lưu/ignore đúng.
12. Đóng Word rồi người dùng tự mở lại, tab vẫn xuất hiện.
13. Word 2010 x86 và Word 16 x64; môi trường chưa chạy phải ghi NOT_RUN, không PASS.

### 16.3. Lệnh gate tối thiểu

```powershell
dotnet restore .\ChuanHoa.slnx
dotnet build .\ChuanHoa.slnx -c Release --no-restore
dotnet test .\ChuanHoa.slnx -c Release --no-build
python .\tools\vsto\validate_vsto_source.py
git diff --check
```

Tạo một clone mới sau sửa và chạy lại ba lệnh dotnet trong clone đó.

Sau khi build installer, kiểm tra payload và fail nếu có:

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
```

## 17. Điều cấm

- Không giữ AI code dưới tên khác.
- Không tăng timeout hoặc thêm sleep để che hang.
- Không gọi Word COM từ background thread.
- Không dùng `Application.DoEvents` trong vùng có thể re-enter command/mutation.
- Không nuốt exception ở persistence/command chính.
- Không báo xóa ignore thành công khi API không làm gì.
- Không sửa generated Ribbon mà không sửa contract canonical.
- Không để lỗi icon làm mất toàn bộ Ribbon tab.
- Không dùng `imageMso` đồng thời với custom `getImage` cho cùng control.
- Không hard-code test output hoặc counts giả.
- Không xóa chức năng ngoài phạm vi để giảm lỗi.
- Không xóa comment trước verification.
- Không thay orientation của người dùng.
- Không đưa `bin`, `obj`, secret, PFX/private key hoặc cache máy build vào Git/installer.
- Không tuyên bố production-ready bằng certificate Development.

## 18. Definition of Done

Chỉ hoàn thành khi đồng thời đạt:

1. `ChuanHoa.Application` được Git theo dõi và tồn tại trong clean clone.
2. Core solution restore/build/test thành công trong clean clone có dependency hợp lệ.
3. Không còn AI engine/IPC/ONNX/model/training trong runtime/build/installer.
4. Spelling checker chỉ dùng rule-based và đạt bộ test lỗi mẫu/false positive.
5. Không auto-scan ở startup/open/activate/getEnabled.
6. Chạy nhiều lệnh liên tiếp không mất document hoặc làm Ribbon mờ.
7. Kiểm tra lần hai và tài liệu lớn không treo; cancellation hoạt động.
8. Line Shape đúng geometry, sửa được và idempotent.
9. 1-Click không xóa comment của lỗi chưa sửa.
10. Comment đúng hai phần và chỉ dẫn sửa cụ thể.
11. Từ điển cá nhân lưu atomic, lỗi được báo thật, ignore theo document hoạt động và clear không còn no-op.
12. Nút Từ điển cá nhân dùng icon custom đã cung cấp, hiển thị rõ ở 16/32 px và lỗi resource không làm mất Ribbon.
13. Ribbon contract có 42 control; QR bị loại bỏ hoàn toàn khỏi source/build/test/installer, ba checkbox Hiển thị vẫn được giữ.
14. Development installer là một EXE rule-only, payload không chứa AI.
15. Word smoke có bằng chứng thật; lane chưa chạy được báo NOT_RUN/BLOCKED.
16. Tài liệu kiến trúc, status, evidence và runbook khớp source hiện tại.
17. Đoạn danh sách có TabStops cũ được 1-Click đưa về hanging indent đồng nhất; có unit test, source gate và Word DOC/DOCX smoke.

## 19. Báo cáo cuối bắt buộc

Trả về:

1. Branch/commit đầu và cuối, `git status`.
2. Root cause `.gitignore` và bằng chứng clean clone đã sửa.
3. Danh sách AI source/project/test/doc/packaging hook đã xóa.
4. Root cause Word hang/ActiveDocument/Ribbon và file đã sửa.
5. Rule-based features/tests đã hoàn thiện.
6. Từ điển cá nhân: persistence, ignore scope, dialog, icon và tests.
7. Ribbon controls trước/sau; xác nhận QR đã retired và ba checkbox Hiển thị vẫn còn.
8. Ma trận lệnh test với exit code.
9. Word smoke theo version/bitness.
10. Installer path, version, size, SHA-256, signature status và payload audit.
11. Blocker production còn lại.
12. Xác nhận không sửa tài liệu Word nguồn của người dùng.

Không kết thúc bằng kế hoạch hoặc mô tả. Hãy triển khai, kiểm chứng và chỉ tuyên bố hoàn thành khi Definition of Done thực sự đạt.
