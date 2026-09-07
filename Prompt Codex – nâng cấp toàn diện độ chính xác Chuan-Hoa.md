Bạn đang làm việc trực tiếp trên repository:

`newstar94/Chuan-Hoa`

## 1. Vai trò và mục tiêu

Hãy đóng vai **Principal Software Engineer / Senior .NET Engineer chuyên xử lý NLP tiếng Việt, Microsoft Word VSTO và hệ thống rule-based document analysis**.

Nhiệm vụ của bạn là:

**Nghiên cứu code mới nhất trên branch `main`, sau đó trực tiếp triển khai một đợt nâng cấp toàn diện nhằm tăng độ chính xác nhận diện, kiểm tra chính tả và chuẩn hóa văn bản hành chính của dự án Chuan-Hoa.**

Không chỉ viết kế hoạch hoặc đề xuất. Hãy:

1. Đọc code thực tế.
2. Xác định trạng thái hiện tại.
3. Phân biệt lỗi nào đã được sửa, sửa một phần và chưa sửa.
4. Trực tiếp sửa code/data/test/CI.
5. Chạy toàn bộ quality gates.
6. Chỉ kết luận hoàn thành khi có bằng chứng test/build/CI tương ứng.

Không được bê nguyên checklist cũ rồi triển khai lại những thay đổi đã có trên `main`.

---

# 2. Các nguyên tắc kiến trúc bắt buộc

Dự án là add-in VSTO/C# cho Microsoft Word.

Phải bảo toàn các nguyên tắc sau:

- Nội dung tài liệu được xử lý **local**.
- Không gửi nội dung văn bản của người dùng ra server.
- Không thêm telemetry chứa nội dung tài liệu.
- Không thêm cloud NLP/API/LLM vào luồng kiểm tra văn bản.
- Kết quả phải deterministic ở cùng input + cùng rule pack.
- Không phá Word VSTO runtime hiện có.
- Không làm thay đổi anchor/offset khiến annotation trỏ sai vị trí.
- Không làm suy yếu security/privacy/supply-chain validation hiện có.
- Không xóa test để làm CI xanh.
- Không giảm threshold quality gate để che lỗi.
- Không vô hiệu hóa rule chỉ để test pass.
- Không sửa unrelated code nếu không cần thiết.

`LocalRulePack` hiện có schema versioning và parsing chặt chẽ. Không tùy tiện thay schema.

Nếu có thể triển khai confidence/bigram/dictionary metadata bên ngoài schema signed rule pack thì ưu tiên cách đó.

Nếu thực sự phải thay schema:
- thiết kế backward compatibility;
- bump schema rõ ràng;
- giữ parser cho version cũ;
- bổ sung migration/backward compatibility tests;
- cập nhật signing/validation tương ứng.

---

# 3. P0 — SỬA CI/TRẠNG THÁI REPOSITORY TRƯỚC

Hiện Source Quality trên `main` đang lỗi ngay tại checkout liên quan tới:

`tmp/deps/hunspell-vi`

Git báo rằng path này là submodule/gitlink nhưng không tìm thấy URL tương ứng trong `.gitmodules`.

### Việc phải làm

Trước tiên chạy và kiểm tra:

```powershell
git status --short
git ls-files -s tmp/deps/hunspell-vi
git submodule status
```

Kiểm tra toàn repo để xác định:

- Đây là submodule thật được thiết kế chủ ý?
- Hay chỉ là gitlink vô tình được commit từ thư mục dependency tạm?
- Trong khi repo đã có dữ liệu Hunspell cần thiết dưới `shared/dictionaries/hunspell-vi`?

### Yêu cầu

Ưu tiên sửa nguyên nhân gốc.

Nếu `tmp/deps/hunspell-vi` chỉ là dependency/cache tạm:
- loại bỏ gitlink khỏi repository;
- đưa `tmp/deps` vào `.gitignore` phù hợp;
- tuyệt đối không phụ thuộc runtime vào `tmp`;
- sử dụng asset hợp lệ được quản lý trong `shared/` hoặc cơ chế dependency chính thức.

Chỉ cấu hình `.gitmodules` nếu repository thực sự có chủ đích dùng submodule và có nguồn/version/license rõ ràng.

Không được “fix CI” bằng cách bỏ checkout submodules một cách mù quáng nếu repository vẫn chứa gitlink lỗi.

### Acceptance P0

GitHub Actions phải checkout được repository bình thường.

---

# 4. P0 — XÂY LẠI DỮ LIỆU ĐƠN VỊ HÀNH CHÍNH VIỆT NAM

## 4.1. Tuyệt đối không dùng giả định 63 tỉnh

Các số liệu:

- 63 tỉnh/thành;
- 705 huyện;
- 10.598 xã;

là dữ liệu cũ và **KHÔNG được sử dụng làm dữ liệu hành chính hiện hành**.

Mô hình hiện hành cần sử dụng:

- **34 đơn vị hành chính cấp tỉnh**;
- **3.321 đơn vị hành chính cấp xã**;
- mô hình đơn vị hành chính/chính quyền địa phương **2 cấp: cấp tỉnh và cấp xã**;
- không coi cấp huyện là tầng hành chính hiện hành.

Nguồn pháp lý tối thiểu phải đối chiếu:

1. Nghị quyết **202/2025/QH15** về sắp xếp đơn vị hành chính cấp tỉnh.
2. Quyết định **19/2025/QĐ-TTg** về Bảng danh mục và mã số các đơn vị hành chính Việt Nam.
3. Các nghị quyết của UBTVQH về sắp xếp đơn vị hành chính cấp xã năm 2025.

Không lấy blog, repository GitHub không rõ nguồn hoặc dataset cũ làm source of truth.

---

# 5. Tách “đơn vị hiện hành” và “địa danh lịch sử”

Đây là yêu cầu rất quan trọng.

Chuan-Hoa xử lý cả văn bản mới lẫn văn bản ban hành trước 01/07/2025. Vì vậy không được xóa hoàn toàn tên:

- huyện;
- quận;
- thị xã;
- thành phố thuộc tỉnh cũ;
- thị trấn;
- tỉnh/thành trước sáp nhập.

Thay vào đó phải thiết kế ít nhất hai lớp:

### A. CurrentAdministrativeUnit

Là danh mục hiện hành, source of truth dùng cho văn bản hiện tại.

Ví dụ metadata hợp lý:

```text
code
name
canonicalName
level
type
provinceCode
effectiveFrom
effectiveTo
status
source
```

Trong đó:

```text
status = Active
level = Province | Commune
type = Province | CentrallyGovernedCity | Commune | Ward | SpecialZone
```

### B. HistoricalAdministrativeAlias

Dùng nhận diện văn bản lịch sử.

Có thể chứa:

```text
name
normalizedName
formerCode
formerLevel
formerProvince
effectiveFrom
effectiveTo
successorCode
successorName
status = Historical
source
```

Historical alias:
- được dùng để tránh false positive khi đọc văn bản cũ;
- được dùng cho search/matching;
- **không được trả về như một đơn vị hành chính hiện hành**.

Không được áp rule kiểu:

> Gặp từ “huyện” => lỗi.

Một văn bản năm 2023 ghi “Ủy ban nhân dân huyện...” có thể hoàn toàn hợp lệ.

Rule phải date/context aware nếu muốn kiểm tra tính hiện hành.

---

# 6. XÂY LẠI `administrative_units.json`

File hiện tại chứa rất nhiều fragment/rác và phải được thay thế.

Không sửa thủ công từng dòng.

Hãy tạo pipeline/generator deterministic, ví dụ dưới:

`tools/data/`

hoặc vị trí phù hợp kiến trúc hiện có.

Pipeline phải:

1. Nhập dữ liệu chính thức.
2. Chuẩn hóa Unicode NFC.
3. Trim whitespace.
4. Không làm mất dấu tiếng Việt.
5. Không tự ý bỏ tiền tố “Tỉnh”, “Thành phố”, “Xã”, “Phường”, “Đặc khu” nếu schema cần type.
6. Detect duplicate.
7. Validate code.
8. Validate parent relationship.
9. Validate effective date.
10. Xuất JSON deterministic với thứ tự ổn định.

### Validation bắt buộc

Dataset active phải xác nhận:

```text
Province-level active count = 34
Commune-level active count = 3321
```

Không chấp nhận:

- string rỗng;
- string chỉ có whitespace;
- fragment một ký tự không phải tên hợp lệ;
- source-code symbol;
- câu ghi chú;
- version string;
- dòng mô tả;
- malformed Unicode;
- duplicate code;
- duplicate record không giải thích được.

Thêm automated validator để CI fail nếu dữ liệu bị corrupt trở lại.

---

# 7. XÂY LẠI `special_capitalizations.json`

File hiện tại bị nhiễm dữ liệu extraction/code và không được dùng làm production source.

Xây lại từ các nhóm có provenance rõ ràng:

- tên 34 tỉnh/thành hiện hành;
- tên 3.321 xã/phường/đặc khu;
- tên cơ quan nhà nước;
- tên cơ quan Đảng nếu HĐ05 cần;
- tổ chức chính trị/xã hội cần thiết;
- tên quốc gia;
- tên ngày lễ/sự kiện có quy tắc viết hoa rõ ràng;
- các special capitalization có căn cứ từ NĐ30/HĐ05.

Không tự động nhập hàng nghìn cụm chưa review chỉ để đạt “target count”.

**Chất lượng > số lượng.**

Thêm validator loại bỏ:
- fragment;
- code symbol;
- câu ghi chú;
- duplicate;
- entry toàn whitespace;
- entry không có chữ;
- entry có nguồn không xác định.

---

# 8. DATA PROVENANCE

Tạo tài liệu ví dụ:

`shared/docs/data/administrative_units_provenance.md`

Ghi rõ:

- nguồn pháp lý;
- ngày hiệu lực;
- ngày snapshot;
- script tạo file;
- checksum nếu phù hợp;
- số lượng record;
- cách phân biệt active/historical;
- cách cập nhật khi Nhà nước thay đổi ĐVHC.

Không hard-code “34” ở hàng chục nơi.

Tạo central metadata/validation constant hoặc manifest để sau này cập nhật dễ dàng.

---

# 9. P0/P1 — RÀ SOÁT CONFUSION DATA

Kiểm tra toàn bộ:

`VietnameseConfusionSets.AdministrativeConfusionPairs`

và dictionary corrections.

Hiện có ít nhất trường hợp dạng:

```text
"công chứng" -> "công chứng"
```

Đây là mapping identity vô nghĩa.

Bổ sung validator:

```text
wrong != replacement
```

sau Unicode/case normalization.

Phát hiện thêm:

- duplicate key;
- correction vòng lặp A -> B và B -> A;
- replacement rỗng;
- correction gây tăng false positive;
- correction chỉ khác whitespace vô nghĩa;
- từ đúng bị đưa sang từ khác không đủ căn cứ.

Không bổ sung hàng loạt confusion pair do AI tự nghĩ ra.

Mỗi correction production phải có ít nhất một trong:
- corpus lỗi thực tế;
- nguồn từ điển/ngôn ngữ tin cậy;
- golden test;
- review thủ công có chủ đích.

---

# 10. P1 — ROLE DETECTION

File trọng tâm:

`DocumentRoleDetector.cs`

Hiện logic vẫn thiên về first-match/ordered `if-else`.

Sai role sẽ cascade sang rất nhiều rule khác.

## 10.1. Multi-signal scoring

Thiết kế deterministic role scoring thay vì chỉ first regex wins.

Mỗi candidate role nên được chấm từ các tín hiệu:

### Text
- exact text;
- regex;
- prefix;
- uppercase ratio;
- legal vocabulary;
- document type vocabulary.

### Formatting
- alignment;
- bold;
- italic;
- font size;
- font family nếu hữu ích.

### Position
- vị trí trong logical document block;
- khoảng cách tới codeNumber;
- khoảng cách tới typeName;
- role trước/sau.

### Structure
- table/cell;
- story type;
- section;
- adjacency;
- block boundary.

Có thể dùng object nội bộ:

```csharp
RoleAssignment
{
    Role
    Confidence
    Evidence
}
```

Không bắt buộc thay public API ngay.

Có thể giữ adapter:

```text
RoleAssignment -> existing Dictionary<int,string>
```

để giảm regression.

---

# 11. SUBJECT DETECTION

Loại bỏ sự phụ thuộc cứng vào logic:

```text
previousRole == typeName &&
(text.Length < 300 || StartsWith("Về việc"))
```

Subject phải nhận diện bằng tổng hợp:

- nằm ngay sau type name;
- style tương đồng với subject NĐ30;
- căn giữa;
- bold;
- font size;
- content pattern;
- vị trí trước legal basis/body;
- multi-paragraph continuity.

Không được hiểu rằng:
- subject luôn <300 ký tự;
- subject luôn bắt đầu bằng “Về việc”;
- subject continuation luôn căn giữa hoặc bold.

Cải thiện `IsSubjectContinuation()` bằng multi-signal/context.

Phải có negative test để tránh nuốt đoạn nội dung đầu tiên thành subject continuation.

---

# 12. LEGAL BASIS

Current `main` đã có logic cho phép bridge/gap giữa các đoạn căn cứ.

**Không rollback thay đổi này.**

Bổ sung regression tests và chỉ sửa nếu tìm thấy case thực tế còn sai.

Các legal basis sau phải nhận diện nhất quán giữa:

- `DocumentRoleDetector.FormalLegalBasis`;
- `CanonicalRuleScanner.Citation`;
- các rule citation liên quan.

Hiện hai vocabulary này chưa đồng nhất.

---

# 13. P1 — CITATION PARSER

Mở rộng parser/regex theo cách maintainable.

Tối thiểu hỗ trợ:

- Hiến pháp
- Bộ luật
- Luật
- Pháp lệnh
- Nghị quyết
- Nghị quyết liên tịch
- Nghị định
- Quyết định
- Chỉ thị
- Thông tư
- Thông tư liên tịch
- Lệnh
- Công văn
- Tờ trình

Không tạo một regex khổng lồ duplicate ở nhiều class.

Ưu tiên central vocabulary hoặc helper parser dùng chung.

Bổ sung test:

- có “số”;
- không có “số” khi cấu trúc hợp lệ;
- số văn bản có `/`, `-`, ký hiệu;
- chữ hoa/chữ thường;
- Unicode;
- nhiều citation trong cùng paragraph.

---

# 14. CODE NUMBER

`CanonicalRuleScanner.CodeNumber` hiện yêu cầu line bắt đầu bằng `Số`.

Rà soát các cấu trúc Word thực tế:

- văn bản thông thường;
- table cell;
- cơ quan và “Số” nằm cùng cell/line;
- placeholder `…/…`;
- có/không colon;
- Word control characters.

Không relax regex đến mức match mọi chữ “số” trong body.

Dùng structural context/table coordinates để giảm false positive.

---

# 15. DATE CHECK PHẢI CONTEXT-AWARE

`CheckBareShortDates()` hiện có nguy cơ cảnh báo `DD/MM/YYYY` quá rộng.

Phân biệt:

- legal basis text;
- bảng biểu;
- header/footer;
- footnote/endnote nếu snapshot có;
- citation trong ngoặc;
- body;
- metadata;
- biểu mẫu.

Không tự động coi mọi `01/07/2025` là lỗi vì thiếu chữ “ngày”.

Viết tests cho cả positive và negative cases.

---

# 16. P1 — TYPOGRAPHY CLEANER

Current code đã có fix cho:

- `1,5`;
- `14:30`;
- một số abbreviation.

Phải giữ regression này.

Nhưng `UrlOrEmailRegex` hiện được khai báo mà chưa thực sự bảo vệ toàn span trước các mutation punctuation khác.

Hãy thiết kế protection mechanism:

```text
Protect spans
→ normalize text outside protected spans
→ restore/reconstruct spans
```

Phải bảo toàn tuyệt đối:

- `https://...`
- `http://...`
- `www.example.com`
- domain hợp lệ nếu hỗ trợ;
- email;
- query strings;
- IPv4 nếu scanner hỗ trợ;
- file/path strings nếu có risk.

Ví dụ không được biến:

```text
example.com
```

thành:

```text
example. com
```

---

# 17. QUOTATION MARKS

`NormalizeQuotationMarks()` hiện dùng state toggle.

Cải thiện trường hợp số dấu `"` lẻ.

Không được để một quote thiếu làm đảo trạng thái toàn bộ phần còn lại.

Có thể:

1. đếm/parity trước;
2. dùng neighbor heuristic;
3. hoặc không auto-fix khi confidence thấp.

Nguyên tắc:

**Không chắc chắn => finding/advisory, không destructive auto-fix.**

---

# 18. TOKENIZER / URL

Kiểm tra:

`VietnameseWordTokenizer.TokenRegex`

URL regex hiện cần được cải thiện.

Không dùng wildcard `.` sai mục đích.

Hỗ trợ hợp lý:

- `http://`
- `https://`
- `www.`
- email
- domain nếu đủ chắc chắn.

Quan trọng nhất:

- offset phải tuyệt đối chính xác;
- tổng token span phải map lại đúng text;
- không làm lệch annotation anchor.

Thêm round-trip/offset tests.

---

# 19. P1 — SPELL CHECKER

## 19.1. Sentence-start capitalization

Hiện unknown title-case token có thể bị bỏ qua quá rộng vì ký tự đầu viết hoa.

Thay:

```text
char.IsUpper(firstLetter) => ignore
```

bằng context-aware decision.

Tokenizer đã có:
- `StartOffset`;
- `SentenceIndex`;
- kind.

Hãy tận dụng chúng.

Cần phân biệt:

- tên người;
- tên cơ quan;
- tên địa danh;
- từ đầu câu;
- từ đầu paragraph;
- từ viết hoa sau `. ! ?`;
- acronym.

Một từ sai chính tả ở đầu câu không được miễn kiểm tra chỉ vì chữ đầu viết hoa.

---

# 20. LEXICON

Rà soát nguồn Hunspell hiện có.

Không tự động tăng dictionary lên một con số tùy ý.

Trước khi redistribute/import asset:
- xác nhận license;
- lưu provenance/license;
- tránh tạo dependency online runtime.

Nếu sử dụng Hunspell-derived word list:
- pipeline build phải deterministic;
- runtime vẫn offline;
- normalization NFC;
- deduplicate case-insensitive;
- test Vietnamese diacritics.

Lưu ý tiếng Việt có cấu trúc âm tiết/từ ghép; chỉ mở rộng syllable dictionary không giải quyết toàn bộ contextual spelling.

---

# 21. PHONETIC-AWARE SUGGESTION

Có thể cải thiện suggestion bằng weighted Vietnamese distance.

Ví dụ nhóm dễ nhầm:

```text
s ↔ x
tr ↔ ch
d ↔ gi ↔ r
l ↔ n
hỏi ↔ ngã
c ↔ t ở âm cuối trong các trường hợp phù hợp
n ↔ ng ở âm cuối trong các trường hợp phù hợp
```

Nhưng:

- chỉ dùng cho ranking suggestion;
- không tự động thay thế khi ambiguous;
- phải deterministic;
- phải có max distance;
- phải có margin giữa candidate #1 và #2;
- nếu confidence không đủ thì chỉ báo “hãy kiểm tra”.

Không biến phonetic generator thành cơ chế tạo false positive hàng loạt.

---

# 22. COMPOUND WORD / CONTEXT

Thiết kế infrastructure cho compound phrase/bigram nhưng không bật aggressive production rule nếu chưa có corpus đủ tin cậy.

Có thể hỗ trợ trước:

- phrase dictionary;
- known administrative compound;
- protected compound spans.

Ví dụ:

```text
quyết định
nghị quyết
cơ quan
ban hành
thực hiện
ủy ban nhân dân
hội đồng nhân dân
```

Nếu triển khai bigram:
- corpus phải có provenance;
- versioned;
- không đưa raw user documents vào repository;
- không thu thập nội dung người dùng;
- threshold conservative;
- feature flag/profile;
- test precision.

Nếu chưa có golden corpus đủ tốt, infrastructure có thể được merge nhưng aggressive bigram finding phải default OFF.

---

# 23. PERFORMANCE

## 23.1 Regex

Current helper `Rx()` vẫn tạo `new Regex(...)`.

Phân loại regex:

### Static pattern
Chuyển thành `static readonly Regex` hoặc cache compile một lần.

### Dynamic pattern từ rule pack
Compile/cache khi load pack, không compile lại mỗi paragraph.

Luôn giữ timeout để tránh ReDoS.

Không compile cache không giới hạn với arbitrary pattern.

---

# 24. NORMALIZED OFFSET MAP

`WholePhraseMatches()` hiện tạo:

```text
NormalizedTextOffsetMap.Create(text)
```

lặp lại theo từng phrase.

Refactor để mỗi paragraph tạo normalized map một lần và reuse cho:

- corrections;
- confusion pairs;
- capitalization;
- personal dictionary phrases;
- các phrase matching khác.

Không làm sai Unicode offset.

Viết tests với:

- NFC;
- NFD;
- tiếng Việt có dấu;
- emoji/surrogate nếu relevant;
- Word paragraph terminators.

---

# 25. KHÔNG REIMPLEMENT NHỮNG FIX ĐÃ CÓ

Current main đã có một số thay đổi so với báo cáo cũ.

Trước mỗi hạng mục:

```text
inspect current implementation
→ inspect tests
→ decide status:
   DONE
   PARTIAL
   NOT DONE
→ implement only missing portion
```

Tối thiểu cần kiểm tra lại:

- legal-basis bridge;
- numeric comma;
- time colon;
- abbreviation handling;
- reuse DetectBlocks;
- confusion set expansion.

Không rollback các regression fix đã đúng.

---

# 26. GOLDEN CORPUS

Tạo framework golden corpus nếu chưa có.

Không commit văn bản mật/nội bộ.

Dùng:

- văn bản công khai;
- synthetic examples;
- fixture đã được phép sử dụng.

Mỗi case nên có:

```text
input
expected roles
expected findings
forbidden findings
document type
regime
reason/source
```

Đánh giá tối thiểu:

```text
precision
recall
false positives
false negatives
```

Ưu tiên giảm **false positive** đối với auto-fix/destructive changes.

Finding dạng advisory có thể nhạy hơn nhưng vẫn phải kiểm soát noise.

---

# 27. TEST BẮT BUỘC CHO ĐỊA GIỚI 2025+

Viết automated tests xác nhận:

```text
active province count == 34
active commune count == 3321
```

Kiểm tra representative codes/names, ví dụ:

```text
01 -> Thành phố Hà Nội
46 -> Thành phố Huế
48 -> Thành phố Đà Nẵng
79 -> Thành phố Hồ Chí Minh
92 -> Thành phố Cần Thơ
```

Không chỉ test count.

Test thêm:

- tất cả active commune có code hợp lệ;
- parent tỉnh tồn tại;
- không duplicate active code;
- không active district-level unit;
- historical unit không bị coi active;
- alias cũ vẫn được tokenizer/spell checker nhận biết phù hợp;
- current spelling/capitalization đúng.

---

# 28. DATA VALIDATION GATE

Tạo validator tích hợp vào:

`tools/validation/run_source_quality_gates.ps1`

Validator phải fail CI nếu:

- administrative count sai;
- malformed JSON;
- fragment/rác tái xuất hiện;
- correction identity;
- duplicate code;
- historical record đánh active sai;
- metadata source/effective date thiếu;
- Unicode normalization sai.

Không chỉ viết unit test; cần một validation gate rõ ràng cho reference data.

---

# 29. TEST TYPOGRAPHY/TOKENIZER

Thêm regression tests cho ít nhất:

```text
1,5
14:30
https://example.com/a?x=1
www.example.com
user@example.com
PGS.TS. Nguyễn Văn A
KT. BỘ TRƯỞNG
TM. ỦY BAN NHÂN DÂN
01/07/2025
"Cụm từ có quote thiếu
```

Đảm bảo cleaner không mutate URL/email/domain.

---

# 30. TEST ROLE DETECTION

Bổ sung:

- subject >300 chars;
- subject không bắt đầu “Về việc”;
- subject không bold;
- subject không centered;
- multi-line subject;
- body ngay sau subject;
- legal basis có 1 bridge;
- legal basis có 2 bridge;
- legal basis kết thúc đúng trước Điều 1;
- nhiều văn bản trong cùng document;
- organ name trong bảng;
- official letter;
- decision;
- report;
- Party/HĐ05 document.

Test cả positive và negative để chống over-detection.

---

# 31. PERFORMANCE TEST

Đo trước/sau trên tài liệu synthetic lớn và representative documents.

Ghi:

```text
paragraph count
rule count
elapsed
findings count
allocations nếu có thể
```

Không tối ưu bằng cách bỏ rule.

Mục tiêu:
- không regression đáng kể;
- ưu tiên cải thiện repeated regex/normalization;
- cancellation vẫn hoạt động;
- không tăng memory không giới hạn.

---

# 32. LOCAL RULE PACK

Current parser hỗ trợ schema v1/v2 và validate cấu trúc khá chặt.

Ưu tiên **không sửa schema** cho các cải tiến có thể nằm trong code/reference assets.

Nếu confidence chỉ là internal result của role detector thì không cần đẩy vào signed rule pack ngay.

Nếu thêm threshold production configurable:
- cân nhắc profile tương tự advisory profile hiện có;
- backward-compatible default;
- malformed config phải fail-safe/conservative.

---

# 33. SECURITY / PRIVACY

Không được vì accuracy mà:

- upload document;
- log raw document;
- lưu paragraph vào telemetry;
- gửi corpus người dùng;
- tạo internet dependency ở runtime.

Quality gate privacy hiện tại phải tiếp tục pass.

Không thêm secret/token vào repository.

---

# 34. THỨ TỰ TRIỂN KHAI

Thực hiện theo thứ tự:

### Phase 0
- Diagnose + fix broken gitlink/submodule.
- Establish clean test baseline.

### Phase 1
- Rebuild administrative data.
- Rebuild special capitalization data.
- Add provenance.
- Add reference-data validators.

### Phase 2
- Fix obvious deterministic bugs:
  - confusion identity;
  - URL protection;
  - tokenizer URL regex;
  - title-case spell exemption;
  - citation vocabulary;
  - context-aware date;
  - regex/map caching.

### Phase 3
- Refactor role detection/scoring.
- Subject detection/continuation.

### Phase 4
- Lexicon/phonetic/context infrastructure.

### Phase 5
- Golden corpus.
- Performance regression.
- Full quality/security verification.

Mỗi phase phải giữ solution buildable.

---

# 35. COMMANDS PHẢI CHẠY

Ít nhất chạy:

```powershell
dotnet test .\ChuanHoa.slnx -c Release
```

và:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass `
  -File .\tools\validation\run_source_quality_gates.ps1
```

Nếu cần chạy offline/local khi NuGet audit bên ngoài không khả dụng, có thể dùng option hiện có để chẩn đoán, nhưng **không được coi PASS_WITH_NOT_RUN_NUGET_AUDIT là bằng chứng tương đương full CI**.

Sau khi push branch/commit, nếu môi trường có GitHub CLI/authentication:
- kiểm tra GitHub Actions;
- xem log nếu fail;
- sửa tới khi Source Quality xanh.

---

# 36. KHÔNG HARD-CODE TEST COUNT CŨ

Không giả định project vẫn là `403/403`.

Test count của báo cáo trước có thể đã stale.

Hãy lấy số thực tế sau khi checkout current main.

Không sửa quality gate test floor xuống chỉ vì test discovery thay đổi.

---

# 37. CẬP NHẬT DOCUMENTATION

Sau khi code pass thực tế mới cập nhật:

- implementation status;
- data provenance;
- architecture notes nếu thay đổi;
- accuracy/golden corpus report.

Không ghi “PASS” nếu chưa chạy.

Không giữ các con số cũ như:

```text
63 tỉnh
705 huyện
10.598 xã
```

ở tài liệu mô tả **hệ thống hiện hành**.

Nếu chúng cần tồn tại để mô tả lịch sử thì phải ghi rõ:

```text
LEGACY / PRE-2025
```

và không được dùng làm current source of truth.

---

# 38. DEFINITION OF DONE

Chỉ coi nhiệm vụ hoàn tất khi thỏa tất cả điều kiện:

### Repository
- Không còn broken gitlink/submodule.
- Clean checkout thành công trên GitHub Actions.

### Administrative data
- 34 active province-level units.
- 3.321 active commune-level units.
- Không có active district layer.
- Historical aliases tách riêng.
- Có provenance/effective dates.
- Không còn fragment/rác trong production dictionary.

### Accuracy
- Role detection tốt hơn và có regression tests.
- Subject dài/không chuẩn vẫn được xử lý phù hợp.
- Legal-basis fix hiện tại không bị regression.
- Citation vocabulary đồng bộ.
- Title-case đầu câu không còn được miễn spell check mù quáng.
- URL/email không bị typography cleaner phá.
- Token offset chính xác.
- Date detection context-aware.
- Confusion mappings được validate.

### Performance
- Không repeated normalization nghiêm trọng.
- Không dynamic regex compilation không cần thiết trong hot loops.
- Regex timeout vẫn được bảo toàn.

### Compatibility
- Local/offline architecture giữ nguyên.
- LocalRulePack compatibility giữ nguyên hoặc có migration đầy đủ.
- Existing Word/VSTO behavior không bị phá.

### Verification
- Full solution tests PASS.
- `run_source_quality_gates.ps1` PASS.
- Source Quality GitHub Actions PASS.
- Không còn warning/error mới bị bỏ qua.

---

# 39. CÁCH BÁO CÁO KẾT QUẢ CHO TÔI

Khi hoàn tất, trả về một báo cáo có đúng các phần sau:

## A. Baseline trước khi sửa
- commit SHA;
- ProductVersion;
- test result;
- CI status;
- các lỗi thực tế xác nhận được.

## B. Những phát hiện từ code
Bảng:

| Severity | File | Function/Area | Vấn đề | Root cause | Đã xử lý |
|---|---|---|---|---|---|

## C. Các file đã thay đổi

Với mỗi file:
- thay đổi gì;
- tại sao;
- ảnh hưởng behavior gì.

## D. Administrative dataset

Báo cáo:

```text
Active provinces:
Active communes:
Historical aliases:
Duplicates:
Invalid records:
Source:
Effective date:
```

## E. Accuracy

So sánh before/after:

```text
Role detection
Spell check
Typography
Citation
False positives
False negatives
```

Nếu chưa có đủ golden corpus để có metric đáng tin cậy, ghi rõ “NOT MEASURED”, tuyệt đối không tự bịa số.

## F. Performance

Before/after nếu đo được.

## G. Tests

Nêu:
- tổng pass;
- fail;
- skipped;
- test mới;
- regression test.

## H. Quality gates

Đưa kết quả chính xác của:

```text
run_source_quality_gates.ps1
GitHub Source Quality
```

## I. Remaining risks

Những vấn đề chưa thể xử lý an toàn phải ghi rõ, không giấu bằng TODO mơ hồ.

## J. Git information

```text
branch
commit SHA
commit message(s)
working tree status
```

---

# 40. NGUYÊN TẮC CUỐI CÙNG

Không hỏi tôi từng bước phải làm gì.

Hãy tự nghiên cứu repository rồi triển khai end-to-end.

Khi gặp khác biệt giữa tài liệu cũ và code hiện tại:

**CODE CURRENT MAIN + TEST + NGUỒN PHÁP LÝ HIỆN HÀNH là source of truth.**

Không coi thông tin “63 tỉnh/thành” là dữ liệu hiện hành.

Không tối ưu theo số lượng dictionary entries.

Không sửa bằng heuristic mới nếu chưa có negative tests.

Không thêm auto-fix khi confidence thấp.

Ưu tiên:

**Correctness → tránh false positive → deterministic behavior → compatibility → performance.**

Bắt đầu bằng việc xác nhận current commit, sửa lỗi checkout CI, dựng baseline, sau đó lần lượt triển khai các phase trên cho tới khi quality gates xanh.