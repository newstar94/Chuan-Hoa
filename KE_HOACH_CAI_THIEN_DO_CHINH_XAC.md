# Kế Hoạch Cải Thiện Độ Chính Xác Nhận Diện Và Chuẩn Hóa Văn Bản Hành Chính

## 1. Tổng quan dự án sau rà soát

Dự án **Chuẩn Hóa** là một add-in VSTO/C# cho Microsoft Word, kiểm tra và chuẩn hóa thể thức văn bản hành chính theo Nghị định 30/2020 (NĐ30) và Hướng dẫn 05 (HĐ05) của Đảng. Toàn bộ xử lý chạy local, không gửi nội dung lên server.

### Kiến trúc hiện tại

| Thành phần | Vai trò | File chính |
|---|---|---|
| **DocumentRoleDetector** | Nhận diện vai trò từng đoạn (Quốc hiệu, Tiêu ngữ, Số hiệu, Căn cứ, …) | `src/ChuanHoa.Client.Core/Scanning/DocumentRoleDetector.cs` |
| **CanonicalRuleScanner** | Quét ~72 route thể thức NĐ30/HĐ05 + 7 route LaTeX | `src/ChuanHoa.Client.Core/Scanning/CanonicalRuleScanner.cs` |
| **HeadingDetector** | Phát hiện heading/tiểu mục theo decimal, roman, alphabet | `src/ChuanHoa.Client.Core/Scanning/HeadingDetector.cs` |
| **VietnameseLexiconSpellChecker** | Kiểm tra chính tả dựa trên 6.710 mục từ | `src/ChuanHoa.Client.Core/Scanning/VietnameseLexiconSpellChecker.cs` |
| **VietnameseConfusionSets** | ~60 cặp nhầm lẫn phụ âm/thanh điệu | `src/ChuanHoa.Client.Core/Lexicon/VietnameseConfusionSets.cs` |
| **VietnameseTypographyCleaner** | Dọn khoảng trắng, dấu câu, ngoặc kép | `src/ChuanHoa.Client.Core/Text/VietnameseTypographyCleaner.cs` |
| **VietnameseWordTokenizer** | Tách từ offset-preserving | `src/ChuanHoa.Client.Core/Lexicon/VietnameseWordTokenizer.cs` |
| **LocalRulePack** | Gói quy tắc có chữ ký, chứa từ điển + cấu hình | `src/ChuanHoa.Client.Core/Rules/LocalRulePack.cs` |

---

## 2. Phân tích điểm yếu về độ chính xác

Sau khi rà soát toàn bộ ~3.500 dòng logic nhận diện và chuẩn hóa, xác định **7 nhóm vấn đề chính** ảnh hưởng đến accuracy:

---

### Nhóm 1. Role Detection: Heuristic cứng, thiếu tín hiệu đa chiều

> Nguồn gốc hàng đầu gây false positive/negative. Sai role sẽ làm sai toàn bộ cascade kiểm tra thể thức tiếp theo.

**Vấn đề hiện tại:**
- `DocumentRoleDetector.DetectBlock()` dùng chuỗi `if-else` tuyến tính, đoạn nào match regex trước được gán role trước, không có scoring/ranking.
- `"subject"` chỉ được gán nếu đoạn trước là `"typeName"` và text ngắn hơn 300 ký tự hoặc bắt đầu bằng "Về việc" — bỏ sót subject dài hoặc không có tiền tố chuẩn.
- `"legalBasis"` phụ thuộc vào cửa sổ trạng thái (`legalBasisWindowOpen`) nhưng window bị đóng sớm nếu gặp đoạn không match bất kỳ role nào → bỏ sót căn cứ pháp lý ở giữa preamble.
- `IsSubjectContinuation()` chỉ dựa vào `Alignment == Center || Bold` — sai khi subject không căn giữa hoặc không đậm.
- Không xét formatting (font size, spacing) để hỗ trợ role detection; chỉ dùng text pattern.

**Cải thiện đề xuất:**
1. **Multi-signal scoring cho role detection:** Mỗi đoạn nhận điểm từ: text pattern, formatting (bold/size/align), vị trí tương đối, khoảng cách đến role đã biết. Chọn role có score cao nhất thay vì first-match.
2. **Mở rộng subject detection:** Bỏ giới hạn 300 ký tự; cho phép subject không bắt đầu bằng "Về việc" nếu đoạn trước là `typeName` và đoạn có style consistent (centered, bold, size 13–14).
3. **Cải thiện legal basis window:** Window không đóng ngay khi gặp đoạn không match — cho phép 1–2 đoạn "bridge" (blank, subject continuation, structural title) giữa các căn cứ.
4. **Formatting-assisted role hints:** Nếu đoạn viết hoa toàn bộ + căn giữa + bold + size 12-13 + nằm trong 15 đoạn đầu → hint rất mạnh cho `organName`/`superiorOrganName`.
5. **Confidence score cho mỗi role:** Trả về `RoleAssignment(role, confidence)` thay vì `string`. Các scanner downstream có thể skip rule khi confidence < threshold.

---

### Nhóm 2. Spell Checker: Từ điển nhỏ, thiếu ngữ cảnh, false positive cao

> 6.710 mục từ chỉ bao phủ ~55–60% âm tiết thường gặp trong văn bản hành chính. Token viết thường không có trong từ điển sẽ bị báo sai.

**Vấn đề hiện tại:**
- `ShouldIgnoreLexiconToken()` bỏ qua token viết hoa đầu (giả sử là tên riêng) → đúng cho tên người nhưng sai cho từ viết hoa đầu câu bị sai chính tả.
- Edit distance dùng Levenshtein cổ điển, không có phonetic awareness → "xử lý" và "sử lý" có distance 1 nhưng "quản lý" và "quảng lý" cũng distance 1 → ambiguous.
- `AdministrativeConfusionPairs` chỉ có ~60 cặp hardcode — chưa bao phủ nhiều lỗi thực tế.
- Không có n-gram/bigram context → không phân biệt được "bổ sung" (đúng) vs "bổ xung" (sai) khi cả hai từ đơn đều hợp lệ.

**Cải thiện đề xuất:**
1. **Mở rộng từ điển lên ≥15.000 mục:** Thu thập từ Hunspell vi-VN (đã có trong `shared/dictionaries/hunspell-vi`), Wiktionary tiếng Việt, và corpus văn bản hành chính.
2. **Phonetic-aware edit distance:** Thay Levenshtein thuần bằng weighted distance: S↔X, TR↔CH, D↔GI↔R, L↔N có weight 0.5 thay vì 1.0; hỏi↔ngã weight 0.3.
3. **Bigram context checker:** Xây bảng bigram frequency từ corpus hành chính (≥50.000 cặp từ). Khi token đúng từ điển nhưng bigram (previous_token, current_token) rất hiếm → cảnh báo nhẹ.
4. **Mở rộng ConfusionPairs lên ≥200 cặp:** Thêm: "kê khai"↔"kể khai", "phối hợp"↔"phối hộp", "đề nghị"↔"đề ngị", "triển khai"↔"chiển khai", "phương án"↔"phương àn", v.v.
5. **Sentence-start capitalization aware:** Không bỏ qua token viết hoa đầu nếu nó đứng ngay sau `.`, `!`, `?` hoặc đầu đoạn — vẫn kiểm tra spelling.

---

### Nhóm 3. Typography Cleaner: Thiếu edge case tiếng Việt

**Vấn đề hiện tại:**
- `SpaceAfterPunctuationRegex` pattern `([,;!?]|:(?!\/\/))([A-Za-zÀ-ỹ])` sẽ thêm space sai trong trường hợp số thập phân dùng dấu phẩy (ví dụ: `1,5` → `1, 5`), thời gian (`14:30` → `14: 30`), hoặc số điện thoại.
- `PeriodFollowedByLetterRegex` negative lookbehind chỉ cover `v.v|tp|gs|ts|pgs|th` — thiếu nhiều viết tắt hành chính: `Ths.`, `PGS.TS.`, `KT.`, `TM.`, `TL.`, `Q.`, `PGĐ.`.
- `NormalizeQuotationMarks()` giả sử dấu `"` luôn đi theo cặp → nếu thiếu một dấu, state `inQuotes` sẽ sai cho toàn bộ phần còn lại.

**Cải thiện đề xuất:**
1. **Exclude numeric contexts từ space insertion:** Thêm negative lookahead/lookbehind cho digits trước/sau dấu phẩy, hai chấm, chấm.
2. **Mở rộng abbreviation allowlist:** Load danh sách viết tắt hành chính từ `non_sentence_ending_abbreviations.json` vào regex.
3. **Robust quote pairing:** Đếm số `"` chẵn/lẻ trước khi normalize; nếu lẻ → skip hoặc dùng heuristic (dấu `"` sát chữ cái = mở, sát dấu câu = đóng).
4. **Bảo toàn URL/email:** Detect và protect URL/email spans trước khi chạy space insertion.

---

### Nhóm 4. Code Number & Citation Parsing: Regex cứng, bỏ sót biến thể

**Vấn đề hiện tại:**
- `CodeNumber` regex `^\s*Số\s*:?\s*\d+` yêu cầu bắt đầu bằng "Số" — bỏ sót khi dòng Số nằm trong bảng hoặc có tiền tố khác.
- `Citation` regex không match "Bộ luật" (hai từ), không match "Lệnh" hoặc "Công văn" khi dùng trong viện dẫn.
- `LegalBasisShortDate` cảnh báo ngày viết dạng số thiếu "ngày" phía trước nhưng không check ngữ cảnh footnote, bảng biểu, hoặc citation đã có format chuẩn.

**Cải thiện đề xuất:**
1. **Mở rộng Citation regex:** Thêm: "Bộ luật", "Lệnh", "Công văn", "Tờ trình", "Nghị quyết liên tịch", "Thông tư liên tịch".
2. **Context-aware date check:** Không cảnh báo ngày dạng DD/MM/YYYY nếu nằm trong bảng biểu, header/footer, hoặc ngoặc đơn citation.
3. **Relaxed code number detection:** Cho phép match khi "Số" đứng sau tên cơ quan trong cùng ô bảng.

---

### Nhóm 5. Dữ liệu tham chiếu: Corrupted và không đầy đủ

> `special_capitalizations.json` chứa dữ liệu bị hỏng — hầu hết là fragment vô nghĩa (`" Ch"`, `" l"`, `"ng"`, `"t"`, ...) thay vì tên riêng đúng chuẩn viết hoa.

**Vấn đề hiện tại:**
- `special_capitalizations.json` (85 entries) chứa ~90% dữ liệu rác.
- `typo_dictionary.json` chỉ có 31 entry — quá ít cho production.
- `administrative_units.json` (1.522 entries) chứa nhiều fragment không đầy đủ.
- `iy_dictionary.json` chỉ có 17 entry — thiếu nhiều cặp i/y quan trọng.

**Cải thiện đề xuất:**
1. **Xây lại `special_capitalizations.json`:** Thu thập từ danh mục chính thức: tên cơ quan nhà nước (Bộ, UBND, Sở, Phòng…), tên đơn vị hành chính (63 tỉnh, 705 huyện, 10.598 xã), tên tổ chức chính trị-xã hội, ngày lễ, tên địa danh. Target: ≥2.000 entry.
2. **Mở rộng `typo_dictionary.json`:** Thu thập từ corpus lỗi thực tế. Target: ≥200 entry.
3. **Xây lại `administrative_units.json`:** Import từ danh mục ĐVHC cấp xã của Tổng cục Thống kê (mã 63/705/10598). Format đúng: tên đầy đủ, không fragment.
4. **Mở rộng `iy_dictionary.json`:** Thêm: "kỷ luật/kỉ luật", "lý lịch/lí lịch", "phương pháp luận/phương pháp luận", "vĩ mô/vĩ mô", v.v. Target: ≥50 entry.

---

### Nhóm 6. Tokenizer: Bỏ sót compound word và hyphenated term

**Vấn đề hiện tại:**
- `VietnameseWordTokenizer` tách theo syllable (âm tiết), không có word segmentation → "kinh tế" thành hai token "kinh" + "tế" → cả hai đúng từ điển → bỏ sót lỗi ở cấp compound word.
- Regex URL pattern `https?://[^\s/$.?#].[^\s]*` có thể match quá rộng hoặc miss URL không có protocol.

**Cải thiện đề xuất:**
1. **Compound word lookup:** Thêm dictionary compound word hành chính (≥5.000 entry). Khi hai syllable liên tiếp tạo compound đúng → skip spell check cả cặp.
2. **Fix URL regex:** Escape `.` sau host; thêm match cho `www.` prefix và bare domain.

---

### Nhóm 7. Performance & Maintainability ảnh hưởng gián tiếp đến accuracy

**Vấn đề hiện tại:**
- `Rx()` helper tạo `new Regex()` mỗi lần gọi trong cả `CanonicalRuleScanner` và `DocumentRoleDetector` — không compile, không cache → chậm trên document lớn → timeout 200ms có thể gây miss.
- `CheckPlaceDate()` tạo `new DocumentRoleDetector()` bên trong → detect blocks lại từ đầu thay vì reuse kết quả đã có.
- `WholePhraseMatches()` tạo `NormalizedTextOffsetMap` mỗi lần gọi cho mỗi paragraph × mỗi rule → O(n×m) normalization.

**Cải thiện đề xuất:**
1. **Cache compiled Regex:** Chuyển `Rx()` inline thành `static readonly Regex` với `RegexOptions.Compiled` cho tất cả pattern dùng nhiều lần.
2. **Reuse DetectBlocks result:** `CheckPlaceDate()` nhận `blocks` parameter thay vì tạo `new DocumentRoleDetector()`.
3. **Cache NormalizedTextOffsetMap:** Tạo map một lần per paragraph, reuse cho tất cả `WholePhraseMatches` call.

---

## 3. Lộ trình triển khai theo thứ tự ưu tiên

| Ưu tiên | Hạng mục | Tác động | Khối lượng | Trạng thái |
|---|---|---|---|---|
| 🔴 **P0** | **5.1** Xây lại `special_capitalizations.json` từ dữ liệu chuẩn | Rất cao | Trung bình | Sẵn sàng triển khai |
| 🔴 **P0** | **1.3** Sửa legal basis window trong `DocumentRoleDetector.cs` | Cao | Thấp | Sẵn sàng triển khai |
| 🔴 **P0** | **5.2** Mở rộng `typo_dictionary.json` | Cao | Thấp | Sẵn sàng triển khai |
| 🔴 **P0** | **5.3** Xây lại `administrative_units.json` | Cao | Trung bình | Sẵn sàng triển khai |
| 🟡 **P1** | **3.1** Exclude numeric contexts trong `VietnameseTypographyCleaner.cs` | Cao | Thấp | Sẵn sàng triển khai |
| 🟡 **P1** | **3.2** Mở rộng abbreviation allowlist trong `VietnameseTypographyCleaner.cs` | Cao | Thấp | Sẵn sàng triển khai |
| 🟡 **P1** | **3.4** Bảo toàn URL/email trong typography cleaner | Trung bình | Thấp | Sẵn sàng triển khai |
| 🟡 **P1** | **7.1** Cache compiled Regex trong scanner và role detector | Trung bình | Thấp | Sẵn sàng triển khai |
| 🟡 **P1** | **7.2** Reuse DetectBlocks trong `CheckPlaceDate()` | Trung bình | Thấp | Sẵn sàng triển khai |
| 🟡 **P1** | **7.3** Cache `NormalizedTextOffsetMap` theo paragraph | Trung bình | Thấp | Sẵn sàng triển khai |
| 🟢 **P2** | **2.4** Mở rộng `VietnameseConfusionSets.AdministrativeConfusionPairs` | Cao | Trung bình | Kế tiếp |
| 🟢 **P2** | **4.1** Mở rộng Citation regex (Bộ luật, Lệnh, Tờ trình...) | Trung bình | Thấp | Kế tiếp |
| 🟢 **P2** | **5.4** Mở rộng `iy_dictionary.json` | Trung bình | Thấp | Kế tiếp |
| 🔵 **P3** | **1.1 & 1.5** Multi-signal scoring cho Role Detection | Rất cao | Cao | Giai đoạn 2 |
| 🔵 **P3** | **2.1 - 2.3** Bigram & Phonetic-aware spell checker | Rất cao | Cao | Giai đoạn 2 |
