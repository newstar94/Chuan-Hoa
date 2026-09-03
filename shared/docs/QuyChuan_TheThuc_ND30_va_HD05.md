# HỆ THỐNG QUY CHUẨN THỂ THỨC VĂN BẢN TOÀN DIỆN
## ĐỐI CHIẾU NGHỊ ĐỊNH 30/2020/NĐ-CP, HƯỚNG DẪN 05-HD/VPTW (ĐẢNG), VIETTEL & 82 QUY TẮC KIỂM TRA

> **Nguồn tài liệu tổng hợp**:
> 1. `Nghị định 30.doc` & `Phụ lục Nghị định 30.doc` (Nghị định 30/2020/NĐ-CP ngày 05/03/2020).
> 2. `Hướng dẫn 05.docx` (Hướng dẫn số 05-HD/VPTW ngày 27/05/2026 của Văn phòng TW Đảng).
> 3. Hệ thống 68 modules mã nguồn VBA (`RuleData.bas`, `ComplianceChecker.bas`, `ComponentDetector.bas`, `ToneNormalizer.bas`, `TableFormatter.bas`...).

---

## 1. MA TRẬN ĐỐI CHIẾU 3 CHẾ ĐỘ THỂ THỨC (REGIMES)

| Thành phần thể thức | 1. Văn bản Hành chính (NĐ 30/2020/NĐ-CP) | 2. Văn bản của Đảng (HD 05-HD/VPTW/2026) | 3. Quy chế Doanh nghiệp (Viettel QĐ 11095) |
| :--- | :--- | :--- | :--- |
| **Phông chữ mặc định** | `Times New Roman`, Unicode NFC, Đen | `Times New Roman`, Unicode NFC, Đen | `Times New Roman`, Unicode NFC, Đen |
| **Căn lề trang A4** | Top/Bottom: 20-25mm; Left: 30-35mm; Right: 15-20mm | Top/Bottom: 20-25mm; Left: 30-35mm; Right: 15-20mm | Top/Bottom: 20-25mm; Left: 30-35mm; Right: 15-20mm |
| **Đánh số trang** | Header căn giữa, cỡ 13-14 đứng, ẩn trang 1 | Header/Footer căn giữa, cỡ 14 đứng, ẩn trang 1 | Header căn giữa, cỡ 13-14 đứng, ẩn trang 1 |
| **1. Tiêu đề / Quốc hiệu** | **CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM**<br>- Cỡ 12-13, In hoa, Đậm.<br>- Tiêu ngữ: *Độc lập - Tự do - Hạnh phúc* (cỡ 13-14, Đậm, kẻ nét liền $100\%$ độ dài Tiêu ngữ). | **ĐẢNG CỘNG SẢN VIỆT NAM**<br>- Cỡ 15, In hoa, Đậm.<br>- Phía dưới có đường kẻ ngang nét liền dài $100\%$ độ dài Tiêu đề. | **CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM**<br>- Cỡ 12, In hoa, Đậm.<br>- Tiêu ngữ: *Độc lập - Tự do - Hạnh phúc* (cỡ 13, Đậm, kẻ nét liền $100\%$). |
| **2. Tên cơ quan ban hành** | - CQ cấp trên: Cỡ 12-13, In hoa, Đứng.<br>- CQ ban hành: Cỡ 12-13, In hoa, Đậm.<br>- Kẻ nét liền $1/3 - 1/2$ độ dài tên cơ quan. | - CQ cấp trên: Cỡ 14, In hoa, Đứng.<br>- CQ ban hành: Cỡ 14, In hoa, Đậm.<br>- Dưới tên CQ có **dấu sao (`*`)** phân cách. | - CQ cấp trên: Cỡ 12, In hoa, Đứng (`TẬP ĐOÀN...`).<br>- CQ ban hành: Cỡ 12, In hoa, Đậm.<br>- Kẻ nét liền $1/3 - 1/2$ độ dài tên cơ quan. |
| **3. Số và Ký hiệu** | Cỡ 13, Đứng.<br>Cú pháp: `Số: 15/QĐ-BNV` (có dấu hai chấm) | Cỡ 14, Đứng.<br>Cú pháp: `Số 05-HD/VPTW` (không có hai chấm) | Cỡ 13, Đứng.<br>Cú pháp: `Số: 1234/VT-TCHC` |
| **4. Địa danh, ngày tháng** | Cỡ 13-14, Nghiêng.<br>Canh giữa dưới Quốc hiệu/Tiêu ngữ. | Cỡ 14, Nghiêng.<br>Canh dưới đường kẻ của Tiêu đề Đảng. | Cỡ 13, Nghiêng.<br>Canh giữa dưới Quốc hiệu/Tiêu ngữ. |
| **5. Tên loại & Trích yếu** | - Tên loại: Cỡ 13-14, In hoa, Đậm.<br>- Trích yếu: Cỡ 13-14, In thường, Đậm.<br>- Kẻ nét liền $1/3 - 1/2$ độ dài trích yếu. | - Tên loại: Cỡ 15-16, In hoa, Đậm.<br>- Trích yếu: Cỡ 14-15, In thường/hoa, Đậm.<br>- Phía dưới có **05 dấu gạch nối (`-----`)**. | - Tên loại: Cỡ 13-14, In hoa, Đậm.<br>- Trích yếu: Cỡ 13-14, In thường, Đậm.<br>- Kẻ nét liền $1/3 - 1/2$ độ dài trích yếu. |
| **5b. Trích yếu Công văn** | Cỡ 12-13, Đứng, sau `V/v ...` (cách 6pt) | Cỡ 12, **Nghiêng**, bắt đầu bằng `về việc ...` | Cỡ 12, Đứng, sau `V/v ...` (cách 6pt) |
| **6. Căn cứ ban hành** | Cỡ 13-14, In thường, Nghiêng. Mỗi căn cứ xuống dòng, cuối dòng `;`, căn cứ cuối `.`. | Mỗi căn cứ 1 dòng riêng, đầu dòng có **dấu gạch ngang (`-`)**, cuối dòng `;`, căn cứ cuối `.`. | Nghị định: In nghiêng; Khác: Đứng thường, kết thúc dòng căn cứ bằng dấu phẩy (`,`). |
| **7. Nội dung văn bản chính** | - Cỡ chữ: 13-14, Đứng, Justify.<br>- Thụt đầu dòng: 1.0 - 1.27 cm.<br>- Giãn dòng: Single đến 1.5 lines. Spacing: $\ge 6\text{pt}$. | - Cỡ chữ: 14-15, Đứng, Justify.<br>- Thụt đầu dòng: ~10 mm.<br>- Giãn dòng: **Exactly 18pt - 22pt**. Spacing: $\ge 6\text{pt}$. | - Cỡ chữ: 13-14, Đứng, Justify.<br>- Thụt đầu dòng: 1.0 - 1.27 cm.<br>- Giãn dòng: **Single line**. Spacing: 6pt. |
| **8. Quyền hạn người ký** | In hoa, cỡ 13-14, Đậm.<br>Cú pháp dấu chấm: `TM.`, `KT.`, `TL.`, `TUQ.`, `Q.` | In hoa, cỡ 14, Đậm.<br>Cú pháp gạch chéo: `T/M`, `K/T`, `T/L`, `Q.` | In hoa, cỡ 13-14, Đậm.<br>Cú pháp: `TM.`, `KT.`, `TL.`, `TUQ.`, `Q.` |
| **9. Chức vụ người ký** | In hoa, cỡ 13-14, **Đứng, Đậm** | In hoa, cỡ 14, **Đứng thường** | In hoa, cỡ 13-14, **Đứng, Đậm** |
| **10. Họ tên người ký** | In thường, cỡ 13-14, Đứng, Đậm. | In thường, cỡ 14, Đứng, Đậm. | In thường, cỡ 13-14, Đứng, Đậm. |
| **11. Kính gửi (Công văn)** | `Kính gửi:` Cỡ 13-14, Đứng thường. | `Kính gửi:` / `Kính trình:` Cỡ 14, **Nghiêng**. | `Kính gửi:` Cỡ 13-14, Đứng thường. |
| **12. Nơi nhận (cuối văn bản)** | - Từ `Nơi nhận:` Cỡ 12, **Nghiêng, Đậm**.<br>- Danh sách: Cỡ 11, Đứng.<br>- Dòng lưu: `- Lưu: VT, [Đơn vị] ([bản]).` | - Từ `Nơi nhận:` Cỡ 14, **Đứng, Gạch chân (`<u>Nơi nhận:</u>`)**.<br>- Danh sách: Cỡ 12, Đứng.<br>- Dòng lưu: `- Lưu [Văn phòng/CQ].` | - Từ `Nơi nhận:` Cỡ 12, **Nghiêng, Đậm**.<br>- Danh sách: Cỡ 11, Đứng.<br>- Dòng lưu: `- Lưu: VT, TCHC.02.` |

---

## 2. BẢNG DANH MỤC 82 QUY TẮC KIỂM TRA (COMPLIANCE CHECKLIST)

### Nhóm 1: Thiết lập trang & Khổ giấy (Page Setup)
- `CHK_PAGE_SIZE`: Bắt buộc khổ A4 ($210 \times 297\text{ mm}$).
- `CHK_PAGE_ORIENTATION`: Hướng giấy Portrait (trừ bảng biểu ngang đặc thù).
- `CHK_PAGE_MARGINS`: Lề trang đúng quy định (Trên 20-25mm, Dưới 20-25mm, Trái 30-35mm, Phải 15-20mm).
- `CHK_PAGE_NUMBERING`: Đánh số trang Header giữa, cỡ 13-14 đứng, không hiện trang 1.
- `CHK_APPENDIX_PAGE_NUM`: Phụ lục đánh số trang riêng từ trang 1.

### Nhóm 2: Kiểu chữ, Đoạn văn & Thụt lề (Typography & Body)
- `CHK_BODY_FONT_NAME`: Phông chữ thống nhất `Times New Roman`.
- `CHK_BODY_FONT_COLOR`: Màu chữ đen (Black / Automatic).
- `CHK_BODY_ALIGNMENT`: Căn đều cả 2 lề (Justify).
- `CHK_BODY_INDENT`: Thụt dòng đầu tiên $1.0\text{ cm} - 1.27\text{ cm}$ ($10\text{ mm}$ với văn bản Đảng).
- `CHK_BODY_SPACE_AFTER`: Giãn đoạn Spacing After tối thiểu $6\text{pt}$.
- `CHK_BODY_LINE_SPACING`: Giãn dòng đúng chuẩn từng Regime (NĐ30: Single - 1.5 lines; Đảng: Exactly 18-22pt; Viettel: Single).
- `CHK_BODY_END_DOT`: Kết thúc nội dung văn bản bắt buộc có dấu chấm câu (`.`).
- `CHK_FONT_SIZE_CONSISTENCY`: Tính nhất quán cỡ chữ theo bộ quy tắc.

### Nhóm 3: 12+ Thành phần Thể thức (Components)
- `CHK_NATIONAL_TITLE` & `CHK_NATIONAL_MOTTO`: Quốc hiệu / Tiêu ngữ / Tiêu đề Đảng.
- `CHK_LINE_SHAPE_LENGTH`: Bắt buộc đối tượng `MsoLineShape` nét liền dưới Tiêu ngữ ($100\%$), Tên cơ quan ($1/3 - 1/2$), Trích yếu ($1/3 - 1/2$); underline, paragraph border và ký tự gạch không được coi là tương đương.
- `CHK_ORGAN_NAMES`: Cơ quan cấp trên (đứng thường) & Cơ quan ban hành (đứng đậm).
- `CHK_CODE_NUMBER_FORMAT` & `CHK_CODE_NUMBER_PAD`: Cú pháp Số/Ký hiệu và chèn số 0 nếu $< 10$.
- `CHK_PLACE_DATE`: Địa danh ngày tháng (chữ nghiêng, ngày $<10$, tháng $<3$ có số 0).
- `CHK_TYPE_NAME_SUBJECT` & `CHK_OFFICIAL_LETTER_SUBJECT`: Tên loại, Trích yếu và Trích yếu công văn (`V/v` / `về việc`).
- `CHK_LEGAL_BASES`: Căn cứ pháp lý (in nghiêng, gạch đầu dòng với Đảng, dấu `;` kết dòng).
- `CHK_SIGN_AUTHORITY`: Viết tắt quyền hạn ký (`TM.`, `KT.`, `TL.`, `TUQ.`, `Q.` vs `T/M`, `K/T`, `T/L`).
- `CHK_SIGN_POSITION` & `CHK_SIGN_FULLNAME`: Chức vụ và Họ tên người ký (Title Case, in đậm, không kèm học hàm học vị trái quy định).
- `CHK_RECIPIENTS_KINH_GUI` & `CHK_RECIPIENTS_NOI_NHAN`: Kính gửi và Nơi nhận cuối văn bản.
- `CHK_RECIPIENT_LUU_LINE`: Dòng lưu văn thư (`- Lưu: VT...`).

### Nhóm 4: Bố cục & Thứ bậc Văn bản (Hierarchy)
- `CHK_PART_CHAPTER`: Phần / Chương (số La Mã, in đậm, căn giữa).
- `CHK_SECTION`: Mục / Tiểu mục (số Ả Rập, in đậm, căn giữa).
- `CHK_ARTICLE`: Điều (thụt đầu dòng, in đậm, số Ả Rập kèm tên điều).
- `CHK_CLAUSE`: Khoản (số Ả Rập, dấu chấm).
- `CHK_POINT` & `CHK_POINT_ORDER`: Điểm (`a)`, `b)`, `c)`...) theo đúng bảng chữ cái tiếng Việt.
- `CHK_LEGAL_CITATION`: Quy cách viện dẫn văn bản quy phạm pháp luật.

### Nhóm 5: Quy tắc Viết hoa (Phụ lục II NĐ 30)
- `CHK_CAP_SENTENCE`: Viết hoa đầu câu và sau các dấu chấm câu `.`, `?`, `!`.
- `CHK_CAP_PERSON_NAME`: Viết hoa tên người Việt Nam và tên nước ngoài (phiên âm Hán-Việt / trực tiếp).
- `CHK_CAP_ADMIN_UNIT`: Viết hoa đơn vị hành chính (`tỉnh Nam Định`, `thành phố Hà Nội`, `Quận 1`, `Phường Điện Biên Phủ`...).
- `CHK_CAP_GEOGRAPHY`: Viết hoa địa danh, sông núi, vùng miền (`Tây Bắc`, `Đông Bắc`, `Bắc Bộ`, `sông Hồng`, `Vịnh Hạ Long`...).
- `CHK_CAP_ORGANIZATION`: Viết hoa cơ quan tổ chức theo từ chỉ loại hình và chức năng.
- `CHK_CAP_SPECIAL_NOUNS`: Viết hoa danh xưng tôn kính và trường hợp đặc biệt (`Nhân dân`, `Nhà nước`, `Đảng`, `Bác`, `Người`).

### Nhóm 6: Chuẩn hóa Tiếng Việt, Chính tả & Bảng biểu (Normalizers & Tables)
- `CHK_TONE_MARK_MIX`: Chuẩn hóa vị trí đặt dấu thanh tiếng Việt (kiểu mới `hòa, thúy` vs kiểu cũ `hoà, thuý`).
- `CHK_IY_MIX`: Chuẩn hóa chính tả `i/y` (kỹ thuật, lý do, thẩm mỹ, xử lý...).
- `CHK_EXTRA_SPACES`: Dọn sạch khoảng trắng kép, khoảng trắng trước dấu câu.
- `CHK_PUNCTUATION_SPACING`: Chuẩn hóa khoảng cách sau dấu phẩy, chấm, hai chấm, chấm phẩy.
- `CHK_DASH_NORMALIZATION`: Chuẩn hóa gạch nối (`-`), en-dash (`–`), em-dash (`—`).
- `CHK_DECIMAL_SEPARATOR`: Chuẩn hóa dấu phẩy số thập phân (`,`) và dấu chấm hàng nghìn (`.`).
- `CHK_SOFT_LINEBREAKS`: Khử ngắt dòng mềm `Shift+Enter` làm rách dòng.
- `CHK_TRAILING_EMPTY_PAGES`: Xóa trang trắng thừa ở cuối tài liệu.
- `CHK_TABLE_HEADER_REPEAT`: **Tự động lặp lại hàng tiêu đề của bảng khi tràn trang (`RepeatHeaderRow = True`)**.
- `CHK_TABLE_ROW_SPLIT`: **Ngăn bảng gãy hàng ngang dở dang giữa 2 trang (`CantSplit = True`)**.
- `CHK_BLANK_FIELD_SPACER`: Chuẩn hóa dòng kẻ chấm điền thông tin (`......`) bằng Tab Stop Dot Leader chuyên nghiệp thay vì gõ chấm thủ công.

---

## 3. DANH MỤC FILE CẤU HÌNH & TỪ ĐIỂN ĐÃ TẠO

Toàn bộ các quy tắc và từ điển trên đã được trích xuất hoàn tất vào hệ sinh thái:
- **Tập luật thể thức**:
  - [rules_nd30.json](file:///d:/chuan-hoa-the-thuc-workspace/shared/rules/rules_nd30.json) (Nghị định 30/2020/NĐ-CP)
  - [rules_party_hd05.json](file:///d:/chuan-hoa-the-thuc-workspace/shared/rules/rules_party_hd05.json) (Hướng dẫn 05-HD/VPTW Đảng)
  - [rules_viettel.json](file:///d:/chuan-hoa-the-thuc-workspace/shared/rules/rules_viettel.json) (Quy chế Viettel QĐ 11095)
  - [rules_compliance_checks.json](file:///d:/chuan-hoa-the-thuc-workspace/shared/rules/rules_compliance_checks.json) (82 mã kiểm tra & phân loại)
- **Từ điển tra cứu**:
  - [administrative_units.json](file:///d:/chuan-hoa-the-thuc-workspace/shared/dictionaries/administrative_units.json) (1.520 đơn vị hành chính)
  - [typo_dictionary.json](file:///d:/chuan-hoa-the-thuc-workspace/shared/dictionaries/typo_dictionary.json) (Từ điển sửa lỗi chính tả hành chính)
  - [iy_dictionary.json](file:///d:/chuan-hoa-the-thuc-workspace/shared/dictionaries/iy_dictionary.json) (Từ điển chuẩn hóa i/y)
  - [non_sentence_ending_abbreviations.json](file:///d:/chuan-hoa-the-thuc-workspace/shared/dictionaries/non_sentence_ending_abbreviations.json) (55 từ viết tắt không kết thúc câu)
  - [doctype_abbreviations.json](file:///d:/chuan-hoa-the-thuc-workspace/shared/dictionaries/doctype_abbreviations.json) (Ký hiệu viết tắt tên loại văn bản)
  - [special_capitalizations.json](file:///d:/chuan-hoa-the-thuc-workspace/shared/dictionaries/special_capitalizations.json) (Danh từ viết hoa đặc biệt)
