# Audit quy tắc thể thức NĐ30 và HD05

Ngày audit: 2026-09-02  
Phạm vi: ba tài liệu do chủ sản phẩm cung cấp, mã scanner local C#, snapshot Word COM và corpus test hiện tại.

## 1. Nguồn đã đọc và kiểm tra

| Nguồn | Trang | Đoạn | Bảng | Shape | Inline shape |
| --- | ---: | ---: | ---: | ---: | ---: |
| `Hướng dẫn 05.docx` | 21 | 842 | 9 | 0 | 3 |
| `Nghị định 30.doc` | 11 | 310 | 2 | 0 | 0 |
| `Phụ lục Nghị định 30.doc` | 62 | 2.778 | 57 | 108 | 2 |

Ba file nguồn được mở chỉ đọc. Audit không Save/SaveAs và không thay đổi timestamp hoặc nội dung nguồn. Toàn bộ 94 trang đã được Word xuất PDF tạm, raster hóa và kiểm tra tổng quan; các trang chứa quy định và mẫu thành phần thể thức được kiểm tra ở độ phân giải đầy đủ.

## 2. Kết luận bắt buộc về đường kẻ

### NĐ30

Phụ lục I quy định:

- Dưới Tiêu ngữ: đường kẻ ngang nét liền, dài bằng dòng Tiêu ngữ.
- Dưới tên cơ quan/tổ chức ban hành: đường kẻ ngang nét liền, dài từ 1/3 đến 1/2 dòng chữ và cân giữa.
- Dưới trích yếu của văn bản có tên loại: đường kẻ ngang nét liền, dài từ 1/3 đến 1/2 dòng chữ và cân giữa.

Định nghĩa triển khai đã khóa: chỉ đối tượng Word `msoLine`/Line Shape mới hợp lệ. Underline của Font, paragraph bottom border và chuỗi ký tự `-----` không được coi là tương đương.

### HD05

- Dưới tiêu đề `ĐẢNG CỘNG SẢN VIỆT NAM`: Line Shape nét liền, dài bằng tiêu đề.
- Dưới tên cơ quan ban hành: dấu sao `*`, không phải Line Shape.
- Dưới trích yếu: đúng 05 dấu gạch nối `-----`, không phải Line Shape.

## 3. Đối chiếu triển khai Line Shape

| Yêu cầu | Detector | Dữ liệu snapshot | Trạng thái |
| --- | --- | --- | --- |
| Line dưới Tiêu ngữ NĐ30 | `ND30-PL1-M2-K1-TN-LINE` | `LocalLineShapeSnapshot` | IMPLEMENTED_LOCAL |
| Line dưới cơ quan ban hành NĐ30 | `ND30-PL1-M2-K2-ORG-LINE` | `LocalLineShapeSnapshot` | IMPLEMENTED_LOCAL |
| Line dưới trích yếu NĐ30 | `ND30-PL1-M2-K5A-SUBJ-LINE` | `LocalLineShapeSnapshot` | IMPLEMENTED_LOCAL |
| Line dưới tiêu đề Đảng HD05 | `HD05-M1-TITLE-LINE` | `LocalLineShapeSnapshot` | IMPLEMENTED_LOCAL |

Snapshot giữ loại Shape, tên/ID, story, section, trang, paragraph neo, vị trí tương đối, tọa độ trang, chiều rộng/cao, độ dày, màu, dash style và hai đầu mũi tên. Shape tham gia fingerprint tài liệu.

Detector chỉ chấp nhận khi đồng thời đạt:

1. `Shape.Type == msoLine` và line đang hiển thị.
2. Nét liền, không có đầu mũi tên và gần ngang.
3. Cùng story, section, trang và neo gần đúng paragraph thành phần.
4. Nằm bên dưới paragraph thành phần.
5. Chiều dài đúng tỷ lệ; tâm line cân với tâm dòng chữ khi Word cung cấp được tọa độ.

Các khoảng tỷ lệ được phát hành trong rule pack RS256 (`mottoLine*`, `organLine*`, `subjectLine*`, `partyTitleLine*`), không còn cố định trong detector. Parser từ chối fail-closed nếu tỷ lệ ngoài biên an toàn hoặc `min > max`.

## 4. Ma trận quy tắc format chính

| Nhóm | NĐ30 | HD05 | Trạng thái C# hiện tại |
| --- | --- | --- | --- |
| Khổ giấy, hướng, lề | A4; dọc; lề trong khoảng quy định | A4; dọc; lề 20/20/30/15 mm, có quy tắc mặt sau | NĐ30 implemented; HD05 mặt sau chưa có metadata/golden evidence |
| Font/màu | Times New Roman, Unicode, đen | Times New Roman, Unicode, đen | Implemented nền; NFC toàn tài liệu chưa có detector hoàn chỉnh |
| Số trang | Ả Rập, giữa phía trên, không hiện trang đầu | Ả Rập, giữa; không hiện trang đầu; phụ lục đánh riêng | Có presence/restart/alignment; chưa đủ font/vị trí/ẩn trang đầu cho mọi story |
| Quốc hiệu/Tiêu ngữ | Cỡ 12–13/13–14, đậm, căn giữa | Không áp dụng | Style implemented; Line Shape implemented |
| Tiêu đề Đảng | Không áp dụng | Cỡ 15, hoa, đứng đậm, căn giữa | Style nền có nhận diện; Line Shape implemented; cần golden HD05 |
| Tên cơ quan | 12–13, cấp trên thường, cơ quan ban hành đậm; có line | 14, cấp trên thường, cơ quan ban hành đậm; có dấu `*` | NĐ30 style + line implemented; HD05 separator/style chưa đạt golden gate |
| Số/ký hiệu | `Số: 05/QĐ-...` | `Số 05-QĐ/...` | NĐ30 implemented; HD05 cú pháp riêng chưa có corpus approved |
| Địa danh/ngày | 13–14 nghiêng | 14 nghiêng | NĐ30 implemented; HD05 cần regime-specific golden |
| Tên loại/trích yếu | 13–14 đậm; trích yếu có line | Tên loại 15–16; trích yếu 14–15; có `-----` | NĐ30 style + line implemented; HD05 separator/style chưa đạt golden gate |
| Trích yếu công văn | `V/v`, 12–13 đứng | `về việc`, cỡ 12 nghiêng | NĐ30 implemented; HD05 riêng chưa đạt golden gate |
| Nội dung | 13–14, justify, thụt 10–12,7 mm, dòng 1–1,5 | 14–15, justify, thụt 10 mm, Exactly 18–22 pt | NĐ30 implemented; HD05 riêng chưa đạt golden gate |
| Căn cứ/cấu trúc | Căn cứ nghiêng; Điều/Khoản/Điểm | Căn cứ đầu `-`; cấu trúc cỡ 14–15 | NĐ30 implemented; HD05 riêng chưa đạt golden gate |
| Ký/Nơi nhận | `TM.`, `KT.`; Nơi nhận nghiêng đậm 11–12 | `T/M`, `K/T`; Nơi nhận cỡ 14 đứng gạch chân | NĐ30 implemented; HD05 riêng chưa đạt golden gate |
| Phụ lục | Số trang liên tục với văn bản chính | Số trang riêng từng phụ lục | NĐ30 implemented; HD05 riêng chưa đạt golden gate |

## 5. Evidence đã chạy

- 103/103 unit/API/domain/rule tests không phụ thuộc cơ sở dữ liệu: PASS.
- Client Core: 58/58 PASS.
- Test riêng xác nhận underline + paragraph border nhưng thiếu Shape vẫn phát lỗi: PASS.
- Test line nét đứt, sai trang, sai độ dài: PASS.
- Test line nét liền đúng neo/vị trí/tỷ lệ: PASS.
- Test gói quy tắc ký chứa tỷ lệ Line Shape và từ chối range min/max không hợp lệ: PASS.
- VSTO Development build bằng Visual Studio MSBuild: 0 warning, 0 error.
- Word COM smoke tạo và đọc Line Shape trên `.doc` và `.docx`: PASS.
- PostgreSQL integration test không được tính trong kết quả trên; lần chạy trực tiếp hiện tại dừng vì thiếu `CHUANHOA_TEST_CONNECTION_STRING`, đúng fail-closed của test harness.

## 6. Kết luận gate

Phần Line Shape đã được sửa đúng theo xác nhận của chủ sản phẩm và nguồn NĐ30/HD05. Không còn dùng underline/border để kết luận ba đường kẻ NĐ30 hoặc đường dưới tiêu đề Đảng là hợp lệ.

Không được tuyên bố toàn bộ format NĐ30 + HD05 đã đạt production/legal gate. Các detector NĐ30 hiện có source/synthetic evidence; các khác biệt regime HD05 liệt kê ở Mục 4 vẫn cần port riêng, golden DOC/DOCX được reviewer nghiệp vụ duyệt và chạy Word 2010/x86. Production signing/KMS/IdP cũng là gate độc lập.
