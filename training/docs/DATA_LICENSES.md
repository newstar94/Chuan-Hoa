# Thẩm định Bản quyền Dữ liệu & Giấy phép (Data Licenses Audit)

Tài liệu này thẩm định tính pháp lý và bản quyền của các tập dữ liệu, mô hình nền tảng phục vụ huấn luyện mô hình chính tả theo yêu cầu của `TASK_ Research, Train, Distill, Quantize and Export a Tiny Vietnamese ONNX Context Model.md`.

---

## 1. Danh mục Thẩm định Pháp lý

| Thành phần | Nguồn gốc | Giấy phép (License) | Cho phép sử dụng thương mại? | Yêu cầu ghi nhận bản quyền (Attribution) | Yêu cầu phân phối lại (Redistribution) | Trạng thái thẩm định |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **VSEC Dataset** | Nghiên cứu học thuật lỗi chính tả tiếng Việt | CC BY-NC 4.0 / Academic | **Chỉ dùng nội bộ nghiên cứu / Validation** | Bắt buộc trích dẫn tác giả | Không phân phối thương mại trực tiếp | **APPROVED_RESEARCH_ONLY** |
| **Viwiki-Spelling** | Wikipedia tiếng Việt chỉnh sửa lỗi chính tả | CC BY-SA 3.0 / 4.0 | Cho phép (Kèm điều kiện dẫn nguồn) | Bắt buộc ghi nhận nguồn Wikipedia | Phải chia sẻ tương tự nếu phát hành dataset | **APPROVED_DERIVATIVE** |
| **Bộ Quy tắc Hành chính (VBQPPL / NĐ 30 / CSDL Quốc gia)** | Cổng thông tin Chính phủ / Thư viện Pháp luật | Văn bản quy phạm pháp luật Nhà nước | **HOÀN TOÀN TỰ DO** (Văn bản pháp quy không thuộc quyền tác giả) | Khuyến khích ghi rõ số hiệu | Tự do phân phối | **FULLY_APPROVED_COMMERCIAL** |
| **PhoBERT Pretrained Weights (Teacher)** | VinAI Research | MIT License | **CÓ (Cho phép thương mại)** | Giữ nguyên thông báo bản quyền MIT | Tự do | **APPROVED_TEACHER_ONLY** (Chỉ dùng lúc train, không đóng gói runtime) |
| **Bộ Sinh lỗi Nhân tạo (Synthetic Generator)** | Tự phát triển nội bộ trong dự án | Độc quyền của sản phẩm | **HOÀN TOÀN TỰ DO** | Thuộc bản quyền dự án Chuẩn Hóa | Toàn quyền kiểm soát | **FULLY_APPROVED_PROPRIETARY** |
| **Tệp Mô hình Đầu ra (Tiny-B ONNX INT8)** | Huấn luyện từ đầu qua chưng cất | Thuộc sở hữu dự án Chuẩn Hóa | **HOÀN TOÀN TỰ DO** | Thuộc sản phẩm Chuẩn Hóa | Đóng gói kèm Add-in | **FULLY_APPROVED_PRODUCTION** |

---

## 2. Kết luận Tuân thủ Bản quyền
1. **Không phân phối dữ liệu thô bị giới hạn**: Không nhúng bất kỳ tệp dữ liệu nào có điều khoản NC (Non-Commercial) vào bộ cài đặt của khách hàng.
2. **Teacher Model**: Chỉ sử dụng `PhoBERT-base-v2` trên máy huấn luyện (Training Environment) để tạo nhãn xác suất (Soft Labels), hoàn toàn không nhúng PhoBERT vào gói cài đặt `.exe` của Add-in Word.
3. **Mô hình Xuất xưởng**: Mô hình sinh viên `Tiny-B` là mô hình hoàn toàn mới do dự án tự sinh, tự lượng tử hóa và giữ toàn quyền sở hữu thương mại.
