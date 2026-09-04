# VIETNAMESE SPELL TEST REPORT: Báo cáo Kiểm thử và Đánh giá Hồi quy

Tài liệu này tổng hợp toàn bộ kết quả kiểm thử đơn vị, kiểm thử tích hợp đầu cuối (End-to-End), và kiểm thử hồi quy (Regression Tests) theo yêu cầu của `NHIỆM VỤ` và `TASK`.

---

## 1. Kết quả Tổng thể Bộ Kiểm thử Tự động (Automated Test Suites)

```bash
dotnet test
```

| Dự án Test | Số lượng Test | Đỗ (Passed) | Hỏng (Failed) | Bỏ qua (Skipped) | Thời gian chạy |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **ChuanHoa.Client.Core.Tests** | 155 | 155 | 0 | 0 | 0.33 s |
| **ChuanHoa.Contracts.Tests** | 4 | 4 | 0 | 0 | 0.10 s |
| **ChuanHoa.Domain.Tests** | 14 | 14 | 0 | 0 | 0.05 s |
| **ChuanHoa.Rules.Tests** | 10 | 10 | 0 | 0 | 0.12 s |
| **ChuanHoa.Application.Tests** | 6 | 6 | 0 | 0 | 0.04 s |
| **ChuanHoa.Api.Tests** | 11 | 11 | 0 | 0 | 0.40 s |
| **TỔNG CỘNG** | **200** | **200** | **0** | **0** | **~1.1 s** |

---

## 2. Kết quả Kiểm thử Hồi quy Nghiệp vụ (Regression Acceptance Tests)

| Kịch bản kiểm thử | Dữ liệu đầu vào | Kết quả mong đợi | Kết quả thực tế | Trạng thái |
| :--- | :--- | :--- | :--- | :--- |
| **Tự động gộp 2 dấu cách** | `"Cộng  hòa  xã  hội"` | Gộp về `"Cộng hòa xã hội"` | Gộp chính xác 1 dấu cách | **PASS** |
| **Xóa khoảng trắng trước dấu câu** | `"Hà Nội , ngày 02 tháng 09"` | Xóa khoảng trắng trước dấu phẩy | `"Hà Nội, ngày 02 tháng 09"` | **PASS** |
| **Lỗi gõ Telex âm tiết** | `"Tôi muốn ngỉ việc."` | Gợi ý `"nghỉ"` | Gợi ý đúng `"nghỉ"` | **PASS** |
| **Lỗi ngữ cảnh (Real-word)** | `"Đồng chí gửi bàn dao hồ sơ"` | Phát hiện Level 3, gợi ý `"bàn giao"` | Gợi ý đúng `"bàn giao"` | **PASS** |
| **Lỗi ngữ cảnh hành chính** | `"sử lý công việc"` | Phát hiện Level 3, gợi ý `"xử lý"` | Gợi ý đúng `"xử lý"` | **PASS** |
| **Lỗi điều khoản hợp đồng** | `"điều khoảng hợp đồng"` | Phát hiện Level 3, gợi ý `"điều khoản"` | Gợi ý đúng `"điều khoản"` | **PASS** |
| **Kiểm thử âm tính (Negative - Tên riêng)** | `"Viettel, BIDV, OpenAI, Windows"` | Không báo lỗi, không gạch chân | Bỏ qua chính xác 100% | **PASS** |
| **Kiểm thử âm tính (URL / Email)** | `"contact@chuanhoa.gov.vn"` | Không báo lỗi | Bỏ qua chính xác 100% | **PASS** |
| **Khắc phục lỗi Range Shape/Field** | Tài liệu chứa Page Number, Shape | Gắn comment an toàn, không ném exception | Không sập Word, comment đầy đủ | **PASS** |
| **Kiểm thử Ngoại tuyến (Offline Check)** | Ngắt toàn bộ kết nối Internet | Toàn bộ chức năng hoạt động bình thường | Hoạt động bình thường 100% | **PASS** |

---

## 3. Thẩm định Mã nguồn VSTO (Invariants & Contract Verification)

```bash
python tools/vsto/validate_vsto_source.py
```
- **Trạng thái**: **`PASS`**
- **Invariants**: 24/24 điều kiện an toàn kiến trúc đều đạt (không dùng persistent taskpane, không static Word Document, không gọi FinalReleaseComObject, không freeze UI thread).
