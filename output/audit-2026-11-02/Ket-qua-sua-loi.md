# Kết quả sửa lỗi — 02/11/2026

## Đã triển khai trong mã nguồn

- Thay vòng quét toàn bộ 6.469 quy tắc trên từng đoạn bằng chỉ mục ứng viên theo từ đầu, vẫn dùng phép so sánh tiếng Việt và ánh xạ offset gốc để kết luận lỗi. Đoạn chứa ký tự đặc biệt có thể ảnh hưởng collation dùng đường đối chiếu đầy đủ. Không bỏ quy tắc, không giới hạn số lỗi/comment.
- Bổ sung kiểm tra hủy trong bước viết hoa và điểm thăm dò Esc từ luồng Word khi quét chính tả. Không đưa Word COM sang luồng nền, không dùng DoEvents; khóa hủy sau khi bắt đầu sửa tài liệu vẫn được giữ.
- Bộ nhận dạng và bộ tách Quốc hiệu–Tiêu ngữ dùng chung chuẩn so sánh, nhưng trả vị trí xuống dòng theo chuỗi UTF-16 gốc: xử lý Unicode NFD, dấu gạch U+2011, ký tự ẩn và thiếu dấu gạch.
- Nhận dạng Tiêu ngữ trong bảng kiểm tra đúng quan hệ cột/hàng, không còn chấp nhận ngay chỉ vì hai đoạn liên tiếp; không còn giới hạn tìm lùi sáu đoạn trong cùng bảng.
- Đọc lại viền dưới và gạch chân của thành phần sở hữu đường kẻ, kể cả sau đoạn 80/150. Chốt tài liệu nặng áp dụng cho cả đo thành phần và đọc tọa độ shape trong lần quét thông thường.
- Đo hình học cho lệnh sửa đường kẻ là một yêu cầu riêng, chỉ bật tại pha đường kẻ sau phân trang và khi hậu kiểm các đường vừa sửa. Nhánh này vẫn phải làm việc với layout Word; chưa có bằng chứng thời gian thực trên file lớn.
- Pha chuẩn bị sửa đường kẻ chỉ quét thể thức, không chạy lại chính tả không cần thiết.
- Mở rộng nhận diện đường cũ neo đúng đoạn thành phần, cùng trang/section và cùng tâm, không mũi tên; tránh để lại đường cũ chỉ vì lệch hơn 18 pt theo chiều dọc. Giới hạn danh sách legacy chuẩn bị trước theo story/section/khối, tránh xóa sang khối khác cùng mã quy tắc.
- Hậu kiểm không còn bỏ qua đường sai thuộc thành phần chỉ vì đã có một đường đúng khác. Có kiểm thử âm cho đường nét đứt gần đó nhưng neo ở đoạn khác.
- Nâng validator nguồn để kiểm tra các đường truyền điều kiện an toàn layout mới, không chỉ sự tồn tại của hàm chốt.

## Bằng chứng hiện tại

| Hạng mục | Kết quả |
| --- | --- |
| Client Core Release | **546/546 đạt**, tăng 11 ca so với trước sửa; không skip |
| Kiểm thử đối chiếu viết hoa | Kết quả/offset của chỉ mục mới trùng phép quét đầy đủ trên corpus tên riêng, Unicode NFD, dấu câu, ký tự ẩn/điều khiển |
| VSTO source validator | **PASS** |
| VSTO `/t:Compile`, Release | **Exit 0**, DLL trung gian được sinh; không phải build/publish bộ cài có ký |
| Probe .NET Framework 4.8 x64 liên kết runtime hiện tại | **Build 0 lỗi, 0 cảnh báo** |
| Thao tác thực tế trong Word | **Bị chặn** ngay lúc lấy giao diện tự động hóa Office: `0x80029C4A / TYPE_E_CANTLOADLIBRARY` |
| Bộ cài đang dùng | **Chưa thay đổi**, chưa tăng phiên bản, chưa cài đặt |

### Đo kiểm tra chính tả ngoài Word

| Số đoạn của cùng fixture | Trước sửa | Sau sửa |
| --- | ---: | ---: |
| 500 | 8.202 ms | 485 ms |
| 2.000 | 34.128 ms | 1.649 ms |
| 8.000 | 121.101 ms | 6.945 ms |

Riêng đối chiếu viết hoa trên 1.000 đoạn: **12.013 ms → 33 ms**. Tổng kiểm tra chính tả trong phép đo tách pha: **13.524 ms → 883 ms**.

Đây là phép đo đơn trên fixture văn bản lặp, không phải p95, không gồm đọc Word/phân trang/ghi comment, không suy ra rằng mọi file lớn đã hết treo.

Probe hủy sau sửa phát yêu cầu ở **107 ms**, scanner trả về ở **115 ms**. Ngưỡng yêu cầu trong probe sau sửa được hạ từ 2.000 xuống 100 ms vì fixture 1.000 đoạn đã hoàn thành trước 2 giây. Kiểm thử hồi quy riêng còn kiểm tra hủy ngay trong vòng quy tắc và callback chạy đúng luồng gọi.

Ca đường kẻ: trước sửa, thêm một đường đúng làm số lỗi đường kẻ giảm từ 1 xuống 0 dù nét đứt cũ còn nguyên; sau sửa số lỗi vẫn là **1**, không báo sạch sai.

## Giới hạn còn lại — chưa coi là hoàn tất kiểm chứng sản phẩm

1. Chưa chạy được sửa đường kẻ trực tiếp trong Word, lưu/đóng/mở lại, hoặc thử file thật người dùng báo treo. Một phiên Word rỗng do phép thử tạo ra đã được đóng; không đụng tài liệu người dùng, không sửa registry/Office.
2. Quốc hiệu–Tiêu ngữ trong Header thật/Text Box vẫn chưa được hỗ trợ bởi detector vai trò. Không mở rộng tự động sang các story này vì toàn bộ đường sửa hiện dựa vào main story; chỉ nhận dạng thêm mà chưa sửa cơ chế neo có thể sửa nhầm nội dung. Đây là phần chưa sửa, không phải đã đạt.
3. Các thao tác Word COM/phân trang/comment vẫn đồng bộ. Đã giảm điểm nghẽn dữ liệu và cải thiện hủy trước mutation, nhưng chưa chứng minh mọi loại file nhiều ảnh/bảng/comment đều đáp ứng tốt. Không cắt bớt comment tại lỗi để làm nhanh.
4. Toàn solution có lỗi API lock file và Domain thiếu contract được ghi trong báo cáo rà soát trước đó; đợt này không sửa dependency/backend không liên quan. Không tuyên bố toàn solution đạt.
5. Chưa tạo bộ cài và chưa cập nhật bản đang cài. Cần khôi phục khả năng kiểm thử Office hoặc kiểm thử trực tiếp trên bản sao tài liệu trước khi kết luận đường kẻ ổn định và phát hành.

## Bằng chứng lưu kèm

- [Đo tốc độ và nhận dạng sau sửa](<D:/Chuẩn Hóa/output/audit-2026-11-02/fixed-core-results.txt>)
- [Đo riêng bước viết hoa](<D:/Chuẩn Hóa/output/audit-2026-11-02/fixed-profile-results.txt>)
- [Hậu kiểm đường kẻ và hủy](<D:/Chuẩn Hóa/output/audit-2026-11-02/fixed-edge-results.txt>)
- [Lỗi tự động hóa Word sau sửa](<D:/Chuẩn Hóa/output/audit-2026-11-02/fixed-word-results.txt>)
- [Báo cáo nguyên nhân trước sửa](<D:/Chuẩn Hóa/output/audit-2026-11-02/Bao-cao-ra-soat.md>)

Không commit, push, merge, sửa tài liệu nguồn hoặc thay bản cài trong đợt này.
