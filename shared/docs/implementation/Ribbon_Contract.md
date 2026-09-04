# Ribbon Contract

## Baseline VBA thực tế (giữ làm provenance)

| Hạng mục | Số lượng |
| --- | ---: |
| Tab | 1 |
| Group | 7 |
| Button | 36 |
| Menu | 4 |
| DropDown | 2 |
| CheckBox | 3 |
| Control tương tác | 45 |

Nguồn layout VBA cũ là `shared/ChuanHoaTheThuc_Full_Ribbon.dotm`. Bản trích xuất có hash và toàn bộ thuộc tính nằm tại `evidence/ribbon_actual.json`; evidence này không bị sửa khi sản phẩm thay đổi.

## Target VSTO đã được chủ sản phẩm phê duyệt

| Hạng mục | Số lượng |
| --- | ---: |
| Tab | 1 |
| Group | 7 |
| Button | 34 |
| Menu | 3 |
| DropDown | 2 |
| CheckBox | 3 |
| Control tương tác | 39 |

- Loại bỏ `btnLuuDocx`; add-in không còn chức năng hoặc callback chuyển/lưu thành DOCX.
- Thực hiện trực tiếp trên tài liệu đã lưu `.doc` hoặc `.docx`, kể cả Compatibility Mode; giữ nguyên định dạng hiện tại khi lưu.
- `.docm`, template, RTF, tài liệu chưa lưu và định dạng khác fail closed cho xử lý tài liệu.
- Loại bỏ menu “i/y” và hai lựa chọn i/y; giữ menu đặt dấu với hai lệnh chủ động đồng nhất oà/uý hoặc òa/úy.
- Loại bỏ hoàn toàn `btnChenQrCode`, callback, dialog, renderer và dependency QR theo quyết định sản phẩm mới nhất; chỉ giữ VBA extracted và evidence baseline làm provenance.
- Nhóm `Hiển thị` và ba checkbox đổi tùy chọn View đã được loại khỏi sản phẩm; người dùng dùng trực tiếp Word Options khi cần.
- `btnTuDienCaNhan` mở không cần tài liệu và dùng PNG custom nhúng trong assembly qua `getImage`.
- Tab sản phẩm hiển thị đúng tên “Chuẩn hóa”, không kèm “thể thức” hoặc năm.

## Correction hành vi bắt buộc

- `btnAutoFixAll2026` hiện trỏ `OnDinhDangTrangGiay`; contract đích phải trỏ handler AutoFix thật.
- Scan thể thức/chính tả phải read-only; không giữ bảy normalizer chạy ngầm của VBA.
- Scanner tự nhận diện loại văn bản từ nội dung; dropdown chỉ phản ánh/lưu context và không được ghi đè kết quả nhận diện rõ ràng.
- Ba checkBox đọc/ghi state theo active window.
- Ba menu là container, không tính là command handler.
- Có 34 command button trong target; control chưa đạt gate phải disable với reason, không callback giả.
- Không có task pane thường trực hoặc tab Ribbon thứ hai.

## Visual gate

Baseline target cũ đã load trên Word 16.0 x64. Sau thay đổi target 39-control, loại QR/nhóm Hiển thị và dùng icon custom, visual smoke mới phải chạy lại trước khi ghi PASS cho release mới.
