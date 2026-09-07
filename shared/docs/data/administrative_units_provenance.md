# Nguồn dữ liệu đơn vị hành chính Việt Nam

## Phạm vi hiện hành

`administrative_units.json` là snapshot danh mục có hiệu lực từ ngày 01/07/2025:

- 34 đơn vị hành chính cấp tỉnh;
- 3.321 đơn vị hành chính cấp xã;
- không có tầng cấp huyện đang hoạt động;
- cấp xã gồm xã, phường và đặc khu.

Nguồn pháp lý chính là Quyết định 19/2025/QĐ-TTg ngày 30/06/2025, có hiệu lực
01/07/2025, ban hành Bảng danh mục và mã số các đơn vị hành chính Việt Nam.
Danh mục cấp tỉnh được đối chiếu với Nghị quyết 202/2025/QH15. Các đơn vị cấp xã
được hình thành theo 34 nghị quyết của Ủy ban Thường vụ Quốc hội năm 2025 và mã
số được tổng hợp chính thức tại Quyết định 19/2025/QĐ-TTg.

Trang Công báo chính thức:

https://congbao.chinhphu.vn/van-ban/quyet-dinh-so-19-2025-qd-ttg-45430/57438.htm

## Pipeline và tính tái lập

Snapshot ngày 07/09/2026 được trích từ hai tệp DOC chính thức đính kèm Công báo
số 919+920 và 921+922. Script
`tools/data/generate_administrative_reference_data.py` chuẩn hóa Unicode NFC,
kiểm tra mã và quan hệ cha, rồi sinh các output theo thứ tự ổn định. Checksum của
snapshot nguồn và output nằm trong `shared/dictionaries/reference_data_manifest.json`.

`tools/validation/validate_reference_data.py` và source-quality gate sẽ thất bại
nếu sai count, trùng mã, sai quan hệ cha, có active district, thiếu metadata,
không phải NFC, có fragment/rác hoặc checksum không khớp.

## Dữ liệu lịch sử

Tên tỉnh trước sắp xếp được lưu riêng trong
`historical_administrative_aliases.json`, có `status = Historical`, ngày kết thúc
và đơn vị kế nhiệm. Dữ liệu này chỉ phục vụ nhận diện/search và tránh false
positive trong văn bản cũ; không bao giờ được trả về như đơn vị hiện hành.

Danh mục huyện/quận/thị xã/thị trấn lịch sử chưa được bulk-import vì chưa có
pipeline nguồn pháp lý đã review ở độ tin cậy tương đương. Không được suy ra rằng
từ “huyện” trong văn bản trước 01/07/2025 là lỗi.

## Cập nhật

Khi có thay đổi địa giới:

1. tải văn bản và phụ lục chính thức mới;
2. lưu provenance, ngày hiệu lực và checksum;
3. chạy generator để tạo snapshot mới;
4. chuyển tên hết hiệu lực sang lớp historical thay vì xóa;
5. chạy full tests và source-quality gates trước khi phát hành.
