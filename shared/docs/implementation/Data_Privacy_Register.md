# Data Privacy Register

| Data class | Nguồn | Mục đích | Nơi xử lý | Retention | Cấm |
| --- | --- | --- | --- | --- | --- |
| Account identity | IdP/user | Login, support, entitlement | Identity Service | Theo policy production chưa chốt | Không log token/magic link |
| Device public key/assurance | VSTO/device | Bind credential, revoke, risk | Device Service | Theo security policy | Không thu private key/raw HWID làm credential |
| Trial/entitlement | Server events | Quyết định quyền | Entitlement Service/PostgreSQL | Append-only history | Không xóa để cấp lại trial |
| Offer/quote/order/payment | Commercial/provider | Mua hàng và đối soát | Commercial Service/PostgreSQL | Theo legal/accounting policy | Không lưu card data; không sửa lịch sử |
| Document snapshot | Word local | Detector, scan và exact-anchor | Chỉ bộ nhớ tiến trình Word local | Hết khi kết thúc scan/context | Cấm upload body, full DOC/DOCX, path, filename, comments, ảnh, chữ ký, con dấu |
| Findings/annotation metadata | Local scanner/client | Hiển thị comment/tô đỏ | Local document và local context | Theo tài liệu/cache local | Không gửi raw finding text lên server mặc định |
| Signed rule pack | Rule publisher | Kiểm tra local | Server phát hành, local cache | Đến expiry có chữ ký | Không chứa dữ liệu người dùng/tài liệu |
| Client telemetry | VSTO | Reliability/update | Telemetry pipeline | Theo privacy policy | Không document text/path/name/comments/QR payload |
| Admin audit | Admin BFF/services | Security/compliance | Append-only/WORM target | Theo audit policy | Không hard-delete hoặc chứa secret |
| Support request | Customer Portal | Hỗ trợ | Support Service | Theo support policy | Không tự đính kèm tài liệu |

Luồng sản phẩm hiện hành không có cloud snapshot/compliance worker. Mọi endpoint mới có khả năng nhận text hoặc file Word phải qua privacy review và ADR mới; không được suy ra quyền upload từ endpoint identity/lease/rule/admin.

`python tools/validation/validate_document_privacy.py --self-test` là regression gate bắt buộc cho VSTO: network API chỉ được tồn tại trong allowlist đã review; endpoint, request field hoặc telemetry sink mới đều fail closed. Gate hiện tại xác nhận bootstrap Development chỉ gửi `deviceThumbprint` và `clientReleaseId`; đây là source gate, không thay thế production traffic observability hoặc pentest.
