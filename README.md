# Chuẩn hóa

Sản phẩm là add-in VSTO/C# cho Microsoft Word trên Windows, hỗ trợ mục tiêu Word 2010 trở lên và xử lý trực tiếp tài liệu `.doc`/`.docx`. Trong Word chỉ có một tab Ribbon tên **Chuẩn hóa**; không dùng task pane thường trực và không phát triển Office Web Add-in như client sản phẩm.

Word được đọc, kiểm tra và chỉnh sửa tại máy người dùng; nội dung, đường dẫn và binary tài liệu không được gửi lên máy chủ. Máy chủ chỉ giữ danh tính, trial/entitlement, giá, thanh toán, release, audit và phát hành **gói quy tắc có chữ ký + giấy phép offline tối đa 7 ngày**. Client chỉ dùng artefact còn hạn, đúng thiết bị, đúng phiên bản và đúng feature.

## Cấu trúc dự án

```text
D:\Chuẩn Hóa
├── src/                             # VSTO, API, application, domain, infrastructure, contracts và rules
├── tests/                           # Unit, contract, API và integration tests
├── database/                        # PostgreSQL migrations và verification
├── shared/                          # Rules, dictionaries, contract và tài liệu triển khai
├── tools/                           # Inventory, build/publish và Word smoke tools
├── backend-api/                     # Prototype TypeScript lưu làm tài liệu đối chiếu, không phải production server
├── client-vsto-csharp/              # Prototype VSTO cũ lưu làm provenance
└── client-web-addin/                # Artefact tham khảo đã retire khỏi client product
```

## Nguồn sự thật

- Kế hoạch triển khai: `shared/docs/KeHoach_TrienKhai_Addin_VSTO_Word_2010_Plus.md`.
- Prompt thực thi: `shared/docs/Prompt_ThucThi_KeHoach_Addin_VSTO_Word_2010_Plus.md`.
- Trạng thái và bằng chứng: `shared/docs/implementation/`.
- VSTO production source: `src/ChuanHoa.AddIn.Vsto/`.
- API production source: `src/ChuanHoa.Api/`.

## Trạng thái hiện tại

Nền VSTO, Ribbon 38 control, signed local rule/lease, scanner local, annotation comment/tô đỏ, snapshot `.doc`/`.docx` và 16 tiện ích Word local đã có code thực. Development Admin quản lý user thử, trial và bảng giá tại `http://127.0.0.1:5206/development/admin`. Các command chưa port hoặc chưa đạt golden/compatibility gate vẫn fail closed. Xem `shared/docs/implementation/Implementation_Status.md`; không suy diễn production readiness từ README.
