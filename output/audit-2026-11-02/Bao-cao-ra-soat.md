# Rà soát treo file lớn, đường kẻ và quốc hiệu–tiêu ngữ

Ngày: 02/11/2026. Mã nguồn: `a5db3ab`, phiên bản khai báo `1.0.0.192`.

## Kết luận

Có các lỗi và điểm nghẽn cụ thể trong mã hiện tại, không chỉ là vấn đề kích thước file. Đã tái hiện được điểm nghẽn kiểm tra chính tả, hủy chậm, nhiều sai lệch nhận dạng đầu văn bản, và việc bộ kiểm tra bỏ qua đường kẻ sai khi có thêm một đường đúng.

Chưa xác nhận nguyên nhân duy nhất gây treo trên file của người dùng: chưa có file tái hiện và việc gọi Word để kiểm thử bị lỗi thư viện tự động hóa Office. Các lỗi đặt/xóa hình học được phân biệt rõ giữa bằng chứng mã nguồn và bằng chứng chạy thực tế bên dưới.

Không sửa mã sản phẩm, không đổi bộ cài/registry, không cài lại add-in và không mở/chỉnh sửa tài liệu người dùng. Chỉ thêm bộ kiểm tra độc lập và báo cáo trong thư mục này.

## 1. P1 — Đối chiếu viết hoa chiếm phần lớn thời gian kiểm tra chính tả

**Đã đo trên mã hiện tại chạy .NET Framework 4.8, ngoài Word.**

`CheckConfiguredCapitalizations` duyệt từng quy tắc rồi từng đoạn văn; mỗi cặp lại tìm cụm từ với phép so sánh tiếng Việt. Bộ quy tắc đi kèm có **6.469 mục**. Với 8.000 đoạn, vòng lặp có khoảng **51,75 triệu cặp quy tắc–đoạn** trước khi tính chi phí tìm chuỗi.

| Fixture văn bản lặp, không có ảnh/bảng | Kiểm tra thể thức | Kiểm tra chính tả |
| --- | ---: | ---: |
| 500 đoạn | 66 ms | 8.202 ms |
| 2.000 đoạn | 110 ms | 34.128 ms |
| 8.000 đoạn | 425 ms | 121.101 ms |

Phép đo tách bước trên tiến trình x64, 1.000 đoạn:

- Riêng đối chiếu viết hoa: **12.013 ms**.
- Toàn bộ kiểm tra chính tả: **13.524 ms**, không phát hiện lỗi chính tả trên fixture.
- Bước viết hoa chiếm khoảng **89%** tổng thời gian trong phép đo này.

Đây là các phép đo đơn trên máy hiện tại, không phải số liệu p95 hay thời gian của một file Word thực. Tuy nhiên, chậm đã xuất hiện khi không có bất kỳ chi phí đọc Word, phân trang hay thêm comment nào.

Luồng Ribbon gọi `Prepare` và scanner đồng bộ. Phần tính toán này vì vậy giữ luồng giao diện Word khi chạy trong add-in. Lệnh Chuẩn hóa toàn bộ còn có thể chạy lại phân tích Full ở các pha xóa comment/tách dòng, đặt đường kẻ và hậu kiểm: tổng cộng 2–4 lượt tùy tài liệu, trước bước ghi comment kết quả.

Nguồn: [vòng lặp viết hoa](<D:/Chuẩn Hóa/src/ChuanHoa.Client.Core/Scanning/CanonicalRuleScanner.cs:1403>), [gọi scanner](<D:/Chuẩn Hóa/src/ChuanHoa.AddIn.Vsto/Runtime/WordDocumentReadRuntime.cs:106>), [luồng Ribbon](<D:/Chuẩn Hóa/src/ChuanHoa.AddIn.Vsto/Runtime/RibbonRuntime.cs:605>), [các lượt phân tích lại](<D:/Chuẩn Hóa/src/ChuanHoa.AddIn.Vsto/Runtime/WordOneClickRuntime.cs:137>).

## 2. P1 — Cơ chế hủy chưa đáp ứng trong bước nặng

**Đã tái hiện bằng yêu cầu hủy từ bộ kiểm tra.**

Với 1.000 đoạn, token được hủy ở **2.004 ms**, nhưng scanner chỉ trả về ở **12.239 ms**: chậm thêm hơn 10 giây sau yêu cầu hủy. `CheckConfiguredCapitalizations` không nhận token và không kiểm tra hủy trong vòng lặp.

Trong sản phẩm, Esc chỉ được thăm dò tại `DocumentOperationSession.Checkpoint`. Trong lúc scanner đồng bộ đang chạy, không có lần thăm dò Esc mới bên trong vòng lặp viết hoa. Chỉ kiểm tra token ở đầu/cuối pha không giải quyết được chuyện này. Sau khi bắt đầu mutation, hủy còn bị vô hiệu hóa có chủ đích để bảo vệ tính toàn vẹn tài liệu; không nên bỏ cơ chế bảo vệ đó một cách đơn giản.

Nguồn: [kiểm tra token sau pha](<D:/Chuẩn Hóa/src/ChuanHoa.Client.Core/Scanning/CanonicalRuleScanner.cs:182>), [thăm dò Esc](<D:/Chuẩn Hóa/src/ChuanHoa.AddIn.Vsto/Runtime/DocumentOperationSession.cs:89>), [khóa hủy sau mutation](<D:/Chuẩn Hóa/src/ChuanHoa.AddIn.Vsto/Runtime/DocumentOperationSession.cs:66>).

## 3. P1 — Nhánh đo đường kẻ đi vòng qua chốt an toàn tài liệu nặng

**Xác nhận bằng luồng gọi trong mã; chưa đo thời gian trực tiếp trong Word.**

`IsPageLayoutCaptureSafe` trả false nếu tài liệu có hơn 10 bảng hoặc hơn 200.000 vị trí ký tự. Nhánh đọc đoạn thông thường tôn trọng kết quả này. Nhưng ngay sau đó `Build` luôn gọi `CaptureRequiredLineLayout`, không truyền hoặc kiểm tra `allowPageLayout`.

Nhánh sau vẫn gọi đo chiều rộng, lấy vị trí dòng và số trang của các thành phần được nhận dạng. Các phép đọc này chạm vào layout Word. Vì thế chốt an toàn hiện tại không bao phủ toàn bộ đường thực thi. Đây là nguy cơ gây phân trang đồng bộ lâu trên file lớn/nhiều bảng; chưa khẳng định đây là điểm treo của file cụ thể người dùng đang gặp.

Bộ kiểm tra nguồn hiện vẫn báo `tableHeavyPaginationGuard=true` vì kiểm tra sự hiện diện của mẫu mã, không chứng minh mọi đường gọi lấy layout đều được bảo vệ.

Nguồn: [chỗ gọi bỏ qua điều kiện](<D:/Chuẩn Hóa/src/ChuanHoa.AddIn.Vsto/Runtime/WordDocumentSnapshotBuilder.cs:475>), [đo layout bắt buộc](<D:/Chuẩn Hóa/src/ChuanHoa.AddIn.Vsto/Runtime/WordDocumentSnapshotBuilder.cs:1763>), [ngưỡng an toàn](<D:/Chuẩn Hóa/src/ChuanHoa.AddIn.Vsto/Runtime/WordDocumentSnapshotBuilder.cs:1854>), [validator](<D:/Chuẩn Hóa/tools/vsto/validate_vsto_source.py:867>).

## 4. P1 — Đường viền cũ ở phần sau tài liệu bị ghi nhận thành “không có”

**Xác nhận bằng mã; fixture Word chưa chạy được.**

Cả hai nhánh đọc đoạn chỉ lấy viền dưới khi `storyParagraphIndex <= 80`. Ngoài ngưỡng này, `HasBottomBorder` giữ giá trị false, dù Word có thể đang chứa viền dưới. Với story hơn 400 đoạn, định dạng đoạn trong bảng sau đoạn 150 còn bị bỏ qua, bao gồm trạng thái gạch chân.

Trong khi đó detector và `CaptureRequiredLineLayout` đã hỗ trợ tìm thành phần của văn bản hành chính xuất hiện ở phía sau. Nhánh đo lại chỉ giữ `p.HasBottomBorder` cũ, không đọc lại viền. Kết quả: có thể nhận ra Tiêu ngữ ở văn bản thứ hai nhưng không biết nó đã có đường viền.

`NormalizeRequiredLine` chỉ xóa viền cũ khi snapshot nói `HasBottomBorder=true`. Vì vậy đường thực thi có thể chèn thêm Line Shape mà giữ nguyên viền cũ, hoặc chèn thêm dưới một Tiêu ngữ đã gạch chân nhưng trạng thái đó không được đọc.

Nguồn: [giới hạn 80 ở nhánh thường](<D:/Chuẩn Hóa/src/ChuanHoa.AddIn.Vsto/Runtime/WordDocumentSnapshotBuilder.cs:794>), [giới hạn 150 trong bảng](<D:/Chuẩn Hóa/src/ChuanHoa.AddIn.Vsto/Runtime/WordDocumentSnapshotBuilder.cs:950>), [giới hạn 80 ở nhánh lớn](<D:/Chuẩn Hóa/src/ChuanHoa.AddIn.Vsto/Runtime/WordDocumentSnapshotBuilder.cs:1001>), [xóa viền theo snapshot](<D:/Chuẩn Hóa/src/ChuanHoa.AddIn.Vsto/Runtime/WordOneClickRuntime.cs:1756>).

## 5. P1 — Bộ hậu kiểm có thể hết báo lỗi dù đường kẻ sai vẫn còn

**Đã tái hiện trong scanner; phần sửa hình học được đối chiếu bằng mã, chưa chạy trong Word.**

Fixture có một đường nét đứt gắn với Tiêu ngữ: scanner báo lỗi đường kẻ. Giữ nguyên đường nét đứt và thêm một đường đúng: scanner trả **0 lỗi đường kẻ**, dù snapshot vẫn chứa cả hai đường.

Nguyên nhân: `EvaluateRequiredLine` lọc bỏ các đường sai kiểu nét trước, sau đó chỉ cần một đường trong nhóm còn lại hợp lệ. Chốt phát hiện nhiều đường chỉ đếm các đường đúng kiểu/đúng hình học, nên không bắt tổ hợp “một đường đúng + một đường sai cũ”.

Nhánh sửa lại dùng tập ứng viên hẹp hơn scanner: đường cũ thông thường phải ở gần tọa độ mới trong 18 pt, hoặc khớp điều kiện legacy dành cho bảng. Một đường nét đứt neo vào Tiêu ngữ nhưng cách tọa độ mới hơn 18 pt ở đoạn ngoài bảng có thể bị bỏ lại. Scanner sau đó không còn cảnh báo khi đường mới hợp lệ đã được thêm.

Ngoài ra, danh sách legacy chuẩn bị trước chỉ lưu `tên shape → mã quy tắc`; lúc xóa lọc theo mã quy tắc mà không giữ danh tính khối văn bản/chủ sở hữu. Đây là rủi ro cần kiểm tra thêm với file ghép nhiều văn bản, chưa nâng thành kết luận lỗi Word đã tái hiện.

Nguồn: [lọc đường sai khỏi hậu kiểm](<D:/Chuẩn Hóa/src/ChuanHoa.Client.Core/Scanning/CanonicalRuleScanner.cs:547>), [điều kiện chọn đường cũ](<D:/Chuẩn Hóa/src/ChuanHoa.AddIn.Vsto/Runtime/WordOneClickRuntime.cs:1795>), [lọc danh sách legacy chỉ theo rule](<D:/Chuẩn Hóa/src/ChuanHoa.AddIn.Vsto/Runtime/WordOneClickRuntime.cs:1807>).

## 6. P2 — Nhận dạng và tách Quốc hiệu–Tiêu ngữ dùng hai chuẩn khác nhau

**Đã tái hiện trên 5 biến thể cùng nội dung.**

| Quốc hiệu và Tiêu ngữ chung đoạn, ngăn bằng Shift+Enter | Nhận ra đầu văn bản | Tìm được điểm tách | Báo lỗi cần tách/đường Tiêu ngữ |
| --- | --- | --- | --- |
| Unicode dựng sẵn, dấu gạch thường | Có | Có | Có |
| Unicode tổ hợp NFD | Có | Không | Không |
| Dấu gạch không ngắt U+2011 | Có | Không | Không |
| Tiêu ngữ thiếu dấu gạch | Có | Không | Không |
| Có ký tự zero-width trong Quốc hiệu | Có | Không | Không |

`ComponentTextNormalizer` chuẩn hóa dấu, khoảng trắng và ký tự ẩn. `CombinedNationalHeader.GetBreakOffsets` lại chạy regex trực tiếp trên chuỗi gốc với điều kiện hẹp hơn. Vì vậy đoạn được gán `nationalTitle`, nhưng không tách được và cũng không có `nationalMotto` riêng để xử lý đường kẻ.

Không nên chỉ thay chuỗi gốc bằng chuỗi đã chuẩn hóa: vị trí ký tự Word phải được bảo toàn. Cần dùng chung kết quả phân tích có ánh xạ về tọa độ gốc.

Nguồn: [bộ nhận dạng](<D:/Chuẩn Hóa/src/ChuanHoa.Client.Core/Text/ComponentTextNormalizer.cs:131>), [bộ tìm điểm tách](<D:/Chuẩn Hóa/src/ChuanHoa.Client.Core/Text/CombinedNationalHeader.cs:11>), [scanner dùng bộ tách khác](<D:/Chuẩn Hóa/src/ChuanHoa.Client.Core/Scanning/CanonicalRuleScanner.cs:370>).

## 7. P2 — Quan hệ ô bảng và vùng chứa đầu văn bản còn bị xử lý thiếu

**Đã tái hiện bằng snapshot không gán sẵn vai trò.**

- Quốc hiệu ở ô trái, Tiêu ngữ ở ô phải cùng hàng, hai đoạn kế nhau: vẫn gán `nationalMotto`, dù không cùng cột hoặc ở ngay dưới Quốc hiệu. Nhánh `position == index - 1` trả true trước khi kiểm tra quan hệ ô.
- Tiêu ngữ ở đúng cột, hàng liền dưới, nhưng ô trái của hàng đó có 7 đoạn: không nhận ra Tiêu ngữ vì chỉ tìm lùi 6 đoạn theo thứ tự đọc Word.
- Đặt hai dòng đúng nội dung trong Header thật của Word hoặc Text Box: không gán hai vai trò. Snapshot có đọc các story này, nhưng detector chỉ dùng `wdMainTextStory`. Đây là giới hạn hỗ trợ hiện tại, cần thông báo rõ hoặc bổ sung luồng riêng, không tự ý coi mọi header/textbox là đầu văn bản hành chính.

Nguồn: [quan hệ Tiêu ngữ–Quốc hiệu](<D:/Chuẩn Hóa/src/ChuanHoa.Client.Core/Scanning/DocumentRoleDetector.cs:408>), [chỉ nhận main story](<D:/Chuẩn Hóa/src/ChuanHoa.Client.Core/Scanning/DocumentRoleDetector.cs:510>).

## Kiểm chứng và giới hạn

- Client Core Release: **535/535 đạt**, không bỏ qua test.
- Bộ kiểm tra nguồn VSTO: **PASS**; bản thân kết quả ghi build/load/visual `NOT_RUN`. PASS này không chứng minh các lỗi hành vi bên trên không tồn tại.
- Bộ probe độc lập liên kết trực tiếp các file runtime hiện tại: build .NET Framework 4.8 x64 **0 lỗi, 0 cảnh báo**. Đây không phải build/cài bộ VSTO phát hành.
- Kiểm thử toàn solution `dotnet test ChuanHoa.slnx -c Release --no-restore` **không đạt**: API có lock file còn `payOS` nhưng project đã bỏ; Domain còn tham chiếu `SignedEnvelope<ExecutionGrantPayload>` không tìm thấy. Core 535, Contracts 4, Rules 10 vẫn chạy đạt. Không sửa các lỗi ngoài phạm vi triệu chứng, không tự cập nhật dependency lock.
- Registry trỏ tới DLL cài tại `DevelopmentInstaller/Current`, metadata DLL báo **1.0.0.192**, giống số phiên bản source. Chưa chứng minh đây là cùng chính xác nội dung binary với source hiện tại.
- Gọi Word để kiểm thử bị chặn ngay trước khi tạo tài liệu bởi `0x80029C4A / TYPE_E_CANTLOADLIBRARY`; xảy ra cả chạy x64 trực tiếp và qua launcher Office. Không thay đổi registry, không sửa/cài lại Office. Lỗi này là blocker của phép kiểm chứng tự động, **không đủ để kết luận nó là nguyên nhân treo khi người dùng bấm Ribbon**.
- Đã đóng đúng bốn tiến trình Word tự động hóa rỗng do probe tạo; không có Word đang chạy trước phép thử và không đụng tài liệu người dùng.
- Chưa kiểm thử trực quan vị trí đường kẻ, lưu/đóng/mở lại, .doc so với .docx, Word 2010/x86, hàng nghìn comment, file nhiều ảnh/ô gộp và file thật gây treo.
- Trạng thái cuối: không có thay đổi tracked source. Validator có tự cập nhật một hash trong evidence; thay đổi tự sinh đó đã được trả về đúng trạng thái trước audit. Chỉ thư mục báo cáo/probe này là mới.

## Thứ tự sửa đề nghị — chưa thực hiện

1. Giải quyết nút thắt viết hoa: lập chỉ mục quy tắc/cụm từ đã chuẩn hóa; giữ đầy đủ quy tắc, cách so sánh tiếng Việt và ánh xạ offset. Thêm đo thời gian riêng từng pha.
2. Tách tính toán thuần dữ liệu khỏi luồng Word, có kiểm tra hủy đủ dày. Giữ Word COM trên luồng STA; nếu chia lô thao tác Word phải có chốt chống vào lại và giải phóng đối tượng trước khi nhường điều khiển. Không chữa bằng `Task.Run` bao quanh Word COM hay `DoEvents` tùy tiện.
3. Đọc đầy đủ trạng thái cần thiết của mọi thành phần đầu văn bản được nhận dạng, dù nằm sau đoạn 80/150. Không biểu diễn “chưa đọc” thành “không có”. Rà toàn bộ đường gọi lấy layout để chốt tài liệu nặng có hiệu lực thực.
4. Dùng chung mô hình nhận dạng đầu văn bản, điểm tách và quan hệ ô; giữ rào chắn tránh nhận nhầm câu trích dẫn trong nội dung.
5. Lập kế hoạch thay đường theo đúng story/section/khối/chủ sở hữu; xác minh đường mới, xóa đúng đường cũ, hậu kiểm cả phần dư sai. Bảo toàn hình và đường người dùng không xác định chắc chắn thuộc thành phần cần sửa.
6. Chạy lại trên bản sao file thật và ma trận tài liệu lớn. Tiêu chí phải gồm thời gian, hủy, số đường đúng/sai, chạy lần hai không sinh thêm đường, và lưu/mở lại vẫn đúng. Giữ yêu cầu comment tại từng lỗi giải được vị trí; không âm thầm cắt bớt comment để làm nhanh.

## Bằng chứng lưu kèm

- [Đo thời gian và ca nhận dạng](<D:/Chuẩn Hóa/output/audit-2026-11-02/core-results.txt>).
- [Tách thời gian riêng bước viết hoa](<D:/Chuẩn Hóa/output/audit-2026-11-02/profile-results.txt>).
- [Đường sai còn lại và độ trễ hủy](<D:/Chuẩn Hóa/output/audit-2026-11-02/edge-results.txt>).
- [Lỗi kiểm thử Office x64](<D:/Chuẩn Hóa/output/audit-2026-11-02/word-appvlp-x64-results.txt>).

Các file probe được giữ lại để tái chạy sau khi có quyền sửa; chúng không nằm trong project sản phẩm hay bộ cài.
