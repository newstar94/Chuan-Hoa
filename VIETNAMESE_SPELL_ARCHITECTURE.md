# VIETNAMESE SPELL ARCHITECTURE: Kiến trúc Hệ thống Sửa Lỗi Chính Tả Tiếng Việt Hybrid Local/Offline

Tài liệu này tổng hợp toàn bộ kiến trúc hệ thống sửa lỗi chính tả tiếng Việt kết hợp quy tắc và mô hình ngữ cảnh AI chạy ngoại tuyến hoàn toàn trên máy người dùng, tích hợp vào Add-in Microsoft Word.

---

## 1. Tổng quan Kiến trúc Tách Tiến trình (Process Isolation)

Nhằm tuân thủ nguyên tắc an toàn cao nhất cho Microsoft Word:
> **Mô hình AI ONNX Runtime không bao giờ chạy trực tiếp trong tiến trình `WINWORD.EXE`.**

Toàn bộ quá trình suy luận AI và xử lý tài nguyên nặng được cách ly sang tiến trình con: **`VietnameseEngine.exe`**.

```text
+-------------------------------------------------------------+
|                 Microsoft Word (WINWORD.EXE)                |
|                                                             |
|   +-------------------+          +----------------------+   |
|   |  ChuanHoa Ribbon  | <------> | WordOneClickRuntime  |   |
|   +-------------------+          +----------------------+   |
|            |                                |               |
|            v                                v               |
|   +-------------------+          +----------------------+   |
|   | FindingAnnotation | <------- | ParagraphContentCache|   |
|   | Adapter (Comments)|          | (LRU SHA-256 Cache)  |   |
|   +-------------------+          +----------------------+   |
|            ^                                |               |
|            |                     +----------------------+   |
|            +-------------------- | VietnameseEngine     |   |
|                                  | IpcClient            |   |
|                                  +----------------------+   |
+---------------------------------------------|---------------+
                                              |
                          Named Pipe IPC      |  \\.\pipe\ChuanHoa_VietnameseEngine
                          (Timeout 2000ms)    |
                                              v
+-------------------------------------------------------------+
|               VietnameseEngine.exe (Background)             |
|                                                             |
|   +-------------------+          +----------------------+   |
|   | Named Pipe Server | -------> | Engine Dispatcher    |   |
|   +-------------------+          +----------------------+   |
|                                             |               |
|                                             v               |
|                                  +----------------------+   |
|                                  | VietnameseWord       |   |
|                                  | Tokenizer (Offsets)  |   |
|                                  +----------------------+   |
|                                             |               |
|        +------------------------------------+               |
|        |                                    |               |
|        v                                    v               |
|   [Level 0 & 1]                       [Level 2 & 3]         |
|   Fast Rule Engine                    Candidate Generator   |
|   (2 dấu cách, dấu câu,               (Confusion Sets,      |
|    Telex cơ bản)                       Levenshtein <= 2)    |
|        |                                    |               |
|        |                                    v               |
|        |                              +------------------+  |
|        |                              | Tiny ONNX Model  |  |
|        |                              | Ranker (INT8 CPU)|  |
|        |                              +------------------+  |
|        |                                    |               |
|        v                                    v               |
|   +-----------------------------------------------------+   |
|   |         Confidence & Score Margin Filter            |   |
|   +-----------------------------------------------------+   |
+-------------------------------------------------------------+
```

---

## 2. Các Tầng Lọc Đa Cấp (Multi-tier Pipeline)

1. **Level 0 (Khoảng trắng, 2 dấu cách, dấu câu, ký tự ẩn)**:
   - Xử lý hoàn toàn bằng Rule Engine nội bộ (< 2ms).
   - Tự động phát hiện và gộp 2+ dấu cách liên tiếp về 1 dấu cách duy nhất.
   - Xóa bỏ khoảng trắng thừa trước các dấu câu `, . ; : ! ?`.
   - Không chuyển sang AI.

2. **Level 1 (Lỗi gõ phím hiển nhiên / Từ điển âm tiết)**:
   - Dò từ vựng tiếng Việt trong từ điển `VietnameseLexiconSpellChecker`.
   - Sử dụng khoảng cách Levenshtein $\le 2$ để tìm từ đúng duy nhất.
   - Không chuyển sang AI.

3. **Level 2 & 3 (Sai mập mờ / Lỗi ngữ cảnh Real-word)**:
   - Áp dụng khi gặp các cặp từ nhầm lẫn trong hành chính: *bàn dao* ➔ *bàn giao*, *sử lý* ➔ *xử lý*, *sát nhập* ➔ *sáp nhập*, *điều khoảng* ➔ *điều khoản*.
   - Sinh danh sách 4–8 ứng viên (luôn bao gồm từ gốc ở vị trí #0).
   - Mô hình Tiny Context Ranker chấm điểm phân phối xác suất Softmax.
   - Chỉ đưa ra khuyến nghị sửa khi điểm số ứng viên $\ge 0.85$ và vượt trội từ gốc $\Delta \ge 0.15$.

---

## 3. Quản lý Từ điển Cá nhân (Personal Dictionary) & Phản hồi Học tập Cục bộ

- Lưu trữ cục bộ tại `%LocalAppData%\ChuanHoa\Dictionaries\`:
  - `user_custom_dictionary.txt`: Lưu các từ chuyên ngành, tên viết tắt (CSGT, BHXH, PC04...) do người dùng chủ động thêm vào.
  - `document_ignore_list.json`: Danh sách các từ người dùng chọn "Bỏ qua trong tài liệu này".
- Phản hồi cục bộ: Lưu tỷ lệ chấp nhận/từ chối của từng cặp gợi ý tại `%LocalAppData%\ChuanHoa\Feedback\spelling_feedback.tsv` để cá nhân hóa trọng số gợi ý.

---

## 4. An toàn Word COM & Quản trị Lỗi (Fault-Tolerance)

- **Không bao giờ làm treo giao diện Word (Non-blocking UI)**: Client gọi IPC hoàn toàn bằng `async/await` với CancellationToken timeout cứng 2000ms.
- **Circuit Breaker**: Nếu tiến trình Engine gặp sự cố 3 lần liên tiếp, Add-in tự động ngắt kết nối tạm thời và chuyển sang chế độ Rule Engine thuần (Fallback).
- **Idle Watchdog**: Tiến trình `VietnameseEngine.exe` tự động giải phóng RAM và thoát khi không có thao tác trong 15 phút.
- **Bảo mật tuyệt đối (Privacy First)**: 100% Offline, không gửi bất kỳ byte dữ liệu nào ra mạng Internet.
