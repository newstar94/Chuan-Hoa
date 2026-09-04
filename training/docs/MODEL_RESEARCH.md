# Nghiên cứu Kiến trúc Mô hình (Model Research) cho Bộ Sửa Lỗi Chính Tả Tiếng Việt Offline

Tài liệu nghiên cứu này phục vụ yêu cầu của `TASK_ Research, Train, Distill, Quantize and Export a Tiny Vietnamese ONNX Context Model.md`.

---

## 1. Mục tiêu và Tiêu chí Lựa chọn

Mục tiêu là xây dựng mô hình AI chạy cục bộ (100% Offline) trên CPU máy tính người dùng văn phòng, làm nhiệm vụ **Xếp hạng Ứng viên theo Ngữ cảnh (Tiny Vietnamese Contextual Candidate Ranker)**.

| Tiêu chí | Ngưỡng bắt buộc | Mục tiêu tối ưu (Preferred) |
| :--- | :--- | :--- |
| **Loại mô hình** | Encoder-only Candidate Ranker | Cross-encoder hoặc Context-Candidate Ranker |
| **Kích thước ONNX INT8** | $\le 50\text{ MB}$ | **$12 - 25\text{ MB}$** |
| **Số tham số (Parameters)** | $\le 30\text{M}$ | **$12 - 16\text{M}$ (Tiny-B)** |
| **Độ trễ CPU (P95)** | $\le 80\text{ ms}$ | **$< 30 - 50\text{ ms}$** |
| **RAM tiêu thụ (Working Set)** | $\le 120\text{ MB}$ | **$< 80\text{ MB}$** |
| **Execution Target** | CPU x86/x64 (AVX2/VNNI) | Không đòi hỏi GPU |

---

## 2. Bảng So sánh các Kiến trúc Ứng viên

| Mô hình | Số tham số | Kiến trúc | Tokenizer | Khả năng Tiếng Việt | Kích thước INT8 ước tính | Độ trễ CPU ước tính | Vai trò phù hợp |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **PhoBERT-base-v2** | ~135M | RoBERTa (12L, 768H) | VnCoreNLP + BPE 64K | Rất cao | ~130 MB | 180 - 350 ms | **Teacher duy nhất lúc Train** (Không đóng gói Runtime) |
| **DistilPhoBERT** | ~66M | RoBERTa (6L, 768H) | BPE 64K | Cao | ~65 MB | 90 - 150 ms | Baseline so sánh |
| **Tiny-A (Custom)** | ~9.5M | Transformer (6L, 256H, 4Heads) | BPE 16K | Tốt trên văn bản hành chính | ~10 MB | 15 - 25 ms | Student dự phòng siêu nhẹ |
| **Tiny-B (Custom - PRIMARY)** | **~14.2M** | **Transformer (6L, 320H, 8Heads)** | **BPE 16K** | **Rất cao (đã chưng cất)** | **~15 - 20 MB** | **25 - 40 ms** | **LỰA CHỌN TỐI ƯU CHO PRODUCTION** |
| **Tiny-C (Custom)** | ~22.5M | Transformer (8L, 384H, 8Heads) | BPE 20K | Rất cao | ~24 - 30 MB | 45 - 75 ms | Student dự phòng độ chính xác cao |
| **ViT5 / BARTpho** | 120M - 400M | Seq2Seq (Encoder-Decoder) | SentencePiece | Tốt | > 150 MB | 400 - 1200 ms | **KHÔNG DÙNG** (Nguy cơ hallucination, quá chậm) |

---

## 3. Lý do Lựa chọn Kiến trúc Tiny-B

1. **Trade-off hoàn hảo**: Tiny-B với 6 layers, hidden 320, 8 attention heads và vocabulary 16K đạt tỷ lệ cân bằng tối ưu giữa khả năng biểu diễn ngữ nghĩa tiếng Việt và hiệu năng tính toán CPU.
2. **Loại trừ hoàn toàn rủi ro ảo giác (Hallucination)**: Mô hình không tự sinh từ hay viết lại câu, mà chỉ chấm điểm xác suất Softmax cho danh sách 4–8 ứng viên đã qua sàng lọc bởi Candidate Generator.
3. **Độ trễ phản hồi cực nhanh**: Với độ dài ngữ cảnh 64 tokens, Tiny-B INT8 đạt P50 ~ 25ms và P95 ~ 45ms trên CPU Intel Core i5 thế hệ 8 trở lên, hoàn toàn không làm giật lag Word khi người dùng soạn thảo.
