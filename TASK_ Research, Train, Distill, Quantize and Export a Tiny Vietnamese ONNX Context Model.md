# TASK: Research, Train, Distill, Quantize and Export a Tiny Vietnamese ONNX Context Model

Bạn là Senior NLP/ML Engineer chuyên về:

- Vietnamese NLP
- Transformer
- knowledge distillation
- spelling correction
- contextual error detection
- ONNX
- ONNX Runtime
- INT8 CPU inference
- model compression
- .NET deployment

Mục tiêu của nhiệm vụ này là xây dựng model AI rất nhỏ để sử dụng trong Vietnamese Spell Checker chạy offline trên Microsoft Word Add-in.

Model cuối cùng KHÔNG phải là generative text correction model.

Không xây một model tự động viết lại câu.

Model cuối cùng phải là:

> Tiny Vietnamese Contextual Candidate Ranker

Nhiệm vụ chính:

```text
Context
+
Current word/error position
+
Candidate words
↓
Tiny Transformer
↓
score(candidate)
```

Ví dụ:

```text
Tôi muốn ___ việc vào cuối tháng.

Candidates:

ngỉ
nghỉ
nghĩ
nghị
```

Output:

```text
ngỉ   0.006
nghỉ  0.982
nghĩ  0.010
nghị  0.002
```

Model phải ưu tiên:

```text
nghỉ
```

Model phải có khả năng xử lý cả real-word error.

Ví dụ:

```text
bàn dao hồ sơ
```

Candidates:

```text
dao
giao
```

Model:

```text
dao    0.01
giao   0.99
```

---

# 1. KHÔNG SỬ DỤNG GENERATIVE MODEL TRONG RUNTIME

Không sử dụng trực tiếp:

- ViT5
- BARTpho
- T5
- GPT
- decoder LLM

làm production inference cho add-in.

Có thể sử dụng model lớn làm:

```text
teacher
dataset labeling
benchmark reference
```

nhưng production model phải là encoder nhỏ.

Không dùng:

```text
incorrect sentence
↓
seq2seq
↓
correct sentence
```

Production architecture phải là:

```text
Candidate Generator
        ↓
candidate list
        ↓
Tiny Context Ranker
        ↓
candidate scores
```

---

# 2. RESEARCH TRƯỚC KHI TRAIN

Trước khi viết training code, hãy nghiên cứu các model/dataset hiện có.

Tạo:

```text
training/docs/MODEL_RESEARCH.md
```

So sánh ít nhất:

- PhoBERT
- PhoBERT-base-v2
- DistilPhoBERT nếu phù hợp
- MiniLM-like architecture
- TinyBERT-like architecture
- small BERT/RoBERTa architecture
- Vietnamese spelling correction models
- BARTpho/ViT5 correction models chỉ như teacher/baseline

Đối với mỗi model ghi:

```text
parameters
architecture
tokenizer
Vietnamese capability
context length
license
model size
estimated CPU latency
ONNX compatibility
suitability as teacher
suitability as student
```

Không chọn model dựa trên độ phổ biến.

Chọn dựa trên mục tiêu:

```text
accuracy / latency / RAM / model size
```

---

# 3. KIỂM TRA LICENSE

Tạo:

```text
training/docs/DATA_LICENSES.md
```

Đối với mọi:

- dataset
- pretrained model
- tokenizer
- corpus

phải ghi:

```text
source
license
commercial-use status
attribution requirement
redistribution requirement
uncertainty
```

Không đưa dataset/model có license không rõ ràng vào production training pipeline một cách âm thầm.

Nếu license chưa rõ:

```text
STATUS = REVIEW_REQUIRED
```

Không giả định rằng:

```text
public on GitHub
```

đồng nghĩa:

```text
commercial use allowed
```

---

# 4. DATASET STRATEGY

Không train chỉ bằng synthetic errors.

Tách dữ liệu thành ba nhóm.

## GOLD

Lỗi thật do người viết tạo và được con người sửa.

Ví dụ có thể nghiên cứu:

```text
VSEC
Viwiki-Spelling
```

Gold data chủ yếu dùng cho:

```text
validation
test
threshold calibration
final fine-tuning
```

Không được contaminate test set.

---

# 5. SYNTHETIC DATA

Xây Vietnamese Error Generator riêng.

Từ sentence chuẩn:

```text
Tôi muốn nghỉ việc vào cuối tháng.
```

tạo:

```text
Tôi muốn ngỉ việc vào cuối tháng.
```

Noise taxonomy phải bao gồm ít nhất:

```text
Telex
VNI
tone error
missing tone
wrong tone
character deletion
character insertion
character substitution
transposition
keyboard neighbor
double character
word merge
word split

s/x
ch/tr
d/gi/r
l/n
c/t
n/ng

regional pronunciation confusion
phonetic confusion
```

Đặc biệt tạo:

```text
REAL_WORD_ERROR
```

Ví dụ:

```text
điều khoản
→ điều khoảng

bàn giao
→ bàn dao

xử lý
→ sử lý
```

Trong trường hợp real-word error:

từ sai vẫn phải có trong Vietnamese dictionary.

---

# 6. KHÔNG ĐỂ SYNTHETIC DATA QUÁ KHÁC DỮ LIỆU THẬT

Noise generator không được làm:

```text
50% token trong mọi câu đều sai
```

nếu dữ liệu thực tế phần lớn chỉ có 1–2 lỗi.

Tạo phân phối:

```text
0 errors
1 error
2 errors
3+ errors
```

theo distribution gần dữ liệu thực.

Phải có nhiều câu:

```text
NO ERROR
```

để model học không báo lỗi bừa.

Đây là yêu cầu cực kỳ quan trọng.

---

# 7. HARD NEGATIVE

Không train chỉ bằng negative ngẫu nhiên.

Ví dụ:

```text
Tôi muốn ___ việc
```

Hard candidates:

```text
nghỉ
nghĩ
nghị
```

tốt hơn:

```text
nghỉ
bầu
xe
cây
```

Candidate negative phải đến từ:

- edit distance
- confusion set
- pronunciation similarity
- Vietnamese syllable similarity
- accent variation
- frequent real words
- candidate generator thực tế

Train distribution phải giống inference distribution.

---

# 8. TRAIN MODEL ĐỂ RANK CANDIDATES

Mỗi training example nên có:

```text
context
errorSpan
candidateSet
correctCandidateIndex
```

Ví dụ:

```json
{
  "context": "Tôi muốn ngỉ việc vào cuối tháng.",
  "error": "ngỉ",
  "candidates": [
    "ngỉ",
    "nghỉ",
    "nghĩ",
    "nghị"
  ],
  "correctIndex": 1
}
```

Quan trọng:

candidate list luôn nên có:

```text
ORIGINAL WORD
```

Model phải được phép quyết định:

```text
NO CHANGE
```

---

# 9. TRAINING INPUT FORMAT

Benchmark tối thiểu hai format.

## Option A — Candidate cross encoder

Mỗi candidate tạo một input:

```text
[CLS]
Tôi muốn
[ERR]
nghỉ
[/ERR]
việc vào cuối tháng.
[SEP]
```

Model output:

```text
score
```

Các candidate chạy trong cùng batch.

---

## Option B — Context + candidate pair

```text
[CLS]
Tôi muốn [MASK_ERR] việc vào cuối tháng.
[SEP]
nghỉ
[SEP]
```

Output:

```text
score
```

Benchmark A và B.

Chọn format:

```text
accuracy
+
ONNX simplicity
+
CPU latency
```

Không chọn chỉ theo accuracy.

---

# 10. LISTWISE RANKING

Ưu tiên listwise loss.

Candidates:

```text
c1
c2
...
ck
```

Model:

```text
s1
s2
...
sk
```

Sau đó:

```text
P(candidate) = softmax(scores)
```

Loss:

```text
CrossEntropy(scores, correctCandidateIndex)
```

Điều này phù hợp hơn binary classification độc lập.

---

# 11. NO-CORRECTION TRAINING

Tạo lượng lớn example câu hoàn toàn đúng.

Ví dụ:

```text
Tôi muốn nghỉ việc.
```

Candidates:

```text
nghỉ
nghĩ
nghị
ngỉ
```

correct:

```text
nghỉ
```

Model phải học rằng original đúng và không được thay đổi.

Đây là một trong những kỹ thuật chính để giảm false positive.

---

# 12. MODEL ARCHITECTURE SEARCH

Không xây duy nhất một configuration.

Train hoặc benchmark tối thiểu:

## Tiny-A

```text
layers = 6
hidden = 256
attention_heads = 4 hoặc 8
ffn = 1024
vocab ≈ 16K
```

Target:

```text
~9–10M parameters
```

---

## Tiny-B — PRIMARY CANDIDATE

```text
layers = 6
hidden = 320
attention_heads = 8
ffn = 1280
vocab = 16K–20K
```

Target:

```text
~12–16M parameters
```

---

## Tiny-C

```text
layers = 8
hidden = 384
attention_heads = 8
ffn = 1536
vocab ≈ 20K
```

Target:

```text
~20–25M parameters
```

Đây là search space ban đầu.

Không cần tuân thủ chính xác nếu benchmark chứng minh cấu hình khác tốt hơn.

---

# 13. MODEL KHÔNG CẦN CONTEXT QUÁ DÀI

Spell checking realtime không cần 512 token.

Benchmark:

```text
max_length = 48
64
96
128
```

Ưu tiên khoảng:

```text
64–96 subword tokens
```

nếu accuracy không giảm đáng kể.

Lấy context chủ yếu quanh error.

Ví dụ:

```text
left context
+
candidate
+
right context
```

Không đưa cả paragraph dài vào model nếu không cần.

---

# 14. TOKENIZER

Không mặc định sử dụng tokenizer của PhoBERT.

Spell checker phải xử lý:

```text
ngỉ
phuowng
ddường
OpenAI
BIDV
Win32
.NET
URL
foreign names
```

Word segmentation có thể thất bại ngay vì input chứa typo.

Vì vậy benchmark tokenizer riêng.

Ưu tiên:

```text
BPE
hoặc
SentencePiece
```

với vocab nhỏ:

```text
8K
12K
16K
20K
```

Recommended starting point:

```text
16K BPE
```

---

# 15. TOKENIZER PHẢI GIỮ TIẾNG VIỆT

Không lowercase vô điều kiện.

Phải giữ:

```text
Đ
đ
Vietnam
Việt Nam
BIDV
OpenAI
```

Unicode normalization:

```text
NFC
```

phải giống 100% giữa:

```text
training
Python inference
.NET inference
```

---

# 16. TOKENIZER PARITY TEST

Đây là acceptance criterion.

Tạo test corpus ít nhất 10.000 input.

So sánh:

```text
Python tokenizer IDs
vs
.NET tokenizer IDs
```

Yêu cầu:

```text
100% identical
```

Nếu không đạt, model chưa được tích hợp production.

Đưa tokenizer artifacts vào:

```text
Models/tokenizer/
```

---

# 17. PRETRAINING

Nếu custom Tiny Transformer train từ đầu, cân nhắc pretrain trước bằng:

```text
Masked Language Modeling
```

trên clean Vietnamese corpus.

Không cần pretrain quá lớn nếu task-specific distillation đạt kết quả tương đương.

Hãy benchmark:

```text
random init + task training
```

vs

```text
MLM pretraining + task training
```

Nếu MLM cải thiện đáng kể thì giữ.

Nếu không, tránh complexity không cần thiết.

---

# 18. TEACHER MODEL

Teacher có thể lớn vì chỉ dùng lúc training.

Primary teacher candidate:

```text
PhoBERT-base-v2
```

hoặc model Vietnamese encoder khác nếu benchmark tốt hơn.

Fine-tune teacher cho cùng ranking task.

Teacher không được trở thành runtime dependency.

Production package KHÔNG chứa teacher.

---

# 19. KNOWLEDGE DISTILLATION

Sau khi teacher đạt accuracy tốt, distill sang Tiny model.

Loss:

```text
L =
α * supervised_listwise_loss
+
β * KD_loss
```

KD:

```text
KL(
softmax(teacher_scores / T),
softmax(student_scores / T)
)
```

Benchmark:

```text
T = 1
2
4
```

Starting point:

```text
T = 2
```

Không hard-code lựa chọn cuối trước benchmark.

---

# 20. KHÔNG CẦN HIDDEN-STATE DISTILLATION BẮT BUỘC

Teacher và student có thể dùng tokenizer khác nhau.

Ưu tiên:

```text
logit distillation
```

vì đơn giản và robust hơn.

Chỉ thêm:

```text
hidden-state alignment
attention distillation
```

nếu benchmark chứng minh cải thiện đáng kể.

---

# 21. TRAINING CURRICULUM

Đề xuất:

## Stage A

Train trên:

```text
large synthetic data
```

để học Vietnamese correction patterns.

## Stage B

Mix:

```text
synthetic
+
real errors
```

với oversampling hợp lý cho gold data.

## Stage C

Fine-tune nhẹ trên high-quality real-world data.

Không overfit gold set.

## Stage D

Distillation + calibration.

---

# 22. DATA SPLIT

Không random split những câu gần giống nhau sau khi synthetic generation.

Split clean source document BEFORE tạo noise.

Ví dụ:

```text
documents
↓
train / validation / test
↓
noise generation
```

Tránh:

```text
same original sentence
```

xuất hiện ở cả train và test với noise khác nhau.

---

# 23. GOLD TEST SET

Final gold test:

Không được sử dụng cho:

- training
- early stopping
- threshold selection
- hyperparameter tuning

Chỉ evaluate khi model candidate đã được chọn.

Có thể có:

```text
dev_gold
final_gold
```

---

# 24. METRICS

Không sử dụng BLEU làm metric chính.

Spell checker phải đo:

```text
Detection Precision
Detection Recall
Detection F1

Correction Precision
Correction Recall
Correction F1

Top-1 Candidate Accuracy
Top-3 Candidate Accuracy

False Positive Rate

No-change Accuracy

Real-word Error Accuracy

MRR
```

Đặc biệt quan tâm:

```text
False Positive Rate
```

---

# 25. PRODUCT KPI

Production Word Add-in cần:

```text
PRECISION > RECALL
```

Thà bỏ sót một số lỗi còn hơn gạch chân câu đúng liên tục.

Tạo benchmark riêng:

```text
correct Vietnamese sentences
```

và measure:

```text
False Positive Per 1000 Words
```

---

# 26. CATEGORY METRICS

Report riêng:

```text
Telex
VNI
tone
keyboard typo
missing char
extra char
transposition
s/x
ch/tr
d/gi/r
l/n
real-word
proper noun
technical term
mixed English
```

Không chỉ báo một F1 tổng.

---

# 27. PROPER NOUN ROBUSTNESS

Dataset phải chứa nhiều:

```text
OpenAI
Microsoft
Windows
Visual Studio
BIDV
VNPT
Viettel
Techcombank
GPT-5
ASP.NET
C#
Win32
```

Không để model có xu hướng Việt hóa hoặc sửa tên riêng.

---

# 28. CODE/MIXED TEXT ROBUSTNESS

Negative examples phải bao gồm:

```text
user@example.com
https://example.com
C:\Program Files\
.NET
System.Text.Json
foo_bar
GPT-5.6
192.168.1.1
```

Thông thường các token này phải được skip bởi upstream Rule Engine.

Nhưng model vẫn phải robust nếu context chứa chúng.

---

# 29. TRAINING PIPELINE STRUCTURE

Tạo:

```text
Training/
│
├── configs/
│   ├── tiny_a.yaml
│   ├── tiny_b.yaml
│   └── tiny_c.yaml
│
├── data/
│   ├── loaders/
│   ├── preprocessing/
│   ├── noise/
│   └── candidate_generation/
│
├── tokenizer/
│   ├── train_tokenizer.py
│   └── validate_tokenizer.py
│
├── models/
│   └── context_ranker.py
│
├── train/
│   ├── pretrain_mlm.py
│   ├── train_teacher.py
│   ├── train_student.py
│   └── distill.py
│
├── evaluation/
│   ├── evaluate.py
│   ├── evaluate_categories.py
│   └── calibration.py
│
├── export/
│   ├── export_onnx.py
│   ├── optimize_onnx.py
│   └── quantize_onnx.py
│
├── benchmark/
│   ├── benchmark_pytorch.py
│   ├── benchmark_onnx.py
│   └── benchmark_dotnet/
│
└── docs/
```

---

# 30. REPRODUCIBILITY

Mỗi training run phải lưu:

```text
git commit
config
random seed
dataset hashes
tokenizer hash
teacher version
student config
learning rate
batch size
epochs
best checkpoint
metrics
```

Có:

```text
experiment_id
```

Không để model cuối cùng không biết train bằng data/config nào.

---

# 31. EXPORT ONNX

Sau khi chọn model:

```text
PyTorch
↓
ONNX FP32
↓
validate
↓
graph optimization
↓
INT8
```

Không quantize trước khi có FP32 baseline.

---

# 32. ONNX OUTPUT

Production ONNX nên nhận:

```text
input_ids
attention_mask
```

Nếu không thực sự cần:

```text
token_type_ids
```

thì bỏ để đơn giản runtime.

Output:

```text
candidate_score
```

Shape ví dụ:

```text
[K]
```

với:

```text
K <= 8
```

---

# 33. BATCH CANDIDATES

Không gọi ONNX 8 lần cho 8 candidate nếu có thể batch.

Input:

```text
batch = candidates
```

Ví dụ:

```text
[8, 64]
```

Một inference:

```text
8 candidate scores
```

Benchmark:

```text
batch 1
batch 4
batch 8
```

---

# 34. ONNX PARITY

Trước quantization:

So sánh:

```text
PyTorch logits
ONNX FP32 logits
```

Yêu cầu:

```text
max_abs_diff
```

trong tolerance hợp lý.

Sau INT8:

so sánh:

```text
accuracy delta
ranking delta
```

Không chỉ kiểm tra model load được.

---

# 35. INT8 QUANTIZATION

Target production:

```text
INT8
```

Benchmark ít nhất:

```text
dynamic quantization
static quantization
```

Ưu tiên CPU-friendly configuration.

Bắt đầu với:

```text
S8S8
QDQ
```

nếu ONNX Runtime/backend hỗ trợ phù hợp.

Nếu accuracy giảm đáng kể:

benchmark phương án khác.

Không giữ INT8 nếu model mất quá nhiều quality.

---

# 36. QUANTIZATION ACCEPTANCE

INT8 model chỉ được chọn nếu:

```text
Top1 drop <= khoảng 1 percentage point
```

hoặc mức giảm nhỏ được chứng minh chấp nhận được bằng product metric.

Đặc biệt:

```text
False Positive Rate
```

không được tăng đáng kể.

---

# 37. MODEL SIZE TARGET

Primary target:

```text
ONNX INT8 model
<= 25 MB
```

Preferred:

```text
12–20 MB
```

Hard upper bound ban đầu:

```text
50 MB
```

Nếu vượt 50 MB phải chứng minh accuracy gain đủ lớn.

---

# 38. LATENCY TARGET

Benchmark CPU.

Context:

```text
64 hoặc 96 tokens
K <= 8
```

Target:

```text
P50 < 30 ms
P95 < 80 ms
```

trên máy văn phòng tương đối phổ biến.

Không tính:

```text
model cold load
```

vào steady-state inference latency.

Báo riêng:

```text
cold start
warm inference
```

---

# 39. MEMORY

Measure:

```text
model file
working set
private bytes
peak inference memory
```

Target:

```text
incremental engine working set
preferably < 100 MB
```

sau khi model đã load.

---

# 40. CPU THREADS

Benchmark:

```text
intra_op_num_threads
1
2
4
```

Không để ONNX chiếm tất cả CPU cores làm Word lag.

Mục tiêu product:

```text
low latency
+
low CPU interference
```

Không phải maximum throughput server-style.

---

# 41. .NET BENCHMARK LÀ BẮT BUỘC

Python benchmark không đủ.

Tạo console benchmark bằng:

```text
C#
.NET
Microsoft.ML.OnnxRuntime
```

và tokenizer production.

Đo end-to-end:

```text
string input
↓
tokenization
↓
tensor construction
↓
ONNX inference
↓
score extraction
```

Đây mới là latency thực tế.

---

# 42. REUSE BUFFERS

Trong .NET benchmark/integration nghiên cứu:

- reuse InferenceSession
- reuse buffers nếu có lợi
- tránh allocation lớn per request
- không recreate tokenizer
- không reload model
- không recreate session

Profile trước khi tối ưu.

---

# 43. CALIBRATION

Softmax score không mặc định là confidence đáng tin.

Thực hiện calibration trên dev set.

Nghiên cứu:

```text
temperature scaling
```

và score margin:

```text
bestAlternativeScore - originalScore
```

Spell checker chỉ flag khi:

```text
alternative confidence cao
AND
margin đủ lớn
```

---

# 44. DETECTION TỪ RANKING

Không nhất thiết cần model detection thứ hai.

Candidates:

```text
original
+
alternatives
```

Nếu:

```text
original wins
```

=> không lỗi.

Nếu alternative thắng rõ:

```text
alternativeScore - originalScore > threshold
```

=> đề xuất correction.

Benchmark cách này trước khi thêm second model.

---

# 45. REAL-WORD ERROR

Tạo evaluation set riêng.

Ví dụ:

```text
bàn dao hồ sơ
điều khoảng hợp đồng
sử lý công việc
chia sẽ thông tin
```

Nhưng phải xác nhận ground truth theo context.

Không hard-code các câu trên vào model.

Chúng chỉ là regression examples.

---

# 46. REGRESSION SET

Tạo:

```text
Tests/SpellModel/regression.jsonl
```

Bao gồm ít nhất:

```text
ngỉ → nghỉ
bàn dao hồ sơ → bàn giao hồ sơ
điều khoảng hợp đồng → điều khoản hợp đồng
```

và negative examples:

```text
Tôi đang nghỉ việc.
BIDV
OpenAI
Microsoft Word
dao
Việt Nam
```

Mỗi lần export model mới chạy regression.

---

# 47. MODEL SELECTION

Không chọn model có accuracy cao nhất một cách máy móc.

Tạo Pareto table:

```text
Model
Parameters
FP32 size
INT8 size
Top1
F1
FPR
real-word accuracy
P50 latency
P95 latency
RAM
```

Ví dụ:

```text
Tiny-A
Tiny-B
Tiny-C
```

Chọn model có tradeoff tốt nhất.

Mặc định ưu tiên Tiny-B nếu:

```text
accuracy gần Tiny-C
nhưng nhanh và nhỏ hơn đáng kể.
```

---

# 48. ACCEPTANCE GATE

Model production phải vượt tất cả:

## Size

```text
INT8 <= 25 MB preferred
```

## Offline

Không dependency Internet.

## CPU

Không GPU bắt buộc.

## Latency

```text
P95 <= 80 ms
```

mục tiêu.

## Quality

False positive thấp.

## ONNX

PyTorch/ONNX parity pass.

## Tokenizer

Python/.NET parity pass.

## Stability

1.000+ inference liên tục không leak đáng kể.

---

# 49. OUTPUT ARTIFACTS

Khi hoàn thành phải có:

```text
Models/
│
├── vietnamese-context-ranker-v1-int8.onnx
├── config.json
├── model_manifest.json
└── tokenizer/
```

`model_manifest.json`:

```json
{
  "modelVersion": "1.0.0",
  "architecture": "TinyVietnameseContextRanker",
  "quantization": "INT8",
  "maxLength": 64,
  "maxCandidates": 8,
  "tokenizerVersion": "...",
  "trainingRun": "...",
  "modelSha256": "..."
}
```

---

# 50. DOCUMENTATION

Tạo:

```text
MODEL_CARD.md
```

bao gồm:

- model architecture
- parameter count
- tokenizer
- intended use
- out-of-scope use
- training data
- licenses
- metrics
- false-positive metrics
- latency
- RAM
- limitations

Tạo:

```text
MODEL_BENCHMARK.md
```

so sánh tất cả candidate models.

Tạo:

```text
MODEL_TRAINING.md
```

để reproduce training.

---

# 51. KHÔNG TRAIN MÙ QUÁ LÂU

Trước full training:

chạy smoke experiment trên subset.

Ví dụ:

```text
1–5% data
```

xác minh:

```text
loss decreases
ranking works
pipeline works
ONNX export works
.NET inference works
```

Chỉ sau đó mới full training.

---

# 52. DEVELOPMENT ORDER

Thực hiện theo thứ tự:

```text
Phase 1
Research + licensing

Phase 2
Dataset pipeline

Phase 3
Candidate generator training format

Phase 4
Tokenizer experiments

Phase 5
Tiny-A quick baseline

Phase 6
Teacher

Phase 7
Tiny-A/B/C experiments

Phase 8
Knowledge distillation

Phase 9
Gold evaluation

Phase 10
ONNX export

Phase 11
INT8

Phase 12
.NET benchmark

Phase 13
Calibration

Phase 14
Select final model
```

Không bắt đầu full-scale training trước khi Phase 1–5 hoạt động.

---

# 53. CUỐI CÙNG HÃY ĐƯA RA KẾT LUẬN RÕ RÀNG

Báo cáo theo format:

```text
# Final Model Decision

## Selected Architecture

## Why This Architecture

## Parameter Count

## Tokenizer

## Training Corpus

## Gold Corpus

## Teacher

## Distillation Strategy

## Accuracy

## False Positive Rate

## Real-word Error Accuracy

## FP32 Size

## INT8 Size

## Python Latency

## .NET Latency

## RAM

## ONNX Settings

## Quantization Settings

## Known Weaknesses

## Data License Review

## Production Recommendation
```

Không nói đơn giản:

```text
model works well
```

Phải có benchmark định lượng.

---

# 54. NGUYÊN TẮC QUAN TRỌNG NHẤT

Mục tiêu không phải tạo model tiếng Việt tổng quát tốt nhất.

Mục tiêu là:

> tạo một contextual ranker rất nhỏ, rất nhanh và có precision cao cho spell checker realtime.

Do đó:

```text
narrow task
> generic NLP capability

precision
> recall

low false positives
> aggressive correction

CPU latency
> benchmark GPU

candidate ranking
> text generation

10–20M specialized model
> 100M+ general model
```

Không tăng kích thước model nếu Candidate Generator hoặc training data có thể giải quyết vấn đề tốt hơn.

Model chỉ là một tầng trong:

```text
Rule Engine
↓
Candidate Generator
↓
Tiny Context Ranker
↓
Confidence Filter
```

chứ không phải toàn bộ spell-checking system.