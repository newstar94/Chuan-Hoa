# NHIỆM VỤ: Tích hợp bộ sửa lỗi chính tả tiếng Việt AI chạy local/offline vào Add-in hiện có

Bạn là Senior Software Architect + Senior .NET/Windows Developer + ML/ONNX Engineer.

Hãy nghiên cứu toàn bộ source code Add-in hiện tại trước khi thực hiện bất kỳ thay đổi nào.

Mục tiêu là tích hợp một hệ thống kiểm tra và sửa lỗi chính tả tiếng Việt mạnh, chạy hoàn toàn trên máy người dùng, không phụ thuộc Internet, sử dụng kiến trúc Hybrid:

- Rule Engine
- Dictionary
- Candidate Generator
- Context Ranking
- Tiny AI Model chạy bằng ONNX
- Personal Dictionary
- Local Feedback Learning

Yêu cầu quan trọng nhất:

> Không được phá vỡ, thay đổi hành vi hoặc làm chậm các chức năng hiện tại của Add-in.

Không được refactor lớn nếu không thực sự cần thiết.

Không được tự ý xóa hoặc thay thế các chức năng đang hoạt động.

Mọi thay đổi phải ưu tiên backward compatibility.

---

# 1. BƯỚC ĐẦU TIÊN: AUDIT SOURCE CODE

Trước khi code, hãy đọc kỹ toàn bộ project.

Phân tích tối thiểu:

- solution/project structure
- công nghệ Add-in đang sử dụng
- .NET version
- VSTO / COM / Office.js nếu có
- Word Object Model đang được sử dụng
- Ribbon
- TaskPane
- event handling
- document change detection
- text selection
- range manipulation
- highlighting/underline
- popup/suggestion UI
- threading
- async code
- timers
- debounce
- caching
- configuration
- logging
- installer/update mechanism
- dependencies
- database/local storage
- existing spell checking logic nếu có

Sau khi audit, tạo tài liệu:

`VIETNAMESE_SPELL_INTEGRATION_ANALYSIS.md`

bao gồm:

1. Kiến trúc hiện tại
2. Luồng xử lý hiện tại
3. Những thành phần có thể tái sử dụng
4. Những thành phần không nên thay đổi
5. Rủi ro khi tích hợp AI
6. Performance bottleneck hiện tại nếu có
7. Memory risk
8. Threading risk
9. Word COM risk
10. Crash risk
11. Kế hoạch tích hợp đề xuất

Không bắt đầu thay đổi lớn trước khi hiểu luồng hiện tại.

---

# 2. MỤC TIÊU KIẾN TRÚC

Không chạy AI trực tiếp trong `WINWORD.EXE` nếu kiến trúc Add-in hiện tại cho phép tách process.

Kiến trúc mục tiêu:

```text
Microsoft Word
      │
      │
      ▼
Existing Add-in
      │
      ├── Document/Range watcher
      ├── Debounce
      ├── Error rendering
      ├── Suggestion UI
      └── IPC Client
              │
              │ Named Pipe / local IPC
              ▼
      VietnameseEngine.exe
              │
              ├── Text Normalizer
              ├── Tokenizer
              ├── Rule Engine
              ├── Dictionary
              ├── Candidate Generator
              ├── Context Analyzer
              ├── Tiny ONNX Model
              ├── Ranking Engine
              ├── Cache
              └── Personal Dictionary
```

Không được dùng HTTP local nếu Named Pipe đơn giản và phù hợp hơn.

Không được yêu cầu Internet.

Không gọi API cloud.

Không gửi nội dung tài liệu ra khỏi máy.

---

# 3. PROCESS ISOLATION

Ưu tiên tạo process:

`VietnameseEngine.exe`

AI inference không chạy trực tiếp trong Word process.

Mục tiêu:

```text
WINWORD.EXE
    │
    │ Named Pipe
    ▼
VietnameseEngine.exe
```

Lợi ích bắt buộc phải đạt được:

- AI crash không làm Word crash
- memory leak engine có thể được xử lý bằng restart engine
- có thể nâng cấp model độc lập
- có thể tái sử dụng engine cho Outlook hoặc ứng dụng khác sau này
- giảm memory pressure lên WINWORD.EXE

Nếu kiến trúc hiện tại khiến việc tách process gây rủi ro lớn, hãy giải thích rõ và đề xuất phương án tương đương an toàn hơn.

---

# 4. ENGINE PHẢI HOẠT ĐỘNG OFFLINE 100%

Các chức năng chính tả không được phụ thuộc:

- Internet
- OpenAI
- Azure
- Gemini
- Claude
- cloud API
- remote server

Engine phải hoạt động ngay cả khi:

```text
Network = unavailable
```

Không có GPU vẫn phải hoạt động.

CPU là execution target mặc định.

---

# 5. PIPELINE XỬ LÝ

Thiết kế pipeline:

```text
Input text
   │
   ▼
Unicode normalization
   │
   ▼
Vietnamese tokenizer
   │
   ▼
Fast Rule Engine
   │
   ▼
Dictionary validation
   │
   ▼
Candidate Generator
   │
   ├── high confidence
   │       ↓
   │   return suggestion
   │
   └── ambiguous
           ↓
       Context AI
           ↓
       Candidate Ranking
           ↓
       confidence filter
           ↓
       result
```

AI không được xử lý mọi từ.

AI chỉ chạy khi cần.

---

# 6. NORMALIZER

Xây dựng Vietnamese Text Normalizer.

Phải hỗ trợ:

- Unicode NFC normalization
- ký tự Unicode tổ hợp
- dấu thanh tiếng Việt
- whitespace
- duplicated whitespace
- punctuation spacing
- malformed Unicode nếu có thể
- common typing artifacts

Không được tự động thay đổi document ở bước normalization.

Normalizer chỉ dùng để phân tích.

---

# 7. VIETNAMESE TOKENIZER

Tokenizer phải:

- hiểu từ tiếng Việt
- giữ đúng position/index trong Word Range
- xử lý punctuation
- xử lý number
- URL
- email
- file path
- acronym
- domain
- hashtag nếu có
- mixed Vietnamese/English text
- company/product names

Kết quả phải giữ được mapping:

```text
token
startOffset
length
sentenceIndex
paragraphIndex
```

Đây là yêu cầu bắt buộc để không sửa nhầm Word Range.

---

# 8. RULE ENGINE

Rule Engine phải cực nhanh.

Không gọi AI cho các lỗi có thể xác định bằng rule.

Bao gồm ít nhất:

## Telex

Ví dụ các pattern liên quan:

```text
aw
aa
dd
ee
oo
ow
uw
w
s
f
r
x
j
```

Nhưng tuyệt đối không được áp dụng naïve replacement làm hỏng tiếng Anh hoặc tên riêng.

Phải sử dụng Vietnamese syllable constraints.

## VNI

Hỗ trợ các lỗi gõ VNI phổ biến nếu có thể xác định chắc chắn.

## Common spelling confusion

Thiết kế confusion sets:

```text
s / x
ch / tr
d / gi / r
l / n
c / t
n / ng
hỏi / ngã
i / y
```

Không tự động sửa chỉ dựa trên confusion pair.

Chỉ tạo candidate.

---

# 9. DICTIONARY

Tạo dictionary layer bao gồm:

```text
Core Vietnamese Dictionary
+
User Dictionary
+
Organization Dictionary
+
Document Ignore List
```

Dictionary phải lookup rất nhanh.

Không parse file JSON khổng lồ mỗi lần lookup.

Có thể nghiên cứu:

- compact binary format
- hash table
- trie
- DAWG
- memory mapped file

Chọn giải pháp phù hợp nhất sau benchmark.

Không over-engineer nếu dictionary hiện tại đã đủ nhanh.

---

# 10. CANDIDATE GENERATOR

Candidate Generator phải tạo các từ thay thế có khả năng đúng dựa trên:

- edit distance
- Vietnamese syllable structure
- Telex errors
- VNI errors
- missing character
- extra character
- substituted character
- transposition
- keyboard proximity
- diacritics
- confusion sets
- dictionary frequency

Không generate hàng nghìn candidate.

Giới hạn candidate trước AI.

Ví dụ:

```text
ngỉ
```

có thể sinh:

```text
nghỉ
nghĩ
nghị
```

nhưng không được đưa hàng trăm từ không liên quan.

Target:

```text
Top candidate count <= 8
```

trừ khi benchmark chứng minh cần nhiều hơn.

---

# 11. CONTEXTUAL ERROR DETECTION

Phải hỗ trợ lỗi "real-word error".

Ví dụ:

```text
bàn dao hồ sơ
```

Từng từ đều hợp lệ nhưng context phải gợi ý:

```text
bàn giao hồ sơ
```

Ví dụ:

```text
điều khoảng hợp đồng
```

gợi ý:

```text
điều khoản hợp đồng
```

Ví dụ:

```text
Tôi muốn ngỉ việc.
```

gợi ý:

```text
nghỉ
```

Không dựa duy nhất vào dictionary.

---

# 12. TINY AI MODEL

Thiết kế hệ thống để sử dụng một model ONNX nhỏ.

Không bắt đầu bằng model 200M–1B parameter.

Target ban đầu:

```text
Parameters:
~10M – 30M

Quantization:
INT8

ONNX size target:
~15 – 50 MB
```

Model ưu tiên nhiệm vụ:

```text
context ranking
```

hơn là:

```text
full sentence rewriting
```

Model không được tự ý viết lại toàn bộ câu.

Input có thể theo dạng:

```text
context
+
error position
+
candidate list
```

Output:

```text
candidate probabilities
```

Ví dụ:

```text
Context:
"Tôi muốn ___ việc vào cuối tháng."

Candidates:
nghỉ
nghĩ
nghị
```

Output:

```text
nghỉ 0.982
nghĩ 0.014
nghị 0.004
```

---

# 13. ONNX INFERENCE

Dùng ONNX Runtime hoặc giải pháp Windows local phù hợp sau khi benchmark.

Ưu tiên:

```text
CPU
INT8
```

Không bắt GPU.

GPU acceleration chỉ là optional optimization.

Thiết lập inference session chỉ một lần.

Không recreate session cho từng câu.

Không reload model cho từng request.

Phải kiểm soát:

- intra-op threads
- inter-op threads
- memory allocation
- session lifetime
- model warm-up

Benchmark trước khi thay đổi thread settings.

---

# 14. LAZY LOAD

Không load AI model ngay khi Word mở.

Luồng mong muốn:

```text
Word start
   ↓
Add-in load
   ↓
AI model NOT loaded
```

Khi lần đầu cần AI:

```text
start VietnameseEngine.exe
↓
load model
↓
warm-up
↓
process request
```

Nếu engine không hoạt động trong khoảng thời gian dài:

có thể cân nhắc:

```text
unload model
```

hoặc:

```text
shutdown process
```

nhưng chỉ thực hiện nếu benchmark cho thấy có lợi.

Không làm engine restart liên tục.

---

# 15. IPC

Ưu tiên:

```text
Named Pipe
```

Thiết kế protocol rõ ràng.

Ví dụ request:

```json
{
  "requestId": "uuid",
  "documentId": "hash",
  "paragraphId": "183",
  "text": "Tôi muốn ngỉ việc vào cuối tháng.",
  "language": "vi-VN"
}
```

Response:

```json
{
  "requestId": "uuid",
  "issues": [
    {
      "start": 9,
      "length": 3,
      "type": "spelling",
      "original": "ngỉ",
      "confidence": 0.982,
      "suggestions": [
        {
          "text": "nghỉ",
          "score": 0.982
        }
      ]
    }
  ]
}
```

Có:

- timeout
- cancellation
- disconnect handling
- process restart
- invalid response handling
- protocol version

Word không được freeze nếu engine không trả lời.

---

# 16. KHÔNG BLOCK WORD UI THREAD

Đây là yêu cầu bắt buộc.

Không được:

```text
Wait()
.Result
Thread.Sleep()
```

trên UI thread.

Không chạy inference trên Word UI thread.

Không giữ COM Range qua background thread nếu không an toàn.

Word Object Model phải được xử lý theo threading constraints phù hợp.

Background worker chỉ xử lý plain text/data.

Sau khi AI trả kết quả mới quay lại Word UI/COM context để render suggestion.

---

# 17. INCREMENTAL ANALYSIS

Tuyệt đối không phân tích lại toàn document mỗi lần người dùng gõ.

Ví dụ document:

```text
2.000 paragraphs
```

người dùng sửa paragraph 183.

Chỉ phân tích:

```text
paragraph 183
```

hoặc context tối thiểu cần thiết:

```text
paragraph 182
paragraph 183
paragraph 184
```

nếu context model cần.

Ưu tiên sentence-level analysis nếu đủ.

---

# 18. DEBOUNCE

Không inference sau mỗi keystroke.

Thiết kế debounce khoảng:

```text
300–500 ms
```

nhưng benchmark để chọn giá trị phù hợp.

Có thể trigger nhanh hơn khi gặp:

```text
Space
.
,
;
:
?
!
Enter
```

Cancel request cũ nếu người dùng tiếp tục gõ.

Ví dụ:

```text
Request A
↓
user continues typing
↓
cancel A
↓
Request B
```

Không để queue inference tồn đọng.

---

# 19. CACHE

Xây cache theo nội dung.

Ví dụ:

```text
SHA256(normalized paragraph)
```

hoặc hash nhanh phù hợp hơn.

Nếu paragraph không thay đổi:

```text
cache hit
```

không chạy AI lại.

Cache cần có:

- size limit
- LRU hoặc equivalent
- invalidation
- document-specific state nếu cần

Không tạo memory leak.

---

# 20. ERROR LEVELS

Phân loại:

```text
LEVEL 0
Unicode / spacing
Rule only

LEVEL 1
Obvious spelling error
Dictionary + Candidate

LEVEL 2
Ambiguous spelling
Candidate + Context AI

LEVEL 3
Real-word contextual error
Context AI
```

Mục tiêu:

AI chỉ chạy cho Level 2 và Level 3.

---

# 21. CONFIDENCE

Không gạch chân khi confidence quá thấp.

Thiết kế threshold.

Ví dụ:

```text
>= 0.90
Strong suggestion

0.75–0.90
Suggestion

< 0.75
Do not show
```

Đây chỉ là giá trị khởi đầu.

Sau này phải calibration bằng validation dataset.

Ưu tiên:

```text
precision > recall
```

Không được spam người dùng bằng false positive.

---

# 22. PERSONAL DICTIONARY

Lưu local, ví dụ:

```text
%LOCALAPPDATA%\VietnameseAssistant\
```

Có thể sử dụng SQLite hoặc format nhẹ hơn.

Hỗ trợ:

```text
Add to dictionary
Ignore once
Ignore in document
Always ignore
```

Ví dụ:

```text
Viettel
VNPT
BIDV
Techcombank
OpenAI
ChatGPT
```

phải có thể được user whitelist.

---

# 23. LOCAL FEEDBACK

Lưu local feedback:

```text
original
suggestion
accepted
rejected
timestamp
optional_context_hash
```

Không lưu full document nếu không cần.

Không gửi telemetry nội dung ra ngoài.

Ranking có thể sử dụng feedback để điều chỉnh suggestion cho user.

Ví dụ:

```text
finalScore =
baseScore
+ contextScore
+ userPreferenceScore
```

Không tự động retrain neural model trong MVP.

---

# 24. PRIVACY

Mặc định:

```text
NO NETWORK REQUEST
```

Không gửi:

- document
- paragraph
- sentence
- words
- filename
- user text

ra Internet.

Nếu Add-in hiện tại có telemetry, tuyệt đối không thêm document content vào telemetry.

Log không được chứa nguyên văn nội dung nhạy cảm mặc định.

---

# 25. UI

Tái sử dụng UI hiện tại nếu có thể.

Không thay giao diện lớn nếu không cần.

Các chức năng cần hỗ trợ:

```text
underline issue
show suggestion
replace
ignore
add to dictionary
```

Nếu có nhiều suggestion:

```text
nghỉ
nghĩ
nghị
```

hiển thị theo score.

Không tự động sửa document nếu user chưa bật chế độ auto-correct rõ ràng.

---

# 26. WORD RANGE SAFETY

Đây là phần rất quan trọng.

Không lưu Word Range lâu dài nếu document có thể thay đổi.

Prefer:

```text
document identifier
paragraph identifier
text hash
relative offset
```

Trước khi replace:

1. kiểm tra text hiện tại
2. kiểm tra original substring vẫn giống error
3. nếu document đã thay đổi, invalidate suggestion
4. không replace offset cũ mù quáng

Ví dụ:

Suggestion sinh ra cho:

```text
ngỉ
```

nhưng user đã sửa thành:

```text
nghỉ
```

thì suggestion cũ phải bị bỏ.

---

# 27. ENGINE LIFECYCLE

Add-in phải có EngineManager.

Trạng thái:

```text
Stopped
Starting
Ready
Busy
Failed
Restarting
```

Nếu engine crash:

```text
detect
↓
restart
↓
continue
```

Không làm Word crash.

Có giới hạn restart để tránh loop.

Ví dụ:

```text
max 3 restart / 5 minutes
```

sau đó disable AI tạm thời và chỉ sử dụng Rule Engine.

---

# 28. FALLBACK

Nếu AI engine lỗi:

Add-in vẫn phải hoạt động.

Fallback:

```text
Rule Engine
+
Dictionary
+
Candidate Generator
```

AI failure không được làm mất toàn bộ chức năng spell checking.

---

# 29. PERFORMANCE KPI

Đặt target:

```text
Rule processing:
< 5 ms

Candidate generation:
< 10 ms

AI inference:
< 30–80 ms

Typical total:
< 100 ms
```

Đây là target, không hard-code giả định.

Phải benchmark trên máy CPU thông thường.

Test ít nhất:

```text
low-end office CPU
mid-range CPU
modern laptop CPU
```

nếu môi trường test cho phép.

---

# 30. MEMORY KPI

Target tham khảo:

```text
Word Add-in overhead:
< 10 MB

Engine idle:
< 20–30 MB

Dictionary:
< 30 MB

ONNX:
15–50 MB

Total active:
preferably < 120 MB
```

Nếu vượt target phải profile và giải thích.

Không tối ưu mù quáng.

---

# 31. MODEL KHÔNG ĐƯỢC HALLUCINATE

Không dùng generative rewriting cho spell checking realtime.

Không cho model tự sinh câu mới tùy ý.

Model chỉ nên:

```text
Detect
Rank
Score
```

candidate.

Việc replace text phải dựa trên candidate rõ ràng.

---

# 32. TEST DATA

Tạo test suite cho ít nhất các nhóm lỗi sau:

```text
Telex
VNI
missing tone
wrong tone
s/x
ch/tr
d/gi/r
l/n
hỏi/ngã
missing character
extra character
transposition
keyboard neighbor
word merge
word split
mixed Vietnamese-English
proper noun
company name
URL
email
number
abbreviation
real-word error
punctuation
Unicode
```

Ví dụ bắt buộc:

```text
ngỉ → nghỉ

Tôi muốn ngỉ việc.
→ nghỉ

bàn dao hồ sơ
→ bàn giao hồ sơ

điều khoảng hợp đồng
→ điều khoản hợp đồng
```

Đồng thời tạo negative test để tránh sửa sai:

```text
dao
Việt
BIDV
OpenAI
Visual Studio
Windows
.NET
GPT
```

---

# 33. FALSE POSITIVE TEST

Đây phải là KPI quan trọng.

Tạo corpus câu đúng.

Engine không được cố tìm lỗi ở mọi câu.

Measure:

```text
Precision
Recall
F1
False Positive Rate
Top-1 Accuracy
Top-3 Accuracy
Latency
Memory
```

Ưu tiên giảm False Positive.

---

# 34. UNIT TEST

Viết unit test cho:

```text
Normalizer
Tokenizer
Telex parser
VNI parser
Dictionary
Candidate Generator
Ranking
Cache
IPC serialization
IPC timeout
Engine restart
Personal Dictionary
```

---

# 35. INTEGRATION TEST

Test:

```text
Word
↓
Add-in
↓
IPC
↓
Engine
↓
Inference
↓
Result
↓
Word UI
```

Kiểm tra:

- typing nhanh
- document dài
- copy/paste lớn
- undo
- redo
- replace
- track changes nếu Add-in hỗ trợ
- document close
- Word exit
- multiple documents
- multiple Word windows
- engine crash

---

# 36. STRESS TEST

Test document:

```text
100 pages
300 pages
1000+ paragraphs
```

Không được scan lại full document theo mỗi keystroke.

Theo dõi:

```text
CPU
RAM
GC
WINWORD CPU
WINWORD RAM
engine RAM
IPC queue
latency
```

---

# 37. LOGGING

Log technical event:

```text
engine_start
engine_stop
engine_crash
model_load_time
inference_time
cache_hit
cache_miss
ipc_timeout
```

Không log document text mặc định.

Có log level:

```text
Error
Warning
Info
Debug
```

Debug sensitive data phải disabled trong production.

---

# 38. CONFIGURATION

Có config nội bộ cho:

```text
SpellCheckEnabled
AiEnabled
DebounceMs
ConfidenceThreshold
MaxCandidates
CacheSize
EngineIdleTimeout
```

Không hard-code các giá trị ở nhiều nơi.

Centralized configuration.

---

# 39. MODEL VERSIONING

Model phải có version.

Ví dụ:

```text
spell-context-v1
```

IPC response có:

```text
engineVersion
modelVersion
```

Cho phép update model sau này mà không ảnh hưởng Add-in.

---

# 40. MODEL UPDATE

MVP không bắt buộc auto-update.

Nhưng kiến trúc phải cho phép:

```text
Models/
    spell-context-v1.onnx
```

sau này thay bằng:

```text
spell-context-v2.onnx
```

không cần rewrite toàn engine.

---

# 41. SECURITY

Kiểm tra:

- Named Pipe permissions
- executable path validation
- DLL hijacking risk
- model tampering
- path traversal
- malformed IPC message
- oversized input
- resource exhaustion
- process spoofing

Named Pipe không được cho process không liên quan tùy ý gửi lệnh nếu có thể hạn chế quyền.

---

# 42. INPUT LIMITS

Không cho engine nhận input vô hạn.

Ví dụ:

```text
Max paragraph chars
Max candidates
Max message size
```

Nếu clipboard paste hàng trăm nghìn ký tự:

chia chunk hợp lý.

Không OOM.

---

# 43. CODE QUALITY

Tuân thủ architecture hiện tại nếu hợp lý.

Ưu tiên:

```text
SOLID
small interfaces
dependency injection khi thực sự có lợi
async/cancellation
clear ownership
testability
```

Không thêm framework lớn chỉ vì tiện.

Không over-engineer.

---

# 44. KHÔNG THAY ĐỔI KHÔNG CẦN THIẾT

Không:

- rename hàng loạt
- format toàn repository
- đổi namespace hàng loạt
- upgrade dependency không liên quan
- đổi framework không cần thiết
- refactor toàn bộ Add-in
- thay UI hoàn toàn
- xóa legacy code chưa chứng minh là unused

Mục tiêu là integration an toàn.

---

# 45. IMPLEMENT THEO PHASE

Thực hiện từng phase.

## Phase 1 — Audit

Không sửa behavior.

Output:

```text
VIETNAMESE_SPELL_INTEGRATION_ANALYSIS.md
```

---

## Phase 2 — Core Rule Engine

Implement:

```text
Normalizer
Tokenizer
Dictionary
Candidate Generator
```

Viết test.

Chưa cần ONNX.

---

## Phase 3 — Local Engine Process

Implement:

```text
VietnameseEngine.exe
Named Pipe Server
EngineManager
Protocol
```

Word Add-in kết nối tới engine.

---

## Phase 4 — Add-in Integration

Implement:

```text
Document watcher
Debounce
Incremental analysis
Cancellation
Suggestion rendering
```

Không ảnh hưởng chức năng cũ.

---

## Phase 5 — ONNX

Thêm:

```text
ContextModel
CandidateRanker
INT8 inference
```

AI chỉ dùng khi cần.

---

## Phase 6 — Cache

Implement:

```text
paragraph/sentence hash
LRU cache
invalidation
```

---

## Phase 7 — Personal Dictionary

Implement:

```text
Add
Remove
Ignore
Persist
```

---

## Phase 8 — Feedback

Implement local:

```text
accepted
rejected
user ranking preference
```

---

## Phase 9 — Optimization

Profile trước.

Sau đó tối ưu:

```text
CPU
allocation
COM calls
IPC
model inference
dictionary lookup
cache
```

Không optimize dựa trên suy đoán.

---

# 46. TRƯỚC MỖI PHASE

Trước khi sửa code:

1. xác định file cần sửa
2. giải thích tại sao
3. xác định regression risk
4. xác định test cần chạy
5. chỉ sau đó mới sửa

---

# 47. SAU MỖI PHASE

Chạy:

```text
build
unit tests
integration tests
```

Nếu repository có test hiện tại:

tất cả test cũ phải tiếp tục pass.

Không được bỏ test để build xanh.

Không disable warning/error chỉ để vượt build.

---

# 48. BACKWARD COMPATIBILITY

Đảm bảo:

```text
existing Add-in feature
+
new spell checker
```

cùng tồn tại.

Nếu spell checker disabled:

Add-in phải hoạt động gần giống trước integration.

---

# 49. FEATURE FLAG

Tạo feature flag:

```text
VietnameseSpellCheckerEnabled
```

Có thể disable toàn bộ module mới.

Nếu xảy ra vấn đề production, module có thể tắt mà không ảnh hưởng chức năng khác.

---

# 50. FINAL DOCUMENTATION

Sau khi hoàn thành tạo:

```text
VIETNAMESE_SPELL_ARCHITECTURE.md
```

Bao gồm:

- architecture
- components
- IPC
- data flow
- cache
- model
- lifecycle
- privacy
- failure handling

Tạo:

```text
VIETNAMESE_SPELL_PERFORMANCE.md
```

Bao gồm:

- startup impact
- engine startup
- model load
- inference latency
- RAM
- CPU
- cache hit rate

Tạo:

```text
VIETNAMESE_SPELL_TEST_REPORT.md
```

Bao gồm:

- test cases
- results
- regression test
- known limitations

---

# 51. FINAL REVIEW

Sau khi code xong, tự review toàn bộ diff.

Đặc biệt tìm:

```text
UI thread blocking
COM misuse
race conditions
deadlock
memory leaks
unbounded cache
unbounded queue
stale Word Range
wrong offsets
engine orphan process
process leak
Named Pipe leak
exception swallowing
invalid cancellation
duplicate checking
excessive inference
```

Fix trước khi kết thúc.

---

# 52. ACCEPTANCE CRITERIA

Chỉ coi nhiệm vụ hoàn thành nếu đáp ứng:

### Offline

```text
Internet disconnected
→ spell checker still works
```

### Word performance

Word không có cảm giác lag đáng kể trong quá trình gõ bình thường.

### Incremental

Một thay đổi nhỏ không kích hoạt scan toàn document.

### AI isolation

AI failure không làm Word crash.

### Safety

Không tự ý sửa nội dung.

### Privacy

Không gửi document content ra Internet.

### Compatibility

Các tính năng Add-in hiện có vẫn hoạt động.

### Quality

Có khả năng xử lý:

```text
ngỉ → nghỉ
bàn dao hồ sơ → bàn giao hồ sơ
điều khoảng hợp đồng → điều khoản hợp đồng
```

mà không chỉ dựa vào dictionary.

### Low false positive

Không báo lỗi liên tục với:

```text
tên riêng
tên công ty
English terms
URL
email
technical terms
```

---

# 53. QUY TẮC QUAN TRỌNG NHẤT

Trong toàn bộ quá trình:

> Đọc code hiện tại trước, tích hợp vào kiến trúc hiện có thay vì áp đặt một kiến trúc mới.

> Không hy sinh độ ổn định của Word để lấy thêm vài phần trăm accuracy.

> Không chạy AI nếu Rule Engine có thể xử lý.

> Không gửi dữ liệu ra Internet.

> Không block UI thread.

> Không scan toàn tài liệu sau mỗi keystroke.

> Không load model khi chưa cần.

> Không tự động sửa nội dung nếu người dùng chưa xác nhận.

> Không refactor những phần không liên quan.

> Mọi optimization phải được benchmark.

---

# 54. CÁCH CODEx PHẢI BÁO CÁO KHI HOÀN THÀNH

Cuối cùng hãy cung cấp summary theo format:

```text
## Architecture

## Files Added

## Files Modified

## Existing Features Affected

## Backward Compatibility

## Spell Engine Pipeline

## AI Model Integration

## IPC

## Performance Before

## Performance After

## Memory Usage

## Tests Added

## Existing Tests

## Security Review

## Privacy Review

## Known Limitations

## Recommended Next Phase
```

Đặc biệt liệt kê rõ từng file đã sửa và lý do sửa.

Nếu có vấn đề chưa thể giải quyết an toàn, không hack workaround.

Hãy ghi rõ:

```text
Problem
Root cause
Risk
Recommended solution
```

rồi chọn phương án ít ảnh hưởng nhất tới hệ thống hiện tại.