# Rule Catalog Reconciliation

| Nguồn/khái niệm | Số đo |
| --- | ---: |
| VBA rule definitions trong RulesPart1–5 | 96 |
| Registered routes | 94 |
| Route trỏ implementation luôn trả `Nothing` | 19 |
| Route có đường logic | 75 |
| Mã từng được tài liệu/metadata tuyên bố | 82 |
| Mã duy nhất trong JSON prototype | 52 |
| Mã duy nhất backend prototype phát ra | 14 |

Hai definition chưa đăng ký route:

- `ND30-PL1-M1-K4-ENC`
- `ND30-PL1-M1-K4-NFC`

Ba implementation luôn trả `Nothing`:

- `CheckComponentUnderline`: 3 routes.
- `CheckComponentNeverDetected`: 5 routes.
- `CheckCapitalizationNotDetectable`: 11 routes.

Không con số nào ở trên được dùng như số rule thương mại. Canonical publication chỉ được phép khi declared, loaded, routed, implemented và reported khớp; mỗi active rule có source, effective period, positive/negative/boundary fixtures và risk/fix policy.

Evidence chi tiết: `evidence/rule_reconciliation.json`.

## Canonical draft hiện hành

- JSON Schema đóng: `shared/contracts/rules/canonical-rule-release.v1.schema.json`.
- Draft sinh máy: `shared/rules/canonical/baseline-draft.v1.json`.
- Parser/validator fail-closed: `src/ChuanHoa.Rules`.
- Evidence: `evidence/canonical_rule_baseline.json`.

Draft giữ đúng 96 definition với phân loại 75 `BaselineLogicPath`, 19 `HardwiredNotChecked` và 2 `Unrouted`. Toàn bộ rule vẫn là `Draft`, legal review `Unreviewed`, fixture trống và fix policy `Blocked`. Vì vậy draft hợp lệ về cấu trúc nhưng tuyệt đối không publishable và không được dùng làm số rule thương mại.

Publication validator chỉ cho rule release `Published` khi từng rule:

- `Active` và có implementation `ImplementedVerified`;
- có route, detector ID và engine version đã kiểm chứng;
- có authority, instrument, provision, reviewer và reviewed time;
- có đủ fixture positive, negative và boundary đã duyệt;
- có fix policy khác `Blocked`.

Corpus contract cũng từ chối tính fixture draft là golden evidence và yêu cầu đủ ba loại fixture đã approved cho từng rule.
