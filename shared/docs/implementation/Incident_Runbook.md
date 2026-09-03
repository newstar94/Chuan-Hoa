# Incident Runbook

## Phân loại

- P0: mất/corrupt tài liệu, signing key compromise, payment cấp quyền sai, cross-tenant leak, malicious release/rule.
- P1: widespread load/update failure, premium outage, trial/entitlement denial sai, severe performance regression.
- P2/P3: phạm vi nhỏ, có workaround và không vi phạm integrity/privacy.

## Quy trình

1. Mở incident với severity, owner, affected cohort, thời điểm và correlation/evidence.
2. Ngăn thiệt hại bằng control nhỏ nhất: pause rollout, disable command, revoke release, rollback rule mapping, disable provider hoặc fail closed premium.
3. Bảo toàn audit, logs đã redaction và artefact; không thu document content ngoài consent/contract.
4. Điều tra source revision, release/ruleset, device/client lane và dependency health.
5. Phục hồi bằng forward release, compatible rule mapping hoặc service fix.
6. Xác minh document safety, entitlement/payment consistency và tenant isolation.
7. Đóng incident với postmortem, corrective actions, owner và deadline nội bộ.

## Break-glass

- Chỉ Security Admin/Super Admin sau step-up MFA.
- Bắt buộc incident ID, reason, scope và TTL.
- Tạo alert và post-review.
- Bật lại global command/release qua approval bình thường.
- Không dùng break-glass để sửa quote/payment/trial history.

