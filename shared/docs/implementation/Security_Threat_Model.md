# Security Threat Model

## Tài sản cần bảo vệ

- Quyền sử dụng, trial, offer, quote, payment và entitlement history.
- Lease/FixPlan/rule signing keys và code-signing certificate.
- Canonical rules, signed rule packages và release metadata.
- Device private key, refresh-token family và admin session.
- Nội dung/snapshot tài liệu chỉ trong bộ nhớ local khi xử lý.
- Chuỗi cung ứng DLL, manifests, installer và Portal artefact.

## Trust boundaries

1. Word/VSTO process trên endpoint không tin cậy hoàn toàn.
2. System browser và Identity Provider.
3. API Gateway/BFF chỉ xử lý identity, entitlement, commercial và artefact metadata.
4. PostgreSQL, queue và KMS/HSM.
5. CI signing pipeline và release distribution.
6. Admin Portal, Customer Portal và managed deployment connectors.

## Threats và controls bắt buộc

| Threat | Mức | Controls | Residual status |
| --- | --- | --- | --- |
| Sửa DLL/manifest | P0 | Authenticode digest validation + signer pin cho outer/owned PE; chữ ký XML và dependency SHA-256 cho VSTO manifest; timestamp; immutable artefact | Development `.110` tamper-negative 7/7 PASS trên bản sao Temp; chưa có production certificate/timestamp |
| Clone/re-signed client dùng account/device hợp lệ | P0 | ADR CLIENT_AUTHENTICITY_BOUNDARY; WDAC/attestation/helper hoặc accepted limitation | BLOCKED_DECISION |
| Copy token/lease sang máy khác | P0 | Device asymmetric key, challenge-response, token rotation, replay cache | Chưa implement |
| Sửa/giả rule pack hoặc kéo dài lease | P0 | RS256, key id/kind binding, device/client/feature binding, hard cap 7 ngày, trusted server time và atomic cache | Development E2E PASS; production KMS/key rotation chưa có |
| Replay ExecutionGrant/FixPlan | P0 | Short expiry, nonce, one-time jti, user/device/document/command binding | Chưa implement |
| Reset trial bằng reinstall/clock | P1 | Server time, immutable grant history, account/device binding | Chưa implement |
| Webhook giả/duplicate/out-of-order | P0 | Provider signature, quote binding, idempotency, outbox, reconciliation | Provider chưa chọn |
| Word mutation làm mất dữ liệu | P0 | Operation allowlist, phase/preflight, fingerprint, cancel trước mutation, recovery copy cho lệnh sửa hàng loạt, exact finding và annotation rollback | Development Word 16 x64 DOC/DOCX/Document1 đã PASS; Word 2010/x86, golden corpus và pilot kéo dài chưa chạy |
| Cross-tenant/BOLA | P0 | Principal-derived organization scope, query/cache/export/job isolation, negative tests | Chưa implement |
| Admin tự duyệt | P0 | Maker-checker, payload hash, ETag, short session, MFA, immutable audit | IdP/role holders chưa chốt |
| Document leakage | P0 | Không có document upload endpoint trong luồng VSTO; snapshot chỉ local; telemetry cấm text/path/name; network seam và payload có source regression gate | `DOCUMENT-PRIVACY-REGRESSION-001` PASS với 4/4 negative self-test; production observability review/pentest vẫn chưa chạy |
| Supply-chain compromise | P0 | Locked dependencies, SBOM, CI provenance, vulnerability/secret scan, separated keys | NuGet locked restore, CycloneDX SBOM, point-in-time NuGet vulnerability audit và current-tree + 798 blob Git-history secret regression local PASS; local CI workflow config có action SHA pin. Remote CI run/branch protection, provenance, production signing và key separation chưa implement |

## Fail-closed rules

- Unknown schema/operation, invalid signature, expired/replayed grant, fingerprint mismatch, revoked release hoặc missing entitlement không được sửa tài liệu.
- Lỗi auth/server/integrity không được làm Word crash hoặc khóa tài liệu.
- Client báo hash/version/trial/payment không bao giờ là nguồn authoritative.
- Gói cache thiếu/sai chữ ký, hết hạn, sai device/version/feature hoặc đồng hồ quay lùi không được dùng.
- Mất mạng không vô hiệu hóa cache hợp lệ; hết lease 7 ngày thì feature có phí khóa cho đến khi gia hạn.
