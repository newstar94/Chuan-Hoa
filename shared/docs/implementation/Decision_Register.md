# Decision Register

| ID | Trạng thái | Quyết định hoặc dữ liệu cần chốt | Giá trị hiện hành | Owner | Ảnh hưởng |
| --- | --- | --- | --- | --- | --- |
| DEC-001 | LOCKED | Client sản phẩm | VSTO/C#, Word 2010+ trên Windows, x86/x64 theo lane | Product | Toàn kiến trúc client |
| DEC-002 | LOCKED | Bề mặt trong Word | Một tab Ribbon; không task pane thường trực | Product | Ribbon/UI |
| DEC-003 | LOCKED_UPDATED | Phân chia xử lý | Toàn bộ đọc/scan/detector/finding/mutation ở local; server chỉ phát hành signed rule pack, lease tối đa 7 ngày và quản lý thương mại | Product/Security | Privacy và chống crack |
| DEC-004 | LOCKED | Trial | Launch trial và personal trial không cộng dồn; người đã dùng launch trial không nhận personal trial | Product | Entitlement |
| DEC-005 | CORRECTED | Rule route baseline | 96 definitions, 94 routes, 19 hardwired not-checked, 75 logic paths, 52 JSON codes, 14 backend codes | Engineering | Canonical rules và marketing count |
| DEC-006 | ACCEPTED_TECHNICAL | Backend framework | ASP.NET Core trên .NET 10 LTS; metadata chính thức ngày 2026-09-01 ghi SDK 10.0.400, EOL 2028-11-14 | Engineering | Server build/runtime |
| DEC-007 | LOCKED | Binary rollback | Forward-fix bằng version cao hơn; không hạ ClickOnce version | Release | Update/incident |
| DEC-008 | LOCKED | Giá client | Không hardcode trong DLL; server chọn versioned offer | Commercial | Quote/payment |
| DEC-009 | BLOCKED_DECISION | Launch trial dates | Chưa có giá trị production | Product Owner | Không thể schedule campaign |
| DEC-010 | BLOCKED_DECISION | Personal trial duration | Chưa có giá trị production; 7 ngày chỉ là đề xuất | Product Owner | Trial policy production |
| DEC-011 | BLOCKED_DECISION | Device limits | Chưa có giá trị production | Product/Security | Device enrollment |
| DEC-012 | LOCKED_PARTIAL | Lease TTL và paid offline grace | Offline lease tối đa 7 ngày đã chốt; grace ngoài lease chưa có | Product/Security | Availability và revocation SLA |
| DEC-013 | BLOCKED_DECISION | Giá, currency, term, renewal, grandfathering | Chưa có giá trị production | Product/Finance | Offer publication |
| DEC-014 | BLOCKED_DECISION | Refund/chargeback policy và approval threshold | Chưa có giá trị production | Finance | Refund workflow |
| DEC-015 | BLOCKED_DECISION | Payment provider và merchant | Chưa chọn | Product/Finance/Security | Production payment adapter |
| DEC-016 | BLOCKED_DECISION | Customer identity provider | Chưa cấu hình tenant/provider production | Security | Login production |
| DEC-017 | BLOCKED_DECISION | Admin IdP, role holders và break-glass SLA | Chưa cấu hình | Security | Admin A production |
| DEC-018 | BLOCKED_DECISION | Code-signing certificate | Chưa có certificate production | Release/Security | VSTO/installer signing |
| DEC-019 | BLOCKED_DECISION | KMS/HSM keys | Chưa cấp lease/FixPlan/rule signing keys | Security | Server signatures |
| DEC-020 | BLOCKED_DECISION | Domain, deployment và callback URLs | Chưa có hostname production | Operations | OAuth/ClickOnce/Portal |
| DEC-021 | BLOCKED_DECISION | Manual entitlement approval thresholds | Chưa chốt | Product/Security | Maker-checker |
| DEC-022 | BLOCKED_DECISION | Organization/tenant, SSO và verified domain boundary | Chưa chốt hợp đồng production | Product/Security | Tenant API/Portal |
| DEC-023 | BLOCKED_DECISION | Windows 32-bit và legacy contractual scope | Chưa có danh sách khách hàng | Product/Support | Compatibility matrix |
| DEC-024 | LOCKED | Document privacy boundary | Không upload snapshot/nội dung tài liệu; xử lý document chỉ tại local | Product/Privacy | Compliance Engine deployment |
| DEC-025 | BLOCKED_DECISION | Legal sources cho HD05/Viettel và reviewer | Chưa có sign-off | Legal/Product | Canonical rule publication |
| DEC-026 | BLOCKED_DECISION | VM/Office licenses | Chưa chuẩn bị | QA/Operations | Compatibility evidence |
| DEC-027 | BLOCKED_DECISION | Portal brand/design/accessibility owner | Chưa phê duyệt | Product/Design | Admin F acceptance |
| DEC-028 | BLOCKED_DECISION | Signed updater riêng cho máy cá nhân | Chưa duyệt; không thuộc MVP mặc định | Product/Security | Zero-click consumer update |
| DEC-029 | BLOCKED_DECISION | Client authenticity boundary | Chưa chọn WDAC/managed endpoint, attestation/helper hoặc accepted consumer limitation | Security/Product | Claim chống client clone/re-sign |
| DEC-030 | BLOCKED_DECISION | Danh sách khách hàng cần on-prem engine | Chưa có | Sales/Product | Deployment topology |

Không mục BLOCKED_DECISION nào được thay bằng giá trị giả trong production config, test evidence hoặc tài liệu marketing.
