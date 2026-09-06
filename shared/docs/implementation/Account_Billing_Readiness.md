# Account and billing implementation status

## Implemented

Product decision: exactly one paid plan, alongside Free. Offers represent price
versions of that single paid plan, not different user-selectable packages. Checkout
must resolve the current published price server-side; no multi-plan selector.

- Free feature contract: FORMAT_SCAN, SPELLING_SCAN, TABLE_IMAGE_TOOLS only.
- Ribbon disabled-state checks and execution-time signed-lease checks.
- Table/image commands accept the narrow capability; other document tools do not.
- payOS SDK 2.1.0 server gateway; environment-only credentials and HTTPS return/cancel URLs.
- Authorized checkout endpoint and signed webhook verification.
- Amount, VND currency, provider link, reference and order-state matching tests.

## Not ready for production

Setting payOS environment values alone does not enable checkout. Production identity
authentication currently fails closed. IPurchaseStore intentionally has no registered
implementation: durable quote/order reservation, retry recovery, reference deduplication,
atomic entitlement activation and signed lease refresh must be implemented against the
existing PostgreSQL schema before accepting money. No in-memory payment store is used.

The checkout endpoint returns 503 while payment credentials or the durable store are
missing. It never accepts client-supplied prices. Return/cancel redirects cannot activate
an entitlement. A verified webhook is still required to match the saved order.

Account login/logout, user profile, device enrollment/revocation and the in-window QR
purchase screen are not completed. Development bootstrap is not a production account.
Do not expose Development Admin as an end-user account portal.

## Cleanup gate

Run `python tools/validation/audit_cleanup_dependencies.py` for a read-only inventory.
Legacy VBA/DOTM currently feeds Ribbon and rule generators; backend-api feeds the
baseline checker inventory. Retain these sources until replacement fixtures reproduce
the same outputs and all build/validation consumers are migrated. Keep migrations,
tests, signed rule verification, packaging and security audits.

Before deleting approved candidates, copy them to a verified backup outside the project,
record SHA-256 hashes, and verify the copies. Never include private keys, local credentials
or customer Word documents in committed cleanup reports or production payloads.
