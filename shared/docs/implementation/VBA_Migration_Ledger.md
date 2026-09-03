# VBA Migration Ledger

Ledger machine-readable đầy đủ nằm tại `evidence/vba_migration_ledger.json`.

| Disposition | Số module |
| --- | ---: |
| PORT | 37 |
| REPLACE | 15 |
| REWRITE | 9 |
| RETIRE | 3 |
| MERGE | 2 |
| MIGRATE | 1 |
| SPLIT | 1 |
| Tổng | 68 |

Baseline có 363 public entry points và 46 module mang ít nhất một side-effect flag. Dependency/caller trong JSON là kết quả phân tích token tĩnh, phải được xác nhận khi port và bằng golden behavior tests.

Mỗi entry đã có hash, line count, public entry points, callers/dependencies suy luận, side-effect flags, đích, disposition và migration status `BASELINED`. Fixture, test ID, reviewer và sign-off chỉ được bổ sung khi có evidence thật.

