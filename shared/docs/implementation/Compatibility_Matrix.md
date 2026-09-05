# Compatibility Matrix

| Word | x86 | x64 | Lane | Evidence hiện tại |
| --- | --- | --- | --- | --- |
| 2010 | NOT_RUN | NOT_RUN | Legacy compatibility | Chưa có VM |
| 2013 | NOT_RUN | NOT_RUN | Legacy compatibility | Chưa có VM |
| 2016 | NOT_RUN | NOT_RUN | Legacy compatibility | Chưa có VM |
| 2019 | NOT_RUN | NOT_RUN | Legacy compatibility | Chưa có VM |
| LTSC 2021 | NOT_RUN | NOT_RUN | Product support | Chưa có VM |
| LTSC 2024 | NOT_RUN | NOT_RUN | Product support | Chưa có VM |
| Microsoft 365 | NOT_RUN | PASS_LOCAL_DEVELOPMENT_RUNTIME_1.0.0.107; PASS_VISUAL_RIBBON_HASH_BASELINE | Product support | Word 16.0.20326.20132 x64: bản cài `.107`, `LoadBehavior=3`; passive startup không đổi cache/backup và không background save/print; 39 control saved/unsaved PASS. Outer + 3/3 owned inner PE signatures, Apps & Features, cached uninstall, repair/uninstall/reinstall và rollback 6/6 PASS trên `.107`. DOC/DOCX, `Document1`, annotation/selected-fix, 1-Click, heavy snapshot và benchmark đã PASS trên cùng runtime/Ribbon source trước đó; người dùng xác nhận visual tab/icon/idle với cùng Ribbon XML hash. |

Mỗi ô chỉ chuyển PASS sau fresh install, Word launch, Ribbon contract, smoke command, update và uninstall trên máy/VM thật. Mock Interop không thay thế evidence này.

`PASS_LOCAL_SMOKE` chỉ chứng minh build/load an toàn trên máy phát triển hiện tại. Nó không thay thế lane fresh-install/update/uninstall, chữ ký production, Word 2010 hoặc Office x86.
