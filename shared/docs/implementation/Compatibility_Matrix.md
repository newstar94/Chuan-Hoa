# Compatibility Matrix

| Word | x86 | x64 | Lane | Evidence hiện tại |
| --- | --- | --- | --- | --- |
| 2010 | NOT_RUN | NOT_RUN | Legacy compatibility | Chưa có VM |
| 2013 | NOT_RUN | NOT_RUN | Legacy compatibility | Chưa có VM |
| 2016 | NOT_RUN | NOT_RUN | Legacy compatibility | Chưa có VM |
| 2019 | NOT_RUN | NOT_RUN | Legacy compatibility | Chưa có VM |
| LTSC 2021 | NOT_RUN | NOT_RUN | Product support | Chưa có VM |
| LTSC 2024 | NOT_RUN | NOT_RUN | Product support | Chưa có VM |
| Microsoft 365 | NOT_RUN | PASS_LOCAL_SMOKE | Product support | Word 16.0.20326.20112 x64: Release build 0 warning/error; `COMAddIn.Connect=true`, `LoadBehavior=3`; đúng 1 tab/7 group; About hoạt động; chưa chạy mutation |

Mỗi ô chỉ chuyển PASS sau fresh install, Word launch, Ribbon contract, smoke command, update và uninstall trên máy/VM thật. Mock Interop không thay thế evidence này.

`PASS_LOCAL_SMOKE` chỉ chứng minh build/load an toàn trên máy phát triển hiện tại. Nó không thay thế lane fresh-install/update/uninstall, chữ ký production, Word 2010 hoặc Office x86.
