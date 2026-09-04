# Compatibility Matrix

| Word | x86 | x64 | Lane | Evidence hiện tại |
| --- | --- | --- | --- | --- |
| 2010 | NOT_RUN | NOT_RUN | Legacy compatibility | Chưa có VM |
| 2013 | NOT_RUN | NOT_RUN | Legacy compatibility | Chưa có VM |
| 2016 | NOT_RUN | NOT_RUN | Legacy compatibility | Chưa có VM |
| 2019 | NOT_RUN | NOT_RUN | Legacy compatibility | Chưa có VM |
| LTSC 2021 | NOT_RUN | NOT_RUN | Product support | Chưa có VM |
| LTSC 2024 | NOT_RUN | NOT_RUN | Product support | Chưa có VM |
| Microsoft 365 | NOT_RUN | PASS_LOCAL_DEVELOPMENT | Product support | Word 16.0.20326 x64: bản Development 1.0.0.91 fresh install, `LoadBehavior=3`, `COMAddIn.Connect=true`; DOC/DOCX local scan, annotation, command chain, 1-Click, quick spelling và heavy snapshot 10/50/100 trang PASS. `Rebuild` giữ registration cài đặt. Visual tab/icon riêng của `.91` chưa chạy |

Mỗi ô chỉ chuyển PASS sau fresh install, Word launch, Ribbon contract, smoke command, update và uninstall trên máy/VM thật. Mock Interop không thay thế evidence này.

`PASS_LOCAL_SMOKE` chỉ chứng minh build/load an toàn trên máy phát triển hiện tại. Nó không thay thế lane fresh-install/update/uninstall, chữ ký production, Word 2010 hoặc Office x86.
