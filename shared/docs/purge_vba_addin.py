import os
import sys
import subprocess
import time

sys.stdout.reconfigure(encoding='utf-8')

print("1. Đang đóng toàn bộ tiến trình Microsoft Word...")
subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True)
time.sleep(1)

appdata = os.environ.get('APPDATA', '')
localappdata = os.environ.get('LOCALAPPDATA', '')

# Tất cả thư mục hệ thống có thể chứa Word Add-in / Template
target_dirs = [
    os.path.join(appdata, r'Microsoft\Word\STARTUP'),
    os.path.join(appdata, r'Microsoft\Templates'),
    os.path.join(appdata, r'Microsoft\AddIns'),
    r'C:\Program Files\Microsoft Office\root\Office16\STARTUP',
    r'C:\Program Files (x86)\Microsoft Office\root\Office16\STARTUP',
    r'C:\Program Files\Microsoft Office\Office16\STARTUP',
    r'C:\Program Files (x86)\Microsoft Office\Office16\STARTUP'
]

deleted_count = 0

for folder in target_dirs:
    if os.path.exists(folder):
        for fname in os.listdir(folder):
            # Xóa tất cả file .dotm / .dot liên quan đến chuẩn hóa thể thức
            lower = fname.lower()
            if lower.endswith(('.dotm', '.dot', '.wll')) or 'chuanhoa' in lower or 'thethuc' in lower:
                full_path = os.path.join(folder, fname)
                try:
                    os.remove(full_path)
                    deleted_count += 1
                    print(f" - Đã xóa vĩnh viễn: {full_path}")
                except Exception as e:
                    print(f" - Lỗi khi xóa {full_path}: {e}")

print(f"\n===> HOÀN TẤT: Đã gỡ bỏ sạch sẽ {deleted_count} tệp Add-in VBA khỏi Word!")
