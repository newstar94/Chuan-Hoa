import os
import sys
import subprocess
import time

sys.stdout.reconfigure(encoding='utf-8')

print("1. Đóng ứng dụng Microsoft Word...")
subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True)
time.sleep(1)

appdata = os.environ.get('APPDATA', '')
localappdata = os.environ.get('LOCALAPPDATA', '')

# Các thư mục chứa add-in Word tiềm năng
target_dirs = [
    os.path.join(appdata, 'Microsoft', 'Word', 'STARTUP'),
    os.path.join(appdata, 'Microsoft', 'Templates'),
    r'C:\Program Files\Microsoft Office\root\Office16\STARTUP',
    r'C:\Program Files (x86)\Microsoft Office\root\Office16\STARTUP',
    r'C:\Program Files\Microsoft Office\Office16\STARTUP',
    r'C:\Program Files (x86)\Microsoft Office\Office16\STARTUP',
    os.path.join(appdata, 'Microsoft', 'AddIns')
]

deleted_files = []

for d in target_dirs:
    if os.path.exists(d):
        for fname in os.listdir(d):
            if 'ChuanHoa' in fname or 'chuan_hoa' in fname.lower() or 'the_thuc' in fname.lower():
                full_path = os.path.join(d, fname)
                try:
                    os.remove(full_path)
                    deleted_files.append(full_path)
                    print(f"Đã xóa vĩnh viễn: {full_path}")
                except Exception as e:
                    print(f"Không thể xóa {full_path}: {e}")

print(f"\nTổng kết: Đã gỡ bỏ hoàn toàn {len(deleted_files)} tệp add-in cũ.")
print("Bây giờ bạn mở lại Word sẽ sạch hoàn toàn không còn tab cũ.")
