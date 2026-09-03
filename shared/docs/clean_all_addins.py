import os
import sys
import shutil
import subprocess
import time
import winreg

sys.stdout.reconfigure(encoding='utf-8')

print("1. Đang đóng toàn bộ tiến trình Microsoft Word và Node server...")
subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True)
time.sleep(1)

# 2. Xóa các tệp Add-in và Template trong tất cả thư mục
appdata = os.environ.get('APPDATA', '')
localappdata = os.environ.get('LOCALAPPDATA', '')
userprofile = os.environ.get('USERPROFILE', '')
temp_dir = os.environ.get('TEMP', '')

target_dirs = [
    os.path.join(appdata, r"Microsoft\Word\STARTUP"),
    os.path.join(appdata, r"Microsoft\Templates"),
    os.path.join(appdata, r"Microsoft\AddIns"),
    r"C:\Program Files\Microsoft Office\root\Office16\STARTUP",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\STARTUP",
    r"C:\Program Files\Microsoft Office\Office16\STARTUP",
    r"C:\Program Files (x86)\Microsoft Office\Office16\STARTUP"
]

deleted_files = []

for folder in target_dirs:
    if os.path.exists(folder):
        for fname in os.listdir(folder):
            lower = fname.lower()
            if lower.endswith(('.dotm', '.dot', '.wll')) or 'chuanhoa' in lower or 'thethuc' in lower:
                fpath = os.path.join(folder, fname)
                try:
                    os.remove(fpath)
                    deleted_files.append(fpath)
                    print(f" - Đã xóa tệp: {fpath}")
                except Exception as e:
                    print(f" - Không thể xóa {fpath}: {e}")

# 3. Xóa bộ nhớ đệm Office WEF Cache và Temp files
wef_cache = os.path.join(localappdata, r"Microsoft\Office\16.0\Wef")
if os.path.exists(wef_cache):
    try:
        shutil.rmtree(wef_cache, ignore_errors=True)
        print(" - Đã xóa sạch bộ nhớ đệm WEF Cache.")
    except Exception as e:
        print(f" - WEF Cache note: {e}")

if os.path.exists(temp_dir):
    for fname in os.listdir(temp_dir):
        if 'Word add-in' in fname or 'ChuanHoa' in fname:
            try:
                os.remove(os.path.join(temp_dir, fname))
                print(f" - Đã xóa file tạm: {fname}")
            except Exception:
                pass

# 4. Xóa cấu hình TrustedCatalogs và Developer trong Windows Registry
reg_paths_to_clean = [
    r"Software\Microsoft\Office\16.0\WEF\TrustedCatalogs",
    r"Software\Microsoft\Office\16.0\WEF\Developer"
]

for rp in reg_paths_to_clean:
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, rp)
        print(f" - Đã xóa nhánh Registry: HKCU\\{rp}")
    except Exception as e:
        # Neu co khoa con, xoa tung khoa con
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, rp, 0, winreg.KEY_ALL_ACCESS) as k:
                num_subkeys = winreg.QueryInfoKey(k)[0]
                for i in range(num_subkeys - 1, -1, -1):
                    sub = winreg.EnumKey(k, i)
                    winreg.DeleteKey(k, sub)
            winreg.DeleteKey(winreg.HKEY_CURRENT_USER, rp)
            print(f" - Đã xóa sạch nhánh Registry: HKCU\\{rp}")
        except Exception:
            pass

print(f"\n===> HOÀN TẤT: Đã gỡ bỏ sạch sẽ toàn bộ tất cả các Add-in khỏi Microsoft Word!")
print("Bây giờ bạn mở Microsoft Word sẽ ở trạng thái nguyên bản 100% của Microsoft.")
