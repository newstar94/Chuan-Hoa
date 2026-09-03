import os
import subprocess
import shutil
import time

print("1. Terminating existing Word processes...")
subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True)
time.sleep(1)

print("2. Syncing Add-in template to Word STARTUP folder...")
appdata = os.environ.get('APPDATA', '')
startup_dir = os.path.join(appdata, 'Microsoft', 'Word', 'STARTUP')
os.makedirs(startup_dir, exist_ok=True)

src = r"C:\Users\newst\Downloads\Compressed\ChuanHoaTheThuc_unlocked.dotm"
dest = os.path.join(startup_dir, "ChuanHoaTheThuc.dotm")

if os.path.exists(src):
    try:
        shutil.copy2(src, dest)
        print(f"Copied template to: {dest}")
    except Exception as ex:
        print(f"Copy note: {ex}")

print("3. Launching Microsoft Word...")
subprocess.Popen(["start", "winword"], shell=True)
print("Microsoft Word restarted successfully!")
