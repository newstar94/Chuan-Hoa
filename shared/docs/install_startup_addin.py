import os
import shutil

appdata = os.environ.get('APPDATA', '')
startup_dir = os.path.join(appdata, 'Microsoft', 'Word', 'STARTUP')
os.makedirs(startup_dir, exist_ok=True)

src = r"C:\Users\newst\Downloads\Compressed\ChuanHoaTheThuc_unlocked.dotm"
dest = os.path.join(startup_dir, "ChuanHoaTheThuc.dotm")

if os.path.exists(src):
    shutil.copy2(src, dest)
    print(f"Successfully installed Add-in to Word STARTUP folder: {dest}")
else:
    print(f"Source file not found: {src}")
