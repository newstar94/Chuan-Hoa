import os
import shutil
import subprocess
import time
import win32com.client as win32

print("1. Closing Word...")
subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True)
time.sleep(1)

print("2. Clearing Office WEF Addin Cache...")
localappdata = os.environ.get('LOCALAPPDATA', '')
wef_cache = os.path.join(localappdata, r'Microsoft\Office\16.0\Wef')
if os.path.exists(wef_cache):
    try:
        shutil.rmtree(wef_cache, ignore_errors=True)
        print("Cleared WEF cache successfully!")
    except Exception as e:
        print(f"Clear cache note: {e}")

print("3. Launching Word fresh...")
word = win32.Dispatch('Word.Application')
word.Visible = True
doc = word.Documents.Add()
print("Word launched fresh with HTTP Add-in ready!")
