import winreg
import os
import subprocess

print("1. Cau hinh WebView2 bo qua loi chung chi localhost...")
# Registry keys for Edge WebView2 to ignore certificate errors on localhost
keys = [
    r"Software\Policies\Microsoft\Edge\WebView2",
    r"Software\Microsoft\Edge\WebView2",
    r"Software\Policies\Microsoft\Office\16.0\WEF",
    r"Software\Microsoft\Office\16.0\WEF"
]

for k in keys:
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, k) as reg_key:
            winreg.SetValueEx(reg_key, "AdditionalBrowserArguments", 0, winreg.REG_SZ, "--ignore-certificate-errors --allow-insecure-localhost")
            print(f" - Da ghi cau hinh vao: HKCU\\{k}")
    except Exception as e:
        print(f" - Note {k}: {e}")

print("\n2. Tu dong import chung chi vao he thong...")
ca_path = os.path.join(os.environ.get('USERPROFILE', ''), r'.office-addin-dev-certs\ca.crt')
if os.path.exists(ca_path):
    cmd = f'certutil -addstore -f "Root" "{ca_path}"'
    res = subprocess.run(cmd, shell=True, capture_output=True, text=True)
    print(f"Certutil output: {res.stdout.strip() or res.stderr.strip()}")

print("\n3. Hoan tat! Bay gio WebView2 va Word se tu dong bo qua kiem tra chung chi.")
