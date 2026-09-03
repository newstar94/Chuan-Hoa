import winreg

def set_reg_keys():
    # 1. WEF TrustedCatalogs
    wef_path = r"Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{D3B07384-D113-4F44-972A-60589A19D826}"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, wef_path) as key:
        winreg.SetValueEx(key, "Url", 0, winreg.REG_SZ, r"d:\chuan-hoa-the-thuc-workspace\client-web-addin")
        winreg.SetValueEx(key, "Flags", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(key, "Id", 0, winreg.REG_SZ, "{D3B07384-D113-4F44-972A-60589A19D826}")
    print("WEF TrustedCatalogs set successfully!")

    # 2. Word Security Trusted Locations
    loc_path = r"Software\Microsoft\Office\16.0\Word\Security\Trusted Locations\VietDocAddin"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, loc_path) as key:
        winreg.SetValueEx(key, "Path", 0, winreg.REG_SZ, r"d:\chuan-hoa-the-thuc-workspace\client-web-addin")
        winreg.SetValueEx(key, "AllowSubFolders", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(key, "Description", 0, winreg.REG_SZ, "VietDoc Standardizer Addin")
    print("Word Trusted Locations set successfully!")

    # 3. WEF Developer
    dev_path = r"Software\Microsoft\Office\16.0\WEF\Developer"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, dev_path) as key:
        winreg.SetValueEx(key, "d3b07384-d113-4f44-972a-60589a19d826", 0, winreg.REG_SZ, r"d:\chuan-hoa-the-thuc-workspace\client-web-addin\manifest.xml")
    print("WEF Developer key set successfully!")

if __name__ == "__main__":
    set_reg_keys()
