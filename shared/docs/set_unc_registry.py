import winreg

def set_unc_trusted_catalog():
    unc_path = r"\\localhost\D$\chuan-hoa-the-thuc-workspace\client-web-addin"
    guid = "{D3B07384-D113-4F44-972A-60589A19D826}"
    
    # 1. Office 16.0 WEF TrustedCatalogs
    wef_path = rf"Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{guid}"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, wef_path) as key:
        winreg.SetValueEx(key, "Url", 0, winreg.REG_SZ, unc_path)
        winreg.SetValueEx(key, "Flags", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(key, "Id", 0, winreg.REG_SZ, guid)
    
    # 2. Add with HTD name as well
    guid2 = "{A1B2C3D4-E5F6-4A5B-8C9D-0E1F2A3B4C5D}"
    wef_path2 = rf"Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{guid2}"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, wef_path2) as key:
        winreg.SetValueEx(key, "Url", 0, winreg.REG_SZ, r"\\HTD\D$\chuan-hoa-the-thuc-workspace\client-web-addin")
        winreg.SetValueEx(key, "Flags", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(key, "Id", 0, winreg.REG_SZ, guid2)

    print("Successfully registered UNC Trusted Catalogs in Windows Registry!")

if __name__ == "__main__":
    set_unc_trusted_catalog()
