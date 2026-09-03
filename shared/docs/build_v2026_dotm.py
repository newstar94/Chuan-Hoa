import os
import sys
import zipfile
import shutil

sys.stdout.reconfigure(encoding='utf-8')

src = r"C:\Users\newst\Downloads\Compressed\ChuanHoaTheThuc_unlocked.dotm"
appdata = os.environ.get('APPDATA', '')
startup_dir = os.path.join(appdata, 'Microsoft', 'Word', 'STARTUP')
os.makedirs(startup_dir, exist_ok=True)
dest = os.path.join(startup_dir, "ChuanHoaTheThuc_v2026.dotm")

if not os.path.exists(src):
    print(f"Source not found: {src}")
    sys.exit(1)

with zipfile.ZipFile(src, 'r') as zin:
    ui_xml = zin.read('customUI/customUI14.xml').decode('utf-8')

# 1. Update Tab Label
ui_xml = ui_xml.replace('label="Chuẩn hóa thể thức"', 'label="CHUẨN HÓA THỂ THỨC (2026)"')

# 2. Add 1-Click Auto-Fix Group
hero_group = """
        <group id="grpAutoFixHero" label="1-Click Auto-Fix">
          <button id="btnAutoFixAll2026" label="CHUẨN HÓA TOÀN BỘ" size="large"
                  imageMso="AutoFormatNow" onAction="OnDinhDangTrangGiay"
                  screentip="1-Click Auto-Fix Toàn Diện (2026)"
                  supertip="Tự động căn lề A4, font Times New Roman, kẻ đường chuẩn NĐ30/HD05, lặp header bảng và xóa trang trắng thừa chỉ với 1 click!" />
        </group>
"""

ui_xml = ui_xml.replace('<group id="grpKhoiDong"', hero_group + '        <group id="grpKhoiDong"')

# 3. Save to output dotm
temp_dotm = r"d:\chuan-hoa-the-thuc-workspace\shared\ChuanHoaTheThuc_v2026.dotm"
with zipfile.ZipFile(src, 'r') as zin, zipfile.ZipFile(temp_dotm, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        if item.filename == 'customUI/customUI14.xml':
            zout.writestr(item, ui_xml.encode('utf-8'))
        else:
            zout.writestr(item, zin.read(item.filename))

shutil.copy2(temp_dotm, dest)
print(f"Built & Installed: {dest}")
