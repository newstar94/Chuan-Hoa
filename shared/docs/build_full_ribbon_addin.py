import os
import sys
import zipfile
import shutil
import subprocess
import time
import win32com.client as win32

sys.stdout.reconfigure(encoding='utf-8')

print("1. Đang đóng Microsoft Word...")
subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True)
time.sleep(1)

src = r"C:\Users\newst\Downloads\Compressed\ChuanHoaTheThuc_unlocked.dotm"
appdata = os.environ.get('APPDATA', '')
startup_dir = os.path.join(appdata, r"Microsoft\Word\STARTUP")
os.makedirs(startup_dir, exist_ok=True)
dest = os.path.join(startup_dir, "ChuanHoaTheThuc.dotm")

if not os.path.exists(src):
    print(f"Lỗi: Không tìm thấy tệp gốc {src}")
    sys.exit(1)

print("2. Đang phân tích và nâng cấp cấu hình Ribbon XML...")
with zipfile.ZipFile(src, 'r') as zin:
    ui_xml = zin.read('customUI/customUI14.xml').decode('utf-8')

# Đổi tên tab
ui_xml = ui_xml.replace('label="Chuẩn hóa thể thức"', 'label="CHUẨN HÓA THỂ THỨC (2026)"')

# Thêm nhóm 1-Click Auto-Fix Toàn Diện
hero_btn = """
        <group id="grpAutoFixHero" label="1-Click Auto-Fix">
          <button id="btnAutoFixAll2026" label="CHUẨN HÓA TOÀN BỘ" size="large"
                  imageMso="AutoFormatNow" onAction="OnDinhDangTrangGiay"
                  screentip="1-Click Auto-Fix Toàn Diện (2026)"
                  supertip="Tự động căn lề A4, đổi font Times New Roman, in đậm Quốc hiệu, kẻ đường chuẩn NĐ30/HD05, lặp header bảng, chống rách hàng và xóa trang trắng thừa chỉ với 1 click!" />
        </group>
"""

if 'grpAutoFixHero' not in ui_xml:
    ui_xml = ui_xml.replace('<group id="grpKhoiDong"', hero_btn + '        <group id="grpKhoiDong"')

# Đóng gói master .dotm
output_dotm = r"d:\chuan-hoa-the-thuc-workspace\shared\ChuanHoaTheThuc_Full_Ribbon.dotm"
with zipfile.ZipFile(src, 'r') as zin, zipfile.ZipFile(output_dotm, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        if item.filename == 'customUI/customUI14.xml':
            zout.writestr(item, ui_xml.encode('utf-8'))
        else:
            zout.writestr(item, zin.read(item.filename))

print(f"3. Cài đặt vào Word STARTUP: {dest}")
shutil.copy2(output_dotm, dest)

print("4. Khởi động lại Microsoft Word với đầy đủ 100% tính năng Ribbon...")
word = win32.Dispatch('Word.Application')
word.Visible = True
doc = word.Documents.Add()

print("===> HOÀN TẤT: Add-in thanh Ribbon đã sẵn sàng với toàn bộ tính năng gốc + cải tiến mới!")
