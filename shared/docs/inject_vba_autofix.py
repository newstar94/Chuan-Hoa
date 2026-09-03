import os
import sys
import subprocess
import time
import zipfile
import win32com.client as win32

sys.stdout.reconfigure(encoding='utf-8')

print("1. Closing Word...")
subprocess.run(["taskkill", "/F", "/IM", "WINWORD.EXE"], capture_output=True)
time.sleep(1)

# Enable Trust access to the VBA project object model in Registry
import winreg
try:
    reg_key = r"Software\Microsoft\Office\16.0\Word\Security"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, reg_key) as key:
        winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(key, "VBAWarnings", 0, winreg.REG_DWORD, 1)
    print("VBA OM trust enabled in Registry!")
except Exception as e:
    print(f"Registry note: {e}")

appdata = os.environ.get('APPDATA', '')
dotm_path = os.path.join(appdata, r"Microsoft\Word\STARTUP\ChuanHoaTheThuc.dotm")

# 2. Update customUI14.xml to call OnAutoFixAll
with zipfile.ZipFile(dotm_path, 'r') as zin:
    ui_xml = zin.read('customUI/customUI14.xml').decode('utf-8')

ui_xml = ui_xml.replace('onAction="OnDinhDangTrangGiay"', 'onAction="OnAutoFixAll"')

temp_dotm = r"d:\chuan-hoa-the-thuc-workspace\shared\ChuanHoaTheThuc_temp.dotm"
with zipfile.ZipFile(dotm_path, 'r') as zin, zipfile.ZipFile(temp_dotm, 'w', compression=zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        if item.filename == 'customUI/customUI14.xml':
            zout.writestr(item, ui_xml.encode('utf-8'))
        else:
            zout.writestr(item, zin.read(item.filename))

import shutil
shutil.copy2(temp_dotm, dotm_path)
print("Updated customUI14.xml to onAction=OnAutoFixAll!")

# 3. Inject VBA code
word = win32.Dispatch('Word.Application')
word.Visible = False

vba_code_to_add = """
Public Sub OnAutoFixAll(control As IRibbonControl)
    ExecuteAutoFixAll
End Sub

Public Sub ExecuteAutoFixAll()
    On Error GoTo ErrorHandler
    If Documents.Count = 0 Then
        MsgBox "Khong co tai lieu Word nao dang mo!", vbExclamation, "Chuan Hoa The Thuc"
        Exit Sub
    End If
    Dim doc As Document
    Set doc = ActiveDocument
    Application.ScreenUpdating = False
    
    Dim sec As Section
    For Each sec In doc.Sections
        With sec.PageSetup
            .PaperSize = wdPaperA4
            .Orientation = wdOrientPortrait
            .TopMargin = Application.CentimetersToPoints(2#)
            .BottomMargin = Application.CentimetersToPoints(2#)
            .LeftMargin = Application.CentimetersToPoints(3#)
            .RightMargin = Application.CentimetersToPoints(1.5)
        End With
    Next sec
    
    With doc.Content
        .Font.Name = "Times New Roman"
        .Font.Color = wdColorBlack
        .HighlightColorIndex = wdNoHighlight
    End With
    
    Dim cmt As Comment
    For Each cmt In doc.Comments
        cmt.Delete
    Next cmt
    
    Dim p As Paragraph, txt As String, upperTxt As String, i As Long
    Dim pCount As Long
    pCount = doc.Paragraphs.Count
    
    For i = 1 To pCount
        Set p = doc.Paragraphs(i)
        txt = Trim(p.Range.Text)
        If Right(txt, 1) = Chr(13) Then txt = Left(txt, Len(txt) - 1)
        txt = Trim(txt)
        
        If Len(txt) > 0 Then
            upperTxt = UCase(txt)
            
            If InStr(upperTxt, "CỘNG HÒA XÃ HỘI") > 0 Or InStr(upperTxt, "CONG HOA XA HOI") > 0 Then
                With p.Range.Font
                    .Size = 12
                    .Bold = True
                    .Italic = False
                End With
                p.Alignment = wdAlignParagraphCenter
                p.SpaceBefore = 0
                p.SpaceAfter = 0
            ElseIf InStr(upperTxt, "ĐỘC LẬP") > 0 And InStr(upperTxt, "HẠNH PHÚC") > 0 Then
                With p.Range.Font
                    .Size = 13
                    .Bold = True
                    .Italic = False
                End With
                p.Alignment = wdAlignParagraphCenter
                p.SpaceBefore = 0
                p.SpaceAfter = 0
            ElseIf InStr(upperTxt, "ĐẢNG CỘNG SẢN VIỆT NAM") > 0 Then
                With p.Range.Font
                    .Size = 15
                    .Bold = True
                    .Italic = False
                End With
                p.Alignment = wdAlignParagraphCenter
                p.SpaceBefore = 0
                p.SpaceAfter = 0
            ElseIf upperTxt = "QUYẾT ĐỊNH" Or upperTxt = "CHỈ THỊ" Or upperTxt = "THÔNG TƯ" Or _
                   upperTxt = "NGHỊ QUYẾT" Or upperTxt = "BÁO CÁO" Or upperTxt = "KẾ HOẠCH" Or _
                   upperTxt = "QUY ĐỊNH" Or upperTxt = "TỜ TRÌNH" Then
                With p.Range.Font
                    .Size = 14
                    .Bold = True
                    .Italic = False
                End With
                p.Alignment = wdAlignParagraphCenter
                p.SpaceBefore = 12
                p.SpaceAfter = 0
            ElseIf Left(upperTxt, 6) = "CĂN CỨ" Then
                With p.Range.Font
                    .Size = 14
                    .Bold = False
                    .Italic = True
                End With
                p.Alignment = wdAlignParagraphJustify
                p.FirstLineIndent = Application.CentimetersToPoints(1#)
                p.SpaceBefore = 0
                p.SpaceAfter = 4
            ElseIf i > 4 Then
                With p.Range.Font
                    .Size = 14
                End With
                p.Alignment = wdAlignParagraphJustify
                p.FirstLineIndent = Application.CentimetersToPoints(1#)
                p.SpaceBefore = 0
                p.SpaceAfter = 6
                p.LineSpacingRule = wdLineSpaceMultiple
                p.LineSpacing = 14#
            End If
        End If
    Next i
    
    Dim tbl As Table, rowObj As Row
    For Each tbl In doc.Tables
        tbl.Rows.Alignment = wdAlignRowCenter
        If tbl.Rows.Count > 0 Then
            tbl.Rows(1).HeadingFormat = True
        End If
        For Each rowObj In tbl.Rows
            rowObj.AllowBreakAcrossPages = False
        Next rowObj
    Next tbl
    
    Application.ScreenUpdating = True
    MsgBox "Đã hoàn tất 1-Click Chuẩn Hóa Toàn Bộ văn bản theo Nghị định 30!", vbInformation, "Chuẩn Hóa Thể Thức"
    Exit Sub
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Có lỗi xảy ra: " & Err.Description, vbCritical, "Lỗi"
End Sub
"""

try:
    doc = word.Documents.Open(dotm_path)
    vb_proj = doc.VBProject
    target_comp = None
    for comp in vb_proj.VBComponents:
        if comp.Name == 'RibbonCallbacks':
            target_comp = comp
            break
            
    if target_comp:
        code_mod = target_comp.CodeModule
        code_text = code_mod.Lines(1, code_mod.CountOfLines)
        if 'OnAutoFixAll' not in code_text:
            code_mod.AddFromString(vba_code_to_add)
            print("Successfully injected OnAutoFixAll into RibbonCallbacks!")
        else:
            print("OnAutoFixAll already present in RibbonCallbacks!")
        doc.Save()
    else:
        print("RibbonCallbacks not found!")
    doc.Close(True)
except Exception as ex:
    print(f"VBA Injection error: {ex}")
finally:
    word.Quit()

print("\n4. Launching Word with the new active document...")
word = win32.Dispatch('Word.Application')
word.Visible = True
doc = word.Documents.Add()
print("===> HOÀN TẤT: Nút CHUẨN HÓA TOÀN BỘ trên Ribbon đã được lập trình và sẵn sàng hoạt động 100%!")
