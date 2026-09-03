Attribute VB_Name = "DocxConverter"
Option Explicit

' ============================================================================
' Nut 1.1 "Luu thanh DOCX". Tra ve True neu da luu thanh.docx thanh cong (RibbonCallbacks dung de
' InvalidateControl lam mo nut), False neu khong lam gi ca (da la.docx, nguoi dung bam Cancel
' trong hop thoai Save As, hoac luu xong nhung nguoi dung tu doi sang dinh dang khac.docx) hoac
' loi (da tu bao MsgBox qua Utils.AbortOperation).
' ============================================================================
Public Function SaveAsDocx() As Boolean
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("DocxConverter.SaveAsDocx")
    On Error GoTo ErrHandler

    ' Da la.docx thi khong lam gi. Ribbon da lam mo nut qua getEnabled
    ' (RibbonCallbacks.GetEnabledLuuDocx) nhung van kiem tra lai o day phong khi goi truc tiep tu
    ' Immediate Window hoac ribbon chua kip InvalidateControl.
    If IsCurrentDocumentDocx() Then
        SaveAsDocx = False
        Exit Function
    End If

    ' Hop thoai Save As goc â€” chon san dinh dang.docx, nguoi dung tu chon noi luu/dat ten. .Show
    ' (khac.Display) VUA hien hop thoai VUA thuc hien luu neu nguoi dung bam OK/Save, dung nhu ban
    ' VBA thong thuong cua Dialogs collection.
    Dim dlg As word.Dialog
    Set dlg = Application.Dialogs(wdDialogFileSaveAs)
    dlg.Format = wdFormatXMLDocument

    Dim shown As Long
    shown = dlg.Show

    ' 0 = nguoi dung bam Cancel (hoac dong bang nut X) - khong lam gi them, khong ghi nhat ky.
    If shown = 0 Then
        SaveAsDocx = False
        Exit Function
    End If

    ' Ghi nhat ky phien lam viec (Utils.BeginOperation/EndOperation) - thao tac ghi file da xong
    ' luc Show tra ve, day chi la buoc ke toan cho nhat ky thao tac.
    Dim opName As String
    opName = "L" & ChrW(&H1B0) & "u th" & ChrW(&HE0) & "nh DOCX"
    Utils.BeginOperation opName

    ' Nguoi dung co the tu doi "Save as type" trong hop thoai sang dinh dang khac.docx - kiem lai
    ' cho chac truoc khi bao thanh cong.
    SaveAsDocx = IsCurrentDocumentDocx()
    Utils.EndOperation IIf(SaveAsDocx, 1, 0), Not SaveAsDocx
    Exit Function

ErrHandler:
    Utils.AbortOperation Err.description
    SaveAsDocx = False
End Function

' Dung o getEnabled cua ribbon (RibbonCallbacks.GetEnabledLuuDocx) - True khi tai lieu dang mo co
' phan mo rong la "docx" (khong phan biet hoa/thuong). Khong co tai lieu dang mo (loi truy cap
' ActiveDocument) thi coi nhu KHONG phai.docx (an toan hon: nut sang, bam vao se tu bao loi qua
' ErrHandler cua SaveAsDocx thay vi im lang lam mo nut sai). "Neu la Document moi chua luu bao gio
' thi duoc tinh la docx"): tai lieu chua tung luu (Path rong) mac dinh se luu thanh.docx (dinh
' dang mac dinh cua Word hien dai tu Document.Add) - coi nhu da thoa dieu kien ".docx" cho MOI
' cong chan cua add-in, khong ep nguoi dung "Luu thanh DOCX" cho mot tai lieu con chua co gi de
' luu (nut do van mo/khong thay doi gi neu bam - xem SaveAsDocx).
Public Function IsCurrentDocumentDocx() As Boolean
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("DocxConverter.IsCurrentDocumentDocx")
    On Error GoTo ErrHandler
    If Len(ActiveDocument.Path) = 0 Then
        IsCurrentDocumentDocx = True
        Exit Function
    End If
    IsCurrentDocumentDocx = (LCase$(GetExtension(ActiveDocument.name)) = "docx")
    Exit Function
ErrHandler:
    IsCurrentDocumentDocx = False
End Function

Private Function GetExtension(ByVal fileName As String) As String
    Dim dotPos As Long
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        GetExtension = Mid$(fileName, dotPos + 1)
    Else
        GetExtension = ""
    End If
End Function
