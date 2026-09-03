Attribute VB_Name = "DataReadState"
'==============================================================
' DataReadState â€” theo doi trang thai "da bam Doc du lieu chua" VA ket qua ma hoa cua LAN Doc
' "Neu van ban da co du lieu (va o dinh dang docx): chi nut Doc du lieu duoc bat, tat ca nut con
' lai mo => ep phai doc du lieu truoc" + "Neu xuat hien TCVN3/VNI thi nut Chuyen doi Unicode moi
' duoc bat").
' Luu qua SessionState.bas (giong DocumentTypeState.bas) - gan VOI CHINH tai lieu, CHI TRONG PHIEN
' VBA HIEN TAI, khong phai ho so nguoi dung (CLAUDE.md muc 5).
' "Bo tat ca ChuanHoaTheThuc_DaDocDuLieu, ChuanHoaTheThuc_CoBangMaCu da ghi an vao file... moi lan
' mo add-in se coi nhu file moi hoan toan"): TRUOC DAY luu qua Document.Variables (ghi that su vao
' file). Nay chuyen sang SessionState.bas â€” vi SessionState da TU dua ve trang thai "trong" cho
' MOI instance tai lieu moi mo (khoa theo ObjPtr, xem chu thich dau module do), hai co o day gio
' LUON bat dau tu False/"0" cho MOI lan mo, KE CA khong co ResetReadData nao chay. ResetReadData
' duoi day VAN GIU LAI va van duoc AppEventsHost.OnAnyDocumentOpen goi â€” nay chi con la lop bao
' hiem tuong minh (an toan hon la chi dua vao ObjPtr khong the trung, xem canh bao dau
' SessionState.bas), khong con la co che DUY NHAT dam bao "tai lieu moi thi trong" nhu truoc.
'==============================================================
Option Explicit

Private Const VAR_HAS_READ As String = "ChuanHoaTheThuc_DaDocDuLieu"
Private Const VAR_NON_UNICODE As String = "ChuanHoaTheThuc_CoBangMaCu"

' RibbonCallbacks.GetEnabledRequiresDataRead/GetEnabledChuyenDoiUnicode/GetEnabledKiemTra doc qua
' day - True CHI SAU KHI DataReader.RunAndReport chay xong khong loi (item 3).
Public Function HasReadData() As Boolean
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("DataReadState.HasReadData")
    HasReadData = (ReadFlag(VAR_HAS_READ) = "1")
End Function

Public Sub MarkReadData()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("DataReadState.MarkReadData")
    WriteFlag VAR_HAS_READ, "1"
End Sub

' RibbonCallbacks.GetEnabledChuyenDoiUnicode doc qua day (item 4) - ket qua CUA LAN "Doc du lieu"
' GAN NHAT (EncodingConverter.DetectEncoding("nonUnicodeCount") > 0).
Public Function HasNonUnicodeEncoding() As Boolean
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("DataReadState.HasNonUnicodeEncoding")
    HasNonUnicodeEncoding = (ReadFlag(VAR_NON_UNICODE) = "1")
End Function

Public Sub SetNonUnicodeEncoding(ByVal value As Boolean)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("DataReadState.SetNonUnicodeEncoding")
    WriteFlag VAR_NON_UNICODE, IIf(value, "1", "0")
End Sub

Public Sub ResetReadData()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("DataReadState.ResetReadData")
    SessionState.ClearValue VAR_HAS_READ
    SessionState.ClearValue VAR_NON_UNICODE
End Sub

Private Function ReadFlag(ByVal varName As String) As String
    Dim v As Variant: v = SessionState.GetValue(varName)
    If IsEmpty(v) Then
        ReadFlag = "0"
    Else
        ReadFlag = CStr(v)
    End If
End Function

Private Sub WriteFlag(ByVal varName As String, ByVal value As String)
    SessionState.SetValue varName, value
End Sub
