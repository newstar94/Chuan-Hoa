Attribute VB_Name = "DocumentTypeState"
'==============================================================
' DocumentTypeState â€” Luu/doc lua chon cua drop-down "Loai van ban" (nhom "Khoi dong", nut 1.4
' Luu qua SessionState.bas â€” gan VOI TUNG TAI LIEU dang mo, CHI TRONG PHIEN VBA HIEN TAI (khong
' con ghi vao Document.Variables/file.docx).
' "Bo tat ca ChuanHoaTheThuc_LoaiVanBanIndex... da ghi an vao file. Toi muon moi lan mo add-in se
' coi nhu file moi hoan toan"): TRUOC DAY luu qua Document.Variables (ghi that su vao file, ton
' tai qua ca luc dong roi mo lai). Nay chuyen sang SessionState.bas (Scripting.Dictionary trong bo
' nho, khong dong cham OOXML) â€” xem chu thich dau file do ve ly do KHONG can code rieng de "reset"
' moi lan mo lai file.
' Danh sach item khop 1-1 voi <dropDown id="ddLoaiVanBan"> trong customUI14.xml: chi so 0 = "Khong
' xac dinh" chi so 1 = "Cong van" chi so 2..28 = 27 ten loai, DUNG THEO THU TU ordinal trong
' shared/rules/chu-viet-tat-ten-loai.json (RuleLoader.GetDocTypeAbbreviations) -> chi so = ordinal
' + 1.
' drop-down nay CHI de HIEN THI + NHAN LUA CHON cua nguoi dung - CHUA anh huong ket qua "Kiem tra"
' (se lam o task rieng sau). FindingReporter.BuildCheckContext VAN tu dong nhan dien loai van ban
' nhu cu (DocumentTypeDetector. DetectDocumentType) de xay dung context kiem tra; gia tri o day
' chi goi AutoDetectAndStore SAU KHI da co ket qua nhan dien, de dong bo hien thi drop-down voi
' ket qua vua quet.
' "khi khoi dong (va doc du lieu) khong tu chon loai van ban trong drop list - toi yeu cau ghi de
' viec nguoi dung chon thu cong"): BO HAN co che "khong ghi de lua chon thu cong" (VAR_MANUAL cu)
' - AutoDetectAndStore nay LUON GHI, moi lan mo tai lieu VA moi lan bam "Kiem tra" deu tu nhan
' dien lai va GHI DE drop-down, bat ke nguoi dung tung tu chon gi truoc do. Nguoi dung van chon
' tay duoc qua drop- down (OnChonLoaiVanBan/SetSelectedIndexManual) - lua chon do CHI giu hieu luc
' cho toi lan tu dong nhan dien TIEP THEO (mo tai lieu khac, hoac bam "Kiem tra" lan nua).
'==============================================================
Option Explicit

Private Const VAR_INDEX As String = "ChuanHoaTheThuc_LoaiVanBanIndex"

' Chi so hien tai cua drop-down cho ActiveDocument - dung lam getSelectedItemIndex
' (RibbonCallbacks.GetSelectedIndexLoaiVanBan). Tai lieu chua tung duoc quet/chon -> 0 ("Khong xac
' dinh").
Public Function GetSelectedIndex() As Long
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("DocumentTypeState.GetSelectedIndex")
    On Error GoTo ErrHandler
    Dim v As Variant
    v = ReadDocVariable(VAR_INDEX)
    If IsEmpty(v) Then
        GetSelectedIndex = 0
    Else
        GetSelectedIndex = CLng(v)
    End If
    Exit Function
ErrHandler:
    GetSelectedIndex = 0
End Function

Public Sub SetSelectedIndexManual(ByVal idx As Long)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("DocumentTypeState.SetSelectedIndexManual")
    On Error GoTo ErrHandler
    WriteDocVariable VAR_INDEX, CStr(idx)
    Exit Sub
ErrHandler:
    ' Khong co ActiveDocument hop le (vi du dang dong tai lieu) - bo qua, khong bao loi cho mot
    ' thao tac phu (drop-down hien thi), tranh lam gian doan luong chinh.
End Sub

' Goi tu FindingReporter.BuildCheckContext NGAY SAU khi co ket qua DocumentTypeDetector.
' DetectDocumentType. LUON GHI (khong con kiem "da tung chon thu cong" - xem ghi chu dau file).
Public Sub AutoDetectAndStore(ByVal documentTypeKey As String, ByVal matchedOrdinal As Long)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("DocumentTypeState.AutoDetectAndStore")
    On Error GoTo ErrHandler
    WriteDocVariable VAR_INDEX, CStr(ResolveIndex(documentTypeKey, matchedOrdinal))
    Exit Sub
ErrHandler:
    ' Nhu SetSelectedIndexManual - khong lam gian doan luong "Kiem tra" vi mot buoc phu.
End Sub

Private Function ResolveIndex(ByVal documentTypeKey As String, ByVal matchedOrdinal As Long) As Long
    If documentTypeKey = "coTenLoai" And matchedOrdinal > 0 Then
        ResolveIndex = matchedOrdinal + 1
    ElseIf documentTypeKey = "congVan" Then
        ResolveIndex = 1
    Else
        ResolveIndex = 0
    End If
End Function

Public Sub DetectAndAutoStore(Optional ByVal snapshot As Object = Nothing)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("DocumentTypeState.DetectAndAutoStore")
    On Error GoTo ErrHandler
    If snapshot Is Nothing Then Set snapshot = DocumentSnapshot.CaptureDocument()
    Dim typeResult As Object: Set typeResult = DocumentTypeDetector.DetectDocumentType(snapshot)
    Dim matchedOrdinal As Long
    matchedOrdinal = DocumentTypeDetector.MatchedTypeNameOrdinal(snapshot("Paragraphs"), typeResult("EvidenceParagraphIndex"))
    AutoDetectAndStore CStr(typeResult("Type")), matchedOrdinal
    RibbonCallbacks.InvalidateLoaiVanBan
    RibbonCallbacks.InvalidateKiemTra
    Exit Sub
ErrHandler:
    ' Tinh nang phu, khong lam gian doan luong goi (mo tai lieu hoac "Kiem tra").
End Sub

Private Function ReadDocVariable(ByVal varName As String) As Variant
    ReadDocVariable = SessionState.GetValue(varName)
End Function

Private Sub WriteDocVariable(ByVal varName As String, ByVal value As String)
    SessionState.SetValue varName, value
End Sub
