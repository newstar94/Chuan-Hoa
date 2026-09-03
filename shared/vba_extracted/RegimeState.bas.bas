Attribute VB_Name = "RegimeState"
'==============================================================
' RegimeState -- Luu/doc "Quy dinh" dang chon (ND30 | VIETTEL | DANG) cho drop-down "Quy dinh"
' (nhom "Khoi dong",, 26/8/2026). CUNG khuon voi DocumentTypeState.bas nhung THEM mot lop nho gia
' tri toan cuc rieng cho "lua chon thu cong da nho" -- day la diem KHAC BIET co chu dich, khong
' phai sao chep may moc:
' - Lua chon HIEN TAI cho ActiveDocument (GetSelectedIndex/SetSelectedIndexManual/
'   AutoDetectAndStore) van luu qua SessionState.bas -- MOI TAI LIEU dang mo co the dang hien mot
'   lua chon khac nhau tren drop-down, dung nhu DocumentTypeState.bas.
' - "Neu nguoi dung tao mot file MOI hoac khong doc duoc Co quan soan thao/Co quan chu quan cua
'   van ban thi su dung gia tri DA NHO" (Bo sung Viettel Dang.md). Nghia la lua chon thu cong tren
'   MOT tai lieu phai con hieu luc lam gia tri du phong cho MOT TAI LIEU KHAC mo sau do trong CUNG
'   phien Word -- day la ly do KHONG dung SessionState (per-tai lieu) cho gia tri nay.
' "Ghi nho lua chon thu cong" CHI TRONG PHIEN WORD -- khong ghi dia, khong ghi vao tai lieu
' (Document.Variables/OOXML). Dong Word la mat, dung y "moi lan mo add-in coi nhu file moi hoan
' toan" va docs/architecture/00-kien-truc.md muc 5 ("khong luu du lieu nguoi dung"). Xem
' docs/design/03-thuat-toan-nhan-dien-va-chinh-ta.md muc B1.
' Mac dinh khi CHUA ai chon gi (ca ActiveDocument lan gia tri nho toan cuc): "ND30" -- dung quy
' dinh pho bien nhat, an toan nhat de khong lam sai lech kiem tra khi chua ro.
'==============================================================
Option Explicit

Private Const VAR_REGIME As String = "ChuanHoaTheThuc_QuyDinh"
Private Const DEFAULT_REGIME As String = "ND30"

' Gia tri TOAN CUC cho ca phien VBA -- xem ghi chu dau file ve ly do KHONG qua SessionState. Empty
' = chua ai TU CHON tay lan nao trong phien nay.
Private mRememberedRegime As Variant

' Ma quy dinh HIEN TAI cua ActiveDocument -- "ND30" neu tai lieu nay chua tung duoc chon/nhan dien
' (drop-down "Quy dinh" moi mo, hoac tai lieu vua tao/mo).
Public Function GetSelectedRegime() As String
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RegimeState.GetSelectedRegime")
    On Error GoTo ErrHandler
    Dim v As Variant
    v = SessionState.GetValue(VAR_REGIME)
    If IsEmpty(v) Then
        GetSelectedRegime = DEFAULT_REGIME
    Else
        GetSelectedRegime = CStr(v)
    End If
    Exit Function
ErrHandler:
    GetSelectedRegime = DEFAULT_REGIME
End Function

' Nguoi dung TU CHON tay tren drop-down "Quy dinh" (RibbonCallbacks.OnChonQuyDinh, Dot 2 - chua
' wire trong dot nay). Ghi CA HAI:
Public Sub SetSelectedRegimeManual(ByVal regimeCode As String)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RegimeState.SetSelectedRegimeManual")
    On Error GoTo ErrHandler
    SessionState.SetValue VAR_REGIME, regimeCode
    mRememberedRegime = regimeCode
    Exit Sub
ErrHandler:
    ' Khong co ActiveDocument hop le - bo qua, khong lam gian doan mot thao tac phu.
End Sub

Public Sub AutoDetectAndStore(ByVal regimeCode As String)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RegimeState.AutoDetectAndStore")
    On Error GoTo ErrHandler
    SessionState.SetValue VAR_REGIME, regimeCode
    Exit Sub
ErrHandler:
    ' Nhu SetSelectedRegimeManual - khong lam gian doan luong goi.
End Sub

' Goi tu RegimeDetector.DetectRegime khi nhan dien KHONG chac chan (confident=False - khong doc
' duoc Co quan soan thao/Co quan chu quan) - dung gia tri da nho TOAN CUC lam du phong, GHI cho
' ActiveDocument (giong AutoDetectAndStore, nhung nguon gia tri khac).
Public Function RememberedRegimeCode() As String
    If IsEmpty(mRememberedRegime) Then
        RememberedRegimeCode = DEFAULT_REGIME
    Else
        RememberedRegimeCode = CStr(mRememberedRegime)
    End If
End Function

Private Function CodeForIndex(ByVal idx As Long) As String
    Select Case idx
        Case 1: CodeForIndex = "VIETTEL"
        Case 2: CodeForIndex = "DANG"
        Case Else: CodeForIndex = "ND30"
    End Select
End Function

Private Function IndexForCode(ByVal code As String) As Long
    Select Case UCase$(code)
        Case "VIETTEL": IndexForCode = 1
        Case "DANG": IndexForCode = 2
        Case Else: IndexForCode = 0
    End Select
End Function

' getSelectedItemIndex cua <dropDown id="ddQuyDinh"> (RibbonCallbacks.GetSelectedIndexQuyDinh).
Public Function GetSelectedIndex() As Long
    GetSelectedIndex = IndexForCode(GetSelectedRegime())
End Function

' onAction cua <dropDown id="ddQuyDinh"> khi nguoi dung tu chon (RibbonCallbacks. OnChonQuyDinh) -
' ghi CA HAI nhu SetSelectedRegimeManual (lua chon rieng cho ActiveDocument + gia tri nho toan cuc
' cho tai lieu khac mo sau, xem ghi chu dau file).
Public Sub SetSelectedIndexManual(ByVal idx As Long)
    SetSelectedRegimeManual CodeForIndex(idx)
End Sub

' Chay nhan dien che do cho ActiveDocument roi tu ghi + ve lai RIENG drop-down "Quy dinh"
' (InvalidateQuyDinh - xem ghi chu 26/8/2026 duoi day ve ly do KHONG con tu ve lai toan bo ribbon
' o day nua). Cac nut phu thuoc che do khac (Co chu 13/14/15, "Loai:") duoc ve lai boi lan
' Invalidate toan bo CUOI CUNG cua DataReader.RunCore, khong can lam o day. Goi tu
' DataReader.RunCore (nut "Doc du lieu"), cung khuon voi DocumentTypeState.DetectAndAutoStore.
' snapshot: Optional - truyen vao neu noi goi DA CO SAN mot ban chup (tranh chup lap).
Public Sub DetectAndAutoStore(Optional ByVal snapshot As Object = Nothing)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RegimeState.DetectAndAutoStore")
    On Error GoTo ErrHandler
    If snapshot Is Nothing Then Set snapshot = DocumentSnapshot.CaptureDocument()
    Dim r As Object: Set r = RegimeDetector.DetectRegime(snapshot)
    If CBool(r("Confident")) Then
        AutoDetectAndStore CStr(r("RegimeCode"))
    Else
        AutoDetectAndStore RememberedRegimeCode()
    End If
    RibbonCallbacks.InvalidateQuyDinh
    Exit Sub
ErrHandler:
    ' Tinh nang phu, khong lam gian doan luong goi ("Doc du lieu").
End Sub
