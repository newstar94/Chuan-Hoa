Attribute VB_Name = "RibbonCallbacks"
Option Explicit

' Module nay KHONG con giu con tro IRibbonUI. Toan bo quyen so huu con tro va moi loi goi
' Invalidate/InvalidateControl da chuyen sang RibbonHandle.bas (xem ghi chu dau file do: ba lop
' kiem soat cho chuoi crash 0xc0000005). build/check-invariants.ps1 lam cong kiem tra luc build:
' khong file nao ngoai RibbonHandle.bas duoc cham vao IRibbonUI nua.

' Chuoi hien thi cua nut 1.2. VBA khong cho goi ChrW trong bieu thuc Const nen phai dung bien cap
' module dien qua EnsureTexts - cung cach frmWarning.frm va EncodingConverter.bas lam. Moi khai
' bao cap module PHAI dung TRUOC toan bo Sub/Function (quy tac VBA, xem ghi chu dau Utils.bas). se
' ra soat lai toan bo chuoi giao dien.
Private mTextsReady As Boolean
Private TITLE_ENCODING As String
Private MSG_ALREADY_UNICODE As String
Private MSG_CONVERTED As String
Private MSG_UNMAPPED_SUFFIX As String
Private MSG_VNI_SKIPPED_PREFIX As String
Private MSG_VNI_SKIPPED_SUFFIX As String
Private MSG_ENCODING_ERROR As String

Private TITLE_KIEU_OA_UY As String
Private TITLE_KIEU_OA_UY2 As String

Private TITLE_KIEU_I As String
Private TITLE_KIEU_Y As String

' Chuoi "Dang kiem tra... (n/tong)" cua thanh trang thai luc "Kiem tra" dang chay.
Private TEXT_CHECKING_PREFIX As String

' Dia chi trang gop y cua san pham. Thuan ASCII nen khai bao duoc bang Const that; chuoi bao loi
' di kem co dau nen phai qua EnsureTexts nhu moi chuoi tieng Viet khac trong file nay.
Public Const FEEDBACK_URL As String = "https://ngoctien.id.vn/chuan-hoa-the-thuc/feedback"
Private MSG_CANNOT_OPEN_FEEDBACK As String

' Dia chi trang chu san pham (nut "Kiem tra phien ban moi") - trang huong dan su dung + thong bao
' phien ban moi, cung mien voi FEEDBACK_URL nhung khong co /feedback.
Public Const HOMEPAGE_URL As String = "https://ngoctien.id.vn/chuan-hoa-the-thuc"
Private MSG_CANNOT_OPEN_HOMEPAGE As String

' Vet dau chay RIENG cho dropdown "Loai:" - Word chet 0xc0000005 trong dung mot dot bung no
' getItemLabel lien tiep khi liet ke dropdown nay (da tai dien hai lan: 23h15 26/8 va 5h36 27/8,
' log ca hai lan deu dung GIUA dot bung no nay, khong phai o RibbonHandle.TouchRibbon - khac chuoi
' crash IRibbonUI da chan o RibbonHandle.bas). KHONG the cham vao IRibbonUI o day (che boi check-
' invariants.ps1), nen dung cung ky thuat vet dau chay nhung o file/bien RIENG: neu phien truoc
' chet giua luc liet ke dropdown nay, phien nay bo qua han vong lap tra cuu
' (RuleLoader.GetDocTypeAbbreviations) cho MOI chi so, chi tra ve nhan rut gon - giam toi da khoi
' luong lam viec trong dung dot bung no da hai lan lam Word chet.
Private Const LOAIVB_GUARD_FILE_NAME As String = "ChuanHoaTheThuc-loaivb-guard.tmp"
Private mLoaiVbGuardReady As Boolean
Private mLoaiVbBlocked As Boolean

Private Sub EnsureTexts()
    If mTextsReady Then Exit Sub

    TITLE_ENCODING = "Chuy" & ChrW(&H1EC3) & "n b" & ChrW(&H1EA3) & "ng m"
    TITLE_ENCODING = TITLE_ENCODING & ChrW(&HE3) & " sang Unicode"

    MSG_ALREADY_UNICODE = "V" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n " & ChrW(&H111) & _
        ChrW(&HE3) & " d" & ChrW(&HF9) & "ng b" & ChrW(&H1EA3) & "ng m" & ChrW(&HE3) & _
        " Unicode"

    MSG_CONVERTED = ChrW(&H110) & ChrW(&HE3) & " chuy" & ChrW(&H1EC3) & "n sang Unicode v" & _
        ChrW(&HE0) & " d" & ChrW(&HF9) & "ng font Times New Roman"

    MSG_UNMAPPED_SUFFIX = " k" & ChrW(&HFD) & " t" & ChrW(&H1EF1) & " kh"
    MSG_UNMAPPED_SUFFIX = MSG_UNMAPPED_SUFFIX & ChrW(&HF4) & "ng c" & ChrW(&HF3) & " trong b" & ChrW(&H1EA3)
    MSG_UNMAPPED_SUFFIX = MSG_UNMAPPED_SUFFIX & "ng " & ChrW(&HE1) & "nh x" & ChrW(&H1EA1) & ", "
    MSG_UNMAPPED_SUFFIX = MSG_UNMAPPED_SUFFIX & ChrW(&H111) & ChrW(&HE3) & " gi" & ChrW(&H1EEF) & " nguy"
    MSG_UNMAPPED_SUFFIX = MSG_UNMAPPED_SUFFIX & ChrW(&HEA) & "n."

    MSG_VNI_SKIPPED_PREFIX = "B" & ChrW(&H1ECF) & " qua "

    MSG_VNI_SKIPPED_SUFFIX = " " & ChrW(&H111) & "o" & ChrW(&H1EA1) & "n d"
    MSG_VNI_SKIPPED_SUFFIX = MSG_VNI_SKIPPED_SUFFIX & ChrW(&HF9) & "ng b" & ChrW(&H1EA3) & "ng m" & ChrW(&HE3)
    MSG_VNI_SKIPPED_SUFFIX = MSG_VNI_SKIPPED_SUFFIX & " VNI " & ChrW(&H2014) & " b" & ChrW(&H1EA3) & "ng m"
    MSG_VNI_SKIPPED_SUFFIX = MSG_VNI_SKIPPED_SUFFIX & ChrW(&HE3) & " n" & ChrW(&HE0) & "y ch" & ChrW(&H1B0)
    MSG_VNI_SKIPPED_SUFFIX = MSG_VNI_SKIPPED_SUFFIX & "a " & ChrW(&H111) & ChrW(&H1B0) & ChrW(&H1EE3) & "c h"
    MSG_VNI_SKIPPED_SUFFIX = MSG_VNI_SKIPPED_SUFFIX & ChrW(&H1ED7) & " tr" & ChrW(&H1EE3) & "."

    MSG_ENCODING_ERROR = ChrW(&H110) & ChrW(&HE3) & " x" & ChrW(&H1EA3) & "y ra l"
    MSG_ENCODING_ERROR = MSG_ENCODING_ERROR & ChrW(&H1ED7) & "i khi chuy" & ChrW(&H1EC3) & "n " & ChrW(&H111)
    MSG_ENCODING_ERROR = MSG_ENCODING_ERROR & ChrW(&H1ED5) & "i b" & ChrW(&H1EA3) & "ng m" & ChrW(&HE3)
    MSG_ENCODING_ERROR = MSG_ENCODING_ERROR & "."

    TITLE_KIEU_OA_UY = "Ki" & ChrW(&H1EC3) & "u o" & ChrW(&HE0) & ", u" & ChrW(&HFD)
    TITLE_KIEU_OA_UY2 = "Ki" & ChrW(&H1EC3) & "u " & ChrW(&HF2) & "a, " & ChrW(&HFA) & "y"

    TITLE_KIEU_I = "Ki" & ChrW(&H1EC3) & "u i"
    TITLE_KIEU_Y = "Ki" & ChrW(&H1EC3) & "u y"

    TEXT_CHECKING_PREFIX = ChrW(&H110) & "ang ki" & ChrW(&H1EC3) & "m tra" & ChrW(&H2026) & " ("

    MSG_CANNOT_OPEN_FEEDBACK = "Kh" & ChrW(&HF4) & "ng m" & ChrW(&H1EDF) & " " & ChrW(&H111) & _
        ChrW(&H1B0) & ChrW(&H1EE3) & "c trang g" & ChrW(&H1EED) & "i ph" & ChrW(&H1EA3) & "n h" & _
        ChrW(&H1ED3) & "i. B" & ChrW(&H1EA1) & "n c" & ChrW(&HF3) & " th" & ChrW(&H1EC3) & " m" & _
        ChrW(&H1EDF) & " th" & ChrW(&H1EE7) & " c" & ChrW(&HF4) & "ng " & ChrW(&H111) & _
        ChrW(&H1ECB) & "a ch" & ChrW(&H1EC9) & " sau:"

    MSG_CANNOT_OPEN_HOMEPAGE = "Kh" & ChrW(&HF4) & "ng m" & ChrW(&H1EDF) & " " & ChrW(&H111) & _
        ChrW(&H1B0) & ChrW(&H1EE3) & "c trang ch" & ChrW(&H1EE7) & " s" & ChrW(&H1EA3) & "n ph" & _
        ChrW(&H1EA9) & "m. B" & ChrW(&H1EA1) & "n c" & ChrW(&HF3) & " th" & ChrW(&H1EC3) & " m" & _
        ChrW(&H1EDF) & " th" & ChrW(&H1EE7) & " c" & ChrW(&HF4) & "ng " & ChrW(&H111) & _
        ChrW(&H1ECB) & "a ch" & ChrW(&H1EC9) & " sau:"

    mTextsReady = True
End Sub

Private Function ProductName() As String
    ProductName = "Chu" & ChrW(&H1EA9) & "n h" & ChrW(&HF3) & "a th" & ChrW(&H1EC3) & _
        " th" & ChrW(&H1EE9) & "c"
End Function

' Goi tu RibbonHandle.RibbonOnLoad ngay sau khi con tro IRibbonUI da duoc cat giu an toan.
' onLoad la callback Ribbon UI extensibility, Microsoft khuyen cao phai TRA VE NHANH - viec nang
' (dang ky su kien, chup tai lieu dang mo...) deu day ra qua Application.OnTime
' (ScheduleOnAnyDocumentOpen), khong chay dong bo ngay trong ham nay.
Public Sub OnRibbonLoaded()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnRibbonLoaded")

    DebugTrace.Log "RibbonCallbacks.OnRibbonLoaded", "Ribbon da nap, dang ky AppEvents + lich OnAnyDocumentOpen"

    ' Tu dong doi don vi do luong hien thi cua Word ("Show measurement in units of") sang Centimet
    ' moi lan Ribbon nap - thuoc tinh cap ung dung (registry Windows cua may, khac cac thiet lap
    ' add-in tu quan luu qua Document.Variables), an toan goi truc tiep dong bo. On Error Resume
    ' Next vi mot so cau hinh Word hiem gap co the tu choi thuoc tinh nay.
    On Error Resume Next
    Application.options.MeasurementUnit = wdCentimeters
    On Error GoTo 0

    ' Viec dau tien luon phai lam khi mo van ban la nhan dien loai van ban - dang ky su kien
    ' Application (AppEvents.cls qua AppEventsHost.bas) cho moi lan mo/tao tai lieu sau thoi diem
    ' nay, roi lich mot lan cho tai lieu dang mo san (neu co), vi thu tu nap add-in khong dam bao
    ' Document_Open cua tai lieu dau tien xay ra sau khi ribbon nap xong.
    AppEventsHost.EnsureWired
    On Error Resume Next
    If Not ActiveDocument Is Nothing Then AppEventsHost.ScheduleOnAnyDocumentOpen
    On Error GoTo 0
End Sub

Public Sub InvalidateRibbon(Optional ByVal callerLabel As String = "InvalidateRibbon")
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.InvalidateRibbon")
    RibbonHandle.RequestInvalidateAll callerLabel
End Sub

Public Sub OnDocDuLieu(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnDocDuLieu")
    DataReader.RunAndReport
End Sub

Public Sub OnLuuDocx(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnLuuDocx")
    If DocxConverter.SaveAsDocx() Then
        InvalidateRibbon "LuuDocx"
    End If
End Sub

Public Sub GetEnabledRequiresDocx(ByVal Control As IRibbonControl, ByRef returnedVal)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.GetEnabledRequiresDocx")
    RibbonHandle.BeginCallback
    On Error Resume Next
    returnedVal = DocxConverter.IsCurrentDocumentDocx()
    On Error GoTo 0
    RibbonHandle.EndCallback
End Sub

' "Cac nut trong cac nhom: Dinh dang, Bang bieu va hinh anh, Chinh ta va so, Hien thi, About se
' khong bi phu thuoc vao viec da 'Doc du lieu' hay chua. Chi can 'co tai lieu dang mo' la cac nut
' nay se khong bi mo"). THAY THE GetEnabledRequiresDataRead o ba nhom do (nhom "Hien thi" da dung
' GetEnabledHienThi tu truoc, nhom "About" von khong gate gi - khong doi). CHI con dieu kien DUY
' NHAT: co it nhat mot tai lieu dang mo - khong doi hoi.docx, khong doi hoi da "Doc du lieu".
' "Loai:" dropdown (ddLoaiVanBan) trong nhom "Khoi dong" KHONG thuoc pham vi nay - van giu
' GetEnabledRequiresDataRead cu.
Public Sub GetEnabledHasDocument(ByVal Control As IRibbonControl, ByRef returnedVal)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.GetEnabledHasDocument")
    RibbonHandle.BeginCallback
    On Error Resume Next
    returnedVal = (Application.Documents.count > 0)
    On Error GoTo 0
    RibbonHandle.EndCallback
End Sub

Public Sub GetEnabledChuyenDoiUnicode(ByVal Control As IRibbonControl, ByRef returnedVal)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.GetEnabledChuyenDoiUnicode")
    RibbonHandle.BeginCallback
    On Error Resume Next
    returnedVal = DocxConverter.IsCurrentDocumentDocx() And DataReadState.HasReadData() _
        And DataReadState.HasNonUnicodeEncoding()
    On Error GoTo 0
    RibbonHandle.EndCallback
End Sub

Public Sub GetEnabledLuuDocx(ByVal Control As IRibbonControl, ByRef returnedVal)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.GetEnabledLuuDocx")
    RibbonHandle.BeginCallback
    On Error Resume Next
    returnedVal = Not DocxConverter.IsCurrentDocumentDocx()
    On Error GoTo 0
    RibbonHandle.EndCallback
End Sub

Public Function IsDocEmpty() As Boolean
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.IsDocEmpty")
    On Error GoTo ErrHandler
    IsDocEmpty = (Len(Trim$(ActiveDocument.Content.text)) <= 1)
    Exit Function
ErrHandler:
    IsDocEmpty = True
End Function

Public Sub OnChuyenDoiUnicode(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnChuyenDoiUnicode")
    On Error GoTo ErrHandler
    EnsureTexts

    Dim detection As Object
    Set detection = EncodingConverter.DetectEncoding()

    If CLng(detection("nonUnicodeCount")) = 0 Then
        MsgBoxW.Show MSG_ALREADY_UNICODE, vbInformation, ProductName()
        Exit Sub
    End If

    Dim warning As Object
    Set warning = SafetyGuard.BuildWarning(SafetyGuard.HIGH_RISK_ENCODING_CONVERSION)

    frmWarning.ShowHighRisk TITLE_ENCODING, CStr(warning("whatWillHappen")), _
        CStr(warning("scope")), CStr(warning("undoability")), CStr(warning("saveReminder"))

    Select Case frmWarning.Result
        Case "saveAndRun"
            ActiveDocument.Save
        Case "runAnyway"
            ' Chay thang, khong luu.
        Case Else ' "cancel" hoac dong bang nut dieu khien he thong cua form
            Exit Sub
    End Select

    Dim Result As Object
    Set Result = EncodingConverter.ConvertToUnicode()

    ' ConvertToUnicode da tu hien MsgBox loi qua Utils.AbortOperation khi that bai giua chung -
    ' khong bao chong them mot lan nua.
    If Not CBool(Result("failed")) Then
        MsgBoxW.Show BuildEncodingSummary(Result), vbInformation, ProductName()
    End If
    Exit Sub
ErrHandler:
    MsgBoxW.Show MSG_ENCODING_ERROR & vbCrLf & Err.description, vbExclamation, ProductName()
End Sub

' Tom tat ket qua cho nguoi dung. so doan da chuyen, so ky tu khong anh xa duoc, so doan VNI bi bo
' qua.
Private Function BuildEncodingSummary(ByVal Result As Object) As String
    Dim msg As String
    msg = MSG_CONVERTED

    If CLng(Result("unmappedCharCount")) > 0 Then
        msg = msg & vbCrLf & CStr(Result("unmappedCharCount")) & MSG_UNMAPPED_SUFFIX
    End If

    If CLng(Result("skippedVniRunCount")) > 0 Then
        msg = msg & vbCrLf & MSG_VNI_SKIPPED_PREFIX & CStr(Result("skippedVniRunCount")) & MSG_VNI_SKIPPED_SUFFIX
    End If

    BuildEncodingSummary = msg
End Function

Public Sub OnKiemTra(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnKiemTra")
    On Error GoTo ErrHandler
    FindingReporter.RunFormatCheck
    Exit Sub
ErrHandler:
    MsgBoxW.Show ChrW(&H110) & ChrW(&HE3) & " x" & ChrW(&H1EA3) & "y ra l" & ChrW(&H1ED7) & "i khi ki" & ChrW(&H1EC3) & "m tra th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c." & vbCrLf & Err.description, _
        vbExclamation, ProductName()
End Sub

' Nut "Kiem tra chinh ta" - CUNG luong voi OnKiemTra o tren, chi khac nhom quy tac duoc chay
' (checklistGroup 13-14) va marker Word Comment - xem FindingReporter.RunSpellingCheck.
Public Sub OnKiemTraChinhTa(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnKiemTraChinhTa")
    On Error GoTo ErrHandler
    FindingReporter.RunSpellingCheck
    Exit Sub
ErrHandler:
    MsgBoxW.Show ChrW(&H110) & ChrW(&HE3) & " x" & ChrW(&H1EA3) & "y ra l" & ChrW(&H1ED7) & "i khi ki" & ChrW(&H1EC3) & "m tra ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3) & "." & vbCrLf & Err.description, _
        vbExclamation, ProductName()
End Sub

Public Sub GetEnabledKiemTra(ByVal Control As IRibbonControl, ByRef returnedVal)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.GetEnabledKiemTra")
    RibbonHandle.BeginCallback
    On Error Resume Next
    Dim requiresDocType As Boolean
    requiresDocType = (RegimeState.GetSelectedRegime() <> "DANG")
    returnedVal = DocxConverter.IsCurrentDocumentDocx() And DataReadState.HasReadData() _
        And (Not FindingReporter.IsScanning()) _
        And ((Not requiresDocType) Or (DocumentTypeState.GetSelectedIndex() <> 0)) _
        And Not IsDocEmpty()
    On Error GoTo 0
    RibbonHandle.EndCallback
End Sub

Public Sub InvalidateKiemTra()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.InvalidateKiemTra")
    RibbonHandle.RequestInvalidateControl "btnKiemTra"
    RibbonHandle.RequestInvalidateControl "btnKiemTraChinhTa"
End Sub

Public Sub GetEnabledLoaiVanBan(ByVal Control As IRibbonControl, ByRef returnedVal)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.GetEnabledLoaiVanBan")
    RibbonHandle.BeginCallback
    On Error Resume Next
    returnedVal = DocxConverter.IsCurrentDocumentDocx() And DataReadState.HasReadData() _
        And (RegimeState.GetSelectedRegime() <> "DANG")
    On Error GoTo 0
    RibbonHandle.EndCallback
End Sub

Private Function MaxDocTypeOrdinalForRegime(ByVal regime As String) As Long
    Dim best As Long: best = 0
    Dim entry As Variant
    For Each entry In RuleLoader.GetDocTypeAbbreviations()("documentTypes")
        Dim regimes As Object: Set regimes = entry("regimes")
        Dim r As Variant
        For Each r In regimes
            If CStr(r) = regime Then
                If CLng(entry("ordinal")) > best Then best = CLng(entry("ordinal"))
                Exit For
            End If
        Next r
    Next entry
    MaxDocTypeOrdinalForRegime = best
End Function

Public Sub GetItemCountLoaiVanBan(ByVal Control As IRibbonControl, ByRef returnedVal)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.GetItemCountLoaiVanBan")
    LoaiVbLoadGuardState
    RibbonHandle.BeginCallback
    On Error Resume Next
    Dim regime As String: regime = RegimeState.GetSelectedRegime()
    If regime = "DANG" Then
        returnedVal = 1 ' dropdown mo o che do DANG (GetEnabledLoaiVanBan=False) - chi can "- Chua xac dinh -".
    Else
        returnedVal = 2 + MaxDocTypeOrdinalForRegime(regime)
    End If
    On Error GoTo 0
    RibbonHandle.EndCallback
    ' Chan doan crash 0xc0000005 luc Word liet ke dropdown nay (23h15 26/8, 5h36 27/8/2026, xem
    ' CLAUDE.md) - xa nhat ky NGAY sau moi loi goi (khac phan lon callback khac): day la mot dot
    ' BUNG NO ~30+ loi goi lien tiep trong DUOI 1 giay khi mo tai lieu/nap ribbon - bo dem 25
    ' dong/0,2s cua DebugTrace co the bo lo dung diem Word chet neu chi xa theo nguong thong
    ' thuong.
    DebugTrace.Flush
End Sub

Public Sub GetItemLabelLoaiVanBan(ByVal Control As IRibbonControl, ByVal Index As Long, ByRef returnedVal)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.GetItemLabelLoaiVanBan")
    LoaiVbLoadGuardState
    RibbonHandle.BeginCallback
    On Error Resume Next
    If Index = 0 Then
        returnedVal = "- Ch" & ChrW(&H1B0) & "a x" & ChrW(&HE1) & "c " & ChrW(&H111) & ChrW(&H1ECB) & "nh -"
    ElseIf Index = 1 Then
        returnedVal = "C" & ChrW(&HF4) & "ng v" & ChrW(&H103) & "n"
    ElseIf mLoaiVbBlocked Then
        ' Phien truoc chet giua dot liet ke dropdown nay - bo han vong lap tra cuu, chi tra ve so
        ' thu tu de dropdown van dung duoc (khong con dung ten day du, xem LoaiVbLoadGuardState).
        returnedVal = "(" & CStr(Index) & ")"
    Else
        LoaiVbArmGuard "GetItemLabelLoaiVanBan Index=" & CStr(Index)
        DebugTrace.Flush
        Dim wantOrdinal As Long: wantOrdinal = Index - 1
        Dim entry As Variant
        For Each entry In RuleLoader.GetDocTypeAbbreviations()("documentTypes")
            If CLng(entry("ordinal")) = wantOrdinal Then
                returnedVal = CStr(entry("typeName"))
                Exit For
            End If
        Next entry
        LoaiVbDisarmGuard
    End If
    On Error GoTo 0
    RibbonHandle.EndCallback
    ' Xem ghi chu o GetItemCountLoaiVanBan - xa nhat ky ngay sau moi chi so de lan chet ke tiep
    ' (neu con) chi ro DUNG chi so cuoi cung da hoan thanh truoc khi Word chet.
    DebugTrace.Flush
End Sub

Private Function LoaiVbGuardFilePath() As String
    On Error Resume Next
    Dim tempDir As String
    tempDir = Environ$("TEMP")
    If Len(tempDir) = 0 Then tempDir = Environ$("TMP")
    If Len(tempDir) = 0 Then Exit Function
    LoaiVbGuardFilePath = tempDir & "\" & LOAIVB_GUARD_FILE_NAME
End Function

' Doc MOT LAN cho ca phien (goi tu ca hai callback, idempotent qua mLoaiVbGuardReady).
Private Sub LoaiVbLoadGuardState()
    If mLoaiVbGuardReady Then Exit Sub
    mLoaiVbGuardReady = True

    On Error Resume Next
    Dim p As String: p = LoaiVbGuardFilePath()
    If Len(p) = 0 Then Exit Sub
    If Len(Dir$(p)) = 0 Then Exit Sub
    Kill p

    mLoaiVbBlocked = True
    DebugTrace.Log "RibbonCallbacks.LoaiVbLoadGuardState", _
        "PHIEN TRUOC Word chet giua luc liet ke dropdown ""Loai:"" - dung ban rut gon (so thu tu) cho ca phien nay"
    DebugTrace.Flush
End Sub

Private Sub LoaiVbArmGuard(ByVal detail As String)
    On Error Resume Next
    Dim p As String: p = LoaiVbGuardFilePath()
    If Len(p) = 0 Then Exit Sub
    Dim fnum As Integer: fnum = FreeFile
    Open p For Output As #fnum
    Print #fnum, detail
    Close #fnum
End Sub

Private Sub LoaiVbDisarmGuard()
    On Error Resume Next
    Dim p As String: p = LoaiVbGuardFilePath()
    If Len(p) = 0 Then Exit Sub
    If Len(Dir$(p)) > 0 Then Kill p
End Sub

' Nut 1.4 â€” drop-down "Loai van ban" (bo sung thang 8/2026). getSelectedItemIndex/onAction la chu
' ky co dinh do Office quy dinh cho <dropDown> (ByRef returnedVal, va selectedId/ selectedIndex
' tuong ung item nguoi dung vua chon) - khong duoc doi ten tham so.
Public Sub GetSelectedIndexLoaiVanBan(ByVal Control As IRibbonControl, ByRef returnedVal)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.GetSelectedIndexLoaiVanBan")
    RibbonHandle.BeginCallback
    On Error Resume Next
    returnedVal = DocumentTypeState.GetSelectedIndex()
    On Error GoTo 0
    RibbonHandle.EndCallback
End Sub

Public Sub OnChonLoaiVanBan(ByVal Control As IRibbonControl, ByVal selectedId As String, ByVal selectedIndex As Long)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnChonLoaiVanBan")
    DocumentTypeState.SetSelectedIndexManual selectedIndex
    ' "Nut kiem tra chi hien khi drop-down duoc chon it nhat mot loai van ban").
    InvalidateKiemTra
End Sub

' FindingReporter goi vao day sau moi lan "Kiem tra" (DocumentTypeState.AutoDetectAndStore da cap
' nhat gia tri) de Ribbon ve lai drop-down dung lua chon vua tu nhan dien - cung ly do voi
' InvalidateKiemTra o tren.
Public Sub InvalidateLoaiVanBan()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.InvalidateLoaiVanBan")
    RibbonHandle.RequestInvalidateControl "ddLoaiVanBan"
End Sub

Public Sub GetSelectedIndexQuyDinh(ByVal Control As IRibbonControl, ByRef returnedVal)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.GetSelectedIndexQuyDinh")
    RibbonHandle.BeginCallback
    On Error Resume Next
    returnedVal = RegimeState.GetSelectedIndex()
    On Error GoTo 0
    RibbonHandle.EndCallback
End Sub

Public Sub OnChonQuyDinh(ByVal Control As IRibbonControl, ByVal selectedId As String, ByVal selectedIndex As Long)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnChonQuyDinh")
    RegimeState.SetSelectedIndexManual selectedIndex
    ' Doi che do co the doi ca dieu kien "Loai:"/"Kiem tra"/ba nut Co chu - ve lai toan bo ribbon
    ' (khong chi rieng control nay), cung nguyen tac voi OnDocDuLieu.
    RibbonCallbacks.InvalidateRibbon "OnChonQuyDinh"
End Sub

' RegimeState.DetectAndAutoStore goi vao day sau khi tu nhan dien (nut "Doc du lieu") de ve lai
' rieng drop-down nay - cung ly do voi InvalidateLoaiVanBan o tren.
Public Sub InvalidateQuyDinh()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.InvalidateQuyDinh")
    RibbonHandle.RequestInvalidateControl "ddQuyDinh"
End Sub

' ND30: 13 sang, 14 sang, 15 mo | VIETTEL: 13 mo, 14 sang, 15 mo | DANG: 13 mo, 14 sang, 15 sang
' Dung MO (getEnabled), KHONG dung AN (getVisible) - chot. Dieu kien nen giong het
' GetEnabledHasDocument (: ba nhom nay chi doi "co tai lieu dang mo") CONG THEM dieu kien che do o
' day - dung CHUNG cho ca hai nut doc lap (btnCoChu1x) LAN muc submenu tuong ung trong "Dung bo
' Style" (btnDungBoStyleCo1x), vi hai cap nut nay LUON phai dong bo voi nhau (: bam "Dung bo
' Style" da tu ap luon co chu cho van ban co san).
Public Sub GetEnabledCoChu13(ByVal Control As IRibbonControl, ByRef returnedVal)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.GetEnabledCoChu13")
    RibbonHandle.BeginCallback
    On Error Resume Next
    returnedVal = (Application.Documents.count > 0) And (RegimeState.GetSelectedRegime() = "ND30")
    On Error GoTo 0
    RibbonHandle.EndCallback
End Sub

Public Sub GetEnabledCoChu14(ByVal Control As IRibbonControl, ByRef returnedVal)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.GetEnabledCoChu14")
    RibbonHandle.BeginCallback
    On Error Resume Next
    ' Ca ba che do deu dung duoc bo "Co chu 14" (set1) - xem fontSizeSetKeys trong
    ' shared/rules/quy-dinh-che-do.json.
    returnedVal = (Application.Documents.count > 0)
    On Error GoTo 0
    RibbonHandle.EndCallback
End Sub

Public Sub GetEnabledCoChu15(ByVal Control As IRibbonControl, ByRef returnedVal)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.GetEnabledCoChu15")
    RibbonHandle.BeginCallback
    On Error Resume Next
    returnedVal = (Application.Documents.count > 0) And (RegimeState.GetSelectedRegime() = "DANG")
    On Error GoTo 0
    RibbonHandle.EndCallback
End Sub

' ComplianceChecker.RunAllChecks goi Application.Run "RibbonCallbacks.NotifyScanProgress"
' (options("OnProgressFunctionName"), xem FindingReporter.RunCheckAndReport) de bao tien do -
' Application.Run chi goi tin cay thu tuc trong STANDARD MODULE (khong goi duoc Sub nam trong
' UserForm), nen cau noi phai dat o day.
Public Sub NotifyScanProgress(ByVal completed As Long, ByVal total As Long)
    On Error Resume Next
    EnsureTexts
    Dim progressText As String
    progressText = TEXT_CHECKING_PREFIX & CStr(completed) & "/" & CStr(total) & ")"
    Application.StatusBar = progressText
    ' Hop thoai "Dang xu ly..." (Utils.BeginOperation da hien san) cap nhat theo tien do quet -
    ' frmProcessing.SetText tu bo qua neu form chua duoc Show (On Error Resume Next o dau ham).
    frmProcessing.SetText progressText
    DoEvents
End Sub

' --- Nhom 2 â€” Dinh dang
Public Sub OnDinhDangTrangGiay(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnDinhDangTrangGiay")
    ' PageFormatter.ApplyPageSetup tu xu ly loi va hien MsgBox qua Utils.AbortOperation - khong
    ' can boc them ErrHandler o day (khac OnKiemTra/OnGioiThieu goi thang UserForm). Truyen che do
    ' dang chon de DANG dung le co dinh 20/20/30/15 mm rieng.
    PageFormatter.ApplyPageSetup RegimeState.GetSelectedRegime()
End Sub

Public Sub OnChenTrangNgang(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnChenTrangNgang")
    PageBreakInserter.InsertLandscapePage
End Sub

Public Sub OnChenTrangDoc(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnChenTrangDoc")
    PageBreakInserter.InsertPortraitPage
End Sub

' Nut "Dung bo Style" la <menu> voi hai muc con "Co chu 13"/"Co chu 14" - chon ngay trong ribbon,
' khong qua hop thoai rieng.
Public Sub OnDungBoStyle14(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnDungBoStyle14")
    BuildStylesForSize "set1"
End Sub

Public Sub OnDungBoStyle13(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnDungBoStyle13")
    BuildStylesForSize "set2"
End Sub

' (26/8/2026): bo "Co chu 15", danh rieng cho che do DANG (xem $regime trong shared/rules/thong-
' so-the-thuc.json, fontSizeSets.set3).
Public Sub OnDungBoStyle15(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnDungBoStyle15")
    BuildStylesForSize "set3"
End Sub

' sizeSetKey: "set1" = Co chu 14, "set2" = Co chu 13, "set3" = Co chu 15 (dong quy uoc voi
' OnCoChu13/OnCoChu14/OnCoChu15 o duoi). StyleBuilder.BuildStyles tu xu ly loi va hien MsgBox qua
' Utils.AbortOperation - khong can boc them cho buoc do (giong OnDinhDangTrangGiay).
Private Sub BuildStylesForSize(ByVal sizeSetKey As String)
    StyleBuilder.BuildStyles sizeSetKey

    ' van ban DA CO NOI DUNG (khong phai tai lieu trang moi tao) thi ap luon co chu toan van ban
    ' theo DUNG bo vua chon o day - dung noi ap dung co chu DUY NHAT
    ' (TextFormatter.ApplyFontSizeWholeDocument), tranh nguoi dung phai bam them nut "Co chu
    ' 13"/"Co chu 14" mot lan nua ngay sau khi Dung bo Styles.
    If Len(Trim$(ActiveDocument.Content.text)) > 0 Then
        TextFormatter.ApplyFontSizeWholeDocument sizeSetKey, RegimeState.GetSelectedRegime()
    End If
End Sub

Public Sub OnChenSoTrang(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnChenSoTrang")
    PageNumberFormatter.InsertPageNumbers
End Sub

' TextFormatter.ApplyFontSizeWholeDocument tu xu ly loi va hien MsgBox qua Utils.AbortOperation
' - khong can boc them ErrHandler o day (giong OnDinhDangTrangGiay). "set2" = Co chu 13
' "set1" = Co chu 14 (dong quy uoc voi OnDungBoStyles).
Public Sub OnCoChu13(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnCoChu13")
    TextFormatter.ApplyFontSizeWholeDocument "set2", RegimeState.GetSelectedRegime()
End Sub

Public Sub OnCoChu14(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnCoChu14")
    TextFormatter.ApplyFontSizeWholeDocument "set1", RegimeState.GetSelectedRegime()
End Sub

' (26/8/2026): "Co chu 15" - bo cho che do DANG, xem ghi chu o OnDungBoStyle15.
Public Sub OnCoChu15(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnCoChu15")
    TextFormatter.ApplyFontSizeWholeDocument "set3", RegimeState.GetSelectedRegime()
End Sub

' Nut "Keep with next" - ParagraphFormatter. ApplyKeepWithNextAtCursor tu xu ly loi va hien MsgBox
' qua Utils.AbortOperation - khong can boc them ErrHandler o day (giong OnDinhDangTrangGiay).
Public Sub OnKeepWithNext(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnKeepWithNext")
    ParagraphFormatter.ApplyKeepWithNextAtCursor
End Sub

Public Sub OnCoChu(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnCoChu")
    TextFormatter.CondenseSpacing
End Sub

Public Sub OnGianChuNormal(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnGianChuNormal")
    TextFormatter.ResetSpacing
End Sub

Public Sub OnGianChuRa(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnGianChuRa")
    TextFormatter.ExpandSpacing
End Sub

Public Sub OnChuanHoaBang(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnChuanHoaBang")
    TableFormatter.NormalizeTables
End Sub

Public Sub OnLapDongTieuDe(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnLapDongTieuDe")
    TableFormatter.RepeatHeaderRows
End Sub

Public Sub OnCanDinhO(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnCanDinhO")
    TableFormatter.AlignCellsTop
End Sub

Public Sub OnCanGiuaO(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnCanGiuaO")
    TableFormatter.AlignCellsCenter
End Sub

' ImageFormatter.NormalizeImages tu xu ly loi va hien MsgBox qua Utils.AbortOperation - khong can
' boc them ErrHandler o day (giong OnChuanHoaBang).
Public Sub OnChuanHoaAnh(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnChuanHoaAnh")
    ImageFormatter.NormalizeImages
End Sub

' ExcelPasteCleaner.CleanExcelPasteWhitespace tu xu ly loi va hien MsgBox qua Utils.AbortOperation
' - khong can boc them ErrHandler o day (giong OnChuanHoaBang).
Public Sub OnXoaKyTuThuaBangExcel(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnXoaKyTuThuaBangExcel")
    ExcelPasteCleaner.CleanExcelPasteWhitespace
End Sub

Public Sub OnKieuOaUy(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnKieuOaUy")
    EnsureTexts
    ToneNormalizer.ApplyToneStyle "toMainVowel", TITLE_KIEU_OA_UY
End Sub

Public Sub OnKieuOaUy2(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnKieuOaUy2")
    EnsureTexts
    ToneNormalizer.ApplyToneStyle "toFirstVowel", TITLE_KIEU_OA_UY2
End Sub

' IyNormalizer.ApplyIyStyle tu xu ly loi va hien MsgBox qua Utils.AbortOperation - khong can boc
' them ErrHandler o day (giong OnKieuOaUy).
Public Sub OnKieuI(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnKieuI")
    EnsureTexts
    IyNormalizer.ApplyIyStyle "toI", TITLE_KIEU_I
End Sub

Public Sub OnKieuY(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnKieuY")
    EnsureTexts
    IyNormalizer.ApplyIyStyle "toY", TITLE_KIEU_Y
End Sub

' TrailingPageRemover.RemoveTrailingPage tu xu ly loi va hien MsgBox qua Utils.AbortOperation hoac
' truc tiep (khong co trang thua / khong xu ly duoc) - khong can boc them ErrHandler o day (giong
' OnChuanHoaBang).
Public Sub OnXoaTrangThua(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnXoaTrangThua")
    TrailingPageRemover.RemoveTrailingPage
End Sub

' DecimalSeparatorConverter.ConvertDecimalSeparators tu xu ly loi va hien MsgBox qua
' Utils.AbortOperation - khong can boc them ErrHandler o day (giong OnChuanHoaBang).
Public Sub OnDoiDauThapPhan(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnDoiDauThapPhan")
    DecimalSeparatorConverter.ConvertDecimalSeparators
End Sub

' --- Nhom 7 â€” Chen QR ---------------------------- QrCodeGenerator.InsertQrCode tu xu ly loi va
' hien MsgBox qua Utils.AbortOperation - khong can boc them ErrHandler o day (giong
' OnDinhDangTrangGiay/OnChenDuongKe).
Public Sub OnChenQrCode(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnChenQrCode")
    QrCodeGenerator.InsertQrCode
End Sub

Public Sub GetEnabledHienThi(ByVal Control As IRibbonControl, ByRef returnedVal)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.GetEnabledHienThi")
    RibbonHandle.BeginCallback
    On Error Resume Next
    returnedVal = ViewOptions.HasActiveWindow()
    On Error GoTo 0
    RibbonHandle.EndCallback
End Sub

Public Sub GetPressedRanhGioiVanBan(ByVal Control As IRibbonControl, ByRef returnedVal)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.GetPressedRanhGioiVanBan")
    RibbonHandle.BeginCallback
    On Error Resume Next
    returnedVal = ViewOptions.GetShowTextBoundaries()
    On Error GoTo 0
    RibbonHandle.EndCallback
End Sub

Public Sub OnRanhGioiVanBan(ByVal Control As IRibbonControl, ByVal pressed As Boolean)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnRanhGioiVanBan")
    ViewOptions.SetShowTextBoundaries pressed
End Sub

Public Sub GetPressedDauGoc(ByVal Control As IRibbonControl, ByRef returnedVal)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.GetPressedDauGoc")
    RibbonHandle.BeginCallback
    On Error Resume Next
    returnedVal = ViewOptions.GetShowCropMarks()
    On Error GoTo 0
    RibbonHandle.EndCallback
End Sub

Public Sub OnDauGoc(ByVal Control As IRibbonControl, ByVal pressed As Boolean)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnDauGoc")
    ViewOptions.SetShowCropMarks pressed
End Sub

Public Sub GetPressedKyHieuSoanThao(ByVal Control As IRibbonControl, ByRef returnedVal)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.GetPressedKyHieuSoanThao")
    RibbonHandle.BeginCallback
    On Error Resume Next
    returnedVal = ViewOptions.GetShowAllMarks()
    On Error GoTo 0
    RibbonHandle.EndCallback
End Sub

Public Sub OnKyHieuSoanThao(ByVal Control As IRibbonControl, ByVal pressed As Boolean)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnKyHieuSoanThao")
    ViewOptions.SetShowAllMarks pressed
End Sub

' --- Nhom 8 â€” About ---------------------------------------------------------- P6 â€” tai dung
' frmWarning.ShowAbout, khong tao form frmAbout rieng.
Public Sub OnGioiThieu(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnGioiThieu")
    On Error GoTo ErrHandler
    frmWarning.ShowAbout
    Exit Sub
ErrHandler:
    MsgBoxW.Show ChrW(&H110) & ChrW(&HE3) & " x" & ChrW(&H1EA3) & "y ra l" & ChrW(&H1ED7) & "i khi m" & ChrW(&H1EDF) & " h" & ChrW(&H1ED9) & "p tho" & ChrW(&H1EA1) & "i gi" & ChrW(&H1EDB) & "i thi" & ChrW(&H1EC7) & "u." & vbCrLf & Err.description, _
        vbExclamation, ProductName()
End Sub

Public Sub OnKiemTraPhienBanMoi(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnKiemTraPhienBanMoi")
    On Error GoTo ErrHandler
    EnsureTexts
    If Not UrlOpener.OpenUrl(HOMEPAGE_URL) Then
        Dim detail As String: detail = UrlOpener.LastFailureDetail()
        Dim msg As String: msg = MSG_CANNOT_OPEN_HOMEPAGE & vbCrLf & HOMEPAGE_URL
        If Len(detail) > 0 Then msg = msg & vbCrLf & "(" & detail & ")"
        MsgBoxW.Show msg, vbExclamation, ProductName()
    End If
    Exit Sub
ErrHandler:
    MsgBoxW.Show MSG_CANNOT_OPEN_HOMEPAGE & vbCrLf & Err.description, vbExclamation, ProductName()
End Sub

Public Sub OnGuiPhanHoi(ByVal Control As IRibbonControl)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonCallbacks.OnGuiPhanHoi")
    On Error GoTo ErrHandler
    EnsureTexts
    If Not UrlOpener.OpenUrl(FEEDBACK_URL) Then
        Dim detail As String: detail = UrlOpener.LastFailureDetail()
        Dim msg As String: msg = MSG_CANNOT_OPEN_FEEDBACK & vbCrLf & FEEDBACK_URL
        If Len(detail) > 0 Then msg = msg & vbCrLf & "(" & detail & ")"
        MsgBoxW.Show msg, vbExclamation, ProductName()
    End If
    Exit Sub
ErrHandler:
    MsgBoxW.Show MSG_CANNOT_OPEN_FEEDBACK & vbCrLf & Err.description, vbExclamation, ProductName()
End Sub
