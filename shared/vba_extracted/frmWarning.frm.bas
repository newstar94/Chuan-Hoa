Attribute VB_Name = "frmWarning"
Attribute VB_Base = "0{0AC05787-D892-48A1-ABCB-7283B48FD404}{4C047AFF-3532-4B64-B6E2-DD01CFB6D8F6}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

'==============================================================
' frmWarning â€” P2, P3 (cong chan), P4 (canh bao rui ro cao â€” bon phan theo ADR-007), P5
' (xac nhan don gian), P6 (Thong tin tien ich) va che do danh sach nhat ky phien.
' UserForm modal (Show vbModal). Moi lan mo la MOT trong nam che do (Gate/HighRisk/
' SimpleConfirm/About/LogList), dung lai HOAN TOAN moi lan Show nen xay dung control ngay trong
' than tung ham BuildXxx, khong can luc UserForm_Initialize.
' Control dung luc chay dung DialogControlSink.cls â€” sink nhan Owner=Me roi goi lien ket muon toi
' HandleControlClick cua CHINH form dang mo.
' MA HOA: moi chuoi co dau tieng Viet trong CODE (khong phai comment) viet bang ChrW(&Hxxxx) noi
' chuoi â€” VBComponents.Import doc file.frm theo ma trang ANSI, KHONG phai UTF-8. Comment (khong
' chay) giu nguyen khong dau.
' ============================================================================
' "Tat ca cac Popup hien tai deu khong dung Fluent design (vay nen qua xau)... Hay tim cach lam
' dep no"). Bon thay doi goc:
' 1. BO THANH TIEU DE CUA HE DIEU HANH, tu ve phan dau hop thoai bang Label
'   (WinApiFormStyle.MakeActiveWindowBorderless, goi tu UserForm_Activate). thanh tieu de UserForm
'   la cua so ANSI nen KHONG the hien dung dau tieng Viet bang bat ky cach nao (xem
'   WinApiFormStyle.bas) - con Label thi Unicode day du. Phan dau tu ve: tieu de CO DAU + nut dong
'   "x", keo duoc de di chuyen cua so (DialogControlSink.DragLabel).
' 2. NUT BAM VE BANG Label CHU KHONG PHAI CommandButton. Label dat duoc
'   BackColor/BorderStyle/BorderColor tu do -> nut phang that su. Moi nut gom HAI Label chong
'   nhau: mot lam nen (dung chieu cao nut) + mot lam chu (canh giua theo chieu doc, vi Label khong
'   co canh giua doc san) - ca hai cung BackColor nen nhin lien mot khoi; ca hai deu gan sink cung
'   mot tag.
' 3. NUT XEP MOT HANG NGANG, CANH PHAI trong dai chan hop thoai mau xam nhat co duong ke manh o
'   tren (dung bo cuc hop thoai Fluent/Office). Nhan tieng Viet dai co the vuot be ngang - do
'   TRUOC (MeasureTextWidth), khong du cho thi tu dong quay ve xep doc day du be ngang nhu truoc,
'   khong bao gio cat chu.
' 4. GAN "lbl.Font.Bold = False" hay "lbl.Font.Italic = False" DEU LAM CHU THANH DAM/NGHIENG (doc
'   lai chinh thuoc tinh do tra ve True!) - moi phep GAN vao Font.Bold/Font.Italic cua
'   MSForms.Label deu cho ket qua True bat ke gan gia tri gi; KHONG gan gi ca thi moi dung la
'   khong dam/khong nghieng. Do chinh la ham AddLabel ban cu ("...Font.Bold = bold" voi bold =
'   False, "...Font.Italic = italic" voi italic = False). QUY TAC TU NAY: chi gan
'   Font.Bold/Font.Italic KHI THAT SU muon dam/nghieng, khong bao gio gan gia tri False. Ban thiet
'   ke moi cung khong dung chu nghieng o bat ky dau.
' ============================================================================
'==============================================================
Option Explicit

' --- Kich thuoc, khoang cach (don vi point)
Private Const FORM_WIDTH_PT As Single = 320
Private Const PAD As Single = 16
Private Const HEADER_HEIGHT As Single = 42
Private Const FOOTER_PAD As Single = 11
Private Const GAP_SMALL As Single = 4
Private Const GAP_LARGE As Single = 12
Private Const BUTTON_HEIGHT As Single = 26
Private Const BUTTON_GAP As Single = 8
Private Const BUTTON_MIN_WIDTH As Single = 76
Private Const BUTTON_TEXT_PAD As Single = 22
Private Const CLOSE_BUTTON_SIZE As Single = 26
Private Const DIVIDER_HEIGHT As Single = 0.75
Private Const LOG_FRAME_MAX_HEIGHT As Single = 180

' --- Bac chu (khong con in nghieng o bat ky bac nao, xem ghi chu dau file)
Private Const FONT_SIZE_HEADER As Single = 12
Private Const FONT_SIZE_TITLE As Single = 11
Private Const FONT_SIZE_BODY As Single = 9
Private Const FONT_SIZE_LABEL As Single = 9
Private Const FONT_SIZE_CITATION As Single = 8
Private Const FONT_SIZE_ABOUT_NAME As Single = 16
Private Const FONT_SIZE_CLOSE As Single = 14
Private Const UI_FONT_NAME As String = "Segoe UI"

Private TEXT_PRODUCT_NAME As String
Private TEXT_BTN_CANCEL As String
Private TEXT_BTN_PROCEED_ANYWAY As String
Private TEXT_BTN_SAVE_AND_RUN As String
Private TEXT_BTN_RUN_ANYWAY As String
Private TEXT_BTN_CONTINUE As String
Private TEXT_BTN_STOP As String
Private TEXT_LABEL_WHAT_WILL_HAPPEN As String
Private TEXT_LABEL_SCOPE As String
Private TEXT_LABEL_UNDOABILITY As String
Private TEXT_LOG_TITLE As String
Private TEXT_LOG_SESSION_NOTICE As String
Private TEXT_LOG_EMPTY As String
Private TEXT_BTN_CLOSE As String
Private TEXT_LOG_WARNING_SUFFIX As String
Private TEXT_INTERNAL_ERROR_PREFIX As String
Private TEXT_ABOUT_TITLE As String
Private TEXT_ABOUT_VERSION_LABEL As String
Private TEXT_ABOUT_AUTHOR_LABEL As String
Private TEXT_ABOUT_AUTHOR_VALUE As String
Private TEXT_ABOUT_UPDATE_NOTICE As String
Private TEXT_CLOSE_GLYPH As String

' "1.0.0, ngay 26/08/2026". Dung ChrW cho "Ă " - xem ghi chu dau file ve ly do khong ghi ky tu co
' dau truc tiep.
Private APP_VERSION As String

' --- Bang mau Fluent (Office/Microsoft 365)
Private mColorPrimary As Long        ' chu chinh   #201F1E
Private mColorSecondary As Long      ' chu phu     #605E5C
Private mColorWarning As Long        ' canh bao    #D83B01
Private mColorAccent As Long         ' nhan manh   #0078D4
Private mColorAccentText As Long     ' chu tren nen nhan manh - trang
Private mColorButtonBg As Long       ' nut phu     trang
Private mColorButtonBorder As Long   ' vien nut phu #8A8886
Private mColorButtonText As Long     ' chu nut phu #201F1E
Private mColorFormBg As Long         ' nen the     trang
Private mColorFooterBg As Long       ' dai chan    #FAF9F8
Private mColorDivider As Long        ' duong ke    #E1DFDD

Public Result As String

Private mDynControlNames As Collection
Private mDynSinks As Collection
Private mNameCounter As Long
Private mCancelResult As String
Private mContentWidth As Single

' ============================================================================
' Vong doi form
' ============================================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler
    InitializeText

    mColorPrimary = RGB(32, 31, 30)
    mColorSecondary = RGB(96, 94, 92)
    mColorWarning = RGB(216, 59, 1)
    mColorAccent = RGB(0, 120, 212)
    mColorAccentText = RGB(255, 255, 255)
    mColorButtonBg = RGB(255, 255, 255)
    mColorButtonBorder = RGB(138, 136, 134)
    mColorButtonText = RGB(32, 31, 30)
    mColorFormBg = RGB(255, 255, 255)
    mColorFooterBg = RGB(250, 249, 248)
    mColorDivider = RGB(225, 223, 221)
    Me.BackColor = mColorFormBg

    ' CenterScreen (2) - VBA runtime tu canh giua MAN HINH VAT LY. Sau khi bo thanh tieu de
    ' (UserForm_Activate) con duoc canh giua lai MOT LAN NUA theo kich thuoc that (xem
    ' WinApiFormStyle.MakeActiveWindowBorderless).
    Me.StartUpPosition = 2 ' 2 = fmStartUpScreen (CenterScreen)
    SetClientWidth FORM_WIDTH_PT
    mContentWidth = FORM_WIDTH_PT - 2 * PAD

    Set mDynControlNames = New Collection
    Set mDynSinks = New Collection
    Exit Sub
ErrHandler:
    HandleFatalError "UserForm_Initialize", Err.description
End Sub

' bo thanh tieu de cua he dieu hanh ngay khi cua so that vua duoc tao (form dang la cua so kich
' hoat) - phan dau hop thoai da duoc tu ve bang Label trong BeginDialog.
Private Sub UserForm_Activate()
    WinApiFormStyle.MakeActiveWindowBorderless
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo ErrHandler
    If CloseMode = vbFormControlMenu Then
        Result = mCancelResult
        Cancel = True
        Me.Hide
    End If
    Exit Sub
ErrHandler:
    ' Loi o day khong duoc chan viec dong form - bo qua im lang.
End Sub

' Gan gia tri cho cac bien chuoi tieng Viet dung ChrW - xem ghi chu dau file.
Private Sub InitializeText()
    TEXT_PRODUCT_NAME = "Chu" & ChrW(&H1EA9) & "n h" & ChrW(&HF3) & "a th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    TEXT_BTN_CANCEL = "H" & ChrW(&H1EE7) & "y"
    TEXT_BTN_PROCEED_ANYWAY = "V" & ChrW(&H1EAB) & "n ki" & ChrW(&H1EC3) & "m tra"
    TEXT_BTN_SAVE_AND_RUN = "L" & ChrW(&H1B0) & "u r" & ChrW(&H1ED3) & "i ch" & ChrW(&H1EA1) & "y"
    TEXT_BTN_RUN_ANYWAY = "V" & ChrW(&H1EAB) & "n ch" & ChrW(&H1EA1) & "y"
    TEXT_BTN_CONTINUE = "Ti" & ChrW(&H1EBF) & "p t" & ChrW(&H1EE5) & "c"
    TEXT_BTN_STOP = ChrW(&H110) & ChrW(&H1EEB) & "ng l" & ChrW(&H1EA1) & "i"
    TEXT_LABEL_WHAT_WILL_HAPPEN = "Vi" & ChrW(&H1EC7) & "c s" & ChrW(&H1EBD) & " l" & ChrW(&HE0) & "m"
    TEXT_LABEL_SCOPE = "Ph" & ChrW(&H1EA1) & "m vi " & ChrW(&H1EA3) & "nh h" & ChrW(&H1B0) & ChrW(&H1EDF) & "ng"
    TEXT_LABEL_UNDOABILITY = "Kh" & ChrW(&H1EA3) & " n" & ChrW(&H103) & "ng ho" & ChrW(&HE0) & "n t" & ChrW(&H1EAF) & "c"
    TEXT_LOG_TITLE = "Nh" & ChrW(&H1EAD) & "t k" & ChrW(&HFD) & " thao t" & ChrW(&HE1) & "c trong phi" & ChrW(&HEA) & "n"
    TEXT_LOG_SESSION_NOTICE = "Nh" & ChrW(&H1EAD) & "t k" & ChrW(&HFD) & " ch" & ChrW(&H1EC9) & " t" & ChrW(&H1ED3) & "n t" & ChrW(&H1EA1) & "i trong phi" & ChrW(&HEA) & "n l" & ChrW(&HE0) & "m vi" & ChrW(&H1EC7) & "c, kh" & ChrW(&HF4) & "ng l" & ChrW(&H1B0) & "u ra file, kh" & ChrW(&HF4) & "ng r" & ChrW(&H1EDD) & "i kh" & ChrW(&H1ECF) & "i m" & ChrW(&HE1) & "y."
    TEXT_LOG_EMPTY = "Ch" & ChrW(&H1B0) & "a c" & ChrW(&HF3) & " thao t" & ChrW(&HE1) & "c n" & ChrW(&HE0) & "o trong phi" & ChrW(&HEA) & "n n" & ChrW(&HE0) & "y."
    TEXT_BTN_CLOSE = ChrW(&H110) & ChrW(&HF3) & "ng"
    TEXT_LOG_WARNING_SUFFIX = ChrW(&H26A0)
    TEXT_INTERNAL_ERROR_PREFIX = ChrW(&H110) & ChrW(&HE3) & " x" & ChrW(&H1EA3) & "y ra l" & ChrW(&H1ED7) & "i n" & ChrW(&H1ED9) & "i b" & ChrW(&H1ED9) & " ("
    TEXT_ABOUT_TITLE = "Th" & ChrW(&HF4) & "ng tin ti" & ChrW(&H1EC7) & "n " & ChrW(&HED) & "ch"
    TEXT_ABOUT_VERSION_LABEL = "Phi" & ChrW(&HEA) & "n b" & ChrW(&H1EA3) & "n: "
    APP_VERSION = "1.0.0, ng" & ChrW(&HE0) & "y 26/08/2026"
    TEXT_ABOUT_AUTHOR_LABEL = "T" & ChrW(&HE1) & "c gi" & ChrW(&H1EA3)
    TEXT_ABOUT_AUTHOR_VALUE = "Nguy" & ChrW(&H1EC5) & "n Ng" & ChrW(&H1ECD) & "c Ti" & ChrW(&H1EBF) & "n (tiennn.ict@gmail.com)"
    TEXT_ABOUT_UPDATE_NOTICE = "Ki" & ChrW(&H1EC3) & "m tra phi" & ChrW(&HEA) & "n b" & ChrW(&H1EA3) & "n m" & ChrW(&H1EDB) & "i t" & ChrW(&H1EA1) & "i https://ngoctien.id.vn"
    ' Dau nhan chuot (U+00D7) - co trong MOI phong chu he thong, khac cac glyph "x" kieu Segoe
    ' MDL2 (U+E8BB...) chi co tren may cai bo phong chu bieu tuong cua Windows 10+.
    TEXT_CLOSE_GLYPH = ChrW(&HD7)
End Sub

' P2, P3 â€” cong chan cua nut "Kiem tra" (muc 3 nhom 1 dac ta giao dien).
Public Sub BuildGate(ByVal titleText As String, ByVal messageText As String, ByVal actionLabel As String)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("frmWarning.BuildGate")
    On Error GoTo ErrHandler
    Dim y As Single
    y = BeginDialog(titleText, "cancel")

    Dim msgLbl As MSForms.LABEL
    Set msgLbl = AddLabel(messageText, y, FONT_SIZE_BODY, False, mColorPrimary)
    y = y + msgLbl.Height

    EndDialog y, Array(actionLabel, TEXT_BTN_PROCEED_ANYWAY, TEXT_BTN_CANCEL), _
        Array("result:action", "result:proceedAnyway", "result:cancel")
    Exit Sub
ErrHandler:
    HandleFatalError "BuildGate", Err.description
End Sub

Public Sub ShowGate(ByVal titleText As String, ByVal messageText As String, ByVal actionLabel As String)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("frmWarning.ShowGate")
    BuildGate titleText, messageText, actionLabel
    Me.Show vbModal
End Sub

' P4 â€” canh bao rui ro cao, bon phan bat buoc theo ADR-007. saveReminder = "" neu tai lieu
' da luu (khong hien dong nhac, khong co nut "Luu roi chay") - dung quy uoc cua
' SafetyGuard.BuildWarning.
Public Sub BuildHighRisk(ByVal titleText As String, ByVal whatWillHappen As String, _
        ByVal scopeText As String, ByVal undoability As String, ByVal saveReminder As String)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("frmWarning.BuildHighRisk")
    On Error GoTo ErrHandler
    Dim y As Single
    y = BeginDialog(titleText, "cancel")

    y = AddSectionBlock(TEXT_LABEL_WHAT_WILL_HAPPEN, whatWillHappen, y)
    y = AddSectionBlock(TEXT_LABEL_SCOPE, scopeText, y)
    y = AddSectionBlock(TEXT_LABEL_UNDOABILITY, undoability, y)
    y = y - GAP_LARGE

    Dim hasSaveReminder As Boolean
    hasSaveReminder = (Len(saveReminder) > 0)
    If hasSaveReminder Then
        ' Dai nhac luu tai lieu: nen do rat nhat + vien trai mau canh bao (mach Fluent
        ' "MessageBar") - de doc hon dong chu tran kem bieu tuong tam giac nhu ban cu.
        y = y + GAP_LARGE
        y = AddNoticeBar(saveReminder, y)
    End If

    If hasSaveReminder Then
        EndDialog y, Array(TEXT_BTN_SAVE_AND_RUN, TEXT_BTN_RUN_ANYWAY, TEXT_BTN_CANCEL), _
            Array("result:saveAndRun", "result:runAnyway", "result:cancel")
    Else
        EndDialog y, Array(TEXT_BTN_RUN_ANYWAY, TEXT_BTN_CANCEL), _
            Array("result:runAnyway", "result:cancel")
    End If
    Exit Sub
ErrHandler:
    HandleFatalError "BuildHighRisk", Err.description
End Sub

Public Sub ShowHighRisk(ByVal titleText As String, ByVal whatWillHappen As String, _
        ByVal scopeText As String, ByVal undoability As String, ByVal saveReminder As String)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("frmWarning.ShowHighRisk")
    BuildHighRisk titleText, whatWillHappen, scopeText, undoability, saveReminder
    Me.Show vbModal
End Sub

Public Sub BuildAbout()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("frmWarning.BuildAbout")
    On Error GoTo ErrHandler
    Dim y As Single
    y = BeginDialog(TEXT_ABOUT_TITLE, "closed")

    Dim nameLbl As MSForms.LABEL
    Set nameLbl = AddLabel(TEXT_PRODUCT_NAME, y, FONT_SIZE_ABOUT_NAME, True, mColorPrimary)
    y = y + nameLbl.Height + GAP_LARGE

    Dim verLbl As MSForms.LABEL
    Set verLbl = AddLabel(TEXT_ABOUT_VERSION_LABEL & APP_VERSION, y, FONT_SIZE_BODY, False, mColorSecondary)
    y = y + verLbl.Height + GAP_SMALL

    Dim authLbl As MSForms.LABEL
    Set authLbl = AddLabel(TEXT_ABOUT_AUTHOR_LABEL & ": " & TEXT_ABOUT_AUTHOR_VALUE, y, FONT_SIZE_BODY, False, mColorSecondary)
    y = y + authLbl.Height + GAP_LARGE

    AddDivider y
    y = y + DIVIDER_HEIGHT + GAP_LARGE

    Dim noticeLbl As MSForms.LABEL
    Set noticeLbl = AddLabel(TEXT_ABOUT_UPDATE_NOTICE, y, FONT_SIZE_CITATION, False, mColorSecondary)
    y = y + noticeLbl.Height

    EndDialog y, Array(TEXT_BTN_CLOSE), Array("result:closed")
    Exit Sub
ErrHandler:
    HandleFatalError "BuildAbout", Err.description
End Sub

Public Sub ShowAbout()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("frmWarning.ShowAbout")
    BuildAbout
    Me.Show vbModal
End Sub

' ============================================================================
' Dieu phoi Click â€” DialogControlSink goi vao day (Owner=Me), xem dau file.
' ============================================================================

Public Sub HandleControlClick(ByVal tag As String)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("frmWarning.HandleControlClick")
    On Error GoTo ErrHandler
    If left$(tag, 7) = "result:" Then
        Result = Mid$(tag, 8)
        Me.Hide
    End If
    Exit Sub
ErrHandler:
    HandleFatalError "HandleControlClick(" & tag & ")", Err.description
End Sub

Private Sub HandleFatalError(ByVal source As String, ByVal description As String)
    On Error Resume Next
    ' MsgBoxW.Show (Win32 MessageBoxW that su) thay cho MsgBox chuan - hien dung dau tren MOI may.
    MsgBoxW.Show TEXT_INTERNAL_ERROR_PREFIX & source & "): " & description, _
        vbCritical, TEXT_PRODUCT_NAME
End Sub

' ============================================================================
' Khung hop thoai dung chung cho nam che do
' ============================================================================

' Dung phan dau hop thoai (tu ve, thay thanh tieu de he dieu hanh - xem ghi chu dau file) va tra
' ve toa do Y de bat dau xep noi dung.
Private Function BeginDialog(ByVal titleText As String, ByVal cancelResult As String) As Single
    ClearDynamic
    Result = ""
    mCancelResult = cancelResult
    ' Van dat Me.Caption (khong dau) lam phuong an du phong: neu WinApiFormStyle that bai tren mot
    ' may nao do thi form quay ve dung thanh tieu de mac dinh, luc do caption nay hien ra.
    Me.caption = Utils.ToUnaccented(titleText)

    ' Nen phan dau - CHINH la vung keo de di chuyen cua so.
    Dim headerBg As MSForms.LABEL
    Set headerBg = NewLabel("hdrBg")
    With headerBg
        .left = 0
        .top = 0
        .width = FORM_WIDTH_PT
        .Height = HEADER_HEIGHT
        .BackStyle = fmBackStyleOpaque
        .BackColor = mColorFormBg
        .BorderStyle = fmBorderStyleNone
        .caption = ""
    End With
    AddDragSink headerBg

    Dim titleLbl As MSForms.LABEL
    Set titleLbl = NewLabel("hdrTitle")
    With titleLbl
        .left = PAD
        .width = FORM_WIDTH_PT - PAD - CLOSE_BUTTON_SIZE - GAP_SMALL
        .Font.name = UI_FONT_NAME
        .Font.size = FONT_SIZE_HEADER
        .Font.bold = True
        .foreColor = mColorPrimary
        .BackStyle = fmBackStyleTransparent
        .WordWrap = False
        .AutoSize = False
        .caption = titleText
        .AutoSize = True
        .AutoSize = False
        .width = FORM_WIDTH_PT - PAD - CLOSE_BUTTON_SIZE - GAP_SMALL
        .top = (HEADER_HEIGHT - .Height) / 2
    End With
    AddDragSink titleLbl

    ' Nut dong "x" goc phai phan dau.
    Dim closeLbl As MSForms.LABEL
    Set closeLbl = NewLabel("hdrClose")
    With closeLbl
        .left = FORM_WIDTH_PT - CLOSE_BUTTON_SIZE - GAP_SMALL
        .top = (HEADER_HEIGHT - CLOSE_BUTTON_SIZE) / 2
        .width = CLOSE_BUTTON_SIZE
        .Height = CLOSE_BUTTON_SIZE
        .Font.name = UI_FONT_NAME
        .Font.size = FONT_SIZE_CLOSE
        .foreColor = mColorSecondary
        .BackStyle = fmBackStyleOpaque
        .BackColor = mColorFormBg
        .BorderStyle = fmBorderStyleNone
        .TextAlign = fmTextAlignCenter
        .caption = TEXT_CLOSE_GLYPH
    End With
    AddClickSink closeLbl, "result:" & cancelResult

    BeginDialog = HEADER_HEIGHT + GAP_SMALL
End Function

' Dai chan hop thoai: duong ke manh + nen xam nhat + hang nut canh phai (tu dong quay ve xep doc
' neu khong du be ngang). Sau do dat lai kich thuoc form theo dung chieu cao that.
Private Sub EndDialog(ByVal contentBottom As Single, ByVal captions As Variant, ByVal tags As Variant)
    Dim footerTop As Single
    footerTop = contentBottom + PAD

    Dim widths() As Single
    Dim i As Long, n As Long
    n = UBound(captions) - LBound(captions) + 1
    ReDim widths(0 To n - 1)
    Dim totalWidth As Single
    For i = 0 To n - 1
        widths(i) = ButtonWidthFor(CStr(captions(LBound(captions) + i)))
        totalWidth = totalWidth + widths(i)
    Next i
    totalWidth = totalWidth + BUTTON_GAP * (n - 1)

    Dim footerHeight As Single
    Dim stacked As Boolean
    stacked = (totalWidth > mContentWidth)
    If stacked Then
        footerHeight = FOOTER_PAD * 2 + BUTTON_HEIGHT * n + BUTTON_GAP * (n - 1)
    Else
        footerHeight = FOOTER_PAD * 2 + BUTTON_HEIGHT
    End If

    AddDivider footerTop

    Dim footerBg As MSForms.LABEL
    Set footerBg = NewLabel("footerBg")
    With footerBg
        .left = 0
        .top = footerTop + DIVIDER_HEIGHT
        .width = FORM_WIDTH_PT
        .Height = footerHeight
        .BackStyle = fmBackStyleOpaque
        .BackColor = mColorFooterBg
        .BorderStyle = fmBorderStyleNone
        .caption = ""
    End With

    Dim y As Single
    y = footerTop + DIVIDER_HEIGHT + FOOTER_PAD
    If stacked Then
        For i = 0 To n - 1
            AddFlatButton CStr(captions(LBound(captions) + i)), CStr(tags(LBound(tags) + i)), _
                PAD, y, mContentWidth, (i = 0)
            y = y + BUTTON_HEIGHT + BUTTON_GAP
        Next i
    Else
        Dim x As Single
        x = FORM_WIDTH_PT - PAD - totalWidth
        For i = 0 To n - 1
            AddFlatButton CStr(captions(LBound(captions) + i)), CStr(tags(LBound(tags) + i)), _
                x, y, widths(i), (i = 0)
            x = x + widths(i) + BUTTON_GAP
        Next i
    End If

    SetClientHeight footerTop + DIVIDER_HEIGHT + footerHeight
End Sub

' Me.Width/Me.Height la kich thuoc NGOAI (ke ca khung + thanh tieu de), con noi dung xep theo kich
' thuoc TRONG (InsideWidth/InsideHeight) - chenh lech giua hai cai do CHINH la be day khung, do
' ngay luc chay thay vi doan mot hang so (ban cu cong tay "+ 24", sai khi doi DPI hoac khi bo
' thanh tieu de).
Private Sub SetClientWidth(ByVal w As Single)
    Dim frame As Single
    frame = Me.width - Me.InsideWidth
    ' Chan truong hop InsideWidth tra ve gia tri vo ly (khong ky vong xay ra, nhung neu xay ra thi
    ' cong nham se lam form phinh to dan sau moi lan mo).
    If frame < 0 Or frame > 40 Then frame = 10
    Me.width = w + frame
End Sub

Private Sub SetClientHeight(ByVal h As Single)
    Dim frame As Single
    frame = Me.Height - Me.InsideHeight
    If frame < 0 Or frame > 60 Then frame = 24
    Me.Height = h + frame
End Sub

Private Sub ClearDynamic()
    Dim nm As Variant
    For Each nm In mDynControlNames
        On Error Resume Next
        Me.Controls.Remove CStr(nm)
        On Error GoTo 0
    Next nm
    Set mDynControlNames = New Collection
    Set mDynSinks = New Collection
    mNameCounter = 0
End Sub

Private Function NextName(ByVal prefix As String) As String
    mNameCounter = mNameCounter + 1
    NextName = "dyn_" & prefix & CStr(mNameCounter)
End Function

Private Function NewLabel(ByVal prefix As String) As MSForms.LABEL
    Dim lbl As MSForms.LABEL
    Set lbl = Me.Controls.Add("Forms.Label.1", NextName(prefix), True)
    mDynControlNames.Add lbl.name
    Set NewLabel = lbl
End Function

Private Sub AddClickSink(ByVal lbl As MSForms.LABEL, ByVal tag As String)
    Dim sink As DialogControlSink
    Set sink = New DialogControlSink
    Set sink.Owner = Me
    sink.tag = tag
    Set sink.actionLabel = lbl
    mDynSinks.Add sink
End Sub

Private Sub AddDragSink(ByVal lbl As MSForms.LABEL)
    Dim sink As DialogControlSink
    Set sink = New DialogControlSink
    Set sink.Owner = Me
    Set sink.DragLabel = lbl
    mDynSinks.Add sink
End Sub

' Dat Caption TRUOC luc AutoSize con False roi MOI bat AutoSize - thu tu nguoc lai lam Height
' phinh to phi thuc do WordWrap boc theo be rong gan-khong-do luc Caption con rong.
Private Sub SetWrapCaption(ByVal lbl As MSForms.LABEL, ByVal text As String, ByVal width As Single)
    lbl.AutoSize = False
    lbl.WordWrap = True
    lbl.width = width
    lbl.caption = text
    lbl.AutoSize = True
    lbl.AutoSize = False
    lbl.width = width
End Sub

' KHONG co tham so "italic" (khac ban cu): xem muc 4 ghi chu dau file - gan Font.Italic = False
' lai lam chu THANH nghieng, va ban thiet ke moi khong dung chu nghieng o dau ca.
Private Function AddLabel(ByVal text As String, ByVal top As Single, ByVal fontSize As Single, _
        ByVal bold As Boolean, ByVal foreColor As Long) As MSForms.LABEL
    Dim lbl As MSForms.LABEL
    Set lbl = NewLabel("lbl")
    With lbl
        .left = PAD
        .top = top
        .Font.name = UI_FONT_NAME
        .Font.size = fontSize
        ' CHI gan Font.Bold khi CAN in dam - gan "= False" lai lam chu THANH IN DAM (loi MSForms,
        ' xem muc 4 ghi chu dau frmWarning.frm).
        If bold Then .Font.bold = True
        .foreColor = foreColor
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
    End With
    SetWrapCaption lbl, text, mContentWidth
    Set AddLabel = lbl
End Function

' Mot khoi "Nhan phu â€” noi dung" theo mau ADR-007 (Viec se lam / Pham vi anh huong / Kha nang hoan
' tac).
Private Function AddSectionBlock(ByVal labelText As String, ByVal bodyText As String, ByVal topIn As Single) As Single
    Dim y As Single
    y = topIn
    Dim lblLbl As MSForms.LABEL
    Set lblLbl = AddLabel(labelText, y, FONT_SIZE_LABEL, True, mColorPrimary)
    y = y + lblLbl.Height + GAP_SMALL
    Dim bodyLbl As MSForms.LABEL
    Set bodyLbl = AddLabel(bodyText, y, FONT_SIZE_BODY, False, mColorSecondary)
    y = y + bodyLbl.Height + GAP_LARGE
    AddSectionBlock = y
End Function

' Dai thong bao kieu Fluent MessageBar: nen rat nhat + mot vach mau dam ben trai.
Private Function AddNoticeBar(ByVal text As String, ByVal topIn As Single) As Single
    Dim textLbl As MSForms.LABEL
    Set textLbl = NewLabel("noticeText")
    With textLbl
        .left = PAD + 3 + 8
        .top = topIn + 8
        .Font.name = UI_FONT_NAME
        .Font.size = FONT_SIZE_BODY
        .foreColor = mColorWarning
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
    End With
    SetWrapCaption textLbl, text, mContentWidth - 3 - 16

    Dim barHeight As Single
    barHeight = textLbl.Height + 16

    Dim bg As MSForms.LABEL
    Set bg = NewLabel("noticeBg")
    With bg
        .left = PAD
        .top = topIn
        .width = mContentWidth
        .Height = barHeight
        .BackStyle = fmBackStyleOpaque
        .BackColor = RGB(253, 243, 238)
        .BorderStyle = fmBorderStyleNone
        .caption = ""
    End With

    Dim stripe As MSForms.LABEL
    Set stripe = NewLabel("noticeStripe")
    With stripe
        .left = PAD
        .top = topIn
        .width = 3
        .Height = barHeight
        .BackStyle = fmBackStyleOpaque
        .BackColor = mColorWarning
        .BorderStyle = fmBorderStyleNone
        .caption = ""
    End With

    ' Dua chu len tren hai lop nen (control them sau nam duoi trong thu tu Z cua MSForms).
    textLbl.ZOrder 0
    AddNoticeBar = topIn + barHeight
End Function

Private Sub AddDivider(ByVal top As Single)
    Dim lbl As MSForms.LABEL
    Set lbl = NewLabel("divider")
    With lbl
        .left = 0
        .top = top
        .width = FORM_WIDTH_PT
        .Height = DIVIDER_HEIGHT
        .BackStyle = fmBackStyleOpaque
        .BackColor = mColorDivider
        .BorderStyle = fmBorderStyleNone
        .caption = ""
    End With
End Sub

' Do be rong that su cua mot doan chu o phong/co chu cua nut - dung Label AutoSize roi xoa ngay
' (khong co API do chu trong MSForms).
Private Function MeasureTextWidth(ByVal text As String, ByVal fontSize As Single, ByVal bold As Boolean) As Single
    Dim lbl As MSForms.LABEL
    Set lbl = Me.Controls.Add("Forms.Label.1", NextName("meas"), False)
    With lbl
        .Font.name = UI_FONT_NAME
        .Font.size = fontSize
        ' CHI gan Font.Bold khi CAN in dam - gan "= False" lai lam chu THANH IN DAM (loi MSForms,
        ' xem muc 4 ghi chu dau frmWarning.frm).
        If bold Then .Font.bold = True
        .AutoSize = False
        .WordWrap = False
        .caption = text
        .AutoSize = True
    End With
    MeasureTextWidth = lbl.width
    On Error Resume Next
    Me.Controls.Remove lbl.name
End Function

Private Function ButtonWidthFor(ByVal caption As String) As Single
    Dim w As Single
    w = MeasureTextWidth(caption, FONT_SIZE_BODY, False) + BUTTON_TEXT_PAD
    If w < BUTTON_MIN_WIDTH Then w = BUTTON_MIN_WIDTH
    ButtonWidthFor = w
End Function

' Nut phang kieu Fluent, ve bang HAI Label chong nhau - xem muc 2 ghi chu dau file.
Private Sub AddFlatButton(ByVal caption As String, ByVal tag As String, ByVal left As Single, _
        ByVal top As Single, ByVal width As Single, ByVal isPrimary As Boolean)
    Dim bgColor As Long, txtColor As Long
    If isPrimary Then
        bgColor = mColorAccent
        txtColor = mColorAccentText
    Else
        bgColor = mColorButtonBg
        txtColor = mColorButtonText
    End If

    Dim bg As MSForms.LABEL
    Set bg = NewLabel("btnBg")
    With bg
        .left = left
        .top = top
        .width = width
        .Height = BUTTON_HEIGHT
        .BackStyle = fmBackStyleOpaque
        .BackColor = bgColor
        .caption = ""
        If isPrimary Then
            .BorderStyle = fmBorderStyleNone
        Else
            .BorderStyle = fmBorderStyleSingle
            .BorderColor = mColorButtonBorder
        End If
    End With
    AddClickSink bg, tag

    Dim txt As MSForms.LABEL
    Set txt = NewLabel("btnTxt")
    With txt
        .Font.name = UI_FONT_NAME
        .Font.size = FONT_SIZE_BODY
        .BackStyle = fmBackStyleOpaque
        .BackColor = bgColor
        .foreColor = txtColor
        .BorderStyle = fmBorderStyleNone
        .TextAlign = fmTextAlignCenter
        .AutoSize = False
        .WordWrap = False
        .caption = caption
        .AutoSize = True
        .AutoSize = False
        .left = left + 1
        .width = width - 2
        .top = top + (BUTTON_HEIGHT - .Height) / 2
    End With
    txt.ZOrder 0
    AddClickSink txt, tag
End Sub
