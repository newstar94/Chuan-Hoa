Attribute VB_Name = "frmQrCode"
Attribute VB_Base = "0{FF2E029B-41B5-4D0F-B245-6896D77650D0}{57685F15-E300-4DB2-959C-058BA5140423}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

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
Private Const TEXTBOX_HEIGHT As Single = 84

Private Const FONT_SIZE_HEADER As Single = 12
Private Const FONT_SIZE_BODY As Single = 9
Private Const FONT_SIZE_HINT As Single = 8
Private Const FONT_SIZE_CLOSE As Single = 14
Private Const UI_FONT_NAME As String = "Segoe UI"

Private TEXT_PRODUCT_NAME As String
Private TEXT_BTN_CANCEL As String
Private TEXT_TITLE As String
Private TEXT_PROMPT As String
Private TEXT_HINT As String
Private TEXT_BTN_CONFIRM As String
Private TEXT_INTERNAL_ERROR_PREFIX As String
Private TEXT_CLOSE_GLYPH As String

Private mColorPrimary As Long
Private mColorSecondary As Long
Private mColorAccent As Long
Private mColorAccentText As Long
Private mColorButtonBg As Long
Private mColorButtonBorder As Long
Private mColorButtonText As Long
Private mColorFormBg As Long
Private mColorFooterBg As Long
Private mColorDivider As Long

Public Result As String
Public Content As String

Private mDynControlNames As Collection
Private mDynSinks As Collection
Private mNameCounter As Long
Private mContentWidth As Single
Private mContentBox As MSForms.TextBox

' ============================================================================
' Vong doi form
' ============================================================================

Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler
    InitializeText

    mColorPrimary = RGB(32, 31, 30)
    mColorSecondary = RGB(96, 94, 92)
    mColorAccent = RGB(0, 120, 212)
    mColorAccentText = RGB(255, 255, 255)
    mColorButtonBg = RGB(255, 255, 255)
    mColorButtonBorder = RGB(138, 136, 134)
    mColorButtonText = RGB(32, 31, 30)
    mColorFormBg = RGB(255, 255, 255)
    mColorFooterBg = RGB(250, 249, 248)
    mColorDivider = RGB(225, 223, 221)
    Me.BackColor = mColorFormBg

    Me.StartUpPosition = 2 ' 2 = fmStartUpScreen (CenterScreen)
    SetClientWidth FORM_WIDTH_PT
    mContentWidth = FORM_WIDTH_PT - 2 * PAD

    Set mDynControlNames = New Collection
    Set mDynSinks = New Collection
    Exit Sub
ErrHandler:
    HandleFatalError "UserForm_Initialize", Err.description
End Sub

' bo thanh tieu de he dieu hanh - xem WinApiFormStyle.bas.
Private Sub UserForm_Activate()
    WinApiFormStyle.MakeActiveWindowBorderless
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error GoTo ErrHandler
    If CloseMode = vbFormControlMenu Then
        Result = "cancel"
        Cancel = True
        Me.Hide
    End If
    Exit Sub
ErrHandler:
    ' Loi o day khong duoc chan viec dong form - bo qua im lang.
End Sub

Private Sub InitializeText()
    TEXT_PRODUCT_NAME = "Chu" & ChrW(&H1EA9) & "n h" & ChrW(&HF3) & "a th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    TEXT_BTN_CANCEL = "H" & ChrW(&H1EE7) & "y"
    TEXT_TITLE = "Ch" & ChrW(&HE8) & "n m" & ChrW(&HE3) & " QR"
    TEXT_PROMPT = "Nh" & ChrW(&H1EAD) & "p n" & ChrW(&H1ED9) & "i dung " & ChrW(&H111) & ChrW(&H1EC3) & " t" & ChrW(&H1EA1) & "o m" & ChrW(&HE3) & " QR"
    TEXT_HINT = "T" & ChrW(&H1ED1) & "i " & ChrW(&H111) & "a 800 k" & ChrW(&H1EF7) & " t" & ChrW(&H1EF1) & "."
    TEXT_BTN_CONFIRM = TEXT_TITLE
    TEXT_INTERNAL_ERROR_PREFIX = ChrW(&H110) & ChrW(&HE3) & " x" & ChrW(&H1EA3) & "y ra l" & ChrW(&H1ED7) & "i n" & ChrW(&H1ED9) & "i b" & ChrW(&H1ED9) & " ("
    TEXT_CLOSE_GLYPH = ChrW(&HD7)
End Sub

Public Sub BuildQrCode()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("frmQrCode.BuildQrCode")
    On Error GoTo ErrHandler
    Dim y As Single
    y = BeginDialog()

    Dim promptLbl As MSForms.LABEL
    Set promptLbl = AddLabel(TEXT_PROMPT, y, FONT_SIZE_BODY, False, mColorPrimary)
    y = y + promptLbl.Height + GAP_SMALL

    Set mContentBox = Me.Controls.Add("Forms.TextBox.1", "dyn_contentBox", True)
    With mContentBox
        .left = PAD
        .top = y
        .width = mContentWidth
        .Height = TEXTBOX_HEIGHT
        .Font.name = UI_FONT_NAME
        .Font.size = FONT_SIZE_BODY
        .MultiLine = True
        .WordWrap = True
        .EnterKeyBehavior = True
        .ScrollBars = fmScrollBarsVertical
        ' O nhap phang kieu Fluent: nen trang, vien mot net mau xam trung tinh - thay hop "chim
        ' xuong" 3D mac dinh cua MSForms.
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = mColorButtonBorder
        .BackColor = mColorFormBg
        .foreColor = mColorPrimary
        .text = ""
    End With
    mDynControlNames.Add mContentBox.name
    y = y + TEXTBOX_HEIGHT + GAP_SMALL

    Dim hintLbl As MSForms.LABEL
    Set hintLbl = AddLabel(TEXT_HINT, y, FONT_SIZE_HINT, False, mColorSecondary)
    y = y + hintLbl.Height

    EndDialog y, Array(TEXT_BTN_CONFIRM, TEXT_BTN_CANCEL), _
        Array("result:confirm", "result:cancel")
    Exit Sub
ErrHandler:
    HandleFatalError "BuildQrCode", Err.description
End Sub

Public Sub ShowQrCode()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("frmQrCode.ShowQrCode")
    BuildQrCode
    Me.Show vbModal
End Sub

' ============================================================================
' Dieu phoi Click -- DialogControlSink goi vao day (Owner=Me).
' ============================================================================

Public Sub HandleControlClick(ByVal tag As String)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("frmQrCode.HandleControlClick")
    On Error GoTo ErrHandler
    If tag = "result:confirm" Then
        ' Chi doi khong rong -- gioi han do dai/dung luong that su kiem tra o
        ' QrCodeGenerator.GenerateQrModuleMatrix sau khi dong form (xem dau file).
        If Len(Trim$(mContentBox.text)) > 0 Then
            Result = "confirm"
            Content = mContentBox.text
            Me.Hide
        End If
    ElseIf tag = "result:cancel" Then
        Result = "cancel"
        Me.Hide
    End If
    Exit Sub
ErrHandler:
    HandleFatalError "HandleControlClick(" & tag & ")", Err.description
End Sub

Private Sub HandleFatalError(ByVal source As String, ByVal description As String)
    On Error Resume Next
    MsgBoxW.Show TEXT_INTERNAL_ERROR_PREFIX & source & "): " & description, _
        vbCritical, TEXT_PRODUCT_NAME
End Sub

Private Function BeginDialog() As Single
    ClearDynamic
    Result = ""
    Content = ""
    Me.caption = Utils.ToUnaccented(TEXT_TITLE)

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
        .caption = TEXT_TITLE
        .AutoSize = True
        .AutoSize = False
        .width = FORM_WIDTH_PT - PAD - CLOSE_BUTTON_SIZE - GAP_SMALL
        .top = (HEADER_HEIGHT - .Height) / 2
    End With
    AddDragSink titleLbl

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
    AddClickSink closeLbl, "result:cancel"

    BeginDialog = HEADER_HEIGHT + GAP_SMALL
End Function

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

Private Sub SetClientWidth(ByVal w As Single)
    Dim frame As Single
    frame = Me.width - Me.InsideWidth
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
    Set mContentBox = Nothing
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

Private Sub SetWrapCaption(ByVal lbl As MSForms.LABEL, ByVal text As String, ByVal width As Single)
    lbl.AutoSize = False
    lbl.WordWrap = True
    lbl.width = width
    lbl.caption = text
    lbl.AutoSize = True
    lbl.AutoSize = False
    lbl.width = width
End Sub

' KHONG co tham so "italic" - xem muc 4 ghi chu dau frmWarning.frm (gan Font.Italic = False lai
' lam chu THANH nghieng).
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
