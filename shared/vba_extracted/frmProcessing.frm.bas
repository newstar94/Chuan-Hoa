Attribute VB_Name = "frmProcessing"
Attribute VB_Base = "0{CFBB93A5-C7A3-40ED-9E50-0C1DF3F5B25E}{7CB43E65-4765-46B1-B5A0-A27196DB74FC}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Private Const FORM_WIDTH_PT As Single = 230
Private Const FORM_HEIGHT_PT As Single = 62
Private Const ACCENT_BAR_HEIGHT As Single = 3
Private Const FONT_SIZE As Single = 10
Private Const UI_FONT_NAME As String = "Segoe UI"

Private mLbl As MSForms.LABEL
Private mTextReady As Boolean
Private TEXT_DEFAULT As String

Private Sub UserForm_Initialize()
    On Error GoTo ErrHandler
    EnsureText

    Me.caption = ""
    Me.StartUpPosition = 2 ' 2 = fmStartUpScreen (CenterScreen) â€” xem ghi chĂº Ä‘áº§u file.
    ' Me.Width/Me.Height lĂ  kĂ­ch thÆ°á»›c NGOĂ€I (ká»ƒ cáº£ khung) - quy vá»� kĂ­ch thÆ°á»›c TRONG Ä‘á»ƒ pháº§n ná»™i
    ' dung Ä‘Ăºng báº±ng FORM_WIDTH_PT Ă— FORM_HEIGHT_PT sau khi khung bá»‹ bá»� á»Ÿ UserForm_Activate.
    SetClientSize FORM_WIDTH_PT, FORM_HEIGHT_PT
    Me.BackColor = RGB(255, 255, 255)

    ' Váº¡ch mĂ u nháº¥n má»�ng suá»‘t mĂ©p trĂªn - dáº¥u hiá»‡u "Ä‘ang cháº¡y" theo máº¡ch Fluent.
    Dim accent As MSForms.LABEL
    Set accent = Me.Controls.Add("Forms.Label.1", "lblAccentBar", True)
    With accent
        .left = 0
        .top = 0
        .width = FORM_WIDTH_PT
        .Height = ACCENT_BAR_HEIGHT
        .BackStyle = fmBackStyleOpaque
        .BackColor = RGB(0, 120, 212)
        .BorderStyle = fmBorderStyleNone
        .caption = ""
    End With

    Set mLbl = Me.Controls.Add("Forms.Label.1", "lblProcessing", True)
    With mLbl
        .left = 0
        .width = FORM_WIDTH_PT
        .TextAlign = fmTextAlignCenter
        .Font.name = UI_FONT_NAME
        .Font.size = FONT_SIZE
        .foreColor = RGB(32, 31, 30)
        .BackStyle = fmBackStyleTransparent
        .BorderStyle = fmBorderStyleNone
        .AutoSize = False
        .WordWrap = False
        .caption = TEXT_DEFAULT
        .AutoSize = True
        .AutoSize = False
        .left = 0
        .width = FORM_WIDTH_PT
        .top = ACCENT_BAR_HEIGHT + (FORM_HEIGHT_PT - ACCENT_BAR_HEIGHT - .Height) / 2
    End With
    Exit Sub
ErrHandler:
    ' Lá»—i lĂºc dá»±ng UI phá»¥ (khĂ´ng pháº£i thao tĂ¡c chĂ­nh trĂªn tĂ i liá»‡u) - bá»� qua im láº·ng, khĂ´ng Ä‘á»ƒ má»™t
    ' lá»—i hiá»ƒn thá»‹ lĂ m giĂ¡n Ä‘oáº¡n thao tĂ¡c Ä‘ang chá»� hiá»‡n.
End Sub

' bá»� khung + thanh tiĂªu Ä‘á»� cá»§a há»‡ Ä‘iá»�u hĂ nh - xem má»¥c 1 ghi chĂº Ä‘áº§u file.
Private Sub UserForm_Activate()
    WinApiFormStyle.MakeActiveWindowBorderless
End Sub

Private Sub SetClientSize(ByVal w As Single, ByVal h As Single)
    Dim frameW As Single, frameH As Single
    frameW = Me.width - Me.InsideWidth
    frameH = Me.Height - Me.InsideHeight
    If frameW < 0 Or frameW > 40 Then frameW = 10
    If frameH < 0 Or frameH > 60 Then frameH = 24
    Me.width = w + frameW
    Me.Height = h + frameH
End Sub

Private Sub EnsureText()
    If mTextReady Then Exit Sub
    TEXT_DEFAULT = ChrW(&H110) & "ang x" & ChrW(&H1EED) & " l" & ChrW(&HFD) & ChrW(&H2026)
    mTextReady = True
End Sub

' ProcessingIndicator.bas gá»�i vĂ o Ä‘Ă¢y khi cáº§n Ä‘á»•i ná»™i dung (vĂ­ dá»¥ "Ä�ang kiá»ƒm tra... (57/100)"
' trong lĂºc FindingReporter Ä‘ang cháº¡y).
Public Sub SetText(ByVal text As String)
    On Error Resume Next
    EnsureText
    If Len(text) = 0 Then text = TEXT_DEFAULT
    If Not mLbl Is Nothing Then mLbl.caption = text
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    On Error Resume Next
    ' KhĂ´ng cĂ³ khung/nĂºt há»‡ Ä‘iá»�u hĂ nh nĂªn ngÆ°á»�i dĂ¹ng KHĂ”NG cĂ³ cĂ¡ch thĂ´ng thÆ°á»�ng Ä‘á»ƒ tá»± Ä‘Ă³ng form
    ' nĂ y - cháº·n cáº£ trÆ°á»�ng há»£p báº¥t ngá»� (vĂ­ dá»¥ Alt+F4) Ä‘á»ƒ trĂ¡nh Unload sai thá»�i Ä‘iá»ƒm giá»¯a chá»«ng má»™t
    ' thao tĂ¡c Ä‘ang cháº¡y; ProcessingIndicator.HideProcessing lĂ  nÆ¡i DUY NHáº¤T Ä‘Æ°á»£c phĂ©p Unload form
    ' nĂ y.
    If CloseMode = vbFormControlMenu Then Cancel = True
End Sub
