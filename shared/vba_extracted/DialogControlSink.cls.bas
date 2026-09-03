Attribute VB_Name = "DialogControlSink"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'==============================================================
' DialogControlSink â€” Bo bat su kien Click cho control dung "Controls.Add" luc chay cua
' bon UserForm popup (frmWarning, frmSaveAsDialog, frmFontSizeDialog, frmTemplateManager). Cung ly
' do ton tai nhu TaskPaneControlSink.cls: VBA khong tu sinh thu tuc su kien cho control them vao
' luc chay, phai co mot lop WithEvents rieng.
' KHAC TaskPaneControlSink.cls o cho Owner la MOT THUOC TINH (Object), khong hard-code goi thang
' mot form co dinh - vi sink nay dung CHUNG cho bon form khac nhau. Moi form chu (Owner) phai tu
' cai dat "Public Sub HandleControlClick(ByVal tag As String)" - VBA goi qua lien ket muon (late
' binding) theo dung ten thu tuc, khong can khai bao Interface.
' Form chu PHAI giu song cac the hien nay trong mot Collection cap module (mDynSinks) - mat
' tham chieu la VBA don rac ngay, su kien Click ngung bat duoc du control van con tren man hinh.
'==============================================================
Option Explicit

Public WithEvents ActionButton As MSForms.CommandButton
Attribute ActionButton.VB_VarHelpID = -1
Public WithEvents OptionBtn As MSForms.OptionButton
Attribute OptionBtn.VB_VarHelpID = -1
Public WithEvents actionLabel As MSForms.LABEL
Attribute actionLabel.VB_VarHelpID = -1
' Label phan dau hop thoai - keo de di chuyen cua so (thay thanh tieu de da bi bo, xem
' WinApiFormStyle.bas).
Public WithEvents DragLabel As MSForms.LABEL
Attribute DragLabel.VB_VarHelpID = -1
Public tag As String
Public Owner As Object

' goi MsgBoxW.Show (Win32 MessageBoxW that su) - hien dung dau tren moi may, khac MsgBox chuan
' (xem ghi chu module MsgBoxW.bas). "Ä�Ă£ xáº£y ra lá»—i ná»™i bá»™ " dung dang chung voi HandleFatalError
' cua cac UserForm khac (frmWarning.frm...).
Private Sub ActionButton_Click()
    On Error GoTo ErrHandler
    Owner.HandleControlClick Me.tag
    Exit Sub
ErrHandler:
    MsgBoxW.Show ChrW(&H110) & ChrW(&HE3) & " x" & ChrW(&H1EA3) & "y ra l" & ChrW(&H1ED7) & _
        "i n" & ChrW(&H1ED9) & "i b" & ChrW(&H1ED9) & " (DialogControlSink.ActionButton_Click): " & _
        Err.description, vbCritical
End Sub

' nut phang ve bang Label - xem ghi chu khai bao ActionLabel o dau file.
Private Sub ActionLabel_Click()
    On Error GoTo ErrHandler
    Owner.HandleControlClick Me.tag
    Exit Sub
ErrHandler:
    MsgBoxW.Show ChrW(&H110) & ChrW(&HE3) & " x" & ChrW(&H1EA3) & "y ra l" & ChrW(&H1ED7) & _
        "i n" & ChrW(&H1ED9) & "i b" & ChrW(&H1ED9) & " (DialogControlSink.ActionLabel_Click): " & _
        Err.description, vbCritical
End Sub

' keo phan dau hop thoai de di chuyen cua so - tu day tro di Windows tu lo phan con lai (xem
' WinApiFormStyle.StartDragActiveWindow). Chi nhan chuot TRAI (Button = 1).
Private Sub DragLabel_MouseDown(ByVal Button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
    On Error Resume Next
    If Button = 1 Then WinApiFormStyle.StartDragActiveWindow
End Sub

Private Sub OptionBtn_Click()
    On Error GoTo ErrHandler
    Owner.HandleControlClick Me.tag
    Exit Sub
ErrHandler:
    MsgBoxW.Show ChrW(&H110) & ChrW(&HE3) & " x" & ChrW(&H1EA3) & "y ra l" & ChrW(&H1ED7) & _
        "i n" & ChrW(&H1ED9) & "i b" & ChrW(&H1ED9) & " (DialogControlSink.OptionBtn_Click): " & _
        Err.description, vbCritical
End Sub
