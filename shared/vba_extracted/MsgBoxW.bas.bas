Attribute VB_Name = "MsgBoxW"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function MessageBoxW Lib "user32" ( _
        ByVal hWnd As LongPtr, ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, _
        ByVal uType As Long) As Long
    Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
#End If

' MB_SETFOREGROUND (&H10000) -- dam bao hop thoai noi len tren cung, cung hanh vi nguoi dung quen
' thuoc voi MsgBox chuan (Word tu dua no len foreground).
Private Const MB_SETFOREGROUND As Long = &H10000

' Thay the truc tiep cho MsgBox chuan -- cung thu tu tham so (text, buttons, title) de doi cho de
' dang tai moi diem goi hien co (MsgBox "...", vbExclamation, "..." -> MsgBoxW.Show("...",
' vbExclamation, "...")). text/title dung CO DAU THAT, KHONG qua Utils.ToUnaccented.
Public Function Show(ByVal text As String, _
        Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, _
        Optional ByVal title As String = "") As VbMsgBoxResult
#If VBA7 Then
    On Error GoTo Fallback
    Dim h As LongPtr
    h = GetForegroundWindow()
    Show = MessageBoxW(h, StrPtr(text), StrPtr(title), CLng(buttons) Or MB_SETFOREGROUND)
    Exit Function
Fallback:
#End If
    ' May khong ho tro duong API nay (hiem gap, gia dinh VBA6) -- lui ve MsgBox chuan cua VBA, ep
    ' khong dau qua Utils.ToUnaccented de it nhat khong hien "?" tung ky tu rieng le.
    Show = MsgBox(Utils.ToUnaccented(text), buttons, Utils.ToUnaccented(title))
End Function
