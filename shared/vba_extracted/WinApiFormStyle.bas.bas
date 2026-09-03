Attribute VB_Name = "WinApiFormStyle"
Option Explicit

Private Const GUARD_FILE_NAME As String = "ChuanHoaTheThuc-winapi-guard.tmp"
Private mGuardReady As Boolean
Private mBlockedLabel As String

#If VBA7 Then

Private Const GWL_STYLE As Long = -16
Private Const WS_DLGFRAME As Long = &H400000
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_BORDER As Long = &H800000

Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_FRAMECHANGED As Long = &H20

Private Const WM_NCLBUTTONDOWN As Long = &HA1
Private Const HTCAPTION As Long = 2

Private Declare PtrSafe Function GetActiveWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function ReleaseCapture Lib "user32" () As Long
Private Declare PtrSafe Function SendMessageW Lib "user32" (ByVal hWnd As LongPtr, ByVal wMsg As Long, _
    ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

#If Win64 Then
    Private Declare PtrSafe Function GetWindowLongPtrW Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtrW Lib "user32" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
    ' Office 32-bit: user32 KHONG xuat GetWindowLongPtrW/SetWindowLongPtrW - dung ban...LongW
    ' (LongPtr = Long tren VBA7 32-bit nen chu ky van khop).
    Private Declare PtrSafe Function GetWindowLongPtrW Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtrW Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If

#End If

Private Function GuardFilePath() As String
    On Error Resume Next
    Dim tempDir As String
    tempDir = Environ$("TEMP")
    If Len(tempDir) = 0 Then tempDir = Environ$("TMP")
    If Len(tempDir) = 0 Then Exit Function
    GuardFilePath = tempDir & "\" & GUARD_FILE_NAME
End Function

Private Sub LoadGuardState()
    If mGuardReady Then Exit Sub
    mGuardReady = True
    On Error Resume Next
    Dim p As String: p = GuardFilePath()
    If Len(p) = 0 Then Exit Sub
    If Len(Dir$(p)) = 0 Then Exit Sub
    Dim fnum As Integer, lineText As String
    fnum = FreeFile
    Open p For Input As #fnum
    If Not EOF(fnum) Then Line Input #fnum, lineText
    Close #fnum
    Kill p
    mBlockedLabel = Trim$(lineText)
    If Len(mBlockedLabel) = 0 Then Exit Sub
    DebugTrace.Log "WinApiFormStyle.LoadGuardState", "PHIEN TRUOC Word chet ngay trong loi goi Win32 """ & _
        mBlockedLabel & """ - VO HIEU HOA diem goi nay cho ca phien nay"
    DebugTrace.Flush
End Sub

' True neu duoc phep goi. False (da bi vo hieu hoa o phien truoc, hoac dang o giua mot lan goi
' khac) thi noi goi PHAI tu bo qua toan bo thao tac Win32, khong chi mot lenh.
' So sanh theo TIEN TO (khong phai bang tuyet doi): mot ham co the arm nhieu nhan con
' "TenHam"/"TenHam:buoc2" - vet dau chay cua BAT KY buoc con nao trong ham do cung phai chan toan
' bo ham o lan sau, khong chi dung buoc con da giet Word truoc do.
Private Function CanCall(ByVal callLabel As String) As Boolean
    LoadGuardState
    If Len(mBlockedLabel) > 0 Then
        If left$(mBlockedLabel, Len(callLabel)) = callLabel Then
            DebugTrace.Log "WinApiFormStyle.CanCall", callLabel & " - BO QUA (da giet Word o phien truoc: " & mBlockedLabel & ")"
            Exit Function
        End If
    End If
    CanCall = True
End Function

Private Sub ArmGuard(ByVal callLabel As String)
    On Error Resume Next
    Dim p As String: p = GuardFilePath()
    If Len(p) = 0 Then Exit Sub
    Dim fnum As Integer: fnum = FreeFile
    Open p For Output As #fnum
    Print #fnum, callLabel
    Close #fnum
    DebugTrace.Log "WinApiFormStyle.ArmGuard", callLabel & " - truoc khi goi Win32"
    DebugTrace.Flush
End Sub

Private Sub DisarmGuard(ByVal callLabel As String)
    On Error Resume Next
    Dim p As String: p = GuardFilePath()
    If Len(p) > 0 Then
        If Len(Dir$(p)) > 0 Then Kill p
    End If
    DebugTrace.Log "WinApiFormStyle.DisarmGuard", callLabel & " - xong"
End Sub

' Goi TU BEN TRONG UserForm_Activate cua form dang mo (luc do cua so Win32 that da ton tai va dang
' la cua so kich hoat cua luong -> GetActiveWindow tra ve dung no). On Error Resume Next: day la
' buoc TRANG DIEM - that bai thi form van hien binh thuong voi thanh tieu de mac dinh, khong duoc
' lam gian doan viec mo form.
Public Sub MakeActiveWindowBorderless()
#If VBA7 Then
    Const LABEL As String = "MakeActiveWindowBorderless"
    If Not CanCall(LABEL) Then Exit Sub

    On Error Resume Next

    Dim h As LongPtr
    h = GetActiveWindow()
    If h = 0 Then Exit Sub

    Dim st As LongPtr
    st = GetWindowLongPtrW(h, GWL_STYLE)
    If st = 0 Then Exit Sub

    st = st And Not CLngPtr(WS_DLGFRAME)
    st = st And Not CLngPtr(WS_SYSMENU)
    st = st And Not CLngPtr(WS_THICKFRAME)
    st = st Or CLngPtr(WS_BORDER)

    ArmGuard LABEL
    SetWindowLongPtrW h, GWL_STYLE, st
    SetWindowPos h, 0, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED
    DisarmGuard LABEL

#End If
End Sub

' Goi tu su kien MouseDown cua Label phan dau hop thoai - tu day tro di Windows tu lo viec keo tha
' cua so y het khi nguoi dung keo thanh tieu de that.
Public Sub StartDragActiveWindow()
#If VBA7 Then
    Const LABEL As String = "StartDragActiveWindow"
    If Not CanCall(LABEL) Then Exit Sub

    On Error Resume Next
    Dim h As LongPtr
    h = GetActiveWindow()
    If h = 0 Then Exit Sub
    ReleaseCapture

    ArmGuard LABEL
    SendMessageW h, WM_NCLBUTTONDOWN, CLngPtr(HTCAPTION), 0
    DisarmGuard LABEL
#End If
End Sub
