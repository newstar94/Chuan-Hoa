Attribute VB_Name = "UrlOpener"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hWnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As LongPtr
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long
#End If

Private Const SW_SHOWNORMAL As Long = 1

' Gia tri tra ve cua ShellExecute > 32 la thanh cong (quy uoc cua Win32).
Private Const SHELL_EXECUTE_MIN_SUCCESS As Long = 32

' Chi tiet loi cua LAN GOI GAN NHAT (rong neu thanh cong) - RibbonCallbacks.OnGuiPhanHoi doc qua
' LastFailureDetail de hien THANG trong hop thoai bao loi, khong can dao log file.
Private mLastFailureDetail As String

Public Function LastFailureDetail() As String
    LastFailureDetail = mLastFailureDetail
End Function

' Tra ve True neu da giao duoc cho he dieu hanh mo (qua ShellExecute, hoac qua duong du phong
' FollowHyperlink neu ShellExecute that bai va co tai lieu dang mo). CHI chap nhan http/https -
' chan moi luoc do khac (file:, javascript:,...) de mot chuoi sai o noi goi khong bao gio bien
' thanh lenh chay mot chuong trinh bat ky.
Public Function OpenUrl(ByVal url As String) As Boolean
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("UrlOpener.OpenUrl")
    On Error GoTo ErrHandler
    mLastFailureDetail = ""

    Dim trimmed As String
    trimmed = Trim$(url)
    If LCase$(left$(trimmed, 7)) <> "http://" And LCase$(left$(trimmed, 8)) <> "https://" Then
        DebugTrace.Log "UrlOpener.OpenUrl", "Tu choi: chi mo duoc http/https"
        mLastFailureDetail = "URL khong hop le (chi ho tro http/https)"
        Exit Function
    End If

    #If VBA7 Then
        Dim ownerHwnd As LongPtr
        Dim rc As LongPtr
    #Else
        Dim ownerHwnd As Long
        Dim rc As Long
    #End If
    ' Word.Application KHONG co thuoc tinh Hwnd truc tiep (khac Excel.Application) - phai qua
    ' ActiveWindow.Hwnd; On Error Resume Next vi co the khong co cua so nao dang mo (Documents
    ' rong), luc do ownerHwnd giu nguyen gia tri mac dinh 0.
    On Error Resume Next
    ownerHwnd = Application.ActiveWindow.hWnd
    On Error GoTo ErrHandler

    rc = ShellExecute(ownerHwnd, "open", trimmed, vbNullString, vbNullString, SW_SHOWNORMAL)
    If rc > SHELL_EXECUTE_MIN_SUCCESS Then
        OpenUrl = True
        Exit Function
    End If
    DebugTrace.Log "UrlOpener.OpenUrl", "ShellExecute that bai, rc=" & CStr(rc)
    mLastFailureDetail = "ShellExecute rc=" & CStr(rc)

    ' Duong du phong: Word.Document.FollowHyperlink - co che KHAC HAN (Word tu xu ly noi bo), chi
    ' thu khi co it nhat mot tai lieu dang mo. Khong dung lam duong CHINH ( da ghi ro ly do: doi
    ' hoi ActiveDocument + ghi vao lich su lien ket cua tai lieu).
    On Error Resume Next
    If Application.Documents.count > 0 Then
        Err.Clear
        Application.ActiveDocument.FollowHyperlink Address:=trimmed, NewWindow:=True
        If Err.number = 0 Then
            OpenUrl = True
            mLastFailureDetail = ""
        Else
            mLastFailureDetail = mLastFailureDetail & "; FollowHyperlink loi " & Err.number
        End If
    End If
    On Error GoTo 0
    Exit Function

ErrHandler:
    DebugTrace.LogErr "UrlOpener.OpenUrl", "Khong mo duoc dia chi", Err.number, Err.description
    mLastFailureDetail = "Loi " & Err.number & ": " & Err.description
End Function
