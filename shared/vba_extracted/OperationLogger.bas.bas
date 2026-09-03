Attribute VB_Name = "OperationLogger"
'==============================================================
' OperationLogger â€” Ghi nhat ky thao tac trong phien lam viec
' FR-SAF-04, FR-SAF-05 â€” xem docs/decisions/ADR-007-an-toan-va-hoan-tac.md.
' RANG BUOC BAT BIEN (NFR-SEC-03): nhat ky CHI ton tai trong phien lam viec hien tai. Khong ghi ra
' file, khong dung Document.Variables/registry cho du lieu nay â€” dong Word la mat, dung chu y,
' Nhat ky CHI ton tai trong phien VBA, khong ghi ra dia va khong luu vao tai lieu.
' GetLastOperation â†” getLast, GetAllOperations â†” getAll. Them ClearLog de phuc vu kiem thu thu
' cong tu Immediate Window theo dung tieu chi hoan thanh cua task.
' Module nay la "handler" duoc Utils.SetOperationLogHandler gan vao â€” Utils dung LIEN KET MUON
' (bien kieu Object) nen PHAI giu dung hai chu ky cong khai duoi day. Doi ten hoac doi tham so ma
' khong sua dong bo Utils.bas se lam Utils cham dut ghi log ma khong co loi bien dich nao bao
' truoc (loi lien ket muon chi lo ra luc CHAY): Public Sub LogOperation(ByVal op As Operation)
' Public Sub LogError(ByVal opName As String, ByVal errDescription As String)
'==============================================================
Option Explicit

Private mOperations As Collection

' Chuoi ASCII, chua co dau tieng Viet day du - cung quy uoc voi cac thong bao khac trong
' RuleLoader.bas/Utils.bas: tranh rui ro sai lech ma khi VBIDE Import file.bas nay. se ra soat lai
' toan bo chuoi giao dien.
' chuyen len day tu vi tri cu o cuoi file (sau UpdateStatusBar) - moi khai bao cap module PHAI
' dung TRUOC toan bo Sub/Function trong file (quy tac VBA).
Private Const LABEL_LOI As String = "loi"
Private Const MARK_WARNING As String = " !"

Private Sub EnsureInitialized()
    If mOperations Is Nothing Then Set mOperations = New Collection
End Sub

' Handler cho Utils.SetOperationLogHandler â€” Utils.EndOperation goi vao day khi mot thao tac hoan
' tat binh thuong. op da duoc Utils dien san Timestamp/Name/AffectedCount/HasWarning.
Public Sub LogOperation(ByVal op As Operation)
    On Error GoTo ErrHandler
    EnsureInitialized
    mOperations.Add op
    UpdateStatusBar op
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "OperationLogger.LogOperation", Err.description
End Sub

' Handler cho Utils.SetOperationLogHandler â€” Utils.AbortOperation goi vao day khi thao tac that
' bai giua chung (FR-SAF-07).
Public Sub LogError(ByVal opName As String, ByVal errDescription As String)
    On Error GoTo ErrHandler
    EnsureInitialized

    Dim op As New Operation
    op.Timestamp = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ' Khong dua errDescription day du vao ten hien thi - co the qua dai hoac chua chi tiet noi bo.
    ' errDescription da duoc AbortOperation dua vao MsgBox rieng cho nguoi dung roi.
    op.name = opName & " - " & LABEL_LOI
    op.affectedCount = 0
    op.hasWarning = True

    mOperations.Add op
    UpdateStatusBar op
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "OperationLogger.LogError", Err.description
End Sub

' ============================================================================
' ============================================================================


Private Sub UpdateStatusBar(ByVal op As Operation)
    ' Lay gio:phut:giay tu chinh Timestamp da ghi (8 ky tu cuoi cua "yyyy-mm-dd hh:nn:ss"), khong
    ' goi lai Now de tranh lech vai mili-giay voi dong da ghi vao nhat ky.
    Dim timePart As String
    timePart = Right$(op.Timestamp, 8)

    Dim line As String
    line = timePart & " " & op.name & " - " & CStr(op.affectedCount)
    If op.hasWarning Then line = line & MARK_WARNING

    Application.StatusBar = line
End Sub
