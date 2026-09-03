Attribute VB_Name = "SafetyGuard"
Option Explicit

' ============================================================================
' Danh sach thao tac rui ro cao co dinh â€” FR-SAF-02. HOP DONG TEN GOI: moi noi goi
' Utils.BeginOperation cho hai nhom thao tac nay PHAI truyen dung mot trong hai hang so duoi day
' lam opName, neu khong AssessRisk se khong nhan ra va ha nham xuong "low". EncodingConverter.bas
' va module doi cau truc bang khi hien thuc hoa phai dung lai dung hai hang so nay.
' ============================================================================

Public Const HIGH_RISK_ENCODING_CONVERSION As String = "encodingConversion"
Public Const HIGH_RISK_TABLE_STRUCTURE_CHANGE As String = "tableStructureChange"

' ============================================================================
' Chuoi hien thi - ASCII, chua co dau day du, xem ghi chu dau file. Dung khi opName khong khop hai
' thao tac rui ro cao co dinh, cung khong tra duoc tu quy-tac-kiem-tra.json.
' chuyen len day tu vi tri cu (ngay truoc DescribeOperation, sau BuildWarning)
' - moi khai bao cap module PHAI dung TRUOC toan bo Sub/Function trong file (quy tac VBA).
' ============================================================================

Private Const MSG_UNDO_UNCERTAIN As String = _
    "Ctrl+Z CO THE khong hoan tac duoc thao tac nay."
' Chi them cho thao tac di qua OOXML (chuyen bang ma, chen mau). Thuc nghiem S3 (do qua
' Range.InsertXML cua COM) cho thay rui ro that KHONG nam o ngan xep Undo ma o cho chu thich va
' dau trang trong pham vi bi thay deu bi xoa - xem muc A15 cua docs/process/cau-hoi-con-mo.md va
' ADR-007. Cau ve Undo giu nguyen.
Private Const MSG_OOXML_DATA_LOSS As String = _
    "Chu thich (comment) va dau trang (bookmark) nam trong pham vi bi thay se bi xoa, " & _
    "khong khoi phuc lai duoc bang Ctrl+Z."
Private Const MSG_SAVE_REMINDER As String = _
    "Tai lieu chua duoc luu. Nen luu truoc khi tiep tuc."
Private Const MSG_GENERIC_WHAT As String = "Ap dung thao tac nay len tai lieu."
Private Const MSG_GENERIC_SCOPE As String = "Co the anh huong nhieu phan cua tai lieu."

Private Const MSG_ENCODING_WHAT As String = _
    "Chuyen cac doan dang dung phong chu roi (TCVN3, VNI) sang Unicode, doi phong ve Times New Roman."
Private Const MSG_ENCODING_SCOPE As String = _
    "Toan bo tai lieu, ke ca bang bieu, dau trang, chan trang."

Private Const MSG_TABLE_STRUCTURE_WHAT As String = _
    "Them, xoa, gop hoac tach hang, cot trong bang."
Private Const MSG_TABLE_STRUCTURE_SCOPE As String = _
    "Bang dang thao tac - co the anh huong dinh dang phuc tap trong bang."

' ============================================================================
' IsDocumentSaved â€” FR-SAF-06
' ============================================================================

' True khi tai lieu da duoc luu it nhat mot lan VA khong con thay doi chua luu.
Public Function IsDocumentSaved() As Boolean
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("SafetyGuard.IsDocumentSaved")
    On Error GoTo ErrHandler
    IsDocumentSaved = (ActiveDocument.Path <> "" And ActiveDocument.Saved)
    Exit Function
ErrHandler:
    Err.Raise Err.number, "SafetyGuard.IsDocumentSaved", Err.description
End Function

' ============================================================================
' BuildWarning â€” FR-SAF-03, FR-SAF-06
' ============================================================================

Public Function BuildWarning(ByVal opName As String) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("SafetyGuard.BuildWarning")
    On Error GoTo ErrHandler

    Dim descriptor As Object
    Set descriptor = DescribeOperation(opName)

    Dim Result As Object
    Set Result = Utils.NewDictionary()
    Result("whatWillHappen") = descriptor("whatWillHappen")
    Result("scope") = descriptor("scope")
    ' Ban Legacy chi co MOT thao tac di qua OOXML: chuyen bang ma (Range.InsertXML).
    If opName = HIGH_RISK_ENCODING_CONVERSION Then
        Result("undoability") = MSG_UNDO_UNCERTAIN & " " & MSG_OOXML_DATA_LOSS
    Else
        Result("undoability") = MSG_UNDO_UNCERTAIN
    End If

    If IsDocumentSaved() Then
        Result("saveReminder") = ""
    Else
        Result("saveReminder") = MSG_SAVE_REMINDER
    End If

    Set BuildWarning = Result
    Exit Function
ErrHandler:
    Err.Raise Err.number, "SafetyGuard.BuildWarning", Err.description
End Function

' Tra Dictionary {whatWillHappen, scope} mo ta mot thao tac - noi dung tinh, khong doc tai lieu.
' Uu tien: hai thao tac rui ro cao co dinh -> ruleCode trung mot CheckRule -> du phong.
Private Function DescribeOperation(ByVal opName As String) As Object
    Dim Result As Object
    Set Result = Utils.NewDictionary()

    Select Case opName
        Case HIGH_RISK_ENCODING_CONVERSION
            Result("whatWillHappen") = MSG_ENCODING_WHAT
            Result("scope") = MSG_ENCODING_SCOPE

        Case HIGH_RISK_TABLE_STRUCTURE_CHANGE
            Result("whatWillHappen") = MSG_TABLE_STRUCTURE_WHAT
            Result("scope") = MSG_TABLE_STRUCTURE_SCOPE

        Case Else
            Dim rule As Object
            Set rule = RuleLoader.GetCheckRule(opName)
            If Not rule Is Nothing Then
                Result("whatWillHappen") = CStr(rule("message"))
                Result("scope") = MSG_GENERIC_SCOPE
            Else
                Result("whatWillHappen") = MSG_GENERIC_WHAT
                Result("scope") = MSG_GENERIC_SCOPE
            End If
    End Select

    Set DescribeOperation = Result
End Function
