Attribute VB_Name = "ViewOptions"
Option Explicit

' --- Doc

Public Function GetShowTextBoundaries() As Boolean
    On Error Resume Next
    GetShowTextBoundaries = ActiveWindow.View.ShowTextBoundaries
End Function

Public Function GetShowCropMarks() As Boolean
    On Error Resume Next
    GetShowCropMarks = ActiveWindow.View.ShowCropMarks
End Function

Public Function GetShowAllMarks() As Boolean
    On Error Resume Next
    GetShowAllMarks = ActiveWindow.View.ShowAll
End Function

' --- Ghi -------------------------------------------------------------------------------------
' KHONG di qua Utils.BeginOperation/EndOperation: day la thiet lap hien thi cua Word, khong sua
' noi dung tai lieu nen khong co gi de Undo va khong can nhat ky thao tac.

Public Sub SetShowTextBoundaries(ByVal pressed As Boolean)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ViewOptions.SetShowTextBoundaries")
    On Error Resume Next
    ActiveWindow.View.ShowTextBoundaries = pressed
End Sub

Public Sub SetShowCropMarks(ByVal pressed As Boolean)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ViewOptions.SetShowCropMarks")
    On Error Resume Next
    ActiveWindow.View.ShowCropMarks = pressed
End Sub

Public Sub SetShowAllMarks(ByVal pressed As Boolean)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ViewOptions.SetShowAllMarks")
    On Error Resume Next
    ActiveWindow.View.ShowAll = pressed
End Sub

' --- Dieu kien bat/mo ------------------------------------------------------------------------ Ba
' cong tac nay KHONG doi hoi "Doc du lieu" da chay hay tai lieu phai la.docx - chung khong cham
' vao noi dung tai lieu, chi can co mot cua so dang mo de doc View.
Public Function HasActiveWindow() As Boolean
    On Error Resume Next
    HasActiveWindow = (Application.Documents.count > 0) And (Not ActiveWindow Is Nothing)
End Function
