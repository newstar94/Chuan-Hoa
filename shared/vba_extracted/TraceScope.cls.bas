Attribute VB_Name = "TraceScope"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private mSource As String
Private mActive As Boolean

' Goi NGAY sau khi tao doi tuong (tu DebugTrace.EnterScope) - ghi dong "BAT DAU".
Public Sub Init(ByVal source As String)
    mSource = source
    mActive = True
    DebugTrace.TraceBegin mSource
End Sub

' Chay TU DONG khi bien cuc bo giu doi tuong nay ra khoi pham vi - tuc la khi thu tuc ket thuc
' theo BAT KY duong nao (End Sub, Exit Sub, hay loi lan nguoc len).
Private Sub Class_Terminate()
    If Not mActive Then Exit Sub
    mActive = False
    On Error Resume Next
    DebugTrace.TraceEnd mSource
    On Error GoTo 0
End Sub
