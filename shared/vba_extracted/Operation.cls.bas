Attribute VB_Name = "Operation"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

' Thoi diem thao tac HOAN TAT, dinh dang "yyyy-mm-dd hh:nn:ss" â€” Utils.EndOperation tu dien gia
' tri nay bang Format$(Now,...), noi goi khong tu dat.
Public Timestamp As String
' Ten thao tac â€” trung voi opName da truyen vao Utils.BeginOperation.
Public name As String
Public affectedCount As Long
Public hasWarning As Boolean

Private Sub Class_Initialize()
    affectedCount = 0
    hasWarning = False
End Sub
