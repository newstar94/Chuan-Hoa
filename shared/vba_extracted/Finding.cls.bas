Attribute VB_Name = "Finding"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public ruleCode As String
Public Group As String
' VBA khong co kieu lieu ke, dung String va doi chieu bang hang so trong Utils.bas khi can so
' sanh.
Public Severity As String
' "ND30" | "QD1989" | "SUY_RA" | "THONG_LE" â€” da chuan hoa qua RuleLoader.NormalizeSourceLabel.
Public SourceLabel As String
Public title As String
Public message As String
' Tuy chon ben TS (citation?) â€” chuoi rong neu khong co, khong dung Null cho truong String.
Public Citation As String
' number | null ben TS â€” Variant de giu duoc Null khi khong xac dinh duoc doan van.
Public paragraphIndex As Variant
' Tuy chon (charOffset?) â€” Variant, Empty neu khong co.
Public charOffset As Variant
' Tuy chon (occurrences?) â€” Variant, Empty neu khong co.
Public occurrences As Variant
Public AutoFixable As Boolean
Public RiskLevel As String
Public RequiresConfirmation As Boolean
' Bat buoc voi moi phat hien loai C (nhom chinh ta) â€” xem ADR-009. Chuoi rong neu khong ap dung.
Public Before As String
Public After As String
Public Checkability As String

Private Sub Class_Initialize()
    paragraphIndex = Null
    charOffset = Empty
    occurrences = Empty
End Sub
