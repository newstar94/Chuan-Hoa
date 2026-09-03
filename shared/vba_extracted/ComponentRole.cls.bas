Attribute VB_Name = "ComponentRole"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Public paragraphIndex As Long
' "nationalTitle" | "nationalMotto" |... (ComponentSpecKey) | "unknown"
Public role As String
' "high" | "medium" | "low".
Public confidence As String
' Ten dau hieu da khop trong shared/rules/dau-hieu-nhan-dien.json (vi du "recipientListClosing") â€”
' phuc vu chan doan/giai thich, KHONG nhat thiet trung voi Role (mot Role co the co nhieu dau
' hieu, xem "recipientList").
Public signalKey As String
