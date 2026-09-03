Attribute VB_Name = "AppEvents"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'==============================================================
' AppEvents â€” Bo bat su kien CAP APPLICATION (mo/tao tai lieu) cho add-in toan cuc (.dotm trong
' STARTUP). "Viec dau tien luon luon phai lam (ngay khi van ban duoc mo) la detect loai van ban" â€”
' tu dong chay DocumentTypeDetector ngay khi mo mot.doc/.docx (hoac tao moi), cap nhat drop-down
' "Loai van ban" (nhom Khoi dong) TRUOC KHI nguoi dung bam "Kiem tra" lan dau.
' PHAI GIU THAM CHIEU SONG cho ca phien Word â€” mot instance Class Module bi giai phong NGAY khi
' khong con bien nao tro toi no, mat het su kien tu do tro di. AppEventsHost.bas (module thuong)
' giu bien Private module-level tro toi instance nay, dam bao song suot phien â€” xem dau file do.
' TRUOC DAY goi THANG AppEventsHost.OnAnyDocumentOpen (chup anh tai lieu + nhan dien loai van ban
' + phat hien ma hoa + chen comment debug â€” nhieu buoc COM nang) NGAY BEN TRONG su kien
' DocumentOpen/ NewDocument cap Application â€” Word tai lieu VAN DANG trong qua trinh mo/nap khi su
' kien nay chay, va theo tai lieu Microsoft, cac API COM nang KHONG an toan goi truc tiep tu day
' (trang thai noi bo cua Word chua on dinh hoan toan). Day la nguyen nhan CO CO SO nhat cho ca
' "khoi dong rat lau" LAN "crash ngay buoc load add-in" ma chu dá»± an bao. uy thac cho
' Application.OnTime (ky thuat chuan cua VBA de "day" mot tac vu nang ra khoi su kien nhay cam,
' chay lai NGAY SAU KHI Word ranh - xem AppEventsHost.ScheduleOnAnyDocumentOpen) - su kien
' DocumentOpen/NewDocument gio CHI lam MOT viec DUY NHAT la dang ky lich, tra ve NGAY LAP TUC.
'==============================================================
Option Explicit

Public WithEvents oApp As word.Application
Attribute oApp.VB_VarHelpID = -1

Private Sub oApp_DocumentOpen(ByVal Doc As Document)
    AppEventsHost.ScheduleOnAnyDocumentOpen
End Sub

Private Sub oApp_NewDocument(ByVal Doc As Document)
    AppEventsHost.ScheduleOnAnyDocumentOpen
End Sub

' "Khi van ban bat dau co noi dung thi moi duoc bat") â€” Ribbon KHONG tu ve lai getVisible khi noi
' dung tai lieu doi, can mot diem kich hoat DONG. Su kien nay CHAY TRONG PHIEN BINH THUONG (Word
' da nap xong, khac DocumentOpen/NewDocument o tren â€” KHONG can defer qua Application.OnTime), goi
' thang AppEventsHost.OnSelectionChange (ham do TU gioi han tan suat ve lai ribbon that su, xem
' ghi chu dau ham).
Private Sub oApp_WindowSelectionChange(ByVal sel As Selection)
    AppEventsHost.OnSelectionChange
End Sub

' Chuyen qua lai giua NHIEU tai lieu dang mo cung la mot cach lam trang thai "rong hay khong" doi
' ma KHONG di qua DocumentOpen/NewDocument - AppEventsHost.OnSelectionChange tu cache MOT gia tri
' DUY NHAT (khong phan biet tai lieu), nen can buoc lai cache o day de lan WindowSelectionChange
' KE TIEP (hoac chinh su kien nay) danh gia lai dung tai lieu VUA active.
Private Sub oApp_WindowActivate(ByVal Doc As Document, ByVal Wn As Window)
    AppEventsHost.OnWindowActivate
End Sub

Private Sub oApp_DocumentBeforeClose(ByVal Doc As Document, Cancel As Boolean)
    If Application.Documents.count <= 1 Then AppEventsHost.OnLastDocumentClosing
End Sub
