Attribute VB_Name = "PageBreakInserter"
'==============================================================
' PageBreakInserter â€” hai nut "Chen trang ngang"/"Chen trang doc" (nhom Dinh dang, dat NGAY SAU
' tai vi tri con tro, ngat "Next Page" (thay cho Layout > Breaks > Next Page), doi Huong trang MOI
' thanh Ngang/Doc, giu DUNG bon le tren/duoi/trai/phai nhu trang truoc do (da doc TRUOC khi ngat),
' chen so trang tiep noi neu chua co.
' "Toi noi la le 'tren, duoi, trai, phai' chu khong noi theo 'canh ngan, canh dai' - ban can hieu
' dung y". Word.PageSetup.Orientation TU DONG HOAN VI bon gia tri Top/Bottom/Left/RightMargin ngay
' khi gan (hanh vi NGAM cua Word object model, khong phai loi cu the cua module nay) - vi du le
' Tren=85.05pt/Trai=99.2pt TRUOC khi doi huong tro thanh Tren=56.7pt/Trai=85.05pt SAU khi doi
' Orientation, DU KHONG HE dong den bon thuoc tinh margin. VI VAY module nay PHAI luu lai bon gia
' tri margin GOC truoc khi doi Orientation, roi GAN LAI (khong doc lai/khong suy luan tu
' PageWidth/PageHeight) sau khi doi xong, moi giu DUNG y nghia "tren/duoi/trai/phai" nguoi dung
' yeu cau (khac "canh ngan/canh dai" ma Word tu hoan vi).
' So trang: goi lai PageNumberFormatter.InsertPageNumbers SAU khi Utils.EndOperation cua rieng
' thao tac chen trang da dong xong (KHONG long BeginOperation - Application.UndoRecord va bien cap
' module cua Utils.bas KHONG tai nhap duoc, xem ghi chu dau TextFormatter.bas/
' ComponentFormatter.bas; dong quy uoc voi RibbonCallbacks.OnDungBoStyles goi hai ham TU BOC rieng
' lien tiep thay vi bao mot BeginOperation chung). InsertPageNumbers da idempotent (xoa moi field
' PAGE cu roi chen lai) va TU DONG lien mach qua moi section (LinkToPrevious, khong
' RestartNumberingAtSection) nen "chen so trang tiep noi neu chua co" duoc thoa MOT CACH TU NHIEN,
' khong can tu kiem tra rieng.
' Day la thao tac soan thao BINH THUONG (ngat section + doi PageSetup deu la API chuan cua Word
' Object Model, Ctrl+Z hoan tac tron ven) â€” KHONG can canh bao rui ro cao qua SafetyGuard (doi
' chieu ADR-007, khac OOXML/chuyen bang ma).
' Moi thu tuc cong khai bat dau bang On Error GoTo ErrHandler.
'==============================================================
Option Explicit

Public Sub InsertLandscapePage()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("PageBreakInserter.InsertLandscapePage")
    InsertPageWithOrientation wdOrientLandscape, "Ch" & ChrW(&H1EBF) & "n trang ngang"
End Sub

Public Sub InsertPortraitPage()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("PageBreakInserter.InsertPortraitPage")
    InsertPageWithOrientation wdOrientPortrait, "Ch" & ChrW(&H1EBF) & "n trang d" & ChrW(&H1ECD) & "c"
End Sub

Private Sub InsertPageWithOrientation(ByVal targetOrientation As WdOrientation, ByVal opName As String)
    On Error GoTo ErrHandler
    Dim opStarted As Boolean
    Utils.BeginOperation opName
    opStarted = True

    ' Doc bon le cua trang dang chua con tro TRUOC khi ngat section - "trang truoc do" theo dung
    ' nghia nguoi dung yeu cau.
    Dim curPs As word.PageSetup
    Set curPs = Selection.PageSetup

    Dim savedTop As Single, savedBottom As Single, savedLeft As Single, savedRight As Single
    savedTop = curPs.TopMargin
    savedBottom = curPs.BottomMargin
    savedLeft = curPs.LeftMargin
    savedRight = curPs.RightMargin

    Selection.InsertBreak Type:=wdSectionBreakNextPage

    Dim newPs As word.PageSetup
    Set newPs = Selection.PageSetup
    newPs.Orientation = targetOrientation

    ' Xem canh bao dau file - Word TU DONG hoan vi bon le khi doi Orientation, gan lai NGUYEN VAN
    ' bon gia tri da luu (KHONG doc lai/suy luan) de dung y "tren/duoi/trai/phai".
    newPs.TopMargin = savedTop
    newPs.BottomMargin = savedBottom
    newPs.LeftMargin = savedLeft
    newPs.RightMargin = savedRight

    Utils.EndOperation 1, False
    opStarted = False

    ' TU BOC rieng (xem ghi chu dau file) - goi SAU khi thao tac chen trang da EndOperation.
    PageNumberFormatter.InsertPageNumbers
    Exit Sub
ErrHandler:
    If opStarted Then Utils.AbortOperation Err.description
End Sub
