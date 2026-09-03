Attribute VB_Name = "ParagraphFormatter"
'==============================================================
' ParagraphFormatter â€” Canh le, thut dau dong, gian dong, gian doan tren MOT vung cu the. Nhan
' rng do TANG GOI truyen vao, KHONG tu quyet dinh pham vi Chon/Toan bo.
' Word.ParagraphFormat.LineSpacingRule PHAI dat TRUOC LineSpacing, neu khong gia tri bi bo qua â€”
' "Luu y rieng cua VBA" #1. FirstLineIndent/SpaceAfter la don vi point â€” dung
' Utils.CmToPoint/MmToPoint, khong nhan chia thu cong.
' Moi thu tuc cong khai bat dau bang On Error GoTo ErrHandler. Dung AscW/ChrW, khong dung Asc/Chr
' (phu thuoc code page).
'==============================================================
Option Explicit

' Don vi noi bo cua Word cho LineSpacingRule "wdLineSpaceMultiple":
' Word.ParagraphFormat.LineSpacing tra ve (boi so x 12), KHONG PHU THUOC co chu -- hang so bieu
' dien CUA WORD, khong phai tham so ND 30, nen KHONG nam trong thong-so-the-thuc.json.
Private Const WORD_MULTIPLE_LINE_SPACING_UNIT_PT As Double = 12

' ============================================================================
' "thay cho cac thao tac: vao Menu Paragraph, chon tab Line and Page Breaks, tich chon Keep with
' next... ap dung cho vi tri con tro chuot hien tai, giong het Keep with next nguyen ban"). KHAC
' cac ham tren (nhan rng do TANG goi truyen vao, khong co nut ribbon rieng) -- ham nay TU DOC
' Selection (chinh la "vi tri con tro chuot hien tai" nguoi dung noi toi), vi day la nut ribbon
' GOI THANG, khong di qua checklist/danh sach doan nao ca. Neu Selection la mot vung chon nhieu
' doan, ap dung cho TAT CA doan trong vung do -- dung HANH VI cua hop thoai Paragraph nguyen ban
' cua Word khi co vung chon.
Public Sub ApplyKeepWithNextAtCursor()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ParagraphFormatter.ApplyKeepWithNextAtCursor")
    On Error GoTo ErrHandler
    Dim opName As String
    opName = "Keep with next"
    Utils.BeginOperation opName

    Dim count As Long: count = 0
    Dim p As word.paragraph
    For Each p In Selection.Range.paragraphs
        p.KeepWithNext = True
        count = count + 1
    Next p

    Utils.EndOperation count, False
    Exit Sub
ErrHandler:
    Utils.AbortOperation Err.description
End Sub
