Attribute VB_Name = "ParagraphSnapshot"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'==============================================================
' ParagraphSnapshot â€” Anh chup MOT doan van, phan tu cua DocumentSnapshot.bas
' Doi ten truong o day thi phai doi ca ben do.
' Dung thuoc tinh cong khai truc tiep, cung kieu voi Finding.cls â€” day la du lieu "struct" thuan,
' khong co logic nghiep vu di kem. Chi doc, khong sua tai lieu.
' khi mot doan co nhieu phong/nhieu co/ nua dam nua khong, Range.Font tra ve "" (Name) hoac
' wdUndefined = 9999999 (Size, Bold, Italic). DocumentSnapshot.bas anh xa ve dung quy uoc chung
' cua ca hai ban:
' - chuoi hon hop -> "" (FontName)
' - so hon hop -> 0 (FontSizePt)
' - luan ly hon hop-> Null (Bold, Italic)
' Cac module doc snapshot nay PHAI kiem ba gia tri nay TRUOC khi so sanh.
'==============================================================
Option Explicit

' Thu tu trong tai lieu, tu 0 (khong phai chi so 1-based cua VBA Paragraphs). Trung voi
' Finding.ParagraphIndex va khoa cua DocumentLayoutMap.
Public Index As Long
' Noi dung doan, DA cat dau ket doan o cuoi (CR/LF, va U+0007 ket o bang) â€” xem StripParagraphMark
' trong DocumentSnapshot.bas. Khoang trang dau/cuoi con lai giu nguyen.
Public text As String
' Ten style hien thi cua Word (NameLocal), khong phai StyleID.
Public styleName As String
' "left" | "center" | "right" | "justify" | "unknown" â€” xem ToParagraphAlignment.
Public alignment As String
' "" khi doan co nhieu phong.
Public fontName As String
' 0 khi doan co nhieu co chu.
Public FontSizePt As Double
Public FontColor As String
' Null khi hon hop, nguoc lai True/False.
Public bold As Variant
Public Italic As Variant
' Chu trong doan toan hoa, tinh theo NOI DUNG (Utils.ToUpperVn(Text) = Text), KHONG doc
' Font.AllCaps â€” nguoi soan thuong go hoa truc tiep chu khong bat thuoc tinh do.
Public AllCaps As Boolean
Public FirstLineIndentPt As Double
Public SpaceBeforePt As Double
Public SpaceAfterPt As Double
' Gian dong quy ra POINT â€” muon biet boi so dong phai chia cho co chu.
Public lineSpacing As Double
Public isInTable As Boolean
' Chi so section chua doan, tu 0 (da quy doi tu Information(wdActiveEndSectionNumber), von 1-based
' ben VBA).
Public sectionIndex As Long
' So trang chua doan (1-based, tu Information(wdActiveEndPageNumber)).
Public PageNumber As Long

Private Sub Class_Initialize()
    bold = Null
    Italic = Null
End Sub
