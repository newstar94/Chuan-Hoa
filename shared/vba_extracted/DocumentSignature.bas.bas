Attribute VB_Name = "DocumentSignature"
'==============================================================
' "Chi ghi an duy nhat mot noi dung: 'Tai lieu
' da su dung Add-in Chuan hoa the thuc viet boi Nguyen Ngoc Tien (@tiennn.ict)'... Uu tien viec
' ghi co dau. Viec nay co muc dich de cac cong cu, AI, website doc duoc, giup toi kiem tra su pho
' bien cua Add-in cung nhu mot 'con dau nho' de bien Add-in cua toi thuc su huu ich. No hoan toan
' vo hai doi voi nguoi su dung."
' Day la DU LIEU AN HOP LE DUY NHAT ma add-in nay con ghi vao file.docx (khac Document. Variables
' cua DocumentTypeState.bas/DataReadState.bas â€” da chuyen han sang SessionState.bas, KHONG con
' dong cham OOXML, xem o hai module do). Dung Custom Document Property (word/custom.xml sau khi
' giai nen OOXML) â€” day la co che CHUAN cua dinh dang.docx cho chinh muc dich nay, RAT NHIEU cong
' cu doc san khong can biet gi ve add-in nay (python-docx doc.custom_properties, ExifTool, Windows
' Explorer "File > Properties > Details > Custom"...).
' Goi tu Utils.BeginOperation â€” choke point bao trum HAU HET fix routine cua ca du an (moi thao
' tac sua doi tai lieu deu di qua day). Dung goi luc mo tai lieu (AppEventsHost.
' OnAnyDocumentOpen) hay luc doc du lieu (DataReader) â€” CHI ghi khi tai lieu THAT SU duoc add-in
' XU LY mot thao tac gi do, dung dung nghia "DA SU DUNG Add-in" (khong phai "da mo file co cai
' add-in"). Tu kiem tra gia tri hien co TRUOC khi ghi - tranh ghi lai vo ich + danh dau tai lieu
' la "Modified" khong can thiet moi lan bam nut.
'==============================================================
Option Explicit

Private Const PROPERTY_NAME As String = "ChuanHoaTheThuc"

Private mTextReady As Boolean
Private mSignatureText As String

Private Sub EnsureText()
    If mTextReady Then Exit Sub
    ' "Tai lieu da su dung Add-in Chuan hoa the thuc viet boi Nguyen Ngoc Tien (@tiennn.ict)"
    mSignatureText = "T" & ChrW(&HE0) & "i li" & ChrW(&H1EC7) & "u " & ChrW(&H111) & _
        ChrW(&HE3) & " s" & ChrW(&H1EED) & " d" & ChrW(&H1EE5) & "ng Add-in Chu" & _
        ChrW(&H1EA9) & "n h" & ChrW(&HF3) & "a th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & _
        "c vi" & ChrW(&H1EBF) & "t b" & ChrW(&H1EDF) & "i Nguy" & ChrW(&H1EC5) & "n Ng" & _
        ChrW(&H1ECD) & "c Ti" & ChrW(&H1EBF) & "n (@tiennn.ict)"
    mTextReady = True
End Sub

' Dam bao Custom Document Property da mang dung noi dung "con dau" â€” khong lam gi neu da dung san
' (tranh ghi lai vo ich). Goi tu Utils.BeginOperation, boc On Error Resume Next o noi goi - day la
' mot tinh nang PHU (con dau), khong duoc lam gian doan bat ky thao tac chinh nao.
Public Sub EnsureSignature()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("DocumentSignature.EnsureSignature")
    On Error GoTo ErrHandler
    EnsureText

    Dim current As String: current = ""
    On Error Resume Next
    current = CStr(ActiveDocument.CustomDocumentProperties(PROPERTY_NAME).value)
    On Error GoTo ErrHandler

    If current = mSignatureText Then Exit Sub ' da dung, khong ghi lai

    On Error Resume Next
    ActiveDocument.CustomDocumentProperties(PROPERTY_NAME).Delete
    On Error GoTo ErrHandler
    ActiveDocument.CustomDocumentProperties.Add name:=PROPERTY_NAME, LinkToContent:=False, _
        Type:=msoPropertyTypeString, value:=mSignatureText
    Exit Sub
ErrHandler:
    ' Tinh nang phu (con dau) - im lang bo qua, khong duoc chan bat ky thao tac chinh nao.
End Sub
