Attribute VB_Name = "DataReader"
Option Explicit

Public Sub RunAndReport()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("DataReader.RunAndReport")
    On Error GoTo ErrHandler
    ' Item 8,: DataReader KHONG di qua Utils.BeginOperation (thao tac CHI DOC, khong can
    ' UndoRecord - xem ghi chu dau file) nen TU goi ProcessingIndicator o day.
    ProcessingIndicator.ResetDepth ' chot an toan cho dem long nhau - xem dau ProcessingIndicator.bas
    ProcessingIndicator.ShowProcessing
    RunCore
    ProcessingIndicator.HideProcessing
    Exit Sub
ErrHandler:
    ProcessingIndicator.HideProcessing
    MsgBoxW.Show ChrW(&H110) & ChrW(&HE3) & " x" & ChrW(&H1EA3) & "y ra l" & ChrW(&H1ED7) & _
        "i khi " & ChrW(&H111) & ChrW(&H1ECD) & "c d" & ChrW(&H1EEF) & " li" & ChrW(&H1EC7) & _
        "u." & vbCrLf & Err.description, vbExclamation, ProductName()
End Sub

' bo Utils.ToUnaccented - dung lam title cho MsgBoxW.Show, khong con qua MsgBox chuan/ANSI nua.
Private Function ProductName() As String
    ProductName = "Chu" & ChrW(&H1EA9) & "n h" & ChrW(&HF3) & "a th" & ChrW(&H1EC3) & _
        " th" & ChrW(&H1EE9) & "c"
End Function

Private Function RunCore() As Object
    Dim Result As Object: Set Result = Utils.NewDictionary()

    Dim snapshot As Object: Set snapshot = DocumentSnapshot.CaptureDocument()

    ' Buoc 1 - Phat hien ma hoa (chi doc, KHONG ghi de tai lieu). Luu lai qua DataReadState de
    ' GetEnabledChuyenDoiUnicode (item 4) doc duoc.
    Dim encDetection As Object: Set encDetection = EncodingConverter.DetectEncoding()
    Result.Add "Encoding", encDetection
    DataReadState.SetNonUnicodeEncoding (CLng(encDetection("nonUnicodeCount")) > 0)

    ' Buoc 2 - nhan dien che do quy dinh (ND30/VIETTEL/DANG), dong bo drop- down "Quy dinh" + ve
    ' lai ribbon lien quan (chay TRUOC buoc nhan dien loai van ban/thanh phan duoi day, vi ca hai
    ' deu can biet regime hien tai).
    RegimeState.DetectAndAutoStore snapshot
    Dim regime As String: regime = RegimeState.GetSelectedRegime()

    ' Buoc 3 - nhan dien loai van ban, dong bo drop-down "Loai van ban" + ve lai ribbon lien quan.
    DocumentTypeState.DetectAndAutoStore snapshot

    ' Buoc 4 - Chen Word Comment cho tung thanh phan nhan dien duoc, CHI khi debug mode dang bat
    ' (DebugAnnotator.bas tu kiem ENABLED o dau AnnotateLayoutMap, tat thi Exit Sub ngay - tinh
    ' nang go loi, giu lai khi phat hanh).
    Dim typeResult As Object: Set typeResult = DocumentTypeDetector.DetectDocumentType(snapshot, regime)
    Dim componentsResult As Object
    Set componentsResult = ComponentDetector.DetectComponents(snapshot, CStr(typeResult("Type")), regime)
    DebugAnnotator.AnnotateLayoutMap snapshot("Paragraphs"), componentsResult("LayoutMap")
    Result.Add "ComponentCount", componentsResult("LayoutMap").count

    ' Buoc 5 - danh dau "da Doc du lieu" cho tai lieu nay, mo khoa moi nut khac tren ribbon
    ' (RibbonCallbacks.GetEnabledRequiresDataRead/GetEnabledKiemTra). Goi InvalidateRibbon truc
    ' tiep/dong bo (khong hoan qua Application.OnTime - xem ghi chu tai do).
    DataReadState.MarkReadData
    RibbonCallbacks.InvalidateRibbon "DocDuLieu"

    Set RunCore = Result
End Function
