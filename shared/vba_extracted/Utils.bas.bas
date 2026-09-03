Attribute VB_Name = "Utils"
Option Explicit

Private mCurrentOpName As String
Private mOpStartTime As Date

' 1 point = 20 twip la dinh nghia CO DINH cua don vi twip trong OOXML (ECMA-376 Part 1, muc 17.3.2
' â€” "twentieths of a point"), khong phai tham so do luong cua ND 30. Vi vay KHONG lay tu
' shared/rules/thong-so-the-thuc.json (units o do chi co mmToPoint/cmToPoint, hai don vi ND 30
' that su dung)
Private Const TWIP_PER_POINT As Double = 20

' ============================================================================
' Bao thao tac â€” dung o dau/cuoi moi thu tuc ribbon co sua doi tai lieu
' ============================================================================

' Tat ScreenUpdating, mo mot UndoRecord tuy chinh, ghi thoi diem bat dau. Goi EndOperation khi
' xong hoac AbortOperation trong ErrHandler â€” khong duoc bo qua ca hai.
Public Sub BeginOperation(ByVal opName As String)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("Utils.BeginOperation")
    On Error GoTo ErrHandler
    mCurrentOpName = opName
    mOpStartTime = Now

    ' hien hop thoai "Dang xu ly..." TRUOC KHI thuc thi
    ' - choke point nay bao trum HAU HET fix routine cua ca du an (moi ham deu di qua
    ' BeginOperation/EndOperation), xem dau ProcessingIndicator.bas.
    ProcessingIndicator.ShowProcessing

    Application.ScreenUpdating = False

    ' Boc trong On Error Resume Next: neu may dich khong ho tro thi thao tac van chay binh thuong,
    ' chi mat kha nang gom thanh MOT buoc Undo duy nhat cho nguoi dung.
    On Error Resume Next
    Application.UndoRecord.StartCustomRecord opName
    On Error GoTo ErrHandler

    ' "con dau" duy nhat con ghi vao file - xem dau DocumentSignature.bas. Goi O DAY vi
    ' BeginOperation bao trum HAU HET fix routine - chi ghi khi tai lieu THAT SU duoc xu ly, dung
    ' nghia "da su dung Add-in". Tinh nang phu, KHONG duoc lam gian doan thao tac chinh dang mo
    ' dau.
    On Error Resume Next
    DocumentSignature.EnsureSignature
    On Error GoTo ErrHandler

    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    ProcessingIndicator.HideProcessing
    Err.Raise Err.number, "Utils.BeginOperation", Err.description
End Sub

' Dong UndoRecord, bat lai ScreenUpdating, ghi log qua OperationLogger.
Public Sub EndOperation(ByVal affectedCount As Long, ByVal hasWarning As Boolean)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("Utils.EndOperation")
    On Error GoTo ErrHandler

    On Error Resume Next
    Application.UndoRecord.EndCustomRecord
    On Error GoTo ErrHandler

    Application.ScreenUpdating = True

    ' Item 8,: dong hop thoai "Dang xu ly..." NGAY KHI xu ly xong (tu dong dong, dung truoc
    ' RecordOperation - ghi log khong can nguoi dung phai cho them).
    ProcessingIndicator.HideProcessing

    RecordOperation affectedCount, hasWarning

    ' Ep xa bo dem nhat ky ra dia NGAY (khong qua nguong thoi gian cua FlushThrottled):
    DebugTrace.Flush

    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    ProcessingIndicator.HideProcessing
    Err.Raise Err.number, "Utils.EndOperation", Err.description
End Sub

' Dung trong ErrHandler cua thu tuc ribbon khi thao tac that bai giua chung. LUON bat lai
' ScreenUpdating truoc tien â€” day la nhanh cuoi cung, khong duoc de man hinh Word treo trang.
Public Sub AbortOperation(ByVal errDescription As String)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("Utils.AbortOperation")
    On Error Resume Next
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    ' Item 8,: dong hop thoai "Dang xu ly..." TRUOC KHI hien MsgBox loi ben duoi - khong de hai
    ' hop thoai chong len nhau.
    ProcessingIndicator.HideProcessing
    On Error GoTo 0

    RecordError errDescription

    ' Ep xa bo dem nhat ky ra dia NGAY (khong qua nguong thoi gian cua FlushThrottled):
    DebugTrace.Flush

    MsgBoxW.Show ChrW(&H110) & ChrW(&HE3) & " x" & ChrW(&H1EA3) & "y ra l" & ChrW(&H1ED7) & _
        "i khi th" & ChrW(&H1EF1) & "c hi" & ChrW(&H1EC7) & "n """ & mCurrentOpName & """." & _
        vbCrLf & errDescription, vbExclamation, ProductNameAccented()
End Sub

' "Chuáº©n hĂ³a thá»ƒ thá»©c" â€” tieu de MsgBoxW dung chung, thay cho hang so ASCII "Chuan hoa the thuc"
' cu. Dung ham (khong phai Const) vi VBA khong cho goi ChrW trong bieu thuc khoi tao Const.
Private Function ProductNameAccented() As String
    ProductNameAccented = "Chu" & ChrW(&H1EA9) & "n h" & ChrW(&HF3) & "a th" & ChrW(&H1EC3) & _
        " th" & ChrW(&H1EE9) & "c"
End Function

Public Function ToUnaccented(ByVal s As String) As String
    On Error GoTo ErrHandler
    Static fromChars As String, toChars As String, ready As Boolean
    If Not ready Then
        Dim fA As String, fE As String, fI As String, fO As String, fU As String, fY As String, fD As String
        fA = ChrW(&H103) & ChrW(&HE2) & ChrW(&HE0) & ChrW(&HE1) & ChrW(&H1EA3) & ChrW(&HE3) & ChrW(&H1EA1) & _
             ChrW(&H1EB1) & ChrW(&H1EAF) & ChrW(&H1EB3) & ChrW(&H1EB5) & ChrW(&H1EB7) & _
             ChrW(&H1EA7) & ChrW(&H1EA5) & ChrW(&H1EA9) & ChrW(&H1EAB) & ChrW(&H1EAD)
        fE = ChrW(&HEA) & ChrW(&HE8) & ChrW(&HE9) & ChrW(&H1EBB) & ChrW(&H1EBD) & ChrW(&H1EB9) & _
             ChrW(&H1EC1) & ChrW(&H1EBF) & ChrW(&H1EC3) & ChrW(&H1EC5) & ChrW(&H1EC7)
        fI = ChrW(&HEC) & ChrW(&HED) & ChrW(&H1EC9) & ChrW(&H129) & ChrW(&H1ECB)
        fO = ChrW(&HF4) & ChrW(&H1A1) & ChrW(&HF2) & ChrW(&HF3) & ChrW(&H1ECF) & ChrW(&HF5) & ChrW(&H1ECD) & _
             ChrW(&H1ED3) & ChrW(&H1ED1) & ChrW(&H1ED5) & ChrW(&H1ED7) & ChrW(&H1ED9) & _
             ChrW(&H1EDD) & ChrW(&H1EDB) & ChrW(&H1EDF) & ChrW(&H1EE1) & ChrW(&H1EE3)
        fU = ChrW(&H1B0) & ChrW(&HF9) & ChrW(&HFA) & ChrW(&H1EE7) & ChrW(&H169) & ChrW(&H1EE5) & _
             ChrW(&H1EEB) & ChrW(&H1EE9) & ChrW(&H1EED) & ChrW(&H1EEF) & ChrW(&H1EF1)
        fY = ChrW(&H1EF3) & ChrW(&HFD) & ChrW(&H1EF7) & ChrW(&H1EF9) & ChrW(&H1EF5)
        fD = ChrW(&H111)

        Dim lowerFrom As String
        lowerFrom = fA & fE & fI & fO & fU & fY & fD
        Dim lowerTo As String
        lowerTo = String$(Len(fA), "a") & String$(Len(fE), "e") & String$(Len(fI), "i") & _
                  String$(Len(fO), "o") & String$(Len(fU), "u") & String$(Len(fY), "y") & String$(Len(fD), "d")

        Dim uA As String, uE As String, uI As String, uO As String, uU As String, uY As String, uD As String
        uA = ChrW(&H102) & ChrW(&HC2) & ChrW(&HC0) & ChrW(&HC1) & ChrW(&H1EA2) & ChrW(&HC3) & ChrW(&H1EA0) & _
             ChrW(&H1EB0) & ChrW(&H1EAE) & ChrW(&H1EB2) & ChrW(&H1EB4) & ChrW(&H1EB6) & _
             ChrW(&H1EA6) & ChrW(&H1EA4) & ChrW(&H1EA8) & ChrW(&H1EAA) & ChrW(&H1EAC)
        uE = ChrW(&HCA) & ChrW(&HC8) & ChrW(&HC9) & ChrW(&H1EBA) & ChrW(&H1EBC) & ChrW(&H1EB8) & _
             ChrW(&H1EC0) & ChrW(&H1EBE) & ChrW(&H1EC2) & ChrW(&H1EC4) & ChrW(&H1EC6)
        uI = ChrW(&HCC) & ChrW(&HCD) & ChrW(&H1EC8) & ChrW(&H128) & ChrW(&H1ECA)
        uO = ChrW(&HD4) & ChrW(&H1A0) & ChrW(&HD2) & ChrW(&HD3) & ChrW(&H1ECE) & ChrW(&HD5) & ChrW(&H1ECC) & _
             ChrW(&H1ED2) & ChrW(&H1ED0) & ChrW(&H1ED4) & ChrW(&H1ED6) & ChrW(&H1ED8) & _
             ChrW(&H1EDC) & ChrW(&H1EDA) & ChrW(&H1EDE) & ChrW(&H1EE0) & ChrW(&H1EE2)
        uU = ChrW(&H1AF) & ChrW(&HD9) & ChrW(&HDA) & ChrW(&H1EE6) & ChrW(&H168) & ChrW(&H1EE4) & _
             ChrW(&H1EEA) & ChrW(&H1EE8) & ChrW(&H1EEC) & ChrW(&H1EEE) & ChrW(&H1EF0)
        uY = ChrW(&H1EF2) & ChrW(&HDD) & ChrW(&H1EF6) & ChrW(&H1EF8) & ChrW(&H1EF4)
        uD = ChrW(&H110)

        Dim upperFrom As String
        upperFrom = uA & uE & uI & uO & uU & uY & uD
        Dim upperTo As String
        upperTo = String$(Len(uA), "A") & String$(Len(uE), "E") & String$(Len(uI), "I") & _
                  String$(Len(uO), "O") & String$(Len(uU), "U") & String$(Len(uY), "Y") & String$(Len(uD), "D")

        fromChars = lowerFrom & upperFrom
        toChars = lowerTo & upperTo
        ready = True
    End If

    Dim Result As String
    Result = s
    Dim i As Long
    For i = 1 To Len(fromChars)
        Result = Replace(Result, Mid$(fromChars, i, 1), Mid$(toChars, i, 1))
    Next i
    ToUnaccented = Result
    Exit Function
ErrHandler:
    ToUnaccented = s ' loi bat ngo thi tra nguyen chuoi goc, khong chan hien thi
End Function

' Khoang cach Levenshtein CHUAN (khop toan bo hai chuoi, khong phai tim doan con) â€” dung cho doi
' chieu GAN DUNG hai nhan/ten ngan (vi du ten loai van ban voi tu dien chu-viet-tat-ten-loai.json).
' Muon tim mot cum GAN DUNG nam trong mot doan van DAI HON thi dung
' ComponentDetector.FuzzySubstringEditDistance (thuat toan khac â€” khop dau/cuoi tu do).
Public Function LevenshteinDistance(ByVal a As String, ByVal b As String) As Long
    Dim n As Long, m As Long
    n = Len(a)
    m = Len(b)
    If n = 0 Then
        LevenshteinDistance = m
        Exit Function
    End If
    If m = 0 Then
        LevenshteinDistance = n
        Exit Function
    End If

    Dim prev() As Long, cur() As Long
    ReDim prev(0 To m)
    ReDim cur(0 To m)
    Dim j As Long
    For j = 0 To m
        prev(j) = j
    Next j

    Dim i As Long
    Dim ach As String, bch As String
    Dim costSub As Long, costDel As Long, costIns As Long
    For i = 1 To n
        cur(0) = i
        ach = Mid$(a, i, 1)
        For j = 1 To m
            bch = Mid$(b, j, 1)
            If ach = bch Then costSub = prev(j - 1) Else costSub = prev(j - 1) + 1
            costDel = prev(j) + 1
            costIns = cur(j - 1) + 1
            cur(j) = costSub
            If costDel < cur(j) Then cur(j) = costDel
            If costIns < cur(j) Then cur(j) = costIns
        Next j
        For j = 0 To m
            prev(j) = cur(j)
        Next j
    Next i

    LevenshteinDistance = prev(m)
End Function

Private Sub RecordOperation(ByVal affectedCount As Long, ByVal hasWarning As Boolean)
    Dim op As New Operation
    op.Timestamp = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    op.name = mCurrentOpName
    op.affectedCount = affectedCount
    op.hasWarning = hasWarning

    ' Loi trong buoc ghi log khong duoc lam hong thao tac chinh â€” thao tac tren tai lieu da xong
    ' va da EndCustomRecord/bat lai ScreenUpdating truoc do roi.
    On Error Resume Next
    OperationLogger.LogOperation op
    On Error GoTo 0

    ' DebugTrace (thang 8/2026, tam thoi â€” xem dau DebugTrace.bas): moi thao tac THANH CONG cung
    ' ghi mot dong, du OperationLogger.bas chi giu trong phien. Choke point nay bao trum HAU HET
    ' fix routine cua ca du an (moi ham deu di qua BeginOperation/EndOperation).
    DebugTrace.Log mCurrentOpName, "OK - so doan anh huong=" & CStr(affectedCount) & _
        IIf(hasWarning, " (co canh bao)", "")
End Sub

Private Sub RecordError(ByVal errDescription As String)
    On Error Resume Next
    OperationLogger.LogError mCurrentOpName, errDescription
    On Error GoTo 0

    ' DebugTrace (thang 8/2026, tam thoi): choke point ghi TAT CA loi da bat duoc qua
    ' AbortOperation, bat ke goi tu module nao â€” xem dau DebugTrace.bas.
    DebugTrace.Log mCurrentOpName, "LOI: " & errDescription
End Sub

' Khong boi den (chi co diem chen, wdSelectionIP) thi tra Nothing â€” khac GetSelectionOrDocument o
' cho khong tu roi ve toan bo tai lieu, de noi goi tu bao nguoi dung "Hay boi den phan can ap"
' thay vi am tham ap toan van ban.
Public Function GetSelectionOnly() As Range
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("Utils.GetSelectionOnly")
    If Application.Selection.Type = wdSelectionIP Then
        Set GetSelectionOnly = Nothing
    Else
        Set GetSelectionOnly = Application.Selection.Range
    End If
End Function

' ============================================================================
' Chuoi va ky tu â€” dung AscW/ChrW, cam Asc/Chr (phu thuoc code page) â€” CLAUDE.md muc 5.
' ============================================================================

Public Function ToUpperVn(ByVal s As String) As String
    ToUpperVn = UCase$(s)
End Function

Public Function ToLowerVn(ByVal s As String) As String
    ToLowerVn = LCase$(s)
End Function

Public Function NormalizeNfc(ByVal s As String) As String
    Dim Result As Object
    Set Result = UnicodeNormalizer.NormalizeNfc(s)
    NormalizeNfc = CStr(Result("text"))
End Function

Public Function CmToPoint(ByVal cm As Double) As Double
    CmToPoint = cm * RuleLoader.GetFormatSpec()("units")("cmToPoint")
End Function

Public Function TwipToPoint(ByVal twip As Double) As Double
    TwipToPoint = twip / TWIP_PER_POINT
End Function

' ============================================================================
' Tu dien
' ============================================================================

' Boc CreateObject("Scripting.Dictionary") mot cho â€” sau nay neu may dich thieu Scripting.Runtime
' thi chi sua ham nay (vi du doi sang mot lop tu cai thay the).
Public Function NewDictionary() As Object
    Set NewDictionary = CreateObject("Scripting.Dictionary")
End Function
