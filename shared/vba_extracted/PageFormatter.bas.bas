Attribute VB_Name = "PageFormatter"
'==============================================================
' PageFormatter â€” Kho giay, le, huong giay - nut 2.1 "Dinh dang trang giay"
' Ban Legacy co loi the:
' QUY UOC CHI SO SECTION: Goi tu noi khac phai luu y lech nay.
' Moi thu tuc cong khai bat dau bang On Error GoTo ErrHandler. Dung AscW/ChrW, khong dung Asc/Chr
' (phu thuoc code page) - CLAUDE.md muc 5.
'==============================================================
Option Explicit

' Dung sai so sanh mm - bu sai so lam tron point<->mm qua hai chieu quy doi.
Private Const MARGIN_TOLERANCE_MM As Double = 0.01

' ============================================================================
' ApplyPageSetup â€” nut 2.1
' ============================================================================

Public Function ApplyPageSetup(Optional ByVal regimeCode As String = "ND30") As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("PageFormatter.ApplyPageSetup")
    On Error GoTo ErrHandler

    Dim opName As String
    opName = ChrW(&H110) & ChrW(&H1ECB) & "nh d" & ChrW(&H1EA1) & "ng trang gi" & _
        ChrW(&H1EA5) & "y"
    Utils.BeginOperation opName

    Dim spec As Object
    Set spec = RuleLoader.GetFormatSpec()("pageSetup")
    Dim factor As Double
    factor = RuleLoader.GetFormatSpec()("units")("mmToPoint")

    Dim margins As Object
    Set margins = MarginsForRegime(regimeCode, spec)

    Dim sectionCount As Long
    sectionCount = ActiveDocument.sections.count

    Dim i As Long
    For i = 1 To sectionCount
        ApplySectionPageSetup ActiveDocument.sections(i).PageSetup, (i = 1), spec, margins, factor
    Next i

    Dim Result As Object
    Set Result = Utils.NewDictionary()
    Result("sectionCount") = sectionCount
    Result("manualPageBreakCount") = CountManualPageBreaks()

    With ActiveDocument.sections(1).PageSetup
        Result("headerDistanceMm") = .HeaderDistance / factor
        Result("footerDistanceMm") = .FooterDistance / factor
    End With

    Utils.EndOperation sectionCount, False
    Set ApplyPageSetup = Result
    Exit Function
ErrHandler:
    Utils.AbortOperation Err.description
    Set ApplyPageSetup = Nothing
End Function

Private Function MarginsForRegime(ByVal regimeCode As String, ByVal defaultSpec As Object) As Object
    Dim regimes As Object
    Set regimes = RuleLoader.GetRegimeConfig()("regimes")

    If regimes.Exists(regimeCode) Then
        Dim regimeEntry As Object
        Set regimeEntry = regimes(regimeCode)
        If regimeEntry.Exists("margins") Then
            Set MarginsForRegime = regimeEntry("margins")
            Exit Function
        End If
    End If

    Set MarginsForRegime = defaultSpec("margins")
End Function

' Ap kho giay/le/titlePg cho MOT section
Private Sub ApplySectionPageSetup(ByVal ps As word.PageSetup, ByVal isFirstSection As Boolean, _
        ByVal spec As Object, ByVal margins As Object, ByVal factor As Double)
    Dim isPortrait As Boolean
    isPortrait = (ps.Orientation = wdOrientPortrait)

    ' Kho giay: A4, kich thuoc khop dung huong DANG CO cua section - khong tu doi huong giay (muc
    ' 6.1/1 cua: nguoi dung tu doi bang Section Break).
    ps.PaperSize = wdPaperA4
    If isPortrait Then
        ps.PageWidth = CDbl(spec("pageWidthMm")) * factor
        ps.PageHeight = CDbl(spec("pageHeightMm")) * factor
    Else
        ps.PageWidth = CDbl(spec("pageHeightMm")) * factor
        ps.PageHeight = CDbl(spec("pageWidthMm")) * factor
    End If

    ' giu nguyen quan he voi mep giay o CA HAI huong giay, khong hoan doi theo huong - quyet dinh
    ' tuong minh 2026-08-11 (docs/design/02-dac-ta-giao-dien.md muc 6.2). Nguon "margins" nay da
    ' duoc chon theo che do o MarginsForRegime.
    ps.TopMargin = ResolveMarginPoint(ps.TopMargin, margins("topMm"), factor)
    ps.BottomMargin = ResolveMarginPoint(ps.BottomMargin, margins("bottomMm"), factor)
    ps.LeftMargin = ResolveMarginPoint(ps.LeftMargin, margins("leftMm"), factor)
    ps.RightMargin = ResolveMarginPoint(ps.RightMargin, margins("rightMm"), factor)

    ' So trang lien mach, dieu kien 1/3 (muc 6.3 dac ta): titlePg CHI o section dau tien.
    ps.DifferentFirstPageHeaderFooter = isFirstSection

    ps.HeaderDistance = CDbl(spec("headerDistanceMm")) * factor

    ps.LayoutMode = wdLayoutModeDefault
End Sub

' Khong tu sua neu gia tri hien tai (point) da quy doi ve mm nam trong dai ND 30 cho phep - tra
' nguyen gia tri hien tai. Ngoai dai thi tra ve mac dinh. Dai lay tu RuleLoader
' (shared/rules/thong-so-the-thuc.json), khong hard-code (CLAUDE.md muc 3.1).
Private Function ResolveMarginPoint(ByVal currentPoint As Double, ByVal rangeSpec As Object, _
        ByVal factor As Double) As Double
    Dim currentMm As Double
    currentMm = currentPoint / factor

    If currentMm >= CDbl(rangeSpec("min")) - MARGIN_TOLERANCE_MM And _
            currentMm <= CDbl(rangeSpec("max")) + MARGIN_TOLERANCE_MM Then
        ResolveMarginPoint = currentPoint
    Else
        ResolveMarginPoint = CDbl(rangeSpec("default")) * factor
    End If
End Function

' ============================================================================
' Dem Page Break thuan
' ============================================================================

Private Function CountManualPageBreaks() As Long
    Dim txt As String
    txt = ActiveDocument.Content.text
    CountManualPageBreaks = Len(txt) - Len(Replace(txt, ChrW(12), ""))
End Function

' Be rong vung trinh bay (mm) cua section thu sectionIndex (1-based) - pageWidth tru hai le
' trai/phai, doc TRANG THAI THUC TE hien tai cua tai lieu (khong gia dinh da bam nut 2.1).
Public Function GetContentWidthMm(ByVal sectionIndex As Long) As Double
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("PageFormatter.GetContentWidthMm")
    On Error GoTo ErrHandler
    Dim ps As word.PageSetup
    Set ps = LoadSectionPageSetup(sectionIndex)

    Dim factor As Double
    factor = RuleLoader.GetFormatSpec()("units")("mmToPoint")

    GetContentWidthMm = (ps.PageWidth - ps.LeftMargin - ps.RightMargin) / factor
    Exit Function
ErrHandler:
    Err.Raise Err.number, "PageFormatter.GetContentWidthMm", Err.description
End Function

Private Function LoadSectionPageSetup(ByVal sectionIndex As Long) As word.PageSetup
    On Error GoTo ErrHandler
    If sectionIndex < 1 Or sectionIndex > ActiveDocument.sections.count Then
        Err.Raise vbObjectError + 515, "PageFormatter.LoadSectionPageSetup", _
            "Khong co section o chi so " & sectionIndex & " (tai lieu chi co " & _
            ActiveDocument.sections.count & " section)."
    End If
    Set LoadSectionPageSetup = ActiveDocument.sections(sectionIndex).PageSetup
    Exit Function
ErrHandler:
    Err.Raise Err.number, "PageFormatter.LoadSectionPageSetup", Err.description
End Function
