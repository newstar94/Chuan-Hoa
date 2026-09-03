Attribute VB_Name = "FindingReporter"
Option Explicit

' Khai bao cap module â€” PHAI dung TRUOC toan bo Sub/Function (quy tac VBA; dat xen giua hai
' Function khien "Debug > Compile Project" bao loi "Only comments may appear after End Sub, End
' Function, or End Property").
Private mIsScanning As Boolean
Private mTextsReady As Boolean
Private TEXT_OP_NAME_FORMAT As String
Private TEXT_OP_NAME_SPELLING As String
Private TEXT_ERROR_PREFIX_FORMAT As String
Private TEXT_ERROR_PREFIX_SPELLING As String

' RibbonCallbacks.GetEnabledKiemTra doc qua day de lam mo CA HAI nut trong luc dang quet, tranh
' bam chong hai lan quet cung luc (ke ca chong giua hai nut khac nhau).
Public Function IsScanning() As Boolean
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("FindingReporter.IsScanning")
    IsScanning = mIsScanning
End Function

Private Sub EnsureTexts()
    If mTextsReady Then Exit Sub
    ' "Kiem tra the thuc" â€” ten thao tac dung cho Utils.BeginOperation/EndOperation (nhat ky
    ' phien, CLAUDE.md muc 3.6).
    TEXT_OP_NAME_FORMAT = "Ki" & ChrW(&H1EC3) & "m tra th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    ' "Kiem tra chinh ta"
    TEXT_OP_NAME_SPELLING = "Ki" & ChrW(&H1EC3) & "m tra ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    TEXT_ERROR_PREFIX_FORMAT = ChrW(&H110) & ChrW(&HE3) & " x" & ChrW(&H1EA3) & "y ra l" & _
        ChrW(&H1ED7) & "i khi ki" & ChrW(&H1EC3) & "m tra th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c."
    TEXT_ERROR_PREFIX_SPELLING = ChrW(&H110) & ChrW(&HE3) & " x" & ChrW(&H1EA3) & "y ra l" & _
        ChrW(&H1ED7) & "i khi ki" & ChrW(&H1EC3) & "m tra ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3) & "."
    mTextsReady = True
End Sub

' checklistGroup 1-12 (moi nhom TRU "spellingConversion"/"spellingLocalFix") â€” xem
' ComplianceChecker.ChecklistGroupIdOrder.
Private Function FormatChecklistGroups() As Variant
    FormatChecklistGroups = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)
End Function

' checklistGroup 13 ("spellingConversion" â€” dau thanh/i-y khong thong nhat toan van ban) + 14
' ("spellingLocalFix" â€” tu dien chinh ta/viet hoa/telex, 29 quy tac, xem ).
Private Function SpellingChecklistGroups() As Variant
    SpellingChecklistGroups = Array(13, 14)
End Function

' Diem vao "Kiem tra the thuc" (RibbonCallbacks.OnKiemTra).
Public Sub RunFormatCheck()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("FindingReporter.RunFormatCheck")
    EnsureTexts
    RunCheckCore FormatChecklistGroups(), FindingAnnotator.MarkerTagTheThuc(), _
        TEXT_OP_NAME_FORMAT, TEXT_ERROR_PREFIX_FORMAT, True
End Sub

' Diem vao "Kiem tra chinh ta" (RibbonCallbacks.OnKiemTraChinhTa, - MOI).
Public Sub RunSpellingCheck()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("FindingReporter.RunSpellingCheck")
    EnsureTexts
    RunCheckCore SpellingChecklistGroups(), FindingAnnotator.MarkerTagChinhTa(), _
        TEXT_OP_NAME_SPELLING, TEXT_ERROR_PREFIX_SPELLING, False
End Sub

Private Sub RunCheckCore(ByVal checklistGroups As Variant, ByVal marker As String, _
        ByVal opName As String, ByVal errorPrefix As String, ByVal includeDominantFontSizeNote As Boolean)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("FindingReporter.RunCheckCore")
    On Error GoTo ErrHandler
    Dim t0 As Double: t0 = Timer
    DebugTrace.Log "FindingReporter.RunCheckCore", "Bat dau: " & opName
    ' THU TU QUAN TRONG: cham vao ribbon (InvalidateKiemTra) TRUOC khi mo frmProcessing. Nhat ky
    ' lan Word chet 22/8/2026 cho thay lenh giet Word la InvalidateControl chay NGAY SAU
    ' ShowProcessing - mo/dong mot UserForm lam cua so tai lieu mat roi lay lai tieu diem
    ' (WindowDeactivate/WindowActivate), dung loai xao tron UI de lam con tro IRibbonUI cu chet
    ' ngam. Goi Invalidate luc van con o trong ngu canh onAction sach se, truoc moi xao tron.
    mIsScanning = True
    RibbonCallbacks.InvalidateKiemTra

    ProcessingIndicator.ResetDepth ' chot an toan cho dem long nhau - xem dau ProcessingIndicator.bas
    ProcessingIndicator.ShowProcessing

    Dim gateBlocked As Object
    If Not CheckGate.RunCheckGates(gateBlocked) Then
        ' Nguoi dung huy, hoac luu.docx that bai â€” giu nguyen tai lieu, khong quet.
        DebugTrace.Log "FindingReporter.RunCheckCore", "Cong chan tra False (nguoi dung huy) - thoat som"
        mIsScanning = False
        RibbonCallbacks.InvalidateKiemTra
        ProcessingIndicator.HideProcessing
        Exit Sub
    End If
    DebugTrace.Log "FindingReporter.RunCheckCore", "CheckGate.RunCheckGates xong, elapsed=" & Format$(Timer - t0, "0.00") & "s"

    ' "Viec chuan hoa xuong dong la hoan toan tu dong", bo han nut ribbon rieng) â€” TU DONG tach
    ' cac doan gop qua Shift+Enter (Quoc hieu/Tieu ngu/co quan/Noi nhan/nguoi ky) TRUOC KHI nhan
    ' dien+kiem tra, de moi buoc sau (BuildCheckContext, RunAllChecks) LUON thay cau truc doan
    ' SACH â€” khong con phai vong qua FindContinuationBreakPos o nhieu noi rieng le nhu truoc. TU
    ' BOC Utils.BeginOperation rieng (LineBreakNormalizer.NormalizeLineBreaks van con nguyen ham
    ' do TU NO, xem dau module do) - KHONG long vao BeginOperation cua chinh "Kiem tra" duoi day
    ' (Application.UndoRecord khong tai nhap duoc, dong quy uoc voi
    ' PageBreakInserter.bas/RibbonCallbacks.OnDungBoStyles). CHAY CHO CA HAI NUT: buoc chuan bi
    ' cau truc tai lieu, khong phu thuoc nhom quy tac.
    On Error Resume Next
    LineBreakNormalizer.NormalizeLineBreaks
    On Error GoTo 0
    DebugTrace.Log "FindingReporter.RunCheckCore", "LineBreakNormalizer.NormalizeLineBreaks xong, elapsed=" & Format$(Timer - t0, "0.00") & "s"

    ' dau cach/tab thua o HAI MEP doan â€” tu dong xoa, khong hoi lai) â€” TU DONG xoa TREN TOAN VAN
    ' BAN, TRUOC KHI nhan dien+kiem tra â€” cung vi tri, cung ly do voi LineBreakNormalizer o tren
    ' (doan phai "on dinh" truoc buoc BuildCheckContext).
    On Error Resume Next
    EdgeWhitespaceTrimmer.TrimEdgeWhitespace
    On Error GoTo 0
    DebugTrace.Log "FindingReporter.RunCheckCore", "EdgeWhitespaceTrimmer.TrimEdgeWhitespace xong, elapsed=" & Format$(Timer - t0, "0.00") & "s"

    On Error Resume Next
    FontVariantNormalizer.NormalizeFontVariants
    On Error GoTo 0
    DebugTrace.Log "FindingReporter.RunCheckCore", "FontVariantNormalizer.NormalizeFontVariants xong, elapsed=" & Format$(Timer - t0, "0.00") & "s"

    ' "Nhieu dau cach lien tiep: Neu chi co 2 thi xoa bot 1, neu nhieu hon 2 thi giu nguyen") â€” TU
    ' DONG gop DUNG 2 ky tu dau cach/tab lien tiep con 1, bo qua hoan toan (khong dung toi) truong
    ' hop 3+. Chay SAU EdgeWhitespaceTrimmer de khoang trang o hai mep doan da duoc don truoc â€”
    ' pham vi con lai chac chan nam O GIUA doan.
    On Error Resume Next
    MultiSpaceCollapser.CollapseDoubleSpaces
    On Error GoTo 0
    DebugTrace.Log "FindingReporter.RunCheckCore", "MultiSpaceCollapser.CollapseDoubleSpaces xong, elapsed=" & Format$(Timer - t0, "0.00") & "s"

    ' "Chua thay the dau ba cham bang dau cham lung") - TU DONG thay MOI cum "..." (3+ dau cham
    ' lien tiep) bang MOT ky tu dau cham lung U+2026, cung vi tri voi cac buoc co hoc khac o tren.
    On Error Resume Next
    EllipsisNormalizer.NormalizeEllipsis
    On Error GoTo 0
    DebugTrace.Log "FindingReporter.RunCheckCore", "EllipsisNormalizer.NormalizeEllipsis xong, elapsed=" & Format$(Timer - t0, "0.00") & "s"

    ' "Thay the tat ca cac bien the cua dau gach ngang (En Dash, Em Dash) bang dau gach noi
    ' (Hyphen -)") - TU DONG, cung vi tri voi cac buoc co hoc khac o tren.
    On Error Resume Next
    DashNormalizer.NormalizeDashes
    On Error GoTo 0
    DebugTrace.Log "FindingReporter.RunCheckCore", "DashNormalizer.NormalizeDashes xong, elapsed=" & Format$(Timer - t0, "0.00") & "s"

    On Error Resume Next
    BlankFieldSpacer.PadBlankFields
    On Error GoTo 0
    DebugTrace.Log "FindingReporter.RunCheckCore", "BlankFieldSpacer.PadBlankFields xong, elapsed=" & Format$(Timer - t0, "0.00") & "s"

    ' muc "Viec phai lam" #6 â€” boc trong Utils.BeginOperation de tat ScreenUpdating va ghi mot
    ' dong nhat ky phien cho thao tac nay (CLAUDE.md muc 3.6). opStarted danh dau ranh gioi de
    ' ErrHandler duoi day biet co can Utils.AbortOperation hay khong.
    Dim opStarted As Boolean
    Utils.BeginOperation opName
    opStarted = True

    Dim context As Object
    Set context = BuildCheckContext()
    DebugTrace.Log "FindingReporter.RunCheckCore", "BuildCheckContext xong, elapsed=" & Format$(Timer - t0, "0.00") & "s"

    Dim options As Object
    Set options = Utils.NewDictionary()
    options("OnProgressFunctionName") = "RibbonCallbacks.NotifyScanProgress"
    options("ChecklistGroupFilter") = checklistGroups

    Dim summary As Object
    Set summary = ComplianceChecker.RunAllChecks(context, options)
    DebugTrace.Log "FindingReporter.RunCheckCore", "ComplianceChecker.RunAllChecks xong, elapsed=" & _
        Format$(Timer - t0, "0.00") & "s, Findings=" & summary("Findings").count

    Application.StatusBar = False

    Dim allFindings As Collection: Set allFindings = summary("Findings")
    Dim keptFindings As New Collection
    Dim f As Finding
    For Each f In allFindings
        If gateBlocked Is Nothing Or Not gateBlocked.Exists(f.Group) Then
            keptFindings.Add f
        End If
    Next f

    ' xoa het comment CU mang DUNG marker cua nut nay (khong dong den comment cua nut kia) o DUY
    ' NHAT MOT DIEM, TRUOC CA HAI buoc chen comment moi duoi day.
    FindingAnnotator.ClearMarker marker

    Dim localFindings As Collection
    Set localFindings = FindingTierAggregator.ApplyTieredFindings(keptFindings, context, marker, includeDominantFontSizeNote)
    DebugTrace.Log "FindingReporter.RunCheckCore", "FindingTierAggregator.ApplyTieredFindings xong, elapsed=" & _
        Format$(Timer - t0, "0.00") & "s, localFindings=" & localFindings.count

    ' MOI phat hien â€” the thuc lan chinh ta (truoc day tach rieng nhom "spellingLocalFix" de
    ' highlight turquoise/do qua SpellingHighlighter.bas, da xoa) â€” nay di CHUNG mot co che DUY
    ' NHAT: Word Comment tai vi tri sai.
    FindingAnnotator.AnnotateFindings localFindings, marker
    DebugTrace.Log "FindingReporter.RunCheckCore", "FindingAnnotator.AnnotateFindings xong, elapsed=" & _
        Format$(Timer - t0, "0.00") & "s, localFindings=" & localFindings.count

    Utils.EndOperation keptFindings.count, False

    mIsScanning = False
    RibbonCallbacks.InvalidateKiemTra
    ProcessingIndicator.HideProcessing
    DebugTrace.Log "FindingReporter.RunCheckCore", "Hoan tat: " & opName & ", tong thoi gian=" & Format$(Timer - t0, "0.00") & "s"
    Exit Sub
ErrHandler:
    DebugTrace.LogErr "FindingReporter.RunCheckCore", "Loi giua chung (" & opName & "), elapsed=" & Format$(Timer - t0, "0.00") & "s", Err.number, Err.description
    Application.StatusBar = False
    mIsScanning = False
    RibbonCallbacks.InvalidateKiemTra
    If opStarted Then
        Utils.AbortOperation Err.description
    Else
        ProcessingIndicator.HideProcessing
        MsgBoxW.Show errorPrefix & vbCrLf & Err.description, vbExclamation, _
            "Chu" & ChrW(&H1EA9) & "n h" & ChrW(&HF3) & "a th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    End If
End Sub

Private Function BuildCheckContext() As Object
    Dim tStep As Double: tStep = Timer

    On Error GoTo StepSnapshot
    Dim snapshot As Object
    Set snapshot = DocumentSnapshot.CaptureDocument()
    DebugTrace.Log "FindingReporter.BuildCheckContext", "CaptureDocument xong, " & Format$(Timer - tStep, "0.00") & "s"

    On Error GoTo StepSpec
    Dim spec As Object
    Set spec = RuleLoader.GetFormatSpec()

    Dim regime As String: regime = RegimeState.GetSelectedRegime()

    On Error GoTo StepDocType
    Dim typeResult As Object
    Set typeResult = DocumentTypeDetector.DetectDocumentType(snapshot, regime)
    DebugTrace.Log "FindingReporter.BuildCheckContext", "DetectDocumentType xong, " & Format$(Timer - tStep, "0.00") & "s"

    On Error GoTo StepComponents
    Dim componentsResult As Object
    Set componentsResult = ComponentDetector.DetectComponents(snapshot, CStr(typeResult("Type")), regime)
    DebugTrace.Log "FindingReporter.BuildCheckContext", "DetectComponents xong, " & Format$(Timer - tStep, "0.00") & "s"

    On Error GoTo 0
    Dim context As Object
    Set context = Utils.NewDictionary()
    context.Add "Snapshot", snapshot
    context.Add "LayoutMap", componentsResult("LayoutMap")
    context.Add "DocumentType", CStr(typeResult("Type"))
    context.Add "Regime", regime
    context.Add "Spec", spec

    ' Dong bo drop-down "Loai van ban" (nhom Khoi dong, nut 1.4) voi ket qua vua nhan dien duoc.
    ' On Error Resume Next: day la buoc PHU, loi o day khong duoc lam gian doan luong "Kiem tra"
    ' chinh.
    On Error Resume Next
    Dim matchedOrdinal As Long
    matchedOrdinal = DocumentTypeDetector.MatchedTypeNameOrdinal(snapshot("Paragraphs"), typeResult("EvidenceParagraphIndex"))
    DocumentTypeState.AutoDetectAndStore CStr(typeResult("Type")), matchedOrdinal
    RibbonCallbacks.InvalidateLoaiVanBan
    RibbonCallbacks.InvalidateKiemTra
    On Error GoTo 0

    ' chen Word Comment "[DEBUG-CHUAN-HOA]" tren MOI doan da duoc gan vai tro â€” xem dau
    ' DebugAnnotator.bas (bat/tat qua DebugAnnotator.ENABLED, giu lai code khi phat hanh).
    On Error Resume Next
    DebugAnnotator.AnnotateLayoutMap snapshot("Paragraphs"), componentsResult("LayoutMap")
    On Error GoTo 0

    Set BuildCheckContext = context
    Exit Function

StepSnapshot:
    Err.Raise Err.number, "FindingReporter.BuildCheckContext", "[DocumentSnapshot.CaptureDocument] " & Err.description
StepSpec:
    Err.Raise Err.number, "FindingReporter.BuildCheckContext", "[RuleLoader.GetFormatSpec] " & Err.description
StepDocType:
    Err.Raise Err.number, "FindingReporter.BuildCheckContext", "[DocumentTypeDetector.DetectDocumentType] " & Err.description
StepComponents:
    Err.Raise Err.number, "FindingReporter.BuildCheckContext", "[ComponentDetector.DetectComponents] " & Err.description
End Function
