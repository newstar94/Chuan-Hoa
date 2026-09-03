Attribute VB_Name = "ComplianceChecker"
Option Explicit

' Bang dinh tuyen nut "Sua" (PHAN 6) -- khai bao cap module PHAI nam o dau module, truoc moi thu
' tuc (VBA: "Only comments may appear after End Sub...").

Private mRegistry As Object
Private mRegistryInitialized As Boolean

' - chan doan loi cua RunAllChecks (xem ghi chu tren dau ham).
Private mCurrentStep As String

Private Const PLACEHOLDER_FUNCTION_NAME As String = "ComplianceChecker.PlaceholderCheck"
' Long.MaxValue - dai dien "vo cung" cho paragraphIndex = Null khi sap xep (xep CUOI nhom).
Private Const MAX_LONG As Long = 2147483647
' Hai hang so cua Phan 2 - PHAI khai bao o day, cung ly do voi ghi chu tren.
Private Const ROUNDING_TOLERANCE_PT As Double = 0.75
Private Const WORD_MULTIPLE_LINE_SPACING_UNIT_PT As Double = 12

' ============================================================================
' So dang ky
' ============================================================================

Private Sub EnsureRegistryInitialized()
    If mRegistryInitialized Then Exit Sub
    Set mRegistry = Utils.NewDictionary()

    Dim rule As Variant
    Dim ruleCode As String
    For Each rule In RuleLoader.GetCheckRules()
        ruleCode = CStr(rule("ruleCode"))
        If Not mRegistry.Exists(ruleCode) Then mRegistry(ruleCode) = PLACEHOLDER_FUNCTION_NAME
    Next rule

    RegisterPageSetupBodyTextChecks
    RegisterComponentChecks
    RegisterStructureCitationChecks
    RegisterCapitalizationSpellingChecks
    mRegistryInitialized = True
End Sub

' ============================================================================
' Dung Finding tu CheckRule (phan tinh) + CheckFindingInput (phan dong, Dictionary voi cac khoa
' "ParagraphIndex"/"Message"/"CharOffset"/"Occurrences"/"Before"/"After").
' ============================================================================

Private Function ChecklistGroupIdOrder() As Variant
    ChecklistGroupIdOrder = Array( _
        "pageAndFont", "nationalTitleAndMotto", "organName", "codeNumber", "placeAndDate", _
        "typeNameAndSubject", "bodyContent", "signer", "recipient", "markingsAndContact", _
        "appendix", "tableAndImage", "spellingConversion", "spellingLocalFix")
End Function

Private Function ChecklistGroupIdFor(ByVal checklistGroup As Long) As String
    Dim order As Variant: order = ChecklistGroupIdOrder()
    Dim count As Long: count = UBound(order) - LBound(order) + 1
    If checklistGroup < 1 Or checklistGroup > count Then
        Err.Raise vbObjectError + 515, "ComplianceChecker.ChecklistGroupIdFor", _
            "checklistGroup " & checklistGroup & " khong hop le (phai trong khoang 1-" & count & ")."
    End If
    ChecklistGroupIdFor = CStr(order(LBound(order) + checklistGroup - 1))
End Function

Private Function BuildFinding(ByVal rule As Object, ByVal inputData As Object, _
        ByVal options As Object) As Finding
    Dim f As New Finding
    f.ruleCode = rule("ruleCode")
    ' f.Group PHAI la ChecklistGroupId (14 gia tri UI, "spellingLocalFix"...) -- KHONG PHAI
    ' rule("group") (CheckRuleGroup, 8 gia tri ky thuat, "spelling"...).
    f.Group = ChecklistGroupIdFor(CLng(rule("checklistGroup")))
    f.Severity = rule("severity")
    f.SourceLabel = rule("sourceLabel")
    f.title = rule("title")
    f.message = inputData("Message")
    f.Citation = rule("citation")
    f.paragraphIndex = inputData("ParagraphIndex")
    If inputData.Exists("CharOffset") Then f.charOffset = inputData("CharOffset")
    If inputData.Exists("Occurrences") Then f.occurrences = inputData("Occurrences")

    ' warnOnly KHONG BAO GIO co nut "Sua" (FR-CHK-14) - ep ve False tai day lam LUOI AN TOAN cuoi
    ' cung, khong phu thuoc tung ham kiem tu nho dung quy tac nay.
    If CStr(rule("checkability")) = "warnOnly" Then
        f.AutoFixable = False
    Else
        f.AutoFixable = CBool(rule("autoFixable"))
    End If
    f.RiskLevel = rule("riskLevel")
    ' Loai C (sua tung cho trong sidebar) luon bat buoc xac nhan - ADR-009. Loai A/B ap truc tiep
    ' qua nut ribbon, khong qua luong xac nhan tung Finding nay.
    f.RequiresConfirmation = (CStr(rule("actionType")) = "C")
    If inputData.Exists("Before") Then f.Before = CStr(inputData("Before"))
    If inputData.Exists("After") Then f.After = CStr(inputData("After"))
    f.Checkability = rule("checkability")

    Set BuildFinding = f
End Function

' ============================================================================
' Chay MOT quy tac â€” dung chung cho ca RunAllChecks va RunSingleCheck. Tra Dictionary voi ba khoa:
' "Status" ("notChecked"|"pass"|"fail"), "Findings" (Collection cua Finding, RONG khi
' notChecked/pass), "FailCount" (Long, CHI dem Finding CHUA bi bo qua).
' ============================================================================

Private Function RunRule(ByVal rule As Object, ByVal context As Object, _
        ByVal options As Object) As Object
    Dim ruleCode As String
    ruleCode = CStr(rule("ruleCode"))

    Dim fnName As String
    If mRegistry.Exists(ruleCode) Then
        fnName = CStr(mRegistry(ruleCode))
    Else
        fnName = PLACEHOLDER_FUNCTION_NAME
    End If

    Dim Result As Object
    Set Result = Utils.NewDictionary()
    Result.Add "RuleCode", ruleCode
    Result.Add "ChecklistGroup", CLng(rule("checklistGroup"))

    ' Quy tac chua co ham kiem tra that (co trong JSON nhung chua RegisterXxx nao gan - xem
    ' EnsureRegistryInitialized) mang ten PLACEHOLDER_FUNCTION_NAME, khong phai ham that -
    ' Application.Run vao ten do se nem loi. Coi nhu "notChecked" ngay tai day.
    If fnName = PLACEHOLDER_FUNCTION_NAME Then
        Result.Add "Status", "notChecked"
        Result.Add "Findings", New Collection
        Result.Add "FailCount", 0
        Set RunRule = Result
        Exit Function
    End If

    On Error GoTo RuleErrHandler
    Dim raw As Object
    Set raw = Application.run(fnName, context)

    If raw Is Nothing Then
        Result.Add "Status", "notChecked"
        Result.Add "Findings", New Collection
        Result.Add "FailCount", 0
        Set RunRule = Result
        Exit Function
    End If

    Dim findings As New Collection
    Dim failCount As Long
    failCount = 0

    Dim item As Variant
    Dim f As Finding
    For Each item In raw
        Set f = BuildFinding(rule, item, options)
        findings.Add f
        failCount = failCount + 1
    Next item

    Result.Add "Status", IIf(failCount > 0, "fail", "pass")
    Result.Add "Findings", findings
    Result.Add "FailCount", failCount
    Set RunRule = Result
    Exit Function

RuleErrHandler:
    ' Ghep ruleCode/fnName vao DAU Err.Description (khong chi Err.Source) - MsgBox cuoi cung
    ' (Utils.AbortOperation) chi hien Description, khong hien Source, xem ghi chu tren dau ham.
    Err.Raise Err.number, "ComplianceChecker.RunRule", _
        "[" & ruleCode & " / " & fnName & "] " & Err.description
End Function

' ============================================================================
' RunAllChecks â€” quet toan bo 100 quy tac
' ============================================================================

Private Function ChecklistGroupFilterOf(ByVal options As Object) As Object
    If options Is Nothing Then
        Set ChecklistGroupFilterOf = Nothing
        Exit Function
    End If
    If Not options.Exists("ChecklistGroupFilter") Then
        Set ChecklistGroupFilterOf = Nothing
        Exit Function
    End If
    Dim src As Variant: src = options("ChecklistGroupFilter")
    If IsEmpty(src) Or IsNull(src) Then
        Set ChecklistGroupFilterOf = Nothing
        Exit Function
    End If
    Dim Result As Object: Set Result = Utils.NewDictionary()
    Dim v As Variant
    For Each v In src
        Result(CStr(CLng(v))) = True
    Next v
    Set ChecklistGroupFilterOf = Result
End Function

Private Function CountRulesInFilter(ByVal rules As Collection, ByVal groupFilter As Object) As Long
    If groupFilter Is Nothing Then
        CountRulesInFilter = rules.count
        Exit Function
    End If
    Dim n As Long: n = 0
    Dim rule As Variant
    For Each rule In rules
        If groupFilter.Exists(CStr(CLng(rule("checklistGroup")))) Then n = n + 1
    Next rule
    CountRulesInFilter = n
End Function

Public Function RunAllChecks(ByVal context As Object, Optional ByVal options As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.RunAllChecks")
    On Error GoTo ErrHandler
    mCurrentStep = "EnsureRegistryInitialized"
    EnsureRegistryInitialized

    mCurrentStep = "RuleLoader.GetCheckRules"
    Dim rules As Collection
    Set rules = RuleLoader.GetCheckRules()

    Dim results As New Collection
    Dim allFindings As New Collection

    Dim groupFilter As Object
    Set groupFilter = ChecklistGroupFilterOf(options)

    Dim total As Long
    total = CountRulesInFilter(rules, groupFilter)
    Dim completed As Long
    completed = 0

    Dim rule As Variant
    Dim ruleResult As Object
    Dim f As Variant
    mCurrentStep = "vong lap RunRule"
    For Each rule In rules
        If Not groupFilter Is Nothing Then
            If Not groupFilter.Exists(CStr(CLng(rule("checklistGroup")))) Then GoTo ContinueLoop
        End If
        Set ruleResult = RunRule(rule, context, options)
        results.Add ruleResult

        For Each f In ruleResult("Findings")
            allFindings.Add f
        Next f

        completed = completed + 1
        mCurrentStep = "CallOnProgress"
        CallOnProgress options, completed, total
        mCurrentStep = "vong lap RunRule"
ContinueLoop:
    Next rule

    mCurrentStep = "SortFindings"
    Dim summary As Object
    Set summary = Utils.NewDictionary()
    summary.Add "Results", results
    summary.Add "Findings", SortFindings(allFindings, rules)
    Set RunAllChecks = summary
    Exit Function
ErrHandler:
    Err.Raise Err.number, "ComplianceChecker.RunAllChecks", "[" & mCurrentStep & "] " & Err.description
End Function

' ============================================================================
' Tien ich dung chung
' ============================================================================

Private Sub CallOnProgress(ByVal options As Object, ByVal completed As Long, ByVal total As Long)
    If options Is Nothing Then Exit Sub
    If Not options.Exists("OnProgressFunctionName") Then Exit Sub
    Dim fnName As String
    fnName = CStr(options("OnProgressFunctionName"))
    If Len(fnName) = 0 Then Exit Sub
    Application.run fnName, completed, total
End Sub

Private Function ParagraphIndexOrMax(ByVal paragraphIndex As Variant) As Long
    If IsNull(paragraphIndex) Then
        ParagraphIndexOrMax = MAX_LONG
    Else
        ParagraphIndexOrMax = CLng(paragraphIndex)
    End If
End Function

Private Function SortFindings(ByVal findings As Collection, ByVal rules As Collection) As Collection
    Dim groupByCode As Object
    Set groupByCode = Utils.NewDictionary()
    Dim rule As Variant
    For Each rule In rules
        groupByCode(CStr(rule("ruleCode"))) = CLng(rule("checklistGroup"))
    Next rule

    Dim n As Long
    n = findings.count
    Dim arr() As Object
    If n > 0 Then ReDim arr(0 To n - 1)

    Dim i As Long
    i = 0
    Dim f As Variant
    For Each f In findings
        Set arr(i) = f
        i = i + 1
    Next f

    Dim j As Long
    Dim key As Object
    Dim keyGroup As Long, keyIndex As Long
    Dim cmpGroup As Long, cmpIndex As Long
    For i = 1 To n - 1
        Set key = arr(i)
        keyGroup = groupByCode(key.ruleCode)
        keyIndex = ParagraphIndexOrMax(key.paragraphIndex)

        j = i - 1
        Do While j >= 0
            cmpGroup = groupByCode(arr(j).ruleCode)
            cmpIndex = ParagraphIndexOrMax(arr(j).paragraphIndex)

            If cmpGroup > keyGroup Or (cmpGroup = keyGroup And cmpIndex > keyIndex) Then
                Set arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        Set arr(j + 1) = key
    Next i

    Dim Result As New Collection
    For i = 0 To n - 1
        Result.Add arr(i)
    Next i
    Set SortFindings = Result
End Function

Private Function PointToMm(ByVal pt As Double, ByVal spec As Object) As Double
    PointToMm = pt / spec("units")("mmToPoint")
End Function

Private Function PointToCm(ByVal pt As Double, ByVal spec As Object) As Double
    PointToCm = pt / spec("units")("cmToPoint")
End Function

Private Function Min2(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then Min2 = a Else Min2 = b
End Function

Private Function Max2(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then Max2 = a Else Max2 = b
End Function

' So kieu Viet Nam: dau phay thap phan, bo so 0 thua -- CHI dung de DUNG THONG DIEP, khong dung de
' so sanh. Tu lam tron bang so nguyen (KHONG dung Format$, vi ky tu thap phan cua Format$ phu
' thuoc locale he thong -- co may la ",", co may la "." -- se lam sai chinh cai dieu ham nay dinh
' chuan hoa).
Private Function FormatVnNumber(ByVal value As Double, ByVal maxDecimals As Long) As String
    Dim factor As Long
    factor = CLng(10 ^ maxDecimals)
    Dim scaled As Long
    scaled = CLng(Int(value * factor + 0.5))

    Dim intPart As Long, fracPart As Long
    intPart = scaled \ factor
    fracPart = scaled Mod factor

    If fracPart = 0 Then
        FormatVnNumber = CStr(intPart)
        Exit Function
    End If

    Dim fracStr As String
    fracStr = CStr(fracPart)
    Do While Len(fracStr) < maxDecimals
        fracStr = "0" & fracStr
    Loop
    Do While Right$(fracStr, 1) = "0"
        fracStr = left$(fracStr, Len(fracStr) - 1)
    Loop
    FormatVnNumber = CStr(intPart) & "," & fracStr
End Function

' Dung Dictionary voi cac khoa "ParagraphIndex"/"Message"/"Occurrences"/"CharOffset" -- dung dinh
' dang inputData ma BuildFinding da doc o Phan 1. charOffset dung cho cac quy tac citation can chi
' ro vi tri khop trong doan (nhieu vien dan tren cung mot dong).
Private Function MakeFindingInput(ByVal paragraphIndex As Variant, ByVal message As String, _
        Optional ByVal occurrences As Variant, Optional ByVal charOffset As Variant, _
        Optional ByVal beforeText As Variant, Optional ByVal afterText As Variant) As Object
    Dim item As Object
    Set item = Utils.NewDictionary()
    item.Add "ParagraphIndex", paragraphIndex
    item.Add "Message", message
    If Not IsMissing(occurrences) Then item.Add "Occurrences", occurrences
    If Not IsMissing(charOffset) Then item.Add "CharOffset", charOffset
    ' Goi cu khong truyen hai tham so nay van dung nguyen, khong can sua.
    If Not IsMissing(beforeText) Then item.Add "Before", CStr(beforeText)
    If Not IsMissing(afterText) Then item.Add "After", CStr(afterText)
    Set MakeFindingInput = item
End Function

' ----------------------------------------------------------------------------
' Nhom pageSetup -- kiem tren Snapshot("Sections")
' ----------------------------------------------------------------------------

Private Function CheckableSections(ByVal snapshot As Object) As Collection
    Dim Result As New Collection
    Dim section As Variant
    For Each section In snapshot("Sections")
        If CBool(section("PageSetupAvailable")) Then Result.Add section
    Next section
    Set CheckableSections = Result
End Function

' ND30-PL1-M1-K1 -- Kho giay A4. So theo CAP (canh ngan, canh dai) de khong phu thuoc huong giay
' -- A4 ngang van la A4, chi hoan doi PageWidth/PageHeight.
Public Function CheckPageSizeA4(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckPageSizeA4")
    Dim spec As Object: Set spec = context("Spec")
    Dim sections As Collection: Set sections = CheckableSections(context("Snapshot"))
    If sections.count = 0 Then
        Set CheckPageSizeA4 = Nothing
        Exit Function
    End If

    Dim pageWidthMm As Double, pageHeightMm As Double
    pageWidthMm = spec("pageSetup")("pageWidthMm")
    pageHeightMm = spec("pageSetup")("pageHeightMm")
    Dim expectedShortMm As Double, expectedLongMm As Double
    expectedShortMm = Min2(pageWidthMm, pageHeightMm)
    expectedLongMm = Max2(pageWidthMm, pageHeightMm)
    Dim toleranceMm As Double
    toleranceMm = PointToMm(ROUNDING_TOLERANCE_PT, spec)

    Dim findings As New Collection
    Dim section As Variant
    For Each section In sections
        Dim widthMm As Double, heightMm As Double
        widthMm = PointToMm(CDbl(section("PageWidthPt")), spec)
        heightMm = PointToMm(CDbl(section("PageHeightPt")), spec)
        Dim shortMm As Double, longMm As Double
        shortMm = Min2(widthMm, heightMm)
        longMm = Max2(widthMm, heightMm)

        If Abs(shortMm - expectedShortMm) > toleranceMm Or Abs(longMm - expectedLongMm) > toleranceMm Then
            Dim msg As String
            msg = "Section " & CStr(CLng(section("Index")) + 1) & ": kh" & ChrW(&H1ED5) & " gi" & _
                ChrW(&H1EA5) & "y hi" & ChrW(&H1EC7) & "n t" & ChrW(&H1EA1) & "i " & FormatVnNumber(widthMm, 1) & _
                " x " & FormatVnNumber(heightMm, 1) & " mm, quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh l" & _
                ChrW(&HE0) & " A4 (" & CStr(pageWidthMm) & " x " & CStr(pageHeightMm) & " mm)."
            findings.Add MakeFindingInput(Null, msg)
        End If
    Next section

    Set CheckPageSizeA4 = findings
End Function

' ND30-PL1-M1-K2 -- Huong giay nam ngang chi la NGHI NGO (severity "warning" trong CheckRule): may
' chi phat hien CHAC CHAN "co dat ngang", con "co hop le khong" (bang/bieu khong tach phu luc
' rieng) la ngu nghia, de nguoi dung tu quyet.
Public Function CheckPageOrientation(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckPageOrientation")
    Dim sections As Collection: Set sections = CheckableSections(context("Snapshot"))
    If sections.count = 0 Then
        Set CheckPageOrientation = Nothing
        Exit Function
    End If

    Dim findings As New Collection
    Dim section As Variant
    For Each section In sections
        If CStr(section("Orientation")) = "landscape" Then
            Dim msg As String
            msg = "Section " & CStr(CLng(section("Index")) + 1) & " " & ChrW(&H111) & ChrW(&H1EB7) & _
                "t ngang. Ch" & ChrW(&H1EC9) & " " & ChrW(&H111) & ChrW(&H1B0) & ChrW(&H1EE3) & "c " & _
                ChrW(&H111) & ChrW(&H1EB7) & "t ngang khi n" & ChrW(&H1ED9) & "i dung c" & ChrW(&HF3) & _
                " b" & ChrW(&H1EA3) & "ng, bi" & ChrW(&H1EC3) & "u kh" & ChrW(&HF4) & "ng l" & ChrW(&HE0) & _
                "m th" & ChrW(&HE0) & "nh ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c ri" & ChrW(&HEA) & "ng."
            findings.Add MakeFindingInput(Null, msg)
        End If
    Next section

    Set CheckPageOrientation = findings
End Function

Private Function MarginSideLabel(ByVal side As String) As String
    Select Case side
        Case "top"
            MarginSideLabel = "tr" & ChrW(&HEA) & "n"
        Case "bottom"
            MarginSideLabel = "d" & ChrW(&H1B0) & ChrW(&H1EDB) & "i"
        Case "left"
            MarginSideLabel = "tr" & ChrW(&HE1) & "i"
        Case "right"
            MarginSideLabel = "ph" & ChrW(&H1EA3) & "i"
    End Select
End Function

Private Sub AddMarginFinding(ByVal findings As Collection, ByVal section As Object, _
        ByVal marginRange As Object, ByVal actualPt As Double, ByVal toleranceMm As Double, _
        ByVal spec As Object, ByVal sideLabel As String)
    Dim actualMm As Double: actualMm = PointToMm(actualPt, spec)
    Dim minMm As Double: minMm = CDbl(marginRange("min"))
    Dim maxMm As Double: maxMm = CDbl(marginRange("max"))

    If actualMm < minMm - toleranceMm Or actualMm > maxMm + toleranceMm Then
        Dim msg As String
        msg = "Section " & CStr(CLng(section("Index")) + 1) & ": l" & ChrW(&H1EC1) & " " & sideLabel & _
            " l" & ChrW(&HE0) & " " & FormatVnNumber(actualMm, 1) & " mm, d" & ChrW(&H1EA3) & "i cho " & _
            "ph" & ChrW(&HE9) & "p " & CStr(minMm) & "-" & CStr(maxMm) & " mm."
        findings.Add MakeFindingInput(Null, msg)
    End If
End Sub

' ND30-PL1-M1-K3 -- Bon le trang trong dai ND 30, KHONG so voi gia tri mac dinh.
Public Function CheckPageMargins(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckPageMargins")
    Dim spec As Object: Set spec = context("Spec")
    Dim sections As Collection: Set sections = CheckableSections(context("Snapshot"))
    If sections.count = 0 Then
        Set CheckPageMargins = Nothing
        Exit Function
    End If

    Dim margins As Object: Set margins = spec("pageSetup")("margins")
    Dim toleranceMm As Double: toleranceMm = PointToMm(ROUNDING_TOLERANCE_PT, spec)

    Dim findings As New Collection
    Dim section As Variant
    For Each section In sections
        AddMarginFinding findings, section, margins("topMm"), CDbl(section("TopMarginPt")), toleranceMm, spec, MarginSideLabel("top")
        AddMarginFinding findings, section, margins("bottomMm"), CDbl(section("BottomMarginPt")), toleranceMm, spec, MarginSideLabel("bottom")
        AddMarginFinding findings, section, margins("leftMm"), CDbl(section("LeftMarginPt")), toleranceMm, spec, MarginSideLabel("left")
        AddMarginFinding findings, section, margins("rightMm"), CDbl(section("RightMarginPt")), toleranceMm, spec, MarginSideLabel("right")
    Next section

    Set CheckPageMargins = findings
End Function

Public Function CheckPageNumbering(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckPageNumbering")
    Set CheckPageNumbering = Nothing
End Function

' ND30-PL1-M3-K1D -- So trang phu luc danh so rieng theo tung phu luc. CHUA HIEN THUC HOA DUOC,
' cung ly do voi CheckPageNumbering o tren.
Public Function CheckAppendixPageNumberingRestart(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckAppendixPageNumberingRestart")
    Set CheckAppendixPageNumberingRestart = Nothing
End Function

' ----------------------------------------------------------------------------
' Nhom bodyText
' ----------------------------------------------------------------------------

' Gom cac doan co gia tri (Font/Color) khac "expected" thanh mot Finding cho moi gia tri sai
' khac nhau (bo qua gia tri rong - hon hop/khong xac dinh). which = "Font" | "Color";
' uppercaseCompare = True cho mau (khong phan biet hoa/thuong).
' useMajorityTier = True: gia tri sai chiem qua nua so doan da quet duoc coi la loi TOAN CUC
' (ParagraphIndex tra Null - FindingAnnotator gop thanh mot comment o dau van ban); thieu so
' van giu vi tri that. useMajorityTier = False (dung cho K4-COLOR): luon giu vi tri doan
' XUAT HIEN DAU TIEN, du la da so hay thieu so.
Private Function GroupWrongParagraphsByValue(ByVal paragraphs As Collection, ByVal which As String, _
        ByVal expected As String, ByVal uppercaseCompare As Boolean, _
        Optional ByVal useMajorityTier As Boolean = True) As Collection
    Dim firstIndexByValue As Object: Set firstIndexByValue = Utils.NewDictionary()
    Dim countByValue As Object: Set countByValue = Utils.NewDictionary()
    Dim orderedValues As New Collection
    Dim totalScanned As Long: totalScanned = 0

    Dim p As ParagraphSnapshot
    Dim value As String
    For Each p In paragraphs
        If which = "Font" Then value = p.fontName Else value = p.FontColor
        If uppercaseCompare Then value = UCase$(value)
        If value <> "" Then
            totalScanned = totalScanned + 1
            If value <> expected Then
                If countByValue.Exists(value) Then
                    countByValue(value) = countByValue(value) + 1
                Else
                    countByValue(value) = 1
                    firstIndexByValue(value) = p.Index
                    orderedValues.Add value
                End If
            End If
        End If
    Next p

    Dim Result As New Collection
    Dim v As Variant
    For Each v In orderedValues
        Dim msg As String
        If which = "Font" Then
            msg = "Ph" & ChrW(&HE1) & "t hi" & ChrW(&H1EC7) & "n " & CStr(countByValue(v)) & " " & _
                ChrW(&H111) & "o" & ChrW(&H1EA1) & "n d" & ChrW(&HF9) & "ng ph" & ChrW(&HF4) & "ng " & CStr(v) & "."
        Else
            msg = "Ph" & ChrW(&HE1) & "t hi" & ChrW(&H1EC7) & "n " & CStr(countByValue(v)) & " " & _
                ChrW(&H111) & "o" & ChrW(&H1EA1) & "n c" & ChrW(&HF3) & " m" & ChrW(&HE0) & "u ch" & ChrW(&H1EEF) & " " & CStr(v) & "."
        End If

        Dim wrongCount As Long: wrongCount = CLng(countByValue(v))
        Dim paraIdxVariant As Variant
        If useMajorityTier And totalScanned > 0 And wrongCount * 2 > totalScanned Then
            paraIdxVariant = Null
        Else
            paraIdxVariant = CLng(firstIndexByValue(v))
        End If
        Result.Add MakeFindingInput(paraIdxVariant, msg, wrongCount)
    Next v

    Set GroupWrongParagraphsByValue = Result
End Function

' ND30-PL1-M1-K4-FONT -- Phong Times New Roman, ap dung TOAN VAN BAN (Muc I khoan 4, khong rieng
' phan loi van) -- vi vay quet Snapshot("Paragraphs") thang, khong loc theo LayoutMap.
Public Function CheckBodyTextFontName(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckBodyTextFontName")
    Dim paragraphs As Collection: Set paragraphs = context("Snapshot")("Paragraphs")
    If paragraphs.count = 0 Then
        Set CheckBodyTextFontName = Nothing
        Exit Function
    End If

    Dim spec As Object: Set spec = context("Spec")
    Set CheckBodyTextFontName = GroupWrongParagraphsByValue(paragraphs, "Font", CStr(spec("font")("name")), False)
End Function

Public Function CheckBodyTextFontColor(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckBodyTextFontColor")
    Dim paragraphs As Collection: Set paragraphs = context("Snapshot")("Paragraphs")
    If paragraphs.count = 0 Then
        Set CheckBodyTextFontColor = Nothing
        Exit Function
    End If

    Dim spec As Object: Set spec = context("Spec")
    Dim expected As String: expected = UCase$(CStr(spec("font")("color")))
    Set CheckBodyTextFontColor = GroupWrongParagraphsByValue(paragraphs, "Color", expected, True, False)
End Function

Private Function BodyTextParagraphs(ByVal snapshot As Object, ByVal layoutMap As Object) As Collection
    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In snapshot("Paragraphs")
        If layoutMap.Exists(p.Index) Then
            If CStr(layoutMap(p.Index)) = "bodyText" Then Result.Add p
        End If
    Next p
    Set BodyTextParagraphs = Result
End Function

' Doan than bai dung de kiem tra canh deu hai ben: khong mang vai tro rieng nao (theo
' LayoutMap), khong nam trong bang, va co noi dung that (bo doan rong, tranh bao nham tren
' tai lieu trong). Noi dung trong hop van ban tu dong bi loai vi DocumentSnapshot.
' CaptureParagraphs chi doc wdMainTextStory, khong doc wdTextFrameStory.
Private Function BodyTextParagraphsForAlignment(ByVal snapshot As Object, ByVal layoutMap As Object) As Collection
    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In snapshot("Paragraphs")
        If Not layoutMap.Exists(p.Index) And Not p.isInTable And Len(Trim$(p.text)) > 0 Then
            Result.Add p
        End If
    Next p
    Set BodyTextParagraphsForAlignment = Result
End Function

' ND30-PL1-M2-K6E-ALIGN -- Canh deu hai le. Bo qua alignment "unknown" (khong chac thi khong sua).
' Mot Finding TONG cho ca tai lieu, ParagraphIndex tro doan sai dau tien.
Public Function CheckBodyTextAlignment(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckBodyTextAlignment")
    Dim paragraphs As Collection
    Set paragraphs = BodyTextParagraphsForAlignment(context("Snapshot"), context("LayoutMap"))
    If paragraphs.count = 0 Then
        Set CheckBodyTextAlignment = Nothing
        Exit Function
    End If

    Dim wrongCount As Long: wrongCount = 0
    Dim firstWrongIndex As Long: firstWrongIndex = -1
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If p.alignment <> "justify" And p.alignment <> "unknown" Then
            wrongCount = wrongCount + 1
            If firstWrongIndex < 0 Then firstWrongIndex = p.Index
        End If
    Next p

    Dim Result As New Collection
    If wrongCount > 0 Then
        Dim msg As String
        msg = "Ph" & ChrW(&HE1) & "t hi" & ChrW(&H1EC7) & "n " & CStr(wrongCount) & " " & ChrW(&H111) & _
            "o" & ChrW(&H1EA1) & "n ch" & ChrW(&H1B0) & "a canh " & ChrW(&H111) & ChrW(&H1EC1) & "u c" & _
            ChrW(&H1EA3) & " hai l" & ChrW(&H1EC1) & "."
        Result.Add MakeFindingInput(CLng(firstWrongIndex), msg, CLng(wrongCount))
    End If
    Set CheckBodyTextAlignment = Result
End Function

Private Function BuildAllowedIndentLabel(ByVal allowedCm As Collection) As String
    Dim Result As String
    Dim v As Variant
    Dim isFirst As Boolean: isFirst = True
    For Each v In allowedCm
        If Not isFirst Then Result = Result & " ho" & ChrW(&H1EB7) & "c "
        Result = Result & FormatVnNumber(CDbl(v), 2) & " cm"
        isFirst = False
    Next v
    BuildAllowedIndentLabel = Result
End Function

' ND30-PL1-M2-K6E-INDENT -- Thut dau dong 1 cm hoac 1,27 cm. MOT Finding MOI doan sai (khac
' ALIGN/FONT/COLOR o tren) vi thong diep mang gia tri {actual} rieng cua tung doan.
Public Function CheckBodyTextFirstLineIndent(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckBodyTextFirstLineIndent")
    Dim paragraphs As Collection
    Set paragraphs = BodyTextParagraphs(context("Snapshot"), context("LayoutMap"))
    If paragraphs.count = 0 Then
        Set CheckBodyTextFirstLineIndent = Nothing
        Exit Function
    End If

    Dim spec As Object: Set spec = context("Spec")
    Dim allowedCm As Collection: Set allowedCm = spec("bodyText")("firstLineIndentCm")("allowed")
    Dim toleranceCm As Double: toleranceCm = PointToCm(ROUNDING_TOLERANCE_PT, spec)
    Dim allowedLabel As String: allowedLabel = BuildAllowedIndentLabel(allowedCm)

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        Dim actualCm As Double: actualCm = PointToCm(p.FirstLineIndentPt, spec)
        Dim matchesAny As Boolean: matchesAny = False
        Dim v As Variant
        For Each v In allowedCm
            If Abs(actualCm - CDbl(v)) <= toleranceCm Then
                matchesAny = True
                Exit For
            End If
        Next v

        If Not matchesAny Then
            Dim msg As String
            msg = "Th" & ChrW(&H1EE5) & "t " & ChrW(&H111) & ChrW(&H1EA7) & "u d" & ChrW(&HF2) & "ng hi" & _
                ChrW(&H1EC7) & "n t" & ChrW(&H1EA1) & "i " & FormatVnNumber(actualCm, 2) & " cm. Quy " & _
                ChrW(&H111) & ChrW(&H1ECB) & "nh l" & ChrW(&HF9) & "i v" & ChrW(&HE0) & "o " & allowedLabel & "."
            Result.Add MakeFindingInput(CLng(p.Index), msg)
        End If
    Next p

    Set CheckBodyTextFirstLineIndent = Result
End Function

' ND30-PL1-M2-K6E-SPACEAFTER -- Gian doan toi thieu, KHONG co tran tren (ND 30 chi quy dinh muc
' toi thieu -- gian 8pt van hop le).
Public Function CheckBodyTextSpaceAfter(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckBodyTextSpaceAfter")
    Dim paragraphs As Collection
    Set paragraphs = BodyTextParagraphs(context("Snapshot"), context("LayoutMap"))
    If paragraphs.count = 0 Then
        Set CheckBodyTextSpaceAfter = Nothing
        Exit Function
    End If

    Dim spec As Object: Set spec = context("Spec")
    Dim minPt As Double: minPt = CDbl(spec("bodyText")("spaceAfterPt")("min"))

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If p.SpaceAfterPt < minPt - ROUNDING_TOLERANCE_PT Then
            Dim msg As String
            msg = "Kho" & ChrW(&H1EA3) & "ng c" & ChrW(&HE1) & "ch hi" & ChrW(&H1EC7) & "n t" & _
                ChrW(&H1EA1) & "i " & FormatVnNumber(p.SpaceAfterPt, 1) & " pt. Quy " & ChrW(&H111) & _
                ChrW(&H1ECB) & "nh t" & ChrW(&H1ED1) & "i thi" & ChrW(&H1EC3) & "u " & CStr(minPt) & "pt."
            Result.Add MakeFindingInput(CLng(p.Index), msg)
        End If
    Next p

    Set CheckBodyTextSpaceAfter = Result
End Function

' spec("bodyText")("lineSpacing")("min"/"max") la NHAN CHUOI ("single"/"1.5lines") -- boi so duoi
' day la doc thang con so DA CO SAN trong chinh nhan do, khong phai mot tham so moi bia ra ngoai
' JSON.
Private Function LineSpacingLabelToMultiple(ByVal LABEL As String) As Double
    Select Case LABEL
        Case "single"
            LineSpacingLabelToMultiple = 1
        Case "1.5lines"
            LineSpacingLabelToMultiple = 1.5
        Case "double"
            LineSpacingLabelToMultiple = 2
        Case Else
            Err.Raise vbObjectError + 1, "ComplianceChecker.LineSpacingLabelToMultiple", _
                "Nhan gian dong """ & LABEL & """ khong nhan dien duoc."
    End Select
End Function

' ND30-PL1-M2-K6E-LINESPACING -- Gian dong toi thieu dong don, toi da 1,5 lines.
Public Function CheckBodyTextLineSpacing(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckBodyTextLineSpacing")
    Dim paragraphs As Collection
    Set paragraphs = BodyTextParagraphs(context("Snapshot"), context("LayoutMap"))
    If paragraphs.count = 0 Then
        Set CheckBodyTextLineSpacing = Nothing
        Exit Function
    End If

    Dim spec As Object: Set spec = context("Spec")
    Dim minMultiple As Double: minMultiple = LineSpacingLabelToMultiple(CStr(spec("bodyText")("lineSpacing")("min")))
    Dim maxMultiple As Double: maxMultiple = LineSpacingLabelToMultiple(CStr(spec("bodyText")("lineSpacing")("max")))
    Const toleranceMultiple As Double = 0.05

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If p.lineSpacing <> 0 Then ' 0 = hon hop/khong doc duoc, bo qua
            Dim multiple As Double: multiple = p.lineSpacing / WORD_MULTIPLE_LINE_SPACING_UNIT_PT
            If multiple < minMultiple - toleranceMultiple Or multiple > maxMultiple + toleranceMultiple Then
                Dim msg As String
                msg = "Gi" & ChrW(&HE3) & "n d" & ChrW(&HF2) & "ng hi" & ChrW(&H1EC7) & "n t" & _
                    ChrW(&H1EA1) & "i " & FormatVnNumber(multiple, 2) & " lines. Quy " & ChrW(&H111) & _
                    ChrW(&H1ECB) & "nh t" & ChrW(&H1ED1) & "i thi" & ChrW(&H1EC3) & "u d" & ChrW(&HF2) & _
                    "ng " & ChrW(&H111) & ChrW(&H1A1) & "n, t" & ChrW(&H1ED1) & "i " & ChrW(&H111) & "a " & _
                    FormatVnNumber(maxMultiple, 2) & " lines."
                Result.Add MakeFindingInput(CLng(p.Index), msg)
            End If
        End If
    Next p

    Set CheckBodyTextLineSpacing = Result
End Function

' ND30-PL1-M2-K6E-DOTSLASH -- doan CUOI CUNG cua noi dung van ban phai ket thuc dung dau ket
' theo che do (RuleLoader.GetContentEndMark).
' Xac dinh "cuoi noi dung": lui tu chi so NHO NHAT trong hai vai tro recipientLabel/
' signerAuthority (vai tro nao den truoc trong tai lieu trong LayoutMap) - dung duoc ca bo cuc
' bang (hai vai tro cung mot bang, vi tri khong doi) lan bo cuc tuan tu (chu ky truoc "Noi
' nhan:", moc lui dung o dau khoi quyen han-chuc vu). Lui ve tu ngay truoc moc do, bo qua doan
' trong bang (lui het bang neu gap) va doan rong, dung lai o doan CO NOI DUNG dau tien gap duoc.
' Loai C (chi bao, khong tu chen): ca hai vai tro neo deu khong nhan dien duoc thi khong bao gi
' ca - khong chac thi khong sua.
Public Function CheckContentEndsWithDotSlash(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckContentEndsWithDotSlash")
    Dim endMark As String: endMark = RuleLoader.GetContentEndMark(RegimeOf(context))

    Dim snapshot As Object: Set snapshot = context("Snapshot")
    Dim layoutMap As Object: Set layoutMap = context("LayoutMap")

    ' "Cuoi noi dung" phai lui truoc CA khoi "Noi nhan" LAN khoi quyen han-chuc vu nguoi ky
    ' (signerAuthority) - lay chi so NHO NHAT (den truoc) trong hai vai tro, khong chi rieng
    ' recipientLabel: bo cuc bang (NÄ�224, "Noi nhan"/chu ky cung mot hang) va bo cuc tuan tu (chu
    ' ky roi moi toi "Noi nhan:" phia duoi) deu gap trong van ban that/fixture - dung mot trong
    ' hai vai tro lam moc se sai o bo cuc con lai (da xac nhan qua ca hai).
    Dim recipients As Collection: Set recipients = ParagraphsWithRole(snapshot, layoutMap, "recipientLabel")
    Dim signers As Collection: Set signers = ParagraphsWithRole(snapshot, layoutMap, "signerAuthority")
    If recipients.count = 0 And signers.count = 0 Then
        Set CheckContentEndsWithDotSlash = Nothing
        Exit Function
    End If

    Dim anchorIndex As Long: anchorIndex = -1
    If recipients.count > 0 Then anchorIndex = CLng(recipients(1).Index)
    If signers.count > 0 Then
        If anchorIndex = -1 Or CLng(signers(1).Index) < anchorIndex Then anchorIndex = CLng(signers(1).Index)
    End If

    Dim paragraphByIndex As Object: Set paragraphByIndex = Utils.NewDictionary()
    Dim p As ParagraphSnapshot
    For Each p In snapshot("Paragraphs")
        Set paragraphByIndex(p.Index) = p
    Next p

    Dim scanIndex As Long: scanIndex = anchorIndex - 1
    Dim lastBody As ParagraphSnapshot: Set lastBody = Nothing
    Do While scanIndex >= 0
        If Not paragraphByIndex.Exists(scanIndex) Then Exit Do
        Dim cur As ParagraphSnapshot: Set cur = paragraphByIndex(scanIndex)
        If cur.isInTable Or Len(Trim$(cur.text)) = 0 Then
            scanIndex = scanIndex - 1
        Else
            Set lastBody = cur
            Exit Do
        End If
    Loop
    If lastBody Is Nothing Then
        Set CheckContentEndsWithDotSlash = Nothing
        Exit Function
    End If

    Dim trimmedText As String: trimmedText = Trim$(lastBody.text)
    Dim Result As New Collection
    Dim msg As String

    If endMark = "." Then
        ' Che do Dang: phai ket thuc bang dau cham "." VA KHONG duoc dung "./.". Ca hai thong bao
        ' deu TRANH dat dau./. trong ngoac kep (khong can escape dau ngoac kep trong VBA).
        If Right$(trimmedText, 3) = "./." Then
            msg = "Ph" & ChrW(&H1EA7) & "n cu" & ChrW(&H1ED1) & "i n" & ChrW(&H1ED9) & "i dung v" & _
                ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n " & ChrW(&H111) & "ang d" & ChrW(&HF9) & _
                "ng d" & ChrW(&H1EA5) & "u ./., ch" & ChrW(&H1EBF) & " " & ChrW(&H111) & ChrW(&H1ED9) & _
                " " & ChrW(&H110) & ChrW(&H1EA3) & "ng ch" & ChrW(&H1EC9) & " k" & ChrW(&H1EBF) & _
                "t th" & ChrW(&HFA) & "c b" & ChrW(&H1EB1) & "ng d" & ChrW(&H1EA5) & "u ch" & _
                ChrW(&H1EA5) & "m, kh" & ChrW(&HF4) & "ng d" & ChrW(&HF9) & "ng ./."
            Result.Add MakeFindingInput(CLng(lastBody.Index), msg)
        ElseIf Right$(trimmedText, 1) <> "." Then
            msg = "Ph" & ChrW(&H1EA7) & "n cu" & ChrW(&H1ED1) & "i n" & ChrW(&H1ED9) & "i dung v" & _
                ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n ch" & ChrW(&H1B0) & "a k" & ChrW(&H1EBF) & _
                "t th" & ChrW(&HFA) & "c b" & ChrW(&H1EB1) & "ng d" & ChrW(&H1EA5) & "u ch" & _
                ChrW(&H1EA5) & "m."
            Result.Add MakeFindingInput(CLng(lastBody.Index), msg)
        End If
    Else
        If Right$(trimmedText, Len(endMark)) <> endMark Then
            msg = "Ph" & ChrW(&H1EA7) & "n cu" & ChrW(&H1ED1) & "i n" & ChrW(&H1ED9) & "i dung v" & _
                ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n ch" & ChrW(&H1B0) & "a k" & ChrW(&H1EBF) & _
                "t th" & ChrW(&HFA) & "c b" & ChrW(&H1EB1) & "ng """ & endMark & """."
            Result.Add MakeFindingInput(CLng(lastBody.Index), msg)
        End If
    End If

    Set CheckContentEndsWithDotSlash = Result
End Function

' ND30-PL1-MV-CT1 -- Co chu giua cac thanh phan the thuc phai THONG NHAT theo mot trong cac bo
' cho phep cua che do dang chon - chu thich 1 Muc V, FR-CHK-10.
' Chi so sanh cac doan DA duoc LayoutMap gan vai tro rieng (do tin cay cao/trung binh) voi co
' chu vai tro do - khong dam toi doan "khong co vai tro" (co the la than bai that, co the la
' thanh phan bi bo sot, khong phan biet duoc chi tu FontSize).
' Voi moi bo co chu KHA DUNG cho chinh che do dang chon (RuleLoader.GetEffectiveFontSizeSet, co
' ghi de rieng theo che do), dem so doan-co-vai-tro lech co chu; chon bo co so lech thap hon lam
' "bo dang dung"; bao loi bang dung so lech cua bo da chon.
Public Function CheckFontSizeConsistency(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckFontSizeConsistency")
    Dim snapshot As Object: Set snapshot = context("Snapshot")
    Dim layoutMap As Object: Set layoutMap = context("LayoutMap")

    Dim regime As String: regime = "ND30"
    If context.Exists("Regime") Then regime = CStr(context("Regime"))

    ' fontSizeSetKeys la mang JSON -> sinh thanh Collection (khong phai VBA array) qua
    ' RuleData.bas - dung For Each/Collection.Item, KHONG dung LBound/UBound.
    Dim setKeys As Collection: Set setKeys = RuleLoader.GetRegimeConfig()("regimes")(regime)("fontSizeSetKeys")

    Dim paragraphByIndex As Object: Set paragraphByIndex = Utils.NewDictionary()
    Dim p As ParagraphSnapshot
    For Each p In snapshot("Paragraphs")
        Set paragraphByIndex(p.Index) = p
    Next p

    ' Nap ban co chu HIEU LUC (co ghi de theo che do) cho MOI khoa kha dung cua che do nay.
    Dim effectiveSets As Object: Set effectiveSets = Utils.NewDictionary()
    Dim mismatchBySetKey As Object: Set mismatchBySetKey = Utils.NewDictionary()
    Dim skVar As Variant
    For Each skVar In setKeys
        Dim sk As String: sk = CStr(skVar)
        Set effectiveSets(sk) = RuleLoader.GetEffectiveFontSizeSet(regime, sk)
        mismatchBySetKey(sk) = 0
    Next skVar

    Dim checkedCount As Long: checkedCount = 0
    Dim key As Variant
    For Each key In layoutMap.Keys
        Dim role As String: role = CStr(layoutMap(key))
        If role <> "unknown" Then
            Dim roleAvailableInAll As Boolean: roleAvailableInAll = True
            For Each skVar In setKeys
                If Not effectiveSets(CStr(skVar)).Exists(role) Then roleAvailableInAll = False
            Next skVar
            If roleAvailableInAll Then
                Dim paragraph As ParagraphSnapshot
                Set paragraph = paragraphByIndex(CLng(key))
                If paragraph.FontSizePt <> 0 Then
                    checkedCount = checkedCount + 1
                    For Each skVar In setKeys
                        Dim skKey As String: skKey = CStr(skVar)
                        If paragraph.FontSizePt <> CDbl(effectiveSets(skKey)(role)) Then
                            mismatchBySetKey(skKey) = CLng(mismatchBySetKey(skKey)) + 1
                        End If
                    Next skVar
                End If
            End If
        End If
    Next key

    ' Khong co doan nao du du lieu de so sanh (LayoutMap rong, hoac tai lieu chua nhan dien duoc
    ' thanh phan nao) - "khong chac thi khong sua" (CLAUDE.md muc 5), tra Nothing.
    If checkedCount = 0 Then
        Set CheckFontSizeConsistency = Nothing
        Exit Function
    End If

    Dim chosenSetKey As String: chosenSetKey = ""
    Dim mismatchCount As Long: mismatchCount = 0
    For Each skVar In setKeys
        Dim candKey As String: candKey = CStr(skVar)
        If Len(chosenSetKey) = 0 Or CLng(mismatchBySetKey(candKey)) < mismatchCount Then
            chosenSetKey = candKey
            mismatchCount = CLng(mismatchBySetKey(candKey))
        End If
    Next skVar
    Dim chosenSet As Object: Set chosenSet = effectiveSets(chosenSetKey)

    Dim Result As New Collection
    If mismatchCount > 0 Then
        Dim msg As String
        msg = "C" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & " ch" & ChrW(&H1B0) & "a th" & ChrW(&H1ED1) & _
            "ng nh" & ChrW(&H1EA5) & "t theo " & CStr(chosenSet("label")) & " (" & CStr(mismatchCount) & _
            " " & ChrW(&H111) & "o" & ChrW(&H1EA1) & "n)."
        Result.Add MakeFindingInput(Null, msg)
    End If

    Set CheckFontSizeConsistency = Result
End Function

' ----------------------------------------------------------------------------
' ----------------------------------------------------------------------------

Private Sub RegisterPageSetupBodyTextChecks()
    mRegistry("ND30-PL1-M1-K1") = "ComplianceChecker.CheckPageSizeA4"
    mRegistry("ND30-PL1-M1-K2") = "ComplianceChecker.CheckPageOrientation"
    mRegistry("ND30-PL1-M1-K3") = "ComplianceChecker.CheckPageMargins"
    mRegistry("ND30-PL1-M1-K7") = "ComplianceChecker.CheckPageNumbering"
    mRegistry("ND30-PL1-M3-K1D") = "ComplianceChecker.CheckAppendixPageNumberingRestart"

    mRegistry("ND30-PL1-M1-K4-FONT") = "ComplianceChecker.CheckBodyTextFontName"
    mRegistry("ND30-PL1-M1-K4-COLOR") = "ComplianceChecker.CheckBodyTextFontColor"
    mRegistry("ND30-PL1-MV-CT1") = "ComplianceChecker.CheckFontSizeConsistency"
    mRegistry("ND30-PL1-M2-K6E-ALIGN") = "ComplianceChecker.CheckBodyTextAlignment"
    mRegistry("ND30-PL1-M2-K6E-INDENT") = "ComplianceChecker.CheckBodyTextFirstLineIndent"
    mRegistry("ND30-PL1-M2-K6E-SPACEAFTER") = "ComplianceChecker.CheckBodyTextSpaceAfter"
    mRegistry("ND30-PL1-M2-K6E-LINESPACING") = "ComplianceChecker.CheckBodyTextLineSpacing"
    mRegistry("ND30-PL1-M2-K6E-DOTSLASH") = "ComplianceChecker.CheckContentEndsWithDotSlash"
End Sub

' ============================================================================
' PHAN 3 -- 43 ham kiem nhom "component". Cung cong thuc, cung regex (xem
' dau-hieu-nhan-dien.json $regexPortability).
' Phan lon muc "partial" dua vao LayoutMap do ComponentDetector.DetectComponents gan: khong co
' doan nao mang vai tro can thiet -> Nothing ("chua nhan dien duoc thanh phan X"), khong phai
' Collection rong. Mot so vai tro (appendixLabel, appendixTitle, appendixReference,
' urgencyStamp, circulationScope, drafterNotation, organContact) khong duoc ComponentDetector
' gan - khong co dau hieu o docs/rules/01 muc 5; bon quy tac phu luc (K1A-REF, K1A-NUM, K1B,
' K1C) tu do bang regex cuc bo thay vi LayoutMap. Bon quy tac con lai (K2B, K2C, K3, K4 Muc
' III) va M4-POS hoan toan khong co cach do -> Nothing vinh vien.
' ============================================================================

' ----------------------------------------------------------------------------
' Tien ich dung chung
' ----------------------------------------------------------------------------

Private Function AlignmentLabelVn(ByVal align As String) As String
    Select Case align
        Case "left"
            AlignmentLabelVn = "tr" & ChrW(&HE1) & "i"
        Case "center"
            AlignmentLabelVn = "gi" & ChrW(&H1EEF) & "a"
        Case "right"
            AlignmentLabelVn = "ph" & ChrW(&H1EA3) & "i"
        Case "justify"
            AlignmentLabelVn = ChrW(&H111) & ChrW(&H1EC1) & "u hai l" & ChrW(&H1EC1)
        Case Else
            AlignmentLabelVn = "kh" & ChrW(&HF4) & "ng x" & ChrW(&HE1) & "c " & ChrW(&H111) & ChrW(&H1ECB) & "nh"
    End Select
End Function

Private Function JoinCollection(ByVal coll As Collection, ByVal sep As String) As String
    Dim Result As String
    Dim v As Variant
    Dim isFirst As Boolean: isFirst = True
    For Each v In coll
        If Not isFirst Then Result = Result & sep
        Result = Result & CStr(v)
        isFirst = False
    Next v
    JoinCollection = Result
End Function

' Moi doan mang "role" trong layoutMap, DUNG THU TU tai lieu -- Collection rong KHAC Nothing: ham
' goi tu quyet dinh Nothing (chua nhan dien duoc) truoc khi dung danh sach nay.
Private Function ParagraphsWithRole(ByVal snapshot As Object, ByVal layoutMap As Object, _
        ByVal role As String) As Collection
    Dim indices As New Collection
    Dim key As Variant
    For Each key In layoutMap.Keys
        If CStr(layoutMap(key)) = role Then indices.Add CLng(key)
    Next key

    Dim Result As New Collection
    Dim n As Long: n = indices.count
    If n = 0 Then
        Set ParagraphsWithRole = Result
        Exit Function
    End If

    Dim arr() As Long
    ReDim arr(0 To n - 1)
    Dim i As Long: i = 0
    Dim v As Variant
    For Each v In indices
        arr(i) = CLng(v)
        i = i + 1
    Next v
    Dim j As Long, keyVal As Long
    For i = 1 To n - 1
        keyVal = arr(i)
        j = i - 1
        Do While j >= 0
            If arr(j) > keyVal Then
                arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        arr(j + 1) = keyVal
    Next i

    Dim paragraphByIndex As Object: Set paragraphByIndex = Utils.NewDictionary()
    Dim p As ParagraphSnapshot
    For Each p In snapshot("Paragraphs")
        Set paragraphByIndex(p.Index) = p
    Next p

    For i = 0 To n - 1
        If paragraphByIndex.Exists(arr(i)) Then Result.Add paragraphByIndex(arr(i))
    Next i
    Set ParagraphsWithRole = Result
End Function

' Liet ke vi pham letterCase/kieu chu/canh le cua MOT doan so voi ComponentSpec -- dung chung cho
' moi quy tac "chua dung the thuc" muc partial. Bo qua Bold/Italic khi Null (hon hop/khong xac
' dinh) -- "khong chac thi khong sua".
Private Function DescribeStyleViolations(ByVal p As ParagraphSnapshot, ByVal componentSpec As Object) As Collection
    Dim problems As New Collection
    Dim styleSpec As Object: Set styleSpec = componentSpec("style")

    If componentSpec.Exists("letterCase") Then
        Dim letterCase As String: letterCase = CStr(componentSpec("letterCase"))
        If letterCase = "upper" And Not p.AllCaps Then
            problems.Add "ch" & ChrW(&H1B0) & "a in hoa"
        End If
        If letterCase = "normal" And p.AllCaps Then
            problems.Add ChrW(&H111) & "ang in hoa to" & ChrW(&HE0) & "n b" & ChrW(&H1ED9) & _
                ", quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh ch" & ChrW(&H1EEF) & " th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
        End If
    End If

    If CBool(styleSpec("bold")) Then
        If Not IsNull(p.bold) Then
            If Not CBool(p.bold) Then problems.Add "ch" & ChrW(&H1B0) & "a in " & ChrW(&H111) & ChrW(&H1EAD) & "m"
        End If
    Else
        If Not IsNull(p.bold) Then
            If CBool(p.bold) Then
                problems.Add ChrW(&H111) & "ang in " & ChrW(&H111) & ChrW(&H1EAD) & "m, quy " & _
                    ChrW(&H111) & ChrW(&H1ECB) & "nh kh" & ChrW(&HF4) & "ng " & ChrW(&H111) & ChrW(&H1EAD) & "m"
            End If
        End If
    End If

    If CBool(styleSpec("italic")) Then
        If Not IsNull(p.Italic) Then
            If Not CBool(p.Italic) Then problems.Add "ch" & ChrW(&H1B0) & "a in nghi" & ChrW(&HEA) & "ng"
        End If
    Else
        If Not IsNull(p.Italic) Then
            If CBool(p.Italic) Then
                problems.Add ChrW(&H111) & "ang in nghi" & ChrW(&HEA) & "ng, quy " & ChrW(&H111) & _
                    ChrW(&H1ECB) & "nh kh" & ChrW(&HF4) & "ng nghi" & ChrW(&HEA) & "ng"
            End If
        End If
    End If

    If componentSpec.Exists("alignment") Then
        Dim expectedAlign As String: expectedAlign = CStr(componentSpec("alignment"))
        ' "canh trai" va "canh deu hai ben" chap nhan lan nhau â€” ND 30 chi noi "sat le trai" cho
        ' khoi Noi nhan, khong phan biet hai gia tri nay. Cac quy dinh alignment KHAC (center,
        ' right) van chi khop dung mot gia tri.
        Dim alignOk As Boolean
        alignOk = (p.alignment = expectedAlign)
        If Not alignOk And expectedAlign = "left" And p.alignment = "justify" Then alignOk = True
        If Not alignOk And expectedAlign = "justify" And p.alignment = "left" Then alignOk = True
        If p.alignment <> "unknown" And Not alignOk Then
            problems.Add "canh l" & ChrW(&H1EC1) & " " & AlignmentLabelVn(p.alignment) & ", quy " & _
                ChrW(&H111) & ChrW(&H1ECB) & "nh canh " & AlignmentLabelVn(expectedAlign)
        End If
    End If

    Set DescribeStyleViolations = problems
End Function

' Che do quy dinh dang chon cho luot kiem tra nay.
Private Function RegimeOf(ByVal context As Object) As String
    RegimeOf = "ND30"
    If context.Exists("Regime") Then RegimeOf = CStr(context("Regime"))
End Function

' Ten loai van ban (noi dung doan mang vai tro "typeName") â€” can cho ghi de theo nhom loai van ban
' (Viettel: can cu cua nghi quyet/quyet dinh/quy che/quy dinh in nghieng).
Private Function DocumentTypeNameOf(ByVal context As Object) As String
    Dim paragraphs As Collection
    Set paragraphs = ParagraphsWithRole(context("Snapshot"), context("LayoutMap"), "typeName")
    If paragraphs.count > 0 Then DocumentTypeNameOf = Trim$(paragraphs(1).text)
End Function

' Dac ta HIEU LUC cua mot component theo che do dang chon â€” moi ham kiem dinh dang component PHAI
' dung ham nay thay vi doc thang context("Spec")("components")(role).
Private Function EffectiveSpecOf(ByVal context As Object, ByVal role As String) As Object
    Set EffectiveSpecOf = RuleLoader.GetEffectiveComponentSpec(RegimeOf(context), role, DocumentTypeNameOf(context))
End Function

' Dung chung cho moi quy tac "chua dung the thuc" THUAN KIEU CHU (khong co dieu kien van ban rieng
' -- xem cac ham tuy bien ben duoi cho K9B-LABEL, K9B-LIST...).
Private Function StyleCheckForRole(ByVal context As Object, ByVal role As String, _
        ByVal labelText As String) As Object
    Dim snapshot As Object: Set snapshot = context("Snapshot")
    Dim layoutMap As Object: Set layoutMap = context("LayoutMap")
    Dim paragraphs As Collection: Set paragraphs = ParagraphsWithRole(snapshot, layoutMap, role)
    If paragraphs.count = 0 Then
        Set StyleCheckForRole = Nothing
        Exit Function
    End If
    Dim componentSpec As Object: Set componentSpec = EffectiveSpecOf(context, role)

    Dim findings As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        Dim problems As Collection: Set problems = DescribeStyleViolations(p, componentSpec)
        If problems.count > 0 Then
            findings.Add MakeFindingInput(CLng(p.Index), labelText & " ch" & ChrW(&H1B0) & "a " & _
                ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c: " & _
                JoinCollection(problems, "; ") & ".")
        End If
    Next p
    Set StyleCheckForRole = findings
End Function

' VBScript.RegExp â€” dong bo voi ComponentDetector.RegexTest (moi module tu boc rieng).
Private Function RegexTestCC(ByVal pattern As String, ByVal s As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern
    regex.IgnoreCase = False
    regex.Global = False
    RegexTestCC = regex.test(s)
End Function

' Tra Match dau tien (co SubMatches cho cac nhom bat) hoac Nothing neu khong khop.
Private Function RegexFirstMatch(ByVal pattern As String, ByVal s As String) As Object
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern
    regex.IgnoreCase = False
    regex.Global = False
    Dim matches As Object
    Set matches = regex.Execute(s)
    If matches.count = 0 Then
        Set RegexFirstMatch = Nothing
    Else
        Set RegexFirstMatch = matches(0)
    End If
End Function

Private Function PadLeftZero(ByVal s As String, ByVal totalLen As Long) As String
    Dim Result As String: Result = s
    Do While Len(Result) < totalLen
        Result = "0" & Result
    Loop
    PadLeftZero = Result
End Function

Private Function StripTypeQualifier(ByVal name As String) As String
    Dim m As Object
    Set m = RegexFirstMatch("^(.*?)\s*\([^)]*\)\s*$", name)
    If m Is Nothing Then
        StripTypeQualifier = Trim$(name)
    Else
        StripTypeQualifier = Trim$(CStr(m.SubMatches(0)))
    End If
End Function

' ----------------------------------------------------------------------------
' O 1 -- Quoc hieu, Tieu ngu (checklistGroup 2)
' ----------------------------------------------------------------------------

Public Function CheckNationalTitleStyle(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckNationalTitleStyle")
    Set CheckNationalTitleStyle = StyleCheckForRole(context, "nationalTitle", _
        "Qu" & ChrW(&H1ED1) & "c hi" & ChrW(&H1EC7) & "u")
End Function

Public Function CheckNationalMottoStyle(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckNationalMottoStyle")
    Set CheckNationalMottoStyle = StyleCheckForRole(context, "nationalMotto", _
        "Ti" & ChrW(&HEA) & "u ng" & ChrW(&H1EEF))
End Function

' ND30-PL1-M2-K1-TN-SEP -- full: mot khi da co doan mang vai tro nationalMotto, so khop NGUYEN VAN
' voi dang chuan "Doc lap - Tu do - Hanh phuc".
Public Function CheckNationalMottoSeparator(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckNationalMottoSeparator")
    Dim paragraphs As Collection
    Set paragraphs = ParagraphsWithRole(context("Snapshot"), context("LayoutMap"), "nationalMotto")
    If paragraphs.count = 0 Then
        Set CheckNationalMottoSeparator = Nothing
        Exit Function
    End If

    Dim expected As String
    expected = ChrW(&H110) & ChrW(&H1ED9) & "c l" & ChrW(&H1EAD) & "p - T" & ChrW(&H1EF1) & " do - H" & _
        ChrW(&H1EA1) & "nh ph" & ChrW(&HFA) & "c"
    Dim msg As String
    msg = "Ti" & ChrW(&HEA) & "u ng" & ChrW(&H1EEF) & " ph" & ChrW(&H1EA3) & "i l" & ChrW(&HE0) & " '" & _
        expected & "' v" & ChrW(&H1EDB) & "i g" & ChrW(&H1EA1) & "ch n" & ChrW(&H1ED1) & "i v" & ChrW(&HE0) & _
        " c" & ChrW(&HF3) & " c" & ChrW(&HE1) & "ch ch" & ChrW(&H1EEF) & " hai b" & ChrW(&HEA) & "n."

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If Trim$(p.text) <> expected Then Result.Add MakeFindingInput(CLng(p.Index), msg)
    Next p
    Set CheckNationalMottoSeparator = Result
End Function

' ND30-PL1-M2-K1-TN-LINE / K2-LINE / K5A-LINE -- duong ke duoi Tieu ngu/ten co quan/trich yeu.
' CHUA HIEN THUC HOA DUOC: DocumentSnapshot khong chup border/duong ke doan van. Tra Nothing mai
' mai, cung ly do voi CheckPageNumbering (Phan 2).
Public Function CheckComponentUnderline(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckComponentUnderline")
    Set CheckComponentUnderline = Nothing
End Function

' ND30-PL1-M2-K1-C -- Quoc hieu/Tieu ngu cach nhau dong don: xap xi bang SpaceBeforePt cua doan
' Tieu ngu xap xi 0 (dung field lineSpacingWithPrevious="single" cua nationalMotto).
Public Function CheckNationalMottoSpacing(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckNationalMottoSpacing")
    Dim paragraphs As Collection
    Set paragraphs = ParagraphsWithRole(context("Snapshot"), context("LayoutMap"), "nationalMotto")
    If paragraphs.count = 0 Then
        Set CheckNationalMottoSpacing = Nothing
        Exit Function
    End If

    Dim msg As String
    msg = "Hai d" & ChrW(&HF2) & "ng Qu" & ChrW(&H1ED1) & "c hi" & ChrW(&H1EC7) & "u v" & ChrW(&HE0) & _
        " Ti" & ChrW(&HEA) & "u ng" & ChrW(&H1EEF) & " ph" & ChrW(&H1EA3) & "i tr" & ChrW(&HEC) & "nh b" & _
        ChrW(&HE0) & "y c" & ChrW(&HE1) & "ch nhau d" & ChrW(&HF2) & "ng " & ChrW(&H111) & ChrW(&H1A1) & "n."

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If p.SpaceBeforePt > ROUNDING_TOLERANCE_PT Then Result.Add MakeFindingInput(CLng(p.Index), msg)
    Next p
    Set CheckNationalMottoSpacing = Result
End Function

' ----------------------------------------------------------------------------
' O 2 -- Ten co quan, to chuc ban hanh (checklistGroup 3)
' ----------------------------------------------------------------------------

Public Function CheckSuperiorOrganNameStyle(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckSuperiorOrganNameStyle")
    Set CheckSuperiorOrganNameStyle = StyleCheckForRole(context, "superiorOrganName", _
        "T" & ChrW(&HEA) & "n c" & ChrW(&H1A1) & " quan ch" & ChrW(&H1EE7) & " qu" & ChrW(&H1EA3) & "n")
End Function

Public Function CheckOrganNameStyle(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckOrganNameStyle")
    Set CheckOrganNameStyle = StyleCheckForRole(context, "organName", _
        "T" & ChrW(&HEA) & "n c" & ChrW(&H1A1) & " quan ban h" & ChrW(&HE0) & "nh")
End Function

' ----------------------------------------------------------------------------
' O 3 -- So, ky hieu van ban (checklistGroup 4)
' ----------------------------------------------------------------------------

Private Function CodeNumberParagraphs(ByVal context As Object) As Collection
    Set CodeNumberParagraphs = ParagraphsWithRole(context("Snapshot"), context("LayoutMap"), "codeNumberNotation")
End Function

' ND30-PL1-M2-K3-PREFIX -- dau hai cham sau "So". Regex nhan dien (dau-hieu-nhan-dien.json) cho
' phep dau hai cham TUY CHON, nen truong hop thieu van duoc gan vai tro va kiem duoc o day.
Public Function CheckCodeNumberColon(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckCodeNumberColon")
    Dim paragraphs As Collection: Set paragraphs = CodeNumberParagraphs(context)
    If paragraphs.count = 0 Then
        Set CheckCodeNumberColon = Nothing
        Exit Function
    End If

    Dim msg As String
    msg = "Sau t" & ChrW(&H1EEB) & " 'S" & ChrW(&H1ED1) & "' ph" & ChrW(&H1EA3) & "i c" & ChrW(&HF3) & _
        " d" & ChrW(&H1EA5) & "u hai ch" & ChrW(&H1EA5) & "m."

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If Not RegexTestCC("^\s*S" & ChrW(&H1ED1) & "\s*:", p.text) Then
            Result.Add MakeFindingInput(CLng(p.Index), msg)
        End If
    Next p
    Set CheckCodeNumberColon = Result
End Function

' ND30-PL1-M2-K3-PAD -- so nho hon 10 phai them so 0.
Public Function CheckCodeNumberPad(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckCodeNumberPad")
    Dim paragraphs As Collection: Set paragraphs = CodeNumberParagraphs(context)
    If paragraphs.count = 0 Then
        Set CheckCodeNumberPad = Nothing
        Exit Function
    End If
    Dim spec As Object: Set spec = context("Spec")
    Dim padBelow As Long: padBelow = 10
    If spec("components")("codeNumberNotation").Exists("padNumberBelow") Then
        padBelow = CLng(spec("components")("codeNumberNotation")("padNumberBelow"))
    End If

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    Dim pattern As String
    pattern = "S" & ChrW(&H1ED1) & "\s*:?\s*(\d+)\s*\/"
    For Each p In paragraphs
        Dim m As Object: Set m = RegexFirstMatch(pattern, p.text)
        If Not m Is Nothing Then
            Dim raw As String: raw = CStr(m.SubMatches(0))
            If CLng(raw) < padBelow And Len(raw) < 2 Then
                Dim expected As String: expected = PadLeftZero(raw, 2)
                Dim msg As String
                msg = "'" & raw & "' ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & _
                    " '" & expected & "'. V" & ChrW(&H1EDB) & "i nh" & ChrW(&H1EEF) & "ng s" & ChrW(&H1ED1) & _
                    " nh" & ChrW(&H1ECF) & " h" & ChrW(&H1A1) & "n 10 ph" & ChrW(&H1EA3) & "i ghi th" & _
                    ChrW(&HEA) & "m s" & ChrW(&H1ED1) & " 0 ph" & ChrW(&HED) & "a tr" & ChrW(&H1B0) & "" & "c."
                Result.Add MakeFindingInput(CLng(p.Index), msg)
            End If
        End If
    Next p
    Set CheckCodeNumberPad = Result
End Function

' Van ban chua duoc ky duyet thi chua co so, chi de cho trong cho sau nay dien tay/danh may - "So:
' Nhan dien: phan TRUOC dau "/" (neu co) khong chua chu so nao - coi la CHU DINH de trong, khong
' phai loi thieu khoang cach.
Private Function IsBlankCodeNumberPlaceholder(ByVal afterColonText As String) As Boolean
    Dim slashPos As Long: slashPos = InStr(afterColonText, "/")
    Dim numberPart As String
    If slashPos > 0 Then
        numberPart = left$(afterColonText, slashPos - 1)
    Else
        numberPart = afterColonText
    End If
    IsBlankCodeNumberPlaceholder = Not RegexTestCC("\d", numberPart)
End Function

' ND30-PL1-M2-K3-SPACE -- khong co dau cach trong toan bo phan so va ky hieu (KHONG tinh dau cach
' bat buoc sau "So:"). Bo qua doan de trong cho so (xem IsBlankCodeNumberPlaceholder).
Public Function CheckCodeNumberNoSpace(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckCodeNumberNoSpace")
    Dim paragraphs As Collection: Set paragraphs = CodeNumberParagraphs(context)
    If paragraphs.count = 0 Then
        Set CheckCodeNumberNoSpace = Nothing
        Exit Function
    End If

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        Dim afterColon As String
        afterColon = Trim$(RegexReplaceFirstPrefix(p.text))
        If IsBlankCodeNumberPlaceholder(afterColon) Then GoTo ContinueLoop
        If InStr(afterColon, " ") > 0 Or InStr(afterColon, vbTab) > 0 Then
            Dim collapsed As String: collapsed = Replace$(Replace$(afterColon, " ", ""), vbTab, "")
            Dim msg As String
            msg = "'" & afterColon & "' ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & _
                " '" & collapsed & "'. Gi" & ChrW(&H1EEF) & "a c" & ChrW(&HE1) & "c nh" & ChrW(&HF3) & _
                "m ch" & ChrW(&H1EEF) & " vi" & ChrW(&H1EBF) & "t t" & ChrW(&H1EAF) & "t kh" & ChrW(&HF4) & _
                "ng c" & ChrW(&HE1) & "ch ch" & ChrW(&H1EEF) & "."
            Result.Add MakeFindingInput(CLng(p.Index), msg)
        End If
ContinueLoop:
    Next p
    Set CheckCodeNumberNoSpace = Result
End Function

' Phan "So:" o dau doan (regex "^\s*So\s*:?\s*"), tra phan CON LAI (chua trim).
Private Function RegexReplaceFirstPrefix(ByVal text As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = "^\s*S" & ChrW(&H1ED1) & "\s*:?\s*"
    regex.IgnoreCase = False
    regex.Global = False
    RegexReplaceFirstPrefix = regex.Replace(text, "")
End Function

' ND30-PL1-M2-K3-SEP -- dau gach cheo giua so/ky hieu, dau gach noi giua cac nhom viet tat.
Public Function CheckCodeNumberSeparators(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckCodeNumberSeparators")
    Dim paragraphs As Collection: Set paragraphs = CodeNumberParagraphs(context)
    If paragraphs.count = 0 Then
        Set CheckCodeNumberSeparators = Nothing
        Exit Function
    End If

    Dim msg As String
    msg = "Gi" & ChrW(&H1EEF) & "a s" & ChrW(&H1ED1) & " v" & ChrW(&HE0) & " k" & ChrW(&HFD) & " hi" & _
        ChrW(&H1EC7) & "u ph" & ChrW(&H1EA3) & "i l" & ChrW(&HE0) & " d" & ChrW(&H1EA5) & "u g" & ChrW(&H1EA1) & _
        "ch ch" & ChrW(&HE9) & "o, gi" & ChrW(&H1EEF) & "a c" & ChrW(&HE1) & "c nh" & ChrW(&HF3) & "m ch" & _
        ChrW(&H1EEF) & " vi" & ChrW(&H1EBF) & "t t" & ChrW(&H1EAF) & "t ph" & ChrW(&H1EA3) & "i l" & ChrW(&HE0) & _
        " d" & ChrW(&H1EA5) & "u g" & ChrW(&H1EA1) & "ch n" & ChrW(&H1ED1) & "i."

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        Dim body As String: body = Trim$(RegexReplaceFirstPrefix(p.text))
        Dim m As Object: Set m = RegexFirstMatch("^(\d+)(.)(.+)$", body)
        If Not m Is Nothing Then
            Dim sep As String: sep = CStr(m.SubMatches(1))
            Dim notation As String: notation = CStr(m.SubMatches(2))
            If sep <> "/" Or Not RegexTestCC("^[A-Za-z" & ChrW(&HC0) & "-" & ChrW(&H1EF9) & ChrW(&H110) & _
                    ChrW(&H111) & "0-9-]+$", notation) Then
                Result.Add MakeFindingInput(CLng(p.Index), msg)
            End If
        End If
    Next p
    Set CheckCodeNumberSeparators = Result
End Function

' ND30-PL1-M2-K3-CASE -- ky hieu (phan sau dau gach cheo) phai viet in hoa. Mot so chu viet
' tat ten loai CHINH THUC theo Phu luc III co chu thuong xen giua chu hoa co chu dich (xem
' shared/rules/chu-viet-tat-ten-loai.json truong "$caseSensitive"), nen khong doi hoi CA CHUOI
' phai in hoa tuyet doi: moi nhom (tach boi dau gach noi, vi du "TTr-VPCC" -> ["TTr", "VPCC"])
' duoc coi la dung neu hoac (a) trung chinh xac (case-sensitive) mot chu viet tat ten loai
' chinh thuc trong file do, hoac (b) tu no da la chu in hoa toan bo (chu viet tat co
' quan/don vi soan thao luon phai in hoa, khong nam trong bang Phu luc III).
Public Function CheckCodeNumberNotationUppercase(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckCodeNumberNotationUppercase")
    Dim paragraphs As Collection: Set paragraphs = CodeNumberParagraphs(context)
    If paragraphs.count = 0 Then
        Set CheckCodeNumberNotationUppercase = Nothing
        Exit Function
    End If

    Dim msg As String
    msg = "K" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u c" & ChrW(&H1EE7) & "a v" & ChrW(&H103) & "n b" & _
        ChrW(&H1EA3) & "n tr" & ChrW(&HEC) & "nh b" & ChrW(&HE0) & "y b" & ChrW(&H1EB1) & "ng ch" & _
        ChrW(&H1EEF) & " in hoa, c" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & " 13, ki" & ChrW(&H1EC3) & "u ch" & _
        ChrW(&H1EEF) & " " & ChrW(&H111) & ChrW(&H1EE9) & "ng."

    Dim knownAbbr As Object: Set knownAbbr = KnownDocTypeAbbreviations()

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        Dim m As Object: Set m = RegexFirstMatch("\/\s*([^/]+)$", Trim$(p.text))
        If Not m Is Nothing Then
            Dim notation As String: notation = CStr(m.SubMatches(0))
            If Not NotationCaseIsCorrect(notation, knownAbbr) Then
                Result.Add MakeFindingInput(CLng(p.Index), msg)
            End If
        End If
    Next p
    Set CheckCodeNumberNotationUppercase = Result
End Function

' Tap hop TAT CA chu viet tat ten loai CHINH THUC (chu-viet-tat-ten-loai.json), GIU NGUYEN dung
' chu hoa/thuong cua file -- dung lam danh sach MIEN TRU cho CheckCodeNumberNotationUppercase.
Private Function KnownDocTypeAbbreviations() As Object
    Dim Result As Object: Set Result = Utils.NewDictionary()
    Dim documentTypes As Collection
    Set documentTypes = RuleLoader.GetDocTypeAbbreviations()("documentTypes")
    Dim d As Variant
    For Each d In documentTypes
        Result(CStr(d("abbreviation"))) = True
    Next d
    Set KnownDocTypeAbbreviations = Result
End Function

' notation dang "TTr-VPCC" (mot hoac nhieu nhom cach nhau dau gach noi). Xem ghi chu dau
' CheckCodeNumberNotationUppercase ve hai dieu kien (a)/(b).
Private Function NotationCaseIsCorrect(ByVal notation As String, ByVal knownAbbr As Object) As Boolean
    Dim parts() As String
    parts = Split(notation, "-")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        Dim part As String: part = Trim$(parts(i))
        If Len(part) > 0 Then
            If Not knownAbbr.Exists(part) Then
                If part <> Utils.ToUpperVn(part) Then
                    NotationCaseIsCorrect = False
                    Exit Function
                End If
            End If
        End If
    Next i
    NotationCaseIsCorrect = True
End Function

' ND30-PL1-M2-K3-ABBR -- chu viet tat ten loai trong ky hieu phai khop Phu luc III. Chi ap dung
' cho van ban "coTenLoai"; thieu du kien -> Nothing (khong doan, ADR-003).
Public Function CheckCodeNumberAbbreviation(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckCodeNumberAbbreviation")
    If CStr(context("DocumentType")) <> "coTenLoai" Then
        Set CheckCodeNumberAbbreviation = Nothing
        Exit Function
    End If
    Dim codeParagraphs As Collection: Set codeParagraphs = CodeNumberParagraphs(context)
    Dim typeParagraphs As Collection
    Set typeParagraphs = ParagraphsWithRole(context("Snapshot"), context("LayoutMap"), "typeName")
    If codeParagraphs.count = 0 Or typeParagraphs.count = 0 Then
        Set CheckCodeNumberAbbreviation = Nothing
        Exit Function
    End If

    Dim documentTypes As Collection
    Set documentTypes = RuleLoader.GetDocTypeAbbreviations()("documentTypes")
    Dim firstType As ParagraphSnapshot
    Set firstType = typeParagraphs(1)
    Dim typeNameText As String
    typeNameText = RegexFirstMatchOrOriginal("^(.*?)[.:]+$", StripTypeQualifier(firstType.text))

    Dim entry As Object: Set entry = Nothing
    Dim d As Variant
    For Each d In documentTypes
        If Utils.ToUpperVn(StripTypeQualifier(CStr(d("typeName")))) = Utils.ToUpperVn(typeNameText) Then
            Set entry = d
            Exit For
        End If
    Next d
    If entry Is Nothing Then
        Set CheckCodeNumberAbbreviation = Nothing
        Exit Function
    End If

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In codeParagraphs
        Dim m As Object: Set m = RegexFirstMatch("\/\s*([^/-]+)", p.text)
        If Not m Is Nothing Then
            Dim actualAbbr As String: actualAbbr = Trim$(CStr(m.SubMatches(0)))
            If actualAbbr <> CStr(entry("abbreviation")) Then
                Dim msg As String
                msg = "K" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u ch" & ChrW(&H1EE9) & "a '" & actualAbbr & _
                    "'. B" & ChrW(&H1EA3) & "ng ch" & ChrW(&H1EEF) & " vi" & ChrW(&H1EBF) & "t t" & _
                    ChrW(&H1EAF) & "t t" & ChrW(&HEA) & "n lo" & ChrW(&H1EA1) & "i quy " & ChrW(&H111) & _
                    ChrW(&H1ECB) & "nh '" & CStr(entry("abbreviation")) & "' cho " & CStr(entry("typeName")) & "."
                Result.Add MakeFindingInput(CLng(p.Index), msg)
            End If
        End If
    Next p
    Set CheckCodeNumberAbbreviation = Result
End Function

Private Function RegexFirstMatchOrOriginal(ByVal pattern As String, ByVal s As String) As String
    Dim m As Object: Set m = RegexFirstMatch(pattern, s)
    If m Is Nothing Then
        RegexFirstMatchOrOriginal = s
    Else
        RegexFirstMatchOrOriginal = CStr(m.SubMatches(0))
    End If
End Function

' ----------------------------------------------------------------------------
' O 4 -- Dia danh va thoi gian ban hanh (checklistGroup 5)
' ----------------------------------------------------------------------------

Public Function CheckPlaceDateStyle(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckPlaceDateStyle")
    Set CheckPlaceDateStyle = StyleCheckForRole(context, "placeAndIssuedDate", _
        ChrW(&H110) & ChrW(&H1ECB) & "a danh v" & ChrW(&HE0) & " th" & ChrW(&H1EDD) & "i gian")
End Function

' ND30-PL1-M2-K4-COMMA -- dau phay sau dia danh, truoc "ngay".
Public Function CheckPlaceDateComma(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckPlaceDateComma")
    Dim paragraphs As Collection
    Set paragraphs = ParagraphsWithRole(context("Snapshot"), context("LayoutMap"), "placeAndIssuedDate")
    If paragraphs.count = 0 Then
        Set CheckPlaceDateComma = Nothing
        Exit Function
    End If

    Dim msg As String
    msg = "Sau " & ChrW(&H111) & ChrW(&H1ECB) & "a danh ph" & ChrW(&H1EA3) & "i c" & ChrW(&HF3) & " d" & _
        ChrW(&H1EA5) & "u ph" & ChrW(&H1EA9) & "y."

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If Not RegexTestCC("^[^,]+,\s*ng" & ChrW(&HE0) & "y", Trim$(p.text)) Then
            Result.Add MakeFindingInput(CLng(p.Index), msg)
        End If
    Next p
    Set CheckPlaceDateComma = Result
End Function

' ND30-PL1-M2-K4-PAD -- ngay < 10 va thang 1, 2 them so 0; thang 3-12 KHONG them.
Public Function CheckPlaceDatePad(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckPlaceDatePad")
    Dim paragraphs As Collection
    Set paragraphs = ParagraphsWithRole(context("Snapshot"), context("LayoutMap"), "placeAndIssuedDate")
    If paragraphs.count = 0 Then
        Set CheckPlaceDatePad = Nothing
        Exit Function
    End If

    Dim spec As Object: Set spec = context("Spec")
    Dim placeSpec As Object: Set placeSpec = spec("components")("placeAndIssuedDate")
    Dim padDayBelow As Long: padDayBelow = 10
    If placeSpec.Exists("padDayBelow") Then padDayBelow = CLng(placeSpec("padDayBelow"))
    Dim padMonths As Object: Set padMonths = Utils.NewDictionary()
    If placeSpec.Exists("padMonthsList") Then
        Dim mv As Variant
        For Each mv In placeSpec("padMonthsList")
            padMonths(CLng(mv)) = True
        Next mv
    Else
        padMonths(1) = True: padMonths(2) = True
    End If

    Dim pattern As String
    pattern = ",\s*ng" & ChrW(&HE0) & "y\s+(\d{1,2})\s+th" & ChrW(&HE1) & "ng\s+(\d{1,2})\s+n" & _
        ChrW(&H103) & "m\s+(\d{4})"

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        Dim m As Object: Set m = RegexFirstMatch(pattern, p.text)
        If Not m Is Nothing Then
            Dim dayStr As String: dayStr = CStr(m.SubMatches(0))
            Dim monthStr As String: monthStr = CStr(m.SubMatches(1))
            Dim dayShouldPad As Boolean: dayShouldPad = (CLng(dayStr) < padDayBelow And Len(dayStr) < 2)
            Dim monthShouldPad As Boolean
            monthShouldPad = (padMonths.Exists(CLng(monthStr)) And Len(monthStr) < 2)
            If dayShouldPad Or monthShouldPad Then
                Dim expDay As String: expDay = IIf(dayShouldPad, PadLeftZero(dayStr, 2), dayStr)
                Dim expMonth As String: expMonth = IIf(monthShouldPad, PadLeftZero(monthStr, 2), monthStr)
                Dim msg As String
                msg = "'ng" & ChrW(&HE0) & "y " & dayStr & " th" & ChrW(&HE1) & "ng " & monthStr & _
                    "' ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " 'ng" & ChrW(&HE0) & _
                    "y " & expDay & " th" & ChrW(&HE1) & "ng " & expMonth & "'. Ng" & ChrW(&HE0) & "y nh" & _
                    ChrW(&H1ECF) & " h" & ChrW(&H1A1) & "n 10 v" & ChrW(&HE0) & " th" & ChrW(&HE1) & "ng 1, " & _
                    "2 ph" & ChrW(&H1EA3) & "i ghi th" & ChrW(&HEA) & "m s" & ChrW(&H1ED1) & " 0 ph" & ChrW(&HED) & _
                    "a tr" & ChrW(&H1B0) & ChrW(&H1EDB) & "c; th" & ChrW(&HE1) & "ng 3 " & ChrW(&H111) & ChrW(&H1EBF) & _
                    "n 12 kh" & ChrW(&HF4) & "ng th" & ChrW(&HEA) & "m."
                Result.Add MakeFindingInput(CLng(p.Index), msg)
            End If
        End If
    Next p
    Set CheckPlaceDatePad = Result
End Function

' ----------------------------------------------------------------------------
' O 5a/5b -- Ten loai, trich yeu (checklistGroup 6)
' ----------------------------------------------------------------------------

Public Function CheckTypeNameStyle(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckTypeNameStyle")
    Set CheckTypeNameStyle = StyleCheckForRole(context, "typeName", _
        "T" & ChrW(&HEA) & "n lo" & ChrW(&H1EA1) & "i v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n")
End Function

Public Function CheckSubjectStyle(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckSubjectStyle")
    Set CheckSubjectStyle = StyleCheckForRole(context, "subject", _
        "Tr" & ChrW(&HED) & "ch y" & ChrW(&H1EBF) & "u n" & ChrW(&H1ED9) & "i dung")
End Function

Public Function CheckSubjectOfficialLetterStyle(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckSubjectOfficialLetterStyle")
    Set CheckSubjectOfficialLetterStyle = StyleCheckForRole(context, "subjectOfficialLetter", _
        "Tr" & ChrW(&HED) & "ch y" & ChrW(&H1EBF) & "u c" & ChrW(&HF4) & "ng v" & ChrW(&H103) & "n")
End Function

' ND30-PL1-M2-K5B-SPACE -- trich yeu cong van cach dong 6pt so voi so, ky hieu.
Public Function CheckSubjectOfficialLetterSpacing(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckSubjectOfficialLetterSpacing")
    Dim paragraphs As Collection
    Set paragraphs = ParagraphsWithRole(context("Snapshot"), context("LayoutMap"), "subjectOfficialLetter")
    If paragraphs.count = 0 Then
        Set CheckSubjectOfficialLetterSpacing = Nothing
        Exit Function
    End If

    Dim spec As Object: Set spec = context("Spec")
    Dim expectedPt As Double: expectedPt = 6
    If spec("components")("subjectOfficialLetter").Exists("spaceBeforePt") Then
        expectedPt = CDbl(spec("components")("subjectOfficialLetter")("spaceBeforePt"))
    End If

    Dim msg As String
    msg = "Tr" & ChrW(&HED) & "ch y" & ChrW(&H1EBF) & "u c" & ChrW(&HF4) & "ng v" & ChrW(&H103) & "n ph" & _
        ChrW(&H1EA3) & "i c" & ChrW(&HE1) & "ch d" & ChrW(&HF2) & "ng 6pt v" & ChrW(&H1EDB) & "i s" & _
        ChrW(&H1ED1) & " v" & ChrW(&HE0) & " k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u v" & ChrW(&H103) & _
        "n b" & ChrW(&H1EA3) & "n."

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If p.SpaceBeforePt < expectedPt - ROUNDING_TOLERANCE_PT Then
            Result.Add MakeFindingInput(CLng(p.Index), msg)
        End If
    Next p
    Set CheckSubjectOfficialLetterSpacing = Result
End Function

' ----------------------------------------------------------------------------
' O 6a -- Can cu ban hanh (checklistGroup 7)
' ----------------------------------------------------------------------------

Public Function CheckLegalBasisStyle(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckLegalBasisStyle")
    Set CheckLegalBasisStyle = StyleCheckForRole(context, "legalBasis", _
        "C" & ChrW(&H103) & "n c" & ChrW(&H1EE9) & " ban h" & ChrW(&HE0) & "nh")
End Function

' ND30-PL1-M2-K6A-PUNCT -- moi dong can cu ket thuc bang ';', dong CUOI CUNG ket thuc bang '.'.
Public Function CheckLegalBasisPunctuation(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckLegalBasisPunctuation")
    Dim paragraphs As Collection
    Set paragraphs = ParagraphsWithRole(context("Snapshot"), context("LayoutMap"), "legalBasis")
    If paragraphs.count = 0 Then
        Set CheckLegalBasisPunctuation = Nothing
        Exit Function
    End If

    ' bulletChar rong (ND30/Viettel) = KHONG duoc co gach noi dau dong; bulletChar "-" (DANG) =
    ' BAT BUOC co gach noi dau dong.
    Dim spec As Object: Set spec = EffectiveSpecOf(context, "legalBasis")
    Dim bullet As String: bullet = ""
    If spec.Exists("bulletChar") Then bullet = CStr(spec("bulletChar"))
    Dim lineEnd As String: lineEnd = CStr(spec("lineEndChar"))
    Dim lastEnd As String: lastEnd = CStr(spec("lastLineEndChar"))

    Dim Result As New Collection
    Dim n As Long: n = paragraphs.count
    Dim i As Long: i = 1
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        Dim problems As Collection: Set problems = New Collection
        Dim trimmed As String: trimmed = Trim$(p.text)
        Dim hasBullet As Boolean: hasBullet = (left$(trimmed, 1) = "-")
        If Len(bullet) > 0 And Not hasBullet Then
            problems.Add "thi" & ChrW(&H1EBF) & "u g" & ChrW(&H1EA1) & "ch " & ChrW(&H111) & _
                ChrW(&H1EA7) & "u d" & ChrW(&HF2) & "ng"
        ElseIf Len(bullet) = 0 And hasBullet Then
            problems.Add "kh" & ChrW(&HF4) & "ng " & ChrW(&H111) & ChrW(&H1B0) & ChrW(&H1EE3) & "c c" & _
                ChrW(&HF3) & " g" & ChrW(&H1EA1) & "ch " & ChrW(&H111) & ChrW(&H1EA7) & "u d" & ChrW(&HF2) & "ng"
        End If
        Dim expectedEnd As String: expectedEnd = IIf(i = n, lastEnd, lineEnd)
        If Right$(trimmed, 1) <> expectedEnd Then
            problems.Add "cu" & ChrW(&H1ED1) & "i d" & ChrW(&HF2) & "ng ph" & ChrW(&H1EA3) & "i l" & _
                ChrW(&HE0) & " '" & expectedEnd & "'"
        End If
        If problems.count > 0 Then
            Result.Add MakeFindingInput(CLng(p.Index), "D" & ChrW(&HF2) & "ng c" & ChrW(&H103) & "n c" & _
                ChrW(&H1EE9) & " ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & _
                ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c: " & JoinCollection(problems, "; ") & ".")
        End If
        i = i + 1
    Next p
    Set CheckLegalBasisPunctuation = Result
End Function

' ----------------------------------------------------------------------------
' O 7 -- Quyen han, chuc vu, ho ten nguoi ky (checklistGroup 8)
' ----------------------------------------------------------------------------

Public Function CheckSignerAuthorityStyle(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckSignerAuthorityStyle")
    Set CheckSignerAuthorityStyle = StyleCheckForRole(context, "signerAuthority", _
        "Quy" & ChrW(&H1EC1) & "n h" & ChrW(&H1EA1) & "n, ch" & ChrW(&H1EE9) & "c v" & ChrW(&H1EE5) & _
        " ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i k" & ChrW(&HFD))
End Function

' ND30-PL1-M2-K7B-AUTH -- chu viet tat quyen han phai la mot trong TM./Q./KT./TL./TUQ.
Public Function CheckSignerAuthorityAbbreviation(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckSignerAuthorityAbbreviation")
    Dim paragraphs As Collection
    Set paragraphs = ParagraphsWithRole(context("Snapshot"), context("LayoutMap"), "signerAuthority")
    If paragraphs.count = 0 Then
        Set CheckSignerAuthorityAbbreviation = Nothing
        Exit Function
    End If

    Dim spec As Object: Set spec = context("Spec")
    Dim prefixesDict As Object: Set prefixesDict = spec("signerAuthorityPrefixes")
    Dim prefixList As String: prefixList = ""
    Dim key As Variant
    Dim isFirst As Boolean: isFirst = True
    For Each key In prefixesDict.Keys
        If Not isFirst Then prefixList = prefixList & ", "
        prefixList = prefixList & CStr(key)
        isFirst = False
    Next key

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        Dim trimmed As String: trimmed = Trim$(p.text)
        Dim matched As Boolean: matched = False
        For Each key In prefixesDict.Keys
            If left$(trimmed, Len(CStr(key))) = CStr(key) Then
                matched = True
                Exit For
            End If
        Next key
        If Not matched Then
            Dim actual As String
            Dim spacePos As Long: spacePos = InStr(trimmed, " ")
            If spacePos = 0 Then actual = trimmed Else actual = left$(trimmed, spacePos - 1)
            Dim msg As String
            msg = "Ph" & ChrW(&HE1) & "t hi" & ChrW(&H1EC7) & "n '" & actual & "'. C" & ChrW(&HE1) & "c " & _
                "ch" & ChrW(&H1EEF) & " vi" & ChrW(&H1EBF) & "t t" & ChrW(&H1EAF) & "t quy " & ChrW(&H111) & _
                ChrW(&H1ECB) & "nh: " & prefixList & "."
            Result.Add MakeFindingInput(CLng(p.Index), msg)
        End If
    Next p
    Set CheckSignerAuthorityAbbreviation = Result
End Function

Private Function SignerNameCandidates(ByVal context As Object) As Collection
    Dim layoutMap As Object: Set layoutMap = context("LayoutMap")

    Dim anchorIndex As Long: anchorIndex = -1
    Dim key As Variant
    Dim r As String
    For Each key In layoutMap.Keys
        r = CStr(layoutMap(key))
        If r = "signerAuthority" Or r = "signerAuthorityTitle" Then
            If CLng(key) > anchorIndex Then anchorIndex = CLng(key)
        End If
    Next key

    Dim Result As New Collection
    If anchorIndex < 0 Then
        Set SignerNameCandidates = Result
        Exit Function
    End If

    Dim p As ParagraphSnapshot
    For Each p In context("Snapshot")("Paragraphs")
        If p.Index = anchorIndex + 1 Then
            If Len(Trim$(p.text)) > 0 And Not p.AllCaps Then Result.Add p
            Exit For
        End If
    Next p
    Set SignerNameCandidates = Result
End Function

' ----------------------------------------------------------------------------
' O 9a/9b -- Kinh gui, Noi nhan (checklistGroup 9)
' ----------------------------------------------------------------------------

' ND30-PL1-M2-K9A-COLON -- doan mo dau bang "Kinh gui"/"Kinh trinh" nhung KHONG duoc gan vai tro
' nao trong hai vai tro chao gui => thieu dau hai cham (ca hai dau hieu nhan dien deu doi hoi).
Public Function CheckRecipientSalutationColon(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckRecipientSalutationColon")
    Dim snapshot As Object: Set snapshot = context("Snapshot")
    Dim layoutMap As Object: Set layoutMap = context("LayoutMap")

    Dim msg As String
    msg = "Sau t" & ChrW(&H1EEB) & " 'K" & ChrW(&HED) & "nh g" & ChrW(&H1EED) & "i' ph" & ChrW(&H1EA3) & _
        "i c" & ChrW(&HF3) & " d" & ChrW(&H1EA5) & "u hai ch" & ChrW(&H1EA5) & "m."

    Dim pattern As String
    pattern = "^\s*K" & ChrW(&HED) & "nh\s+(g" & ChrW(&H1EED) & "i|tr" & ChrW(&HEC) & "nh)"

    Dim Result As New Collection
    Dim role As String
    Dim p As ParagraphSnapshot
    For Each p In snapshot("Paragraphs")
        If RegexTestCC(pattern, p.text) Then
            role = ""
            If layoutMap.Exists(p.Index) Then role = CStr(layoutMap(p.Index))
            If role <> "recipientSalutation" And role <> "recipientSalutationInline" Then
                Result.Add MakeFindingInput(CLng(p.Index), msg)
            End If
        End If
    Next p
    Set CheckRecipientSalutationColon = Result
End Function

' ND30-PL1-M2-K9A-LAYOUT -- "Kinh gui:" dung mot minh mot dong thi PHAI co danh sach gach dau dong
' ngay ben duoi; gui mot noi thi viet lien tren cung dong.
Public Function CheckRecipientSalutationLayout(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckRecipientSalutationLayout")
    Dim snapshot As Object: Set snapshot = context("Snapshot")
    Dim layoutMap As Object: Set layoutMap = context("LayoutMap")

    Dim standalone As Collection
    Set standalone = ParagraphsWithRole(snapshot, layoutMap, "recipientSalutation")
    If standalone.count = 0 Then
        Set CheckRecipientSalutationLayout = Nothing
        Exit Function
    End If

    Dim listLines As Collection
    Set listLines = ParagraphsWithRole(snapshot, layoutMap, "recipientSalutationList")
    If listLines.count > 0 Then
        Set CheckRecipientSalutationLayout = New Collection
        Exit Function
    End If

    Dim msg As String
    msg = "G" & ChrW(&H1EED) & "i m" & ChrW(&H1ED9) & "t n" & ChrW(&H1A1) & "i th" & ChrW(&HEC) & " tr" & _
        ChrW(&HEC) & "nh b" & ChrW(&HE0) & "y tr" & ChrW(&HEA) & "n c" & ChrW(&HF9) & "ng m" & ChrW(&H1ED9) & _
        "t d" & ChrW(&HF2) & "ng; g" & ChrW(&H1EED) & "i t" & ChrW(&H1EEB) & " hai n" & ChrW(&H1A1) & "i tr" & _
        ChrW(&H1EDF) & " l" & ChrW(&HEA) & "n th" & ChrW(&HEC) & " xu" & ChrW(&H1ED1) & "ng d" & ChrW(&HF2) & _
        "ng, m" & ChrW(&H1ED7) & "i n" & ChrW(&H1A1) & "i m" & ChrW(&H1ED9) & "t d" & ChrW(&HF2) & "ng ri" & _
        ChrW(&HEA) & "ng, " & ChrW(&H111) & ChrW(&H1EA7) & "u d" & ChrW(&HF2) & "ng c" & ChrW(&HF3) & " g" & _
        ChrW(&H1EA1) & "ch " & ChrW(&H111) & ChrW(&H1EA7) & "u d" & ChrW(&HF2) & "ng."

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In standalone
        Result.Add MakeFindingInput(CLng(p.Index), msg)
    Next p
    Set CheckRecipientSalutationLayout = Result
End Function

' ND30-PL1-M2-K9A-PUNCT -- danh sach cac noi kinh gui: luon bat dau bang gach noi, ket thuc moi
' dong la dau cham phay, rieng dong cuoi ket thuc bang dau cham.
Public Function CheckRecipientSalutationPunctuation(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckRecipientSalutationPunctuation")
    Dim lines As Collection
    Set lines = ParagraphsWithRole(context("Snapshot"), context("LayoutMap"), "recipientSalutationList")
    If lines.count = 0 Then
        Set CheckRecipientSalutationPunctuation = Nothing
        Exit Function
    End If

    Dim spec As Object: Set spec = EffectiveSpecOf(context, "recipientSalutationList")

    Dim Result As New Collection
    Dim n As Long: n = lines.count
    Dim i As Long: i = 1
    Dim line As ParagraphSnapshot
    For Each line In lines
        Dim problems As Collection: Set problems = New Collection
        Dim trimmed As String: trimmed = Trim$(line.text)
        If spec.Exists("bulletChar") Then
            Dim bullet As String: bullet = CStr(spec("bulletChar"))
            If Len(bullet) > 0 And left$(trimmed, Len(bullet)) <> bullet Then
                problems.Add "thi" & ChrW(&H1EBF) & "u g" & ChrW(&H1EA1) & "ch " & ChrW(&H111) & _
                    ChrW(&H1EA7) & "u d" & ChrW(&HF2) & "ng"
            End If
        End If
        Dim expectedEnd As String
        expectedEnd = IIf(i = n, CStr(spec("lastItemEndChar")), CStr(spec("itemEndChar")))
        If Right$(trimmed, 1) <> expectedEnd Then
            problems.Add "cu" & ChrW(&H1ED1) & "i d" & ChrW(&HF2) & "ng ph" & ChrW(&H1EA3) & "i l" & _
                ChrW(&HE0) & " '" & expectedEnd & "'"
        End If
        If problems.count > 0 Then
            Result.Add MakeFindingInput(CLng(line.Index), "Danh s" & ChrW(&HE1) & "ch n" & ChrW(&H1A1) & _
                "i k" & ChrW(&HED) & "nh g" & ChrW(&H1EED) & "i ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & _
                ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c: " & _
                JoinCollection(problems, "; ") & ".")
        End If
        i = i + 1
    Next line
    Set CheckRecipientSalutationPunctuation = Result
End Function

' ND30-PL1-M2-K9A-INLINE-END -- noi nhan viet lien sau "Kinh gui:" tren cung dong thi cuoi dong
' phai co dau cham.
Public Function CheckRecipientSalutationInlineEnd(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckRecipientSalutationInlineEnd")
    Dim paragraphs As Collection
    Set paragraphs = ParagraphsWithRole(context("Snapshot"), context("LayoutMap"), "recipientSalutationInline")
    If paragraphs.count = 0 Then
        Set CheckRecipientSalutationInlineEnd = Nothing
        Exit Function
    End If

    Dim spec As Object: Set spec = EffectiveSpecOf(context, "recipientSalutationInlineContent")
    Dim expectedEnd As String: expectedEnd = "."
    If spec.Exists("lastItemEndChar") Then expectedEnd = CStr(spec("lastItemEndChar"))

    Dim msg As String
    msg = "Khi n" & ChrW(&H1A1) & "i nh" & ChrW(&H1EAD) & "n " & ChrW(&H111) & ChrW(&H1B0) & ChrW(&H1EE3) & _
        "c vi" & ChrW(&H1EBF) & "t ngay sau 'K" & ChrW(&HED) & "nh g" & ChrW(&H1EED) & "i:' tr" & ChrW(&HEA) & _
        "n c" & ChrW(&HF9) & "ng m" & ChrW(&H1ED9) & "t d" & ChrW(&HF2) & "ng th" & ChrW(&HEC) & " cu" & _
        ChrW(&H1ED1) & "i d" & ChrW(&HF2) & "ng ph" & ChrW(&H1EA3) & "i c" & ChrW(&HF3) & " d" & _
        ChrW(&H1EA5) & "u ch" & ChrW(&H1EA5) & "m."

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If Right$(Trim$(p.text), 1) <> expectedEnd Then Result.Add MakeFindingInput(CLng(p.Index), msg)
    Next p
    Set CheckRecipientSalutationInlineEnd = Result
End Function

' ND30-PL1-M2-K9B-LABEL -- tu "Noi nhan" kem dau hai cham + kieu chu.
Public Function CheckRecipientLabelStyle(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckRecipientLabelStyle")
    Dim paragraphs As Collection
    Set paragraphs = ParagraphsWithRole(context("Snapshot"), context("LayoutMap"), "recipientLabel")
    If paragraphs.count = 0 Then
        Set CheckRecipientLabelStyle = Nothing
        Exit Function
    End If

    Dim componentSpec As Object: Set componentSpec = EffectiveSpecOf(context, "recipientLabel")

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        Dim problems As Collection: Set problems = DescribeStyleViolations(p, componentSpec)
        If Not RegexTestCC("^\s*N" & ChrW(&H1A1) & "i\s+nh" & ChrW(&H1EAD) & "n\s*:", p.text) Then
            problems.Add "thi" & ChrW(&H1EBF) & "u d" & ChrW(&H1EA5) & "u hai ch" & ChrW(&H1EA5) & "m sau '" & _
                "N" & ChrW(&H1A1) & "i nh" & ChrW(&H1EAD) & "n'"
        End If
        If problems.count > 0 Then
            Dim msg As String
            msg = "T" & ChrW(&H1EEB) & " 'N" & ChrW(&H1A1) & "i nh" & ChrW(&H1EAD) & "n' ch" & ChrW(&H1B0) & _
                "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c: " & _
                JoinCollection(problems, "; ") & "."
            Result.Add MakeFindingInput(CLng(p.Index), msg)
        End If
    Next p
    Set CheckRecipientLabelStyle = Result
End Function

' ND30-PL1-M2-K9B-LIST -- kieu chu + gach dau dong + dau cham phay cuoi dong. Dong CUOI CUNG la
' dong "Luu..." co quy dinh rieng (K9B-LUU) nen KHONG kiem dau ket thuc o day.
Public Function CheckRecipientListStyle(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckRecipientListStyle")
    Dim paragraphs As Collection
    Set paragraphs = ParagraphsWithRole(context("Snapshot"), context("LayoutMap"), "recipientList")
    If paragraphs.count = 0 Then
        Set CheckRecipientListStyle = Nothing
        Exit Function
    End If

    Dim componentSpec As Object: Set componentSpec = EffectiveSpecOf(context, "recipientList")

    Dim Result As New Collection
    Dim n As Long: n = paragraphs.count
    Dim i As Long: i = 1
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        Dim problems As Collection: Set problems = DescribeStyleViolations(p, componentSpec)
        Dim trimmed As String: trimmed = Trim$(p.text)
        If left$(trimmed, 1) <> CStr(componentSpec("bulletChar")) Then
            problems.Add "thi" & ChrW(&H1EBF) & "u g" & ChrW(&H1EA1) & "ch " & ChrW(&H111) & ChrW(&H1EA7) & _
                "u d" & ChrW(&HF2) & "ng s" & ChrW(&HE1) & "t l" & ChrW(&H1EC1) & " tr" & ChrW(&HE1) & "i"
        End If
        If i < n Then
            If Right$(trimmed, 1) <> CStr(componentSpec("itemEndChar")) Then
                problems.Add "cu" & ChrW(&H1ED1) & "i d" & ChrW(&HF2) & "ng thi" & ChrW(&H1EBF) & "u d" & _
                    ChrW(&H1EA5) & "u ch" & ChrW(&H1EA5) & "m ph" & ChrW(&H1EA9) & "y"
            End If
        End If
        If problems.count > 0 Then
            Dim msg As String
            msg = "Danh s" & ChrW(&HE1) & "ch n" & ChrW(&H1A1) & "i nh" & ChrW(&H1EAD) & "n ch" & _
                ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & _
                "c: " & JoinCollection(problems, "; ") & "."
            Result.Add MakeFindingInput(CLng(p.Index), msg)
        End If
        i = i + 1
    Next p
    Set CheckRecipientListStyle = Result
End Function

' ND30-PL1-M2-K9B-LUU -- chi kiem dong CUOI CUNG cua danh sach noi nhan. Cau truc lay tu
' components.recipientList.archiveLinePattern (khac nhau giua ND30/VIETTEL va DANG).
Public Function CheckRecipientListLuuLine(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckRecipientListLuuLine")
    Dim paragraphs As Collection
    Set paragraphs = ParagraphsWithRole(context("Snapshot"), context("LayoutMap"), "recipientList")
    If paragraphs.count = 0 Then
        Set CheckRecipientListLuuLine = Nothing
        Exit Function
    End If

    Dim componentSpec As Object: Set componentSpec = EffectiveSpecOf(context, "recipientList")
    If Not componentSpec.Exists("archiveLinePattern") Then
        Set CheckRecipientListLuuLine = Nothing
        Exit Function
    End If

    Dim last As ParagraphSnapshot: Set last = paragraphs(paragraphs.count)
    Dim Result As New Collection
    If Not RegexTestCC(CStr(componentSpec("archiveLinePattern")), Trim$(last.text)) Then
        Dim msg As String
        msg = "D" & ChrW(&HF2) & "ng 'L" & ChrW(&H1B0) & "u' ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & _
            ChrW(&HFA) & "ng c" & ChrW(&H1EA5) & "u tr" & ChrW(&HFA) & "c quy " & ChrW(&H111) & _
            ChrW(&H1ECB) & "nh: " & ArchiveLineHintVn(RegimeOf(context)) & "."
        Result.Add MakeFindingInput(CLng(last.Index), msg)
    End If
    Set CheckRecipientListLuuLine = Result
End Function

Private Function ArchiveLineHintVn(ByVal regime As String) As String
    Select Case regime
        Case "DANG"
            ArchiveLineHintVn = "L" & ChrW(&H1B0) & "u <" & ChrW(&H111) & ChrW(&H1A1) & "n v" & ChrW(&H1ECB) & _
                ">.<t" & ChrW(&HEA) & "n ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i so" & ChrW(&H1EA1) & "n>-<s" & _
                ChrW(&H1ED1) & " l" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng>"
        Case "VIETTEL"
            ArchiveLineHintVn = "L" & ChrW(&H1B0) & "u: VT, <" & ChrW(&H111) & ChrW(&H1A1) & "n v" & _
                ChrW(&H1ECB) & ">.<t" & ChrW(&HEA) & "n ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i so" & _
                ChrW(&H1EA1) & "n><s" & ChrW(&H1ED1) & " l" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng>."
        Case Else
            ArchiveLineHintVn = "L" & ChrW(&H1B0) & "u: VT, <" & ChrW(&H111) & ChrW(&H1A1) & "n v" & _
                ChrW(&H1ECB) & ">.<t" & ChrW(&HEA) & "n ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i so" & _
                ChrW(&H1EA1) & "n>.(<s" & ChrW(&H1ED1) & " l" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng>)"
    End Select
End Function

' ----------------------------------------------------------------------------
' Muc III -- Phu luc (checklistGroup 11) -- KHONG co vai tro trong LayoutMap, tu do bang regex.
' ----------------------------------------------------------------------------

Private Function AppendixLabelParagraphs(ByVal snapshot As Object) As Collection
    Dim Result As New Collection
    Dim pattern As String
    pattern = "^\s*Ph" & ChrW(&H1EE5) & "\s*l" & ChrW(&H1EE5) & "c(\s+([IVXLCDM]+))?\s*\.?\s*$"
    Dim p As ParagraphSnapshot
    For Each p In snapshot("Paragraphs")
        If RegexTestCC(pattern, Trim$(p.text)) Then Result.Add p
    Next p
    Set AppendixLabelParagraphs = Result
End Function

' ND30-PL1-M3-K1A-REF -- warning, khong autoFixable.
Public Function CheckAppendixReferenceMentioned(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckAppendixReferenceMentioned")
    Dim snapshot As Object: Set snapshot = context("Snapshot")
    Dim labels As Collection: Set labels = AppendixLabelParagraphs(snapshot)
    If labels.count = 0 Then
        Set CheckAppendixReferenceMentioned = Nothing
        Exit Function
    End If
    Dim firstAppendixIndex As Long: firstAppendixIndex = labels(1).Index

    Dim mentioned As Boolean: mentioned = False
    Dim p As ParagraphSnapshot
    For Each p In snapshot("Paragraphs")
        If p.Index < firstAppendixIndex Then
            If RegexTestCC("ph" & ChrW(&H1EE5) & "\s*l" & ChrW(&H1EE5) & "c", Utils.ToLowerVn(p.text)) Then
                mentioned = True
                Exit For
            End If
        End If
    Next p

    Dim Result As New Collection
    If Not mentioned Then
        Dim msg As String
        msg = "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng h" & ChrW(&H1EE3) & "p v" & ChrW(&H103) & "n b" & _
            ChrW(&H1EA3) & "n c" & ChrW(&HF3) & " ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c k" & _
            ChrW(&HE8) & "m theo th" & ChrW(&HEC) & " trong v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n ph" & _
            ChrW(&H1EA3) & "i c" & ChrW(&HF3) & " ch" & ChrW(&H1EC9) & " d" & ChrW(&H1EAB) & "n v" & ChrW(&H1EC1) & _
            " ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c " & ChrW(&H111) & ChrW(&HF3) & "."
        Result.Add MakeFindingInput(Null, msg)
    End If
    Set CheckAppendixReferenceMentioned = Result
End Function

' ND30-PL1-M3-K1A-NUM -- chi bat buoc danh so khi co TU HAI phu luc tro len.
Public Function CheckAppendixRomanNumbering(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckAppendixRomanNumbering")
    Dim labels As Collection: Set labels = AppendixLabelParagraphs(context("Snapshot"))
    If labels.count = 0 Then
        Set CheckAppendixRomanNumbering = Nothing
        Exit Function
    End If
    Dim Result As New Collection
    If labels.count < 2 Then
        Set CheckAppendixRomanNumbering = Result
        Exit Function
    End If

    Dim msg As String
    msg = "V" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n c" & ChrW(&HF3) & " t" & ChrW(&H1EEB) & " hai ph" & _
        ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c tr" & ChrW(&H1EDF) & " l" & ChrW(&HEA) & "n th" & ChrW(&HEC) & _
        " c" & ChrW(&HE1) & "c ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c ph" & ChrW(&H1EA3) & "i " & _
        ChrW(&H111) & ChrW(&H1B0) & ChrW(&H1EE3) & "c " & ChrW(&H111) & ChrW(&HE1) & "nh s" & ChrW(&H1ED1) & _
        " th" & ChrW(&H1EE9) & " t" & ChrW(&H1EF1) & " b" & ChrW(&H1EB1) & "ng ch" & ChrW(&H1EEF) & " s" & _
        ChrW(&H1ED1) & " La M" & ChrW(&HE3) & "."

    Dim p As ParagraphSnapshot
    For Each p In labels
        Dim m As Object
        Set m = RegexFirstMatch("^\s*Ph" & ChrW(&H1EE5) & "\s*l" & ChrW(&H1EE5) & "c\s+([^\s.]+)", Trim$(p.text))
        Dim ok As Boolean: ok = False
        If Not m Is Nothing Then
            If RegexTestCC("^[IVXLCDM]+$", CStr(m.SubMatches(0))) Then ok = True
        End If
        If Not ok Then Result.Add MakeFindingInput(CLng(p.Index), msg)
    Next p
    Set CheckAppendixRomanNumbering = Result
End Function

' ND30-PL1-M3-K1B -- dong "Phu luc <so>" + tieu de phu luc (doan NGAY DUOI, doan Index+1).
Public Function CheckAppendixTitleStyle(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckAppendixTitleStyle")
    Dim snapshot As Object: Set snapshot = context("Snapshot")
    Dim labels As Collection: Set labels = AppendixLabelParagraphs(snapshot)
    If labels.count = 0 Then
        Set CheckAppendixTitleStyle = Nothing
        Exit Function
    End If

    Dim spec As Object: Set spec = context("Spec")
    Dim labelSpec As Object: Set labelSpec = spec("components")("appendixLabel")
    Dim titleSpec As Object: Set titleSpec = spec("components")("appendixTitle")
    Dim paragraphByIndex As Object: Set paragraphByIndex = Utils.NewDictionary()
    Dim pp As ParagraphSnapshot
    For Each pp In snapshot("Paragraphs")
        Set paragraphByIndex(pp.Index) = pp
    Next pp

    Dim Result As New Collection
    Dim LABEL As ParagraphSnapshot
    For Each LABEL In labels
        Dim labelProblems As Collection: Set labelProblems = DescribeStyleViolations(LABEL, labelSpec)
        If labelProblems.count > 0 Then
            Dim labelMsg As String
            labelMsg = "T" & ChrW(&H1EEB) & " 'Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c' v" & _
                ChrW(&HE0) & " s" & ChrW(&H1ED1) & " th" & ChrW(&H1EE9) & " t" & ChrW(&H1EF1) & " ch" & _
                ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & _
                ChrW(&H1EE9) & "c: " & JoinCollection(labelProblems, "; ") & "."
            Result.Add MakeFindingInput(CLng(LABEL.Index), labelMsg)
        End If
        If paragraphByIndex.Exists(LABEL.Index + 1) Then
            Dim title As ParagraphSnapshot: Set title = paragraphByIndex(LABEL.Index + 1)
            Dim titleProblems As Collection: Set titleProblems = DescribeStyleViolations(title, titleSpec)
            If titleProblems.count > 0 Then
                Dim titleMsg As String
                titleMsg = "T" & ChrW(&HEA) & "n ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c ch" & _
                    ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & _
                    ChrW(&H1EE9) & "c: " & JoinCollection(titleProblems, "; ") & "."
                Result.Add MakeFindingInput(CLng(title.Index), titleMsg)
            End If
        End If
    Next LABEL
    Set CheckAppendixTitleStyle = Result
End Function

' ND30-PL1-M3-K1C -- thong tin chi dan "(Kem theo Van ban so...)" ngay duoi ten phu luc.
Public Function CheckAppendixReferenceInfoStyle(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckAppendixReferenceInfoStyle")
    Dim snapshot As Object: Set snapshot = context("Snapshot")
    Dim pattern As String
    pattern = "^\s*\(\s*K" & ChrW(&HE8) & "m\s+theo\b"

    Dim matches As New Collection
    Dim p As ParagraphSnapshot
    For Each p In snapshot("Paragraphs")
        If RegexTestCC(pattern, p.text) Then matches.Add p
    Next p
    If matches.count = 0 Then
        Set CheckAppendixReferenceInfoStyle = Nothing
        Exit Function
    End If

    Dim spec As Object: Set spec = context("Spec")
    Dim componentSpec As Object: Set componentSpec = spec("components")("appendixReference")

    Dim Result As New Collection
    For Each p In matches
        Dim problems As Collection: Set problems = DescribeStyleViolations(p, componentSpec)
        If problems.count > 0 Then
            Dim msg As String
            msg = "Th" & ChrW(&HF4) & "ng tin ch" & ChrW(&H1EC9) & " d" & ChrW(&H1EAB) & "n k" & _
                ChrW(&HE8) & "m theo ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c ch" & ChrW(&H1B0) & "a " & _
                ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c: " & _
                JoinCollection(problems, "; ") & "."
            Result.Add MakeFindingInput(CLng(p.Index), msg)
        End If
    Next p
    Set CheckAppendixReferenceInfoStyle = Result
End Function

' ND30-PL1-M3-K2B / K2C / K3 / K4 (dau khan, pham vi luu hanh, ky hieu soan thao, thong tin lien
' he) va ND30-PL1-M4-POS (vi tri 14 o theo so do) -- CHUA HIEN THUC HOA DUOC: bon vai tro dau
' khong co dau hieu nhan dien o docs/rules/01 muc 5 nen KHONG BAO GIO co trong LayoutMap; M4-POS
' can toa do/vung trang ma DocumentSnapshot khong chup. Tra Nothing mai mai.
Public Function CheckComponentNeverDetected(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckComponentNeverDetected")
    Set CheckComponentNeverDetected = Nothing
End Function

' ----------------------------------------------------------------------------
' Dang ky -- goi tu EnsureRegistryInitialized, GHI THANG vao mRegistry.
' ----------------------------------------------------------------------------

Private Sub RegisterComponentChecks()
    mRegistry("ND30-PL1-M2-K1-QH") = "ComplianceChecker.CheckNationalTitleStyle"
    mRegistry("ND30-PL1-M2-K1-TN") = "ComplianceChecker.CheckNationalMottoStyle"
    mRegistry("ND30-PL1-M2-K1-TN-SEP") = "ComplianceChecker.CheckNationalMottoSeparator"
    mRegistry("ND30-PL1-M2-K1-TN-LINE") = "ComplianceChecker.CheckComponentUnderline"
    mRegistry("ND30-PL1-M2-K1-C") = "ComplianceChecker.CheckNationalMottoSpacing"

    mRegistry("ND30-PL1-M2-K2-SUP") = "ComplianceChecker.CheckSuperiorOrganNameStyle"
    mRegistry("ND30-PL1-M2-K2-ORG") = "ComplianceChecker.CheckOrganNameStyle"
    mRegistry("ND30-PL1-M2-K2-LINE") = "ComplianceChecker.CheckComponentUnderline"

    mRegistry("ND30-PL1-M2-K3-PREFIX") = "ComplianceChecker.CheckCodeNumberColon"
    mRegistry("ND30-PL1-M2-K3-PAD") = "ComplianceChecker.CheckCodeNumberPad"
    mRegistry("ND30-PL1-M2-K3-SPACE") = "ComplianceChecker.CheckCodeNumberNoSpace"
    mRegistry("ND30-PL1-M2-K3-SEP") = "ComplianceChecker.CheckCodeNumberSeparators"
    mRegistry("ND30-PL1-M2-K3-CASE") = "ComplianceChecker.CheckCodeNumberNotationUppercase"
    mRegistry("ND30-PL1-M2-K3-ABBR") = "ComplianceChecker.CheckCodeNumberAbbreviation"

    mRegistry("ND30-PL1-M2-K4-STYLE") = "ComplianceChecker.CheckPlaceDateStyle"
    mRegistry("ND30-PL1-M2-K4-COMMA") = "ComplianceChecker.CheckPlaceDateComma"
    mRegistry("ND30-PL1-M2-K4-PAD") = "ComplianceChecker.CheckPlaceDatePad"

    mRegistry("ND30-PL1-M2-K5A-TYPE") = "ComplianceChecker.CheckTypeNameStyle"
    mRegistry("ND30-PL1-M2-K5A-SUBJ") = "ComplianceChecker.CheckSubjectStyle"
    mRegistry("ND30-PL1-M2-K5A-LINE") = "ComplianceChecker.CheckComponentUnderline"
    mRegistry("ND30-PL1-M2-K5B-STYLE") = "ComplianceChecker.CheckSubjectOfficialLetterStyle"
    mRegistry("ND30-PL1-M2-K5B-SPACE") = "ComplianceChecker.CheckSubjectOfficialLetterSpacing"

    mRegistry("ND30-PL1-M2-K6A-STYLE") = "ComplianceChecker.CheckLegalBasisStyle"
    mRegistry("ND30-PL1-M2-K6A-PUNCT") = "ComplianceChecker.CheckLegalBasisPunctuation"

    mRegistry("ND30-PL1-M2-K7B-AUTH") = "ComplianceChecker.CheckSignerAuthorityAbbreviation"
    mRegistry("ND30-PL1-M2-K7D-STYLE") = "ComplianceChecker.CheckSignerAuthorityStyle"

    mRegistry("ND30-PL1-M2-K9A-COLON") = "ComplianceChecker.CheckRecipientSalutationColon"
    mRegistry("ND30-PL1-M2-K9A-LAYOUT") = "ComplianceChecker.CheckRecipientSalutationLayout"
    mRegistry("ND30-PL1-M2-K9A-PUNCT") = "ComplianceChecker.CheckRecipientSalutationPunctuation"
    mRegistry("ND30-PL1-M2-K9A-INLINE-END") = "ComplianceChecker.CheckRecipientSalutationInlineEnd"
    mRegistry("ND30-PL1-M2-K9B-LABEL") = "ComplianceChecker.CheckRecipientLabelStyle"
    mRegistry("ND30-PL1-M2-K9B-LIST") = "ComplianceChecker.CheckRecipientListStyle"
    mRegistry("ND30-PL1-M2-K9B-LUU") = "ComplianceChecker.CheckRecipientListLuuLine"

    mRegistry("ND30-PL1-M3-K1A-REF") = "ComplianceChecker.CheckAppendixReferenceMentioned"
    mRegistry("ND30-PL1-M3-K1A-NUM") = "ComplianceChecker.CheckAppendixRomanNumbering"
    mRegistry("ND30-PL1-M3-K1B") = "ComplianceChecker.CheckAppendixTitleStyle"
    mRegistry("ND30-PL1-M3-K1C") = "ComplianceChecker.CheckAppendixReferenceInfoStyle"

    mRegistry("ND30-PL1-M3-K2B") = "ComplianceChecker.CheckComponentNeverDetected"
    mRegistry("ND30-PL1-M3-K2C") = "ComplianceChecker.CheckComponentNeverDetected"
    mRegistry("ND30-PL1-M3-K3") = "ComplianceChecker.CheckComponentNeverDetected"
    mRegistry("ND30-PL1-M3-K4") = "ComplianceChecker.CheckComponentNeverDetected"
    mRegistry("ND30-PL1-M4-POS") = "ComplianceChecker.CheckComponentNeverDetected"
End Sub

' bo hoan toan khoi code moi noi lien quan phan/chuong/muc/tieu muc).
Public Function VnDieu() As String
    VnDieu = ChrW(&H110) & "i" & ChrW(&H1EC1) & "u"
End Function

Public Function VnKhoan() As String
    VnKhoan = "Kho" & ChrW(&H1EA3) & "n"
End Function

Public Function VnDiem() As String
    VnDiem = ChrW(&H110) & "i" & ChrW(&H1EC3) & "m"
End Function

' Docs/rules/01 muc 5 ghi lop ky tu cua regex "Diem" la [a-z...] kem dai "a-z" o dau -- TU MAU
' THUAN voi chinh "Rat de sai" ngay tren no (muc 6.3) khang dinh bang chu cai tieng Viet KHONG co
' f, j, w, z. Day la loi danh may trong tai lieu. Dung lop ky tu TRUC TIEP tu
' spec("vietnameseAlphabet") thay vi regex hard-code cua docs -- vua tranh loi do vua dung "Rang
' buoc" (khong so sanh chuoi mac dinh, luon tra JSON).
Private Function EscapeRegExpLiteral(ByVal text As String) As String
    Dim specials As String: specials = ".*+?^$(){}|[]\"
    Dim Result As String: Result = ""
    Dim i As Long
    For i = 1 To Len(text)
        Dim ch As String: ch = Mid$(text, i, 1)
        If InStr(specials, ch) > 0 Then
            Result = Result & "\" & ch
        Else
            Result = Result & ch
        End If
    Next i
    EscapeRegExpLiteral = Result
End Function

Public Function BuildPointPattern(ByVal vietnameseAlphabet As Collection) As String
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.BuildPointPattern")
    Dim charClass As String: charClass = ""
    Dim v As Variant
    For Each v In vietnameseAlphabet
        charClass = charClass & EscapeRegExpLiteral(CStr(v))
    Next v
    BuildPointPattern = "^\s*([" & charClass & "])\)\s"
End Function

' bo hoan toan khoi code moi noi lien quan phan/chuong/muc/tieu muc).

Public Function ArticleKeywordPattern() As String
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.ArticleKeywordPattern")
    ArticleKeywordPattern = "^\s*" & VnDieu() & "\s+(\d+)\s*\.\s*"
End Function

' Doan khop "Dieu <so A Rap>." -- do tin cay Cao (docs/rules/01 muc 5). loai tru doan nam trong
' bang.
Public Function ArticleParagraphs(ByVal snapshot As Object) As Collection
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.ArticleParagraphs")
    Dim pattern As String: pattern = ArticleKeywordPattern()
    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In snapshot("Paragraphs")
        If Not p.isInTable Then
            Dim m As Object: Set m = RegexFirstMatch(pattern, p.text)
            If Not m Is Nothing Then
                Dim item As Object: Set item = Utils.NewDictionary()
                item.Add "Paragraph", p
                item.Add "Number", CLng(m.SubMatches(0))
                Result.Add item
            End If
        End If
    Next p
    Set ArticleParagraphs = Result
End Function

' Xuat rieng chuoi pattern (khac ArticleKeywordPattern -- da co tu truoc) de
' StructureFormatter.bas tai dung NGUYEN VEN khi do ranh gioi reset cua RenumberPoints, tranh mot
' ban sao lech.
Public Function clausePattern() As String
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.ClausePattern")
    clausePattern = "^\s*(\d+)\.\s"
End Function

' Doan khop "<so A Rap>. " -- do tin cay TRUNG BINH (docs/rules/01 muc 5), loai tru doan da khop
' "Dieu". CheckClauseFormat con DOI HOI tai lieu da co it nhat mot Dieu truoc khi coi cac doan nay
' la Khoan -- tu than regex khong du tin cay de khang dinh mot danh sach so bat ky trong van ban
' la cap "Khoan" (ADR-003).
Public Function ClauseParagraphs(ByVal snapshot As Object) As Collection
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.ClauseParagraphs")
    Dim articlePattern As String: articlePattern = ArticleKeywordPattern()
    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In snapshot("Paragraphs")
        If Not RegexTestCC(articlePattern, p.text) Then
            Dim m As Object: Set m = RegexFirstMatch(clausePattern(), p.text)
            If Not m Is Nothing Then
                Dim item As Object: Set item = Utils.NewDictionary()
                item.Add "Paragraph", p
                item.Add "Number", CLng(m.SubMatches(0))
                Result.Add item
            End If
        End If
    Next p
    Set ClauseParagraphs = Result
End Function

' Doan khop "<chu cai tieng Viet>) " -- do tin cay TRUNG BINH; lop ky tu dung 29 chu cai tieng
' Viet (dung tu spec.vietnameseAlphabet, xem BuildPointPattern) nen so thu tu tach ra LUON hop le,
' khong can kiem lai bang StructureNumeralValid.
Public Function PointParagraphs(ByVal snapshot As Object, ByVal vietnameseAlphabet As Collection) As Collection
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.PointParagraphs")
    Dim pattern As String: pattern = BuildPointPattern(vietnameseAlphabet)
    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In snapshot("Paragraphs")
        Dim m As Object: Set m = RegexFirstMatch(pattern, p.text)
        If Not m Is Nothing Then
            Dim item As Object: Set item = Utils.NewDictionary()
            item.Add "Paragraph", p
            item.Add "Letter", CStr(m.SubMatches(0))
            Result.Add item
        End If
    Next p
    Set PointParagraphs = Result
End Function

' bo hoan toan khoi code moi noi lien quan phan/chuong/muc/tieu muc).

Private Function RegexReplaceFirst(ByVal pattern As String, ByVal s As String) As String
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern
    regex.IgnoreCase = False
    regex.Global = False
    RegexReplaceFirst = regex.Replace(s, "")
End Function

' Tra Match toan bo (Global=True) -- dung cho cac quy tac citation can nhieu vien dan tren cung
' mot doan (RegexFirstMatch/RegexTestCC deu Global=False, khong du cho truong hop nay).
Private Function RegexAllMatches(ByVal pattern As String, ByVal s As String) As Object
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern
    regex.IgnoreCase = False
    regex.Global = True
    Set RegexAllMatches = regex.Execute(s)
End Function

Public Function CheckStructureTitlePresence(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckStructureTitlePresence")
    Dim snapshot As Object: Set snapshot = context("Snapshot")

    Dim Result As New Collection
    Dim anyFound As Boolean: anyFound = False

    Dim articles As Collection: Set articles = ArticleParagraphs(snapshot)
    If articles.count > 0 Then
        anyFound = True
        Dim a As Variant
        For Each a In articles
            Dim ap As ParagraphSnapshot: Set ap = a("Paragraph")
            Dim articleNumber As Long: articleNumber = CLng(a("Number"))
            Dim afterKeyword As String: afterKeyword = Trim$(RegexReplaceFirst(ArticleKeywordPattern(), ap.text))
            If afterKeyword = "" Then
                Result.Add MakeFindingInput(CLng(ap.Index), VnDieu() & " " & CStr(articleNumber) & " thi" & _
                    ChrW(&H1EBF) & "u ti" & ChrW(&HEA) & "u " & ChrW(&H111) & ChrW(&H1EC1) & ".")
            End If
        Next a
    End If

    If Not anyFound Then
        Set CheckStructureTitlePresence = Nothing
    Else
        Set CheckStructureTitlePresence = Result
    End If
End Function

' "khong kiem tra, khong xu ly tat ca nhung van de lien quan den phan, chuong, muc, tieu muc - bo
' hoan toan khoi code") - ham StructureKeywordTitleCheck/
' CheckPartChapterFormat/CheckSectionSubsectionFormat cu da xoa, cung StructureFormatter.
' FixPartChapterFormat/FixSectionSubsectionFormat (xem StructureFormatter.bas).

' ND30-PL1-M2-K6D-ARTICLE -- full: lui dau dong, kieu chu, VA lien tuc so thu tu (Dieu 1, 2, 3...
' khong nhay so -- muc "Viec phai lam" #3 cua; khong co ruleCode rieng cho viec nay nen gop vao
' day). Dau cham sau so thu tu da duoc chinh regex nhan dien doi hoi, khong kiem lai.
Public Function CheckArticleFormat(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckArticleFormat")
    Dim snapshot As Object: Set snapshot = context("Snapshot")
    Dim articles As Collection: Set articles = ArticleParagraphs(snapshot)
    If articles.count = 0 Then
        Set CheckArticleFormat = Nothing
        Exit Function
    End If

    Dim spec As Object: Set spec = context("Spec")
    Dim levelSpec As Object: Set levelSpec = spec("structureLevels")("article")
    Dim allowedIndentCm As Collection: Set allowedIndentCm = levelSpec("indentCm")("allowed")
    Dim toleranceCm As Double: toleranceCm = PointToCm(ROUNDING_TOLERANCE_PT, spec)

    Dim Result As New Collection
    Dim expectedNext As Variant: expectedNext = Null
    Dim a As Variant
    For Each a In articles
        Dim p As ParagraphSnapshot: Set p = a("Paragraph")
        Dim number As Long: number = CLng(a("Number"))
        Dim problems As New Collection

        Dim actualIndentCm As Double: actualIndentCm = PointToCm(p.FirstLineIndentPt, spec)
        Dim indentOk As Boolean: indentOk = False
        Dim v As Variant
        For Each v In allowedIndentCm
            If Abs(actualIndentCm - CDbl(v)) <= toleranceCm Then indentOk = True
        Next v
        If Not indentOk Then
            problems.Add "l" & ChrW(&HF9) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u d" & ChrW(&HF2) & "ng " & _
                FormatVnNumber(actualIndentCm, 2) & " cm, quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh 1 cm ho" & _
                ChrW(&H1EB7) & "c 1,27 cm"
        End If

        If levelSpec.Exists("style") Then
            Dim styleProblems As Collection: Set styleProblems = DescribeStyleViolations(p, levelSpec)
            Dim sv As Variant
            For Each sv In styleProblems
                problems.Add sv
            Next sv
        End If

        If Not IsNull(expectedNext) Then
            If number <> CLng(expectedNext) Then
                problems.Add "s" & ChrW(&H1ED1) & " th" & ChrW(&H1EE9) & " t" & ChrW(&H1EF1) & " nh" & _
                    ChrW(&H1EA3) & "y sang " & VnDieu() & " " & CStr(number) & ", ph" & ChrW(&H1EA3) & "i li" & _
                    ChrW(&HEA) & "n t" & ChrW(&H1EE5) & "c l" & ChrW(&HE0) & " " & VnDieu() & " " & CStr(expectedNext)
            End If
        End If
        expectedNext = number + 1

        If problems.count > 0 Then
            Result.Add MakeFindingInput(CLng(p.Index), VnDieu() & " " & CStr(number) & " ch" & ChrW(&H1B0) & _
                "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c: " & _
                JoinCollection(problems, "; ") & ".")
        End If
    Next a
    Set CheckArticleFormat = Result
End Function

Public Function CheckClauseFormat(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckClauseFormat")
    Dim snapshot As Object: Set snapshot = context("Snapshot")
    If ArticleParagraphs(snapshot).count = 0 Then
        Set CheckClauseFormat = Nothing
        Exit Function
    End If
    Dim clauses As Collection: Set clauses = ClauseParagraphs(snapshot)
    If clauses.count = 0 Then
        Set CheckClauseFormat = Nothing
        Exit Function
    End If

    Dim Result As New Collection
    Dim c As Variant
    For Each c In clauses
        Dim p As ParagraphSnapshot: Set p = c("Paragraph")
        Dim number As Long: number = CLng(c("Number"))
        If Not IsNull(p.Italic) Then
            If CBool(p.Italic) Then
                Result.Add MakeFindingInput(CLng(p.Index), VnKhoan() & " " & CStr(number) & " ch" & ChrW(&H1B0) & _
                    "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c: " & _
                    ChrW(&H111) & "ang in nghi" & ChrW(&HEA) & "ng, quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh " & _
                    ChrW(&H111) & ChrW(&H1EE9) & "ng.")
            End If
        End If
    Next c
    Set CheckClauseFormat = Result
End Function

' ND30-PL1-M2-K6D-POINT -- full: kieu chu (khong dam, khong nghieng). So thu tu CHAC CHAN hop le
' vi regex nhan dien chi khop dung 29 chu cai tieng Viet; thu tu cac diem VOI NHAU do
' ND30-PL1-M2-K6D-ALPHABET kiem rieng ngay duoi day.
Public Function CheckPointFormat(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckPointFormat")
    Dim snapshot As Object: Set snapshot = context("Snapshot")
    Dim spec As Object: Set spec = context("Spec")
    Dim points As Collection: Set points = PointParagraphs(snapshot, spec("vietnameseAlphabet"))
    If points.count = 0 Then
        Set CheckPointFormat = Nothing
        Exit Function
    End If

    Dim levelSpec As Object: Set levelSpec = spec("structureLevels")("point")
    Dim Result As New Collection
    Dim item As Variant
    For Each item In points
        Dim p As ParagraphSnapshot: Set p = item("Paragraph")
        Dim letter As String: letter = CStr(item("Letter"))
        Dim problems As Collection
        If levelSpec.Exists("style") Then
            Set problems = DescribeStyleViolations(p, levelSpec)
        Else
            Set problems = New Collection
        End If
        If problems.count > 0 Then
            Result.Add MakeFindingInput(CLng(p.Index), VnDiem() & " " & letter & ") ch" & ChrW(&H1B0) & "a " & _
                ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c: " & _
                JoinCollection(problems, "; ") & ".")
        End If
    Next item
    Set CheckPointFormat = Result
End Function

Private Function VietnameseAlphabetIndexOf(ByVal vietnameseAlphabet As Collection, ByVal letter As String) As Long
    Dim i As Long: i = 0
    Dim v As Variant
    For Each v In vietnameseAlphabet
        If CStr(v) = letter Then
            VietnameseAlphabetIndexOf = i
            Exit Function
        End If
        i = i + 1
    Next v
    VietnameseAlphabetIndexOf = -1
End Function

' ND30-PL1-M2-K6D-ALPHABET -- full: thu tu diem trong CUNG mot day phai TANG DAN theo VI TRI trong
' bang chu cai tieng Viet (29 chu, tra tu thong-so-the-thuc.json -- CAM so sanh chuoi mac dinh cua
' nen tang, muc "Rang buoc" ), khong duoc dung yen hay lui lai. CHI doi hoi TANG DAN, khong doi
' hoi lien tuc tuyet doi tung chu mot -- chinh vi du trong ND 30 (docs/rules/01 muc 6.3: Day diem
' RESET ve dau moi khi gap mot Dieu/Khoan moi xen giua. Danh sach cac diem bi bo sot giua hai chi
' so, dang "'c)', 'd)'"
Private Function MissingPointLetters(ByVal alphabet As Collection, ByVal fromIndex As Long, _
        ByVal toIndex As Long) As String
    Dim parts As String
    Dim i As Long
    For i = fromIndex + 1 To toIndex - 1
        ' Collection cua VBA la 1-based, chi so alphabet la 0-based.
        If Len(parts) > 0 Then parts = parts & ", "
        parts = parts & "'" & CStr(alphabet(i + 1)) & ")'"
    Next i
    MissingPointLetters = parts
End Function

Public Function CheckPointAlphabetOrder(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckPointAlphabetOrder")
    Dim snapshot As Object: Set snapshot = context("Snapshot")
    Dim spec As Object: Set spec = context("Spec")
    Dim vietnameseAlphabet As Collection: Set vietnameseAlphabet = spec("vietnameseAlphabet")
    ' Day DAT TEN DIEM cua Nghi dinh (26 chu, bo a-breve/a-circumflex/e-circumflex) - KHAC
    ' vietnameseAlphabet (29 chu, dung cho muc dich khac). Dung nham bang 29 chu thi chinh day
    ' trong Nghi dinh (a) roi b)) bi coi la "bo sot" - xem muc A16 cua cau-hoi-con-mo.md.
    Dim pointOrderAlphabet As Collection: Set pointOrderAlphabet = spec("pointOrderAlphabet")
    Dim points As Collection: Set points = PointParagraphs(snapshot, vietnameseAlphabet)
    If points.count = 0 Then
        Set CheckPointAlphabetOrder = Nothing
        Exit Function
    End If

    Dim structureBreakIndexes As New Collection
    Dim a As Variant
    For Each a In ArticleParagraphs(snapshot)
        Dim ap As ParagraphSnapshot: Set ap = a("Paragraph")
        structureBreakIndexes.Add CLng(ap.Index)
    Next a
    Dim cl As Variant
    For Each cl In ClauseParagraphs(snapshot)
        Dim cp As ParagraphSnapshot: Set cp = cl("Paragraph")
        structureBreakIndexes.Add CLng(cp.Index)
    Next cl

    Dim Result As New Collection
    Dim previousIndex As Variant: previousIndex = Null
    Dim previousLetter As String: previousLetter = ""
    Dim lastPointIndex As Long: lastPointIndex = -1
    Dim item As Variant
    For Each item In points
        Dim p As ParagraphSnapshot: Set p = item("Paragraph")
        Dim letter As String: letter = CStr(item("Letter"))

        Dim hasBreakBetween As Boolean: hasBreakBetween = False
        Dim bi As Variant
        For Each bi In structureBreakIndexes
            If CLng(bi) > lastPointIndex And CLng(bi) < p.Index Then
                hasBreakBetween = True
                Exit For
            End If
        Next bi

        Dim alphabetIndex As Long
        alphabetIndex = VietnameseAlphabetIndexOf(pointOrderAlphabet, Utils.ToLowerVn(letter))

        If Not hasBreakBetween Then
            If Not IsNull(previousIndex) Then
                If alphabetIndex <= CLng(previousIndex) Then
                    Result.Add MakeFindingInput(CLng(p.Index), "Ph" & ChrW(&HE1) & "t hi" & ChrW(&H1EC7) & "n " & _
                        VnDiem() & " '" & letter & ")'. B" & ChrW(&H1EA3) & "ng ch" & ChrW(&H1EEF) & " c" & _
                        ChrW(&HE1) & "i ti" & ChrW(&H1EBF) & "ng Vi" & ChrW(&H1EC7) & "t kh" & ChrW(&HF4) & _
                        "ng c" & ChrW(&HF3) & " ch" & ChrW(&H1EEF) & " f, j, w, z v" & ChrW(&HE0) & " c" & _
                        ChrW(&HF3) & " ch" & ChrW(&H1EEF) & " " & ChrW(&H111) & "; sau e l" & ChrW(&HE0) & " g.")
                ElseIf alphabetIndex > CLng(previousIndex) + 1 Then
                    Result.Add MakeFindingInput(CLng(p.Index), "Sau " & VnDiem() & " '" & previousLetter & _
                        ")' l" & ChrW(&HE0) & " " & VnDiem() & " '" & letter & ")', b" & ChrW(&H1ECF) & _
                        " s" & ChrW(&HF3) & "t " & MissingPointLetters(pointOrderAlphabet, CLng(previousIndex), alphabetIndex) & ".")
                End If
            End If
        End If
        previousIndex = alphabetIndex
        previousLetter = Utils.ToLowerVn(letter)
        lastPointIndex = p.Index
    Next item
    Set CheckPointAlphabetOrder = Result
End Function

' ----------------------------------------------------------------------------
' Nhom citation -- vien dan van ban (shared/rules/vien-dan-van-ban.json qua
' RuleLoader.GetCitationRules).
' ----------------------------------------------------------------------------

Private Function VnLuat() As String
    VnLuat = "Lu" & ChrW(&H1EAD) & "t"
End Function

Private Function VnBoLuat() As String
    VnBoLuat = "B" & ChrW(&H1ED9) & " lu" & ChrW(&H1EAD) & "t"
End Function

Private Function VnPhapLenh() As String
    VnPhapLenh = "Ph" & ChrW(&HE1) & "p l" & ChrW(&H1EC7) & "nh"
End Function

Private Function CitationTypeNamesSortedByLength() As Collection
    Dim citationRules As Object: Set citationRules = RuleLoader.GetCitationRules()
    Dim excluded As New Collection
    excluded.Add VnLuat()
    excluded.Add VnBoLuat()
    excluded.Add VnPhapLenh()

    Dim all As New Collection
    Dim v As Variant
    For Each v In citationRules("documentTypes")("legislative")
        all.Add CStr(v)
    Next v
    For Each v In citationRules("documentTypes")("administrative")
        all.Add CStr(v)
    Next v

    Dim filtered As New Collection
    Dim t As Variant
    For Each t In all
        Dim isExcluded As Boolean: isExcluded = False
        Dim e As Variant
        For Each e In excluded
            If CStr(e) = CStr(t) Then isExcluded = True
        Next e
        If Not isExcluded Then filtered.Add CStr(t)
    Next t

    ' Sap XA den NGAN -- bubble sort don gian, danh muc ten loai chi khoang 30 phan tu.
    Dim n As Long: n = filtered.count
    Dim arr() As String
    ReDim arr(1 To n)
    Dim i As Long
    For i = 1 To n
        arr(i) = CStr(filtered(i))
    Next i
    Dim j As Long, tmp As String
    For i = 1 To n - 1
        For j = 1 To n - i
            If Len(arr(j)) < Len(arr(j + 1)) Then
                tmp = arr(j)
                arr(j) = arr(j + 1)
                arr(j + 1) = tmp
            End If
        Next j
    Next i

    Dim Result As New Collection
    For i = 1 To n
        Result.Add arr(i)
    Next i
    Set CitationTypeNamesSortedByLength = Result
End Function

' Khop "<ten loai> [so] <so/ky hieu>" -- nhom bat thu hai ("so\s+") CO MAT khi da dung, VANG MAT
' khi thieu chu "so" (ND30-PL1-M2-K6B-SO do dung truong hop vang mat nay). Khong dung named
' group/lookbehind -- VBScript.RegExp khong ho tro, giu regex don gian.
Private Function CitationTypeAndCodePattern(ByVal typeNames As Collection) As String
    Dim alternation As String: alternation = ""
    Dim v As Variant
    Dim isFirst As Boolean: isFirst = True
    For Each v In typeNames
        If Not isFirst Then alternation = alternation & "|"
        alternation = alternation & EscapeRegExpLiteral(CStr(v))
        isFirst = False
    Next v
    CitationTypeAndCodePattern = "(" & alternation & ")\s+(s" & ChrW(&H1ED1) & "\s+)?(\d+[\/-][A-Za-z" & _
        ChrW(&HC0) & "-" & ChrW(&H1EF9) & ChrW(&H110) & ChrW(&H111) & "0-9\/-]+)"
End Function

' Doan CO THE chua vien dan -- loai bang bieu (FR-CIT-05) va loai doan dang mang vai tro "so, ky
' hieu cua chinh van ban dang soan" (o so 3 -- khong phai vien dan; tren thuc te mau "So: ..." da
' tu nhien khong khop CitationTypeAndCodePattern vi thieu ten loai dung ngay truoc, loc them o day
' lam LUOI AN TOAN thu hai, dung "Rang buoc" ).
Private Function CitationScanParagraphs(ByVal snapshot As Object, ByVal layoutMap As Object) As Collection
    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In snapshot("Paragraphs")
        If Not p.isInTable Then
            Dim isCodeNumberRole As Boolean: isCodeNumberRole = False
            If layoutMap.Exists(p.Index) Then
                If CStr(layoutMap(p.Index)) = "codeNumberNotation" Then isCodeNumberRole = True
            End If
            If Not isCodeNumberRole Then Result.Add p
        End If
    Next p
    Set CitationScanParagraphs = Result
End Function

' ND30-PL1-M2-K6B-SO -- full: thieu chu "so" sau ten loai van ban khi vien dan.
Public Function CheckCitationMissingSoKeyword(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckCitationMissingSoKeyword")
    Dim snapshot As Object: Set snapshot = context("Snapshot")
    Dim layoutMap As Object: Set layoutMap = context("LayoutMap")
    Dim paragraphs As Collection: Set paragraphs = CitationScanParagraphs(snapshot, layoutMap)
    If paragraphs.count = 0 Then
        Set CheckCitationMissingSoKeyword = Nothing
        Exit Function
    End If

    Dim pattern As String: pattern = CitationTypeAndCodePattern(CitationTypeNamesSortedByLength())
    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        Dim matches As Object: Set matches = RegexAllMatches(pattern, p.text)
        Dim m As Object
        For Each m In matches
            Dim typeName As String: typeName = CStr(m.SubMatches(0))
            Dim soKeyword As String: soKeyword = CStr(m.SubMatches(1))
            Dim code As String: code = CStr(m.SubMatches(2))
            If Trim$(soKeyword) = "" Then
                Dim msg As String
                msg = "'" & typeName & " " & code & "' ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t l" & _
                    ChrW(&HE0) & " '" & typeName & " s" & ChrW(&H1ED1) & " " & code & "'. Sau t" & ChrW(&HEA) & _
                    "n lo" & ChrW(&H1EA1) & "i v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n ph" & ChrW(&H1EA3) & _
                    "i c" & ChrW(&HF3) & " ch" & ChrW(&H1EEF) & " 's" & ChrW(&H1ED1) & "'."
                Result.Add MakeFindingInput(CLng(p.Index), msg, , CLng(m.firstIndex))
            End If
        Next m
    Next p
    Set CheckCitationMissingSoKeyword = Result
End Function

' ND30-PL1-M2-K6B-DATE -- full: ngay thang khi vien dan viet tat dang dd/mm/yyyy thay vi day du
' "ngay... thang... nam...". Quy tac nay CHI ap dung cho doan mang vai tro legalBasis ("Can
' cu..."); vien dan ngay thang o than van ban (ngoai "Can cu") duoc phep viet dang dd/mm/yyyy.
' Nguong them so 0 lay tu vien-dan-van-ban.json (padDayBelow/padMonthsList) - khong hard-code.
Public Function CheckCitationAbbreviatedDate(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckCitationAbbreviatedDate")
    Dim snapshot As Object: Set snapshot = context("Snapshot")
    Dim layoutMap As Object: Set layoutMap = context("LayoutMap")

    Dim paragraphs As Collection: Set paragraphs = ParagraphsWithRole(snapshot, layoutMap, "legalBasis")
    Dim p As ParagraphSnapshot
    If paragraphs.count = 0 Then
        Set CheckCitationAbbreviatedDate = Nothing
        Exit Function
    End If

    Dim citationRules As Object: Set citationRules = RuleLoader.GetCitationRules()
    Dim padDayBelow As Long: padDayBelow = 10
    Dim padMonths As New Collection
    padMonths.Add 1
    padMonths.Add 2
    Dim r As Variant
    For Each r In citationRules("rules")
        If CStr(r("name")) = "requireFullDateForm" Then
            If r.Exists("padDayBelow") Then padDayBelow = CLng(r("padDayBelow"))
            If r.Exists("padMonthsList") Then
                Set padMonths = New Collection
                Dim pm As Variant
                For Each pm In r("padMonthsList")
                    padMonths.Add CLng(pm)
                Next pm
            End If
            Exit For
        End If
    Next r

    Dim pattern As String
    pattern = "ng" & ChrW(&HE0) & "y\s+(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})"

    Dim Result As New Collection
    For Each p In paragraphs
        Dim matches As Object: Set matches = RegexAllMatches(pattern, p.text)
        Dim m As Object
        For Each m In matches
            Dim dayStr As String: dayStr = CStr(m.SubMatches(0))
            Dim monthStr As String: monthStr = CStr(m.SubMatches(1))
            Dim yearStr As String: yearStr = CStr(m.SubMatches(2))

            Dim dayPad As String: dayPad = dayStr
            If CLng(dayStr) < padDayBelow Then dayPad = PadLeftZero(dayStr, 2)

            Dim monthInList As Boolean: monthInList = False
            Dim pmv As Variant
            For Each pmv In padMonths
                If CLng(pmv) = CLng(monthStr) Then monthInList = True
            Next pmv
            Dim monthPad As String: monthPad = monthStr
            If monthInList Then monthPad = PadLeftZero(monthStr, 2)

            Dim msg As String
            msg = "'ng" & ChrW(&HE0) & "y " & dayStr & "/" & monthStr & "/" & yearStr & "' ph" & ChrW(&H1EA3) & _
                "i vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " 'ng" & ChrW(&HE0) & "y " & dayPad & " th" & _
                ChrW(&HE1) & "ng " & monthPad & " n" & ChrW(&H103) & "m " & yearStr & "'. Th" & ChrW(&H1EDD) & _
                "i gian ban h" & ChrW(&HE0) & "nh ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t " & ChrW(&H111) & _
                ChrW(&H1EA7) & "y " & ChrW(&H111) & ChrW(&H1EE7) & " d" & ChrW(&H1EA1) & "ng 'ng" & ChrW(&HE0) & _
                "y ... th" & ChrW(&HE1) & "ng ... n" & ChrW(&H103) & "m ...'."
            Result.Add MakeFindingInput(CLng(p.Index), msg, , CLng(m.firstIndex))
        Next m
    Next p
    Set CheckCitationAbbreviatedDate = Result
End Function

Private Function CitationHasDate(ByVal text As String) As Boolean
    If RegexTestCC("ng" & ChrW(&HE0) & "y\s+\d{1,2}[^" & vbLf & "]{0,20}n" & _
            ChrW(&H103) & "m\s+\d{4}", text) Then
        CitationHasDate = True
        Exit Function
    End If
    CitationHasDate = RegexTestCC("ng" & ChrW(&HE0) & "y\s+\d{1,2}[/.\-]\d{1,2}[/.\-]\d{4}", text)
End Function

Private Function CitationHasOrgan(ByVal text As String) As Boolean
    CitationHasOrgan = RegexTestCC("\bc" & ChrW(&H1EE7) & "a\s+\S", text)
End Function

Public Function CheckCitationFullFirstCitation(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckCitationFullFirstCitation")
    Dim snapshot As Object: Set snapshot = context("Snapshot")
    Dim layoutMap As Object: Set layoutMap = context("LayoutMap")
    Dim paragraphs As Collection: Set paragraphs = CitationScanParagraphs(snapshot, layoutMap)
    If paragraphs.count = 0 Then
        Set CheckCitationFullFirstCitation = Nothing
        Exit Function
    End If

    Dim pattern As String: pattern = CitationTypeAndCodePattern(CitationTypeNamesSortedByLength())
    Dim seenCodes As Object: Set seenCodes = Utils.NewDictionary()
    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        Dim matches As Object: Set matches = RegexAllMatches(pattern, p.text)
        Dim m As Object
        For Each m In matches
            Dim typeName As String: typeName = CStr(m.SubMatches(0))
            Dim codeRaw As String: codeRaw = CStr(m.SubMatches(2))
            Dim code As String: code = RegexReplaceFirst("[.,;:]+$", codeRaw)

            If Not seenCodes.Exists(code) Then
                seenCodes.Add code, True
                Dim missing As New Collection
                If Not CitationHasDate(p.text) Then
                    missing.Add "th" & ChrW(&H1EDD) & "i gian ban h" & ChrW(&HE0) & "nh"
                End If
                If Not CitationHasOrgan(p.text) Then
                    missing.Add "t" & ChrW(&HEA) & "n c" & ChrW(&H1A1) & " quan ban h" & ChrW(&HE0) & _
                        "nh ('c" & ChrW(&H1EE7) & "a ...')"
                End If
                If missing.count > 0 Then
                    Dim msg As String
                    msg = "'" & typeName & " " & codeRaw & "' vi" & ChrW(&H1EC7) & "n d" & ChrW(&H1EAB) & "n l" & _
                        ChrW(&H1EA7) & "n " & ChrW(&H111) & ChrW(&H1EA7) & "u c" & ChrW(&HF2) & "n thi" & _
                        ChrW(&H1EBF) & "u " & JoinCollection(missing, ", ") & ". Khi vi" & ChrW(&H1EC7) & "n d" & _
                        ChrW(&H1EAB) & "n l" & ChrW(&H1EA7) & "n " & ChrW(&H111) & ChrW(&H1EA7) & "u ph" & _
                        ChrW(&H1EA3) & "i ghi " & ChrW(&H111) & ChrW(&H1EA7) & "y " & ChrW(&H111) & ChrW(&H1EE7) & _
                        " t" & ChrW(&HEA) & "n lo" & ChrW(&H1EA1) & "i, s" & ChrW(&H1ED1) & ", k" & ChrW(&HFD) & _
                        " hi" & ChrW(&H1EC7) & "u, th" & ChrW(&H1EDD) & "i gian ban h" & ChrW(&HE0) & "nh, t" & _
                        ChrW(&HEA) & "n c" & ChrW(&H1A1) & " quan ban h" & ChrW(&HE0) & "nh v" & ChrW(&HE0) & _
                        " tr" & ChrW(&HED) & "ch y" & ChrW(&H1EBF) & "u n" & ChrW(&H1ED9) & "i dung."
                    Result.Add MakeFindingInput(CLng(p.Index), msg, , CLng(m.firstIndex))
                End If
            End If
        Next m
    Next p
    Set CheckCitationFullFirstCitation = Result
End Function

' ----------------------------------------------------------------------------
' Dang ky -- goi tu EnsureRegistryInitialized, GHI THANG vao mRegistry, dung 10 ma tren.
' ----------------------------------------------------------------------------

Private Sub RegisterStructureCitationChecks()
    mRegistry("ND30-PL1-M2-K6D-TITLE") = "ComplianceChecker.CheckStructureTitlePresence"
    mRegistry("ND30-PL1-M2-K6D-ARTICLE") = "ComplianceChecker.CheckArticleFormat"
    mRegistry("ND30-PL1-M2-K6D-CLAUSE") = "ComplianceChecker.CheckClauseFormat"
    mRegistry("ND30-PL1-M2-K6D-POINT") = "ComplianceChecker.CheckPointFormat"
    mRegistry("ND30-PL1-M2-K6D-ALPHABET") = "ComplianceChecker.CheckPointAlphabetOrder"

    mRegistry("ND30-PL1-M2-K6B-SO") = "ComplianceChecker.CheckCitationMissingSoKeyword"
    mRegistry("ND30-PL1-M2-K6B-DATE") = "ComplianceChecker.CheckCitationAbbreviatedDate"
    mRegistry("ND30-PL1-M2-K6B-CITE") = "ComplianceChecker.CheckCitationFullFirstCitation"
End Sub

' ----------------------------------------------------------------------------
' Tien ich dung chung PHAN 5
' ----------------------------------------------------------------------------

' Dai &H00C0-&H1EF9 gom ca chu hoa/thuong tieng Viet co dau (Latin Extended-A/B), cung quy uoc da
' dung o CitationTypeAndCodePattern (PHAN 4).
Private Function IsLetterChar(ByVal ch As String) As Boolean
    If ch = "" Then
        IsLetterChar = False
        Exit Function
    End If
    Dim code As Long: code = AscW(ch)
    If (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Then
        IsLetterChar = True
    ElseIf code >= &HC0 And code <= &H1EF9 Then
        IsLetterChar = True
    Else
        IsLetterChar = False
    End If
End Function

Private Function IsWordOrDigitChar(ByVal ch As String) As Boolean
    If ch = "" Then
        IsWordOrDigitChar = False
    ElseIf IsLetterChar(ch) Then
        IsWordOrDigitChar = True
    Else
        Dim code As Long: code = AscW(ch)
        IsWordOrDigitChar = (code >= 48 And code <= 57)
    End If
End Function

' Phan biet hoa/thuong bang UCase$/LCase$ (Utils.ToUpperVn/ToLowerVn) -- KHONG dung \p{Lu}/ \p{Ll}
' cua regex (VBScript.RegExp khong ho tro).
Private Function IsLowercaseLetterChar(ByVal ch As String) As Boolean
    IsLowercaseLetterChar = IsLetterChar(ch) And (Utils.ToUpperVn(ch) <> ch)
End Function

Private Function IsUppercaseLetterChar(ByVal ch As String) As Boolean
    IsUppercaseLetterChar = IsLetterChar(ch) And (Utils.ToLowerVn(ch) <> ch)
End Function

' Tim moi vi tri khop NGUYEN CUM phrase trong text, khong phan biet hoa/thuong, bien an toan cho
' tieng Viet (xem IsWordOrDigitChar) -- doi chieu findWholePhraseCaseInsensitive ben TS. Tra
' Collection cua Dictionary {"Index" (Long, 0-based, dong bo CharOffset), "Actual" (String, chuoi
' khop THAT trong van ban goc, giu nguyen hoa/thuong)}.
Private Function FindWholePhraseCI(ByVal text As String, ByVal phrase As String) As Collection
    Dim Result As New Collection
    If Len(phrase) = 0 Then
        Set FindWholePhraseCI = Result
        Exit Function
    End If
    Dim lowerText As String: lowerText = Utils.ToLowerVn(text)
    Dim lowerPhrase As String: lowerPhrase = Utils.ToLowerVn(phrase)
    Dim fromPos As Long: fromPos = 1
    Dim foundPos As Long
    Do
        foundPos = InStr(fromPos, lowerText, lowerPhrase, vbBinaryCompare)
        If foundPos = 0 Then Exit Do
        Dim beforeCh As String: beforeCh = ""
        If foundPos > 1 Then beforeCh = Mid$(text, foundPos - 1, 1)
        Dim afterPos As Long: afterPos = foundPos + Len(phrase)
        Dim afterCh As String: afterCh = ""
        If afterPos <= Len(text) Then afterCh = Mid$(text, afterPos, 1)
        If Not IsWordOrDigitChar(beforeCh) And Not IsWordOrDigitChar(afterCh) Then
            Dim item As Object: Set item = Utils.NewDictionary()
            item.Add "Index", foundPos - 1
            item.Add "Actual", Mid$(text, foundPos, Len(phrase))
            Result.Add item
        End If
        fromPos = foundPos + 1
    Loop
    Set FindWholePhraseCI = Result
End Function

Private Function CapitalizeFirstLetterVn(ByVal word As String) As String
    If Len(word) = 0 Then
        CapitalizeFirstLetterVn = word
    Else
        CapitalizeFirstLetterVn = Utils.ToUpperVn(left$(word, 1)) & Mid$(word, 2)
    End If
End Function

' Viet hoa chu cai dau MOI tu (cach nhau boi dung MOT dau cach) -- dung cho goi y sua dia danh
' (K4-CASE) va don vi hanh chinh co chu so (M3-K1B).
Private Function CapitalizeEachWordVn(ByVal text As String) As String
    Dim words() As String: words = Split(text, " ")
    Dim i As Long
    For i = LBound(words) To UBound(words)
        words(i) = CapitalizeFirstLetterVn(words(i))
    Next i
    CapitalizeEachWordVn = Join(words, " ")
End Function

Private Function FreeTextSkipIndexes(ByVal context As Object, Optional ByVal extraRolesCsv As String = "") As Object
    Dim layoutMap As Object: Set layoutMap = context("LayoutMap")
    Dim typoDict As Object: Set typoDict = RuleLoader.GetTypoDictionary()
    Dim configuredRoles As Object: Set configuredRoles = typoDict("skipContexts")("componentRoles")

    Dim skipRoles As Object: Set skipRoles = Utils.NewDictionary()
    Dim skipSigner As Boolean: skipSigner = False
    Dim r As Variant
    For Each r In configuredRoles
        Dim roleName As String: roleName = CStr(r)
        If roleName = "signerName" Then
            skipSigner = True
        ElseIf roleName = "placeName" Then
            skipRoles("placeAndIssuedDate") = True
        Else
            skipRoles(roleName) = True
        End If
    Next r
    If Len(extraRolesCsv) > 0 Then
        Dim extras() As String: extras = Split(extraRolesCsv, ",")
        Dim i As Long
        For i = LBound(extras) To UBound(extras)
            skipRoles(extras(i)) = True
        Next i
    End If

    Dim skip As Object: Set skip = Utils.NewDictionary()
    Dim key As Variant
    For Each key In layoutMap.Keys
        If skipRoles.Exists(CStr(layoutMap(key))) Then skip(CLng(key)) = True
    Next key
    If skipSigner Then
        Dim p As ParagraphSnapshot
        For Each p In SignerNameCandidates(context)
            skip(p.Index) = True
        Next p
    End If
    Set FreeTextSkipIndexes = skip
End Function

' Doan CO THE ra quet chinh ta/viet hoa toan van ban -- loai bang bieu VA cac vai tro the thuc co
' noi dung co dinh. Collection rong KHAC Nothing -- ham goi tu quyet dinh Nothing truoc.
Private Function ScannableParagraphs(ByVal context As Object) As Collection
    Dim skip As Object: Set skip = FreeTextSkipIndexes(context)
    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In context("Snapshot")("Paragraphs")
        If Not p.isInTable Then
            If Not skip.Exists(p.Index) Then Result.Add p
        End If
    Next p
    Set ScannableParagraphs = Result
End Function

' Doan CO THE ra QD1989-IY-MIX -- giong ScannableParagraphs nhung loai them vai tro
' codeNumberNotation (i-y-hai-kieu.json/excludeAbsolute.componentRoles.roles co them ma nay).
Private Function IyScannableParagraphs(ByVal context As Object) As Collection
    Dim skip As Object: Set skip = FreeTextSkipIndexes(context, "codeNumberNotation")
    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In context("Snapshot")("Paragraphs")
        If Not p.isInTable Then
            If Not skip.Exists(p.Index) Then Result.Add p
        End If
    Next p
    Set IyScannableParagraphs = Result
End Function

' ----------------------------------------------------------------------------
' Nhom capitalization -- Phu luc II.
' ----------------------------------------------------------------------------

' ND30-PL1-M2-K4-CASE -- full: chu cai dau dia danh (phan truoc dau phay cua o 4) phai viet hoa
' tung am tiet.
Public Function CheckPlaceNameLetterCase(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckPlaceNameLetterCase")
    Dim snapshot As Object: Set snapshot = context("Snapshot")
    Dim layoutMap As Object: Set layoutMap = context("LayoutMap")
    Dim paragraphs As Collection: Set paragraphs = ParagraphsWithRole(snapshot, layoutMap, "placeAndIssuedDate")
    If paragraphs.count = 0 Then
        Set CheckPlaceNameLetterCase = Nothing
        Exit Function
    End If

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        Dim commaPos As Long: commaPos = InStr(p.text, ",")
        Dim placePart As String
        If commaPos > 0 Then
            placePart = Trim$(left$(p.text, commaPos - 1))
        Else
            placePart = Trim$(p.text)
        End If
        If Len(placePart) > 0 Then
            Dim hasLower As Boolean: hasLower = False
            Dim words() As String: words = Split(placePart, " ")
            Dim i As Long
            For i = LBound(words) To UBound(words)
                If Len(words(i)) > 0 Then
                    If IsLowercaseLetterChar(left$(words(i), 1)) Then hasLower = True
                End If
            Next i
            If hasLower Then
                Result.Add MakeFindingInput(CLng(p.Index), "'" & placePart & "' ch" & ChrW(&H1B0) & "a vi" & ChrW(&H1EBF) & "t hoa " & ChrW(&H111) & ChrW(&HFA) & "ng. " & _
                    "C" & ChrW(&HE1) & "c ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u c" & ChrW(&H1EE7) & "a " & ChrW(&H111) & ChrW(&H1ECB) & "a danh ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t hoa.", , , placePart, CapitalizeEachWordVn(placePart))
            End If
        End If
    Next p
    Set CheckPlaceNameLetterCase = Result
End Function

Private Function NonSentenceEndingAbbreviationForms() As Collection
    Dim Result As New Collection
    Dim entries As Object: Set entries = RuleLoader.GetNonSentenceEndingAbbreviations()("entries")
    Dim e As Variant
    For Each e In entries
        Result.Add CStr(e("abbreviation"))
    Next e
    Set NonSentenceEndingAbbreviationForms = Result
End Function

' ND30-PL2-M1 -- full: viet hoa dau cau. Dieu kien an toan (docs/rules/02 muc I): dau ket cau theo
' sau boi DUNG MOT dau cach, ky tu tiep theo la CHU CAI THUONG (so thap phan/ngay thang dang so tu
' dong loai vi ky tu sau khong phai chu cai). Loai tru khi tu ngay truoc dau cham la mot chu viet
' tat khong ket cau (TM., KT., TL., TUQ., Q., v.v., TP....).
Public Function CheckSentenceCapitalization(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckSentenceCapitalization")
    Dim paragraphs As Collection: Set paragraphs = ScannableParagraphs(context)
    If paragraphs.count = 0 Then
        Set CheckSentenceCapitalization = Nothing
        Exit Function
    End If
    Dim abbreviations As Collection: Set abbreviations = NonSentenceEndingAbbreviationForms()

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        Dim text As String: text = p.text
        Dim i As Long
        For i = 1 To Len(text) - 2
            Dim c0 As String: c0 = Mid$(text, i, 1)
            If c0 = "." Or c0 = "!" Or c0 = "?" Then
                If Mid$(text, i + 1, 1) = " " Then
                    Dim letterCh As String: letterCh = Mid$(text, i + 2, 1)
                    If IsLowercaseLetterChar(letterCh) Then
                        Dim beforeWord As String: beforeWord = LastWordEndingAt(text, i)
                        Dim isAbbrev As Boolean: isAbbrev = False
                        Dim a As Variant
                        For Each a In abbreviations
                            If EndsWithVn(beforeWord, CStr(a)) Then isAbbrev = True
                        Next a
                        If Not isAbbrev Then
                            Result.Add MakeFindingInput(CLng(p.Index), "'" & letterCh & "' ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t hoa. " & _
                                "Vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u " & ChrW(&HE2) & "m ti" & ChrW(&H1EBF) & "t th" & ChrW(&H1EE9) & " nh" & ChrW(&H1EA5) & "t c" & ChrW(&H1EE7) & "a m" & ChrW(&H1ED9) & "t c" & ChrW(&HE2) & "u ho" & ChrW(&HE0) & "n ch" & ChrW(&H1EC9) & "nh.", , CLng(i + 1), _
                                letterCh, CapitalizeFirstLetterVn(letterCh))
                        End If
                    End If
                End If
            End If
        Next i
    Next p
    Set CheckSentenceCapitalization = Result
End Function

' Tu (chuoi khong khoang trang) KET THUC tai vi tri endPos (bao gom endPos) trong text.
Private Function LastWordEndingAt(ByVal text As String, ByVal endPos As Long) As String
    Dim startPos As Long: startPos = endPos
    Do While startPos > 1
        If Mid$(text, startPos - 1, 1) = " " Then Exit Do
        startPos = startPos - 1
    Loop
    LastWordEndingAt = Mid$(text, startPos, endPos - startPos + 1)
End Function

Private Function EndsWithVn(ByVal s As String, ByVal suffix As String) As Boolean
    If Len(suffix) = 0 Or Len(suffix) > Len(s) Then
        EndsWithVn = (Len(suffix) = 0)
        Exit Function
    End If
    EndsWithVn = (Right$(s, Len(suffix)) = suffix)
End Function

' True neu chuoi "words" (mang tach boi dau cach) co >=2 tu bat dau bang chu cai va IT NHAT mot tu
' trong so do bat dau bang chu THUONG -- dau hieu ten rieng chua viet hoa dung.
Private Function HasUncapitalizedNameWord(words() As String) As Boolean
    Dim letterWordCount As Long: letterWordCount = 0
    Dim hasLower As Boolean: hasLower = False
    Dim i As Long
    For i = LBound(words) To UBound(words)
        If Len(words(i)) > 0 Then
            If IsLetterChar(left$(words(i), 1)) Then
                letterWordCount = letterWordCount + 1
                If IsLowercaseLetterChar(left$(words(i), 1)) Then hasLower = True
            End If
        End If
    Next i
    HasUncapitalizedNameWord = (letterWordCount >= 2 And hasLower)
End Function

Private Function PersonNameCapitalizationMsg(ByVal nameText As String) As String
    PersonNameCapitalizationMsg = "'" & nameText & "' " & ChrW(&H2014) & " vi" & ChrW(&H1EBF) & "t hoa ch" & _
        ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u t" & ChrW(&H1EA5) & "t c" & _
        ChrW(&H1EA3) & " c" & ChrW(&HE1) & "c " & ChrW(&HE2) & "m ti" & ChrW(&H1EBF) & "t c" & ChrW(&H1EE7) & _
        "a danh t" & ChrW(&H1EEB) & " ri" & ChrW(&HEA) & "ng ch" & ChrW(&H1EC9) & " t" & ChrW(&HEA) & "n ng" & _
        ChrW(&H1B0) & ChrW(&H1EDD) & "i."
End Function

' Vi tri (0-based) ky tu dau tien SAU mot cum "Ong"/"Ba" (kem bien the "Ong (ba)") gioi thieu ten
' nguoi -- mau nay dung san trong Mau 1.7/1.8/1.10 Phu luc III NÄ� 30 (Giay moi/Giay gioi
' thieu/Giay nghi phep: "Ong (ba)..."), nen day la dau hieu AN TOAN de mo rong quet ten nguoi ra
' CA than van ban ma KHONG phai doan van hoa toan bo cau (tranh bao sai tran lan tren van xuoi
' thuong). Tra Collection cac Long (0-based) trong PHAM VI mot doan.
Private Function HonorificNameStartPositions(ByVal text As String) As Collection
    Dim Result As New Collection
    Dim pattern As String
    pattern = "(" & ChrW(&HD4) & "ng|B" & ChrW(&HE0) & ")(\s*\(\s*(" & ChrW(&HF4) & "ng|b" & ChrW(&HE0) & _
        ")\s*\))?\s*:?\s+"
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.pattern = pattern
    re.IgnoreCase = False
    re.Global = True
    Dim mc As Object: Set mc = re.Execute(text)
    Dim m As Object
    For Each m In mc
        Result.Add CLng(m.firstIndex + m.length)
    Next m
    Set HonorificNameStartPositions = Result
End Function

' Lay toi da 6 "tu" lien tiep (chi chu cai, cach nhau dung mot dau cach) bat dau tu vi tri 0-based
' startPos0 -- dung gian day la ten nguoi, dung khi gap dau cham/phay/xuong dong/so. Dung THEM
' ngay khi mot tu co KIEU CHU (hoa/thuong o ky tu dau) khac tu dau tien da lay -- vi du "Nguyen
' Van A cong tac..." dung dung sau "A" (tranh gom nham cau van phia sau vao ten nguoi, gay bao sai
' tren ten DA viet hoa dung).
Private Function ExtractLetterWordRun(ByVal text As String, ByVal startPos0 As Long) As String
    Dim i As Long: i = startPos0 + 1
    Dim Result As String: Result = ""
    Dim wordCount As Long: wordCount = 0
    Dim curWord As String: curWord = ""
    Dim firstIsLower As Boolean, firstIsLowerSet As Boolean: firstIsLowerSet = False
    Do While i <= Len(text) And wordCount < 6
        Dim ch As String: ch = Mid$(text, i, 1)
        If IsLetterChar(ch) Then
            curWord = curWord & ch
        ElseIf ch = " " Then
            If Len(curWord) = 0 Then Exit Do
            Dim wordIsLower As Boolean: wordIsLower = IsLowercaseLetterChar(left$(curWord, 1))
            If Not firstIsLowerSet Then
                firstIsLower = wordIsLower
                firstIsLowerSet = True
            ElseIf wordIsLower <> firstIsLower Then
                Exit Do
            End If
            If Len(Result) > 0 Then Result = Result & " "
            Result = Result & curWord
            wordCount = wordCount + 1
            curWord = ""
        Else
            Exit Do
        End If
        i = i + 1
    Loop
    If Len(curWord) > 0 Then
        Dim lastIsLower As Boolean: lastIsLower = IsLowercaseLetterChar(left$(curWord, 1))
        If Not firstIsLowerSet Or lastIsLower = firstIsLower Then
            If Len(Result) > 0 Then Result = Result & " "
            Result = Result & curWord
        End If
    End If
    ExtractLetterWordRun = Result
End Function

' ND30-PL2-M2-K1 -- warnOnly. Hai nguon ung vien ten nguoi: (1) doan "ho ten nguoi ky"
' (SignerNameCandidates, do tin cay cao vi vi tri co dinh sau khoi quyen han/chuc vu) -- kiem tra
' CA doan. (2) ten nguoi duoc gioi thieu qua "Ong"/"Ba" (Mau 1.7/1.8/1.10 NÄ� 30) o BAT KY dau
' trong than van ban -- CHI kiem tra cum tu ngay sau, khong ra ca cau/doan van, tranh bao sai tren
' van xuoi thuong (moi cau tieng Viet thuong chi viet hoa chu dau, khong phai ten rieng).
Public Function CheckPersonNameCapitalizationWarn(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckPersonNameCapitalizationWarn")
    Dim Result As New Collection

    Dim p As ParagraphSnapshot
    For Each p In SignerNameCandidates(context)
        Dim words() As String: words = Split(Trim$(p.text), " ")
        If HasUncapitalizedNameWord(words) Then
            Result.Add MakeFindingInput(CLng(p.Index), PersonNameCapitalizationMsg(Trim$(p.text)))
        End If
    Next p

    For Each p In ScannableParagraphs(context)
        Dim starts As Collection: Set starts = HonorificNameStartPositions(p.text)
        Dim startPos As Variant
        For Each startPos In starts
            Dim nameSpan As String: nameSpan = ExtractLetterWordRun(p.text, CLng(startPos))
            If Len(nameSpan) > 0 Then
                Dim nameWords() As String: nameWords = Split(nameSpan, " ")
                If HasUncapitalizedNameWord(nameWords) Then
                    Result.Add MakeFindingInput(CLng(p.Index), PersonNameCapitalizationMsg(nameSpan), , CLng(startPos))
                End If
            End If
        Next startPos
    Next p

    Set CheckPersonNameCapitalizationWarn = Result
End Function

' Danh tu chung chi don vi hanh chinh -- LOP TU DONG (liet ke dich danh trong docs/rules/02- viet-
' hoa.md muc III.1), khong phai "tu dien" can tra JSON -- cung tien le VnPhan/VnDieu o PHAN 4.
Private Function AdminUnitKeywords() As Collection
    Dim c As New Collection
    c.Add "th" & ChrW(&HE0) & "nh ph" & ChrW(&H1ED1)
    c.Add "t" & ChrW(&H1EC9) & "nh"
    c.Add "qu" & ChrW(&H1EAD) & "n"
    c.Add "huy" & ChrW(&H1EC7) & "n"
    c.Add "th" & ChrW(&H1ECB) & " x" & ChrW(&HE3)
    c.Add "ph" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "x" & ChrW(&HE3)
    c.Add "th" & ChrW(&H1ECB) & " tr" & ChrW(&H1EA5) & "n"
    Set AdminUnitKeywords = c
End Function

Public Function CheckAdministrativeUnitNameWarn(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckAdministrativeUnitNameWarn")
    Dim paragraphs As Collection: Set paragraphs = ScannableParagraphs(context)
    If paragraphs.count = 0 Then
        Set CheckAdministrativeUnitNameWarn = Nothing
        Exit Function
    End If

    Dim gazetteer As Object: Set gazetteer = AdministrativeUnitNameLookup()

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        Dim kw As Variant
        For Each kw In AdminUnitKeywords()
            Dim keyword As String: keyword = CStr(kw)
            Dim searchFrom As Long: searchFrom = 1
            Do
                Dim pos As Long: pos = InStr(searchFrom, p.text, keyword, vbBinaryCompare)
                If pos = 0 Then Exit Do
                Dim afterPos As Long: afterPos = pos + Len(keyword)
                If afterPos <= Len(p.text) And Mid$(p.text, afterPos, 1) = " " Then
                    Dim namePart As String: namePart = CaptureNamePart(p.text, afterPos + 1)
                    If Len(namePart) > 0 Then
                        If Not IsDigitChar(left$(namePart, 1)) Then
                            Dim canonical As String: canonical = LongestGazetteerMatch(namePart, gazetteer)
                            If Len(canonical) > 0 Then
                                Dim actualSpan As String: actualSpan = left$(namePart, Len(canonical))
                                If actualSpan <> canonical Then
                                    Result.Add MakeFindingInput(CLng(p.Index), "'" & keyword & " " & actualSpan & _
                                        "' " & ChrW(&H2014) & " t" & ChrW(&HEA) & "n ri" & ChrW(&HEA) & "ng " & ChrW(&H111) & ChrW(&H1A1) & "n v" & ChrW(&H1ECB) & " h" & ChrW(&HE0) & "nh ch" & ChrW(&HED) & "nh ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t hoa " & _
                                        "ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u t" & ChrW(&H1EEB) & "ng " & ChrW(&HE2) & "m ti" & ChrW(&H1EBF) & "t.", , CLng(pos - 1), _
                                        keyword & " " & actualSpan, keyword & " " & canonical)
                                End If
                            End If
                        End If
                    End If
                End If
                searchFrom = pos + 1
            Loop
        Next kw
    Next p
    Set CheckAdministrativeUnitNameWarn = Result
End Function

' Dictionary tra cuu KHONG PHAN BIET HOA THUONG (key = Utils.ToUpperVn(ten)) -> ten DUNG CHINH TA
' theo don-vi-hanh-chinh.json (gop chung tinh/thanh + xa/phuong, vi o day chi can biet "co phai
' ten dia danh hop le hay khong", khong can phan biet cap hanh chinh).
Private Function AdministrativeUnitNameLookup() As Object
    Dim Result As Object: Set Result = Utils.NewDictionary()
    Dim data As Object: Set data = RuleLoader.GetAdministrativeUnitNames()
    Dim v As Variant
    For Each v In data("provinces")
        Result(Utils.ToUpperVn(CStr(v))) = CStr(v)
    Next v
    For Each v In data("communes")
        Dim key As String: key = Utils.ToUpperVn(CStr(v))
        If Not Result.Exists(key) Then Result(key) = CStr(v)
    Next v
    Set AdministrativeUnitNameLookup = Result
End Function

' Tim TIEN TO DAI NHAT cua namePart (cat theo tu) trung mot dia danh trong lookup (khong phan biet
' hoa/thuong) -- tra ve DUNG chinh ta chinh thuc neu tim thay, chuoi rong neu khong tim thay tien
' to nao ca (namePart khong co trong danh sach thi KHONG lay lam can cu bao loi).
Private Function LongestGazetteerMatch(ByVal namePart As String, ByVal lookup As Object) As String
    Dim words() As String: words = Split(namePart, " ")
    Dim n As Long: n = UBound(words) - LBound(words) + 1
    Dim wc As Long
    For wc = n To 1 Step -1
        Dim candidate As String: candidate = words(LBound(words))
        Dim i As Long
        For i = 1 To wc - 1
            candidate = candidate & " " & words(LBound(words) + i)
        Next i
        Dim key As String: key = Utils.ToUpperVn(candidate)
        If lookup.Exists(key) Then
            LongestGazetteerMatch = CStr(lookup(key))
            Exit Function
        End If
    Next wc
    LongestGazetteerMatch = ""
End Function

Private Function IsDigitChar(ByVal ch As String) As Boolean
    If ch = "" Then
        IsDigitChar = False
    Else
        Dim code As Long: code = AscW(ch)
        IsDigitChar = (code >= 48 And code <= 57)
    End If
End Function

Private Function CaptureNamePart(ByVal text As String, ByVal startPos As Long) As String
    Dim Result As String: Result = ""
    Dim pos As Long: pos = startPos
    Dim wordCount As Long: wordCount = 0
    Do While pos <= Len(text) And wordCount < 5
        Dim ch As String: ch = Mid$(text, pos, 1)
        If ch = "," Or ch = "." Or ch = ";" Or ch = ":" Or ch = ")" Then Exit Do
        If ch = " " Then
            If Right$(Result, 1) = " " Then Exit Do ' hai dau cach lien tiep -- ket thuc cum
            Result = Result & ch
            wordCount = wordCount + 1
        ElseIf IsLetterChar(ch) Or IsDigitChar(ch) Then
            Result = Result & ch
        Else
            Exit Do
        End If
        pos = pos + 1
    Loop
    CaptureNamePart = Trim$(Result)
End Function

' ND30-PL2-M3-K1B -- partial, autoFixable: don vi hanh chinh dat theo chu so ("quan 1" phai la
' "Quan 1") -- chi khop dang CHUA viet hoa.
Public Function CheckAdministrativeUnitNumeralCase(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckAdministrativeUnitNumeralCase")
    Dim paragraphs As Collection: Set paragraphs = ScannableParagraphs(context)
    If paragraphs.count = 0 Then
        Set CheckAdministrativeUnitNumeralCase = Nothing
        Exit Function
    End If

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        Dim kw As Variant
        For Each kw In AdminUnitKeywords()
            Dim keyword As String: keyword = CStr(kw)
            Dim searchFrom As Long: searchFrom = 1
            Do
                Dim pos As Long: pos = InStr(searchFrom, p.text, keyword, vbBinaryCompare)
                If pos = 0 Then Exit Do
                Dim afterPos As Long: afterPos = pos + Len(keyword)
                If afterPos <= Len(p.text) And Mid$(p.text, afterPos, 1) = " " Then
                    Dim digitsEnd As Long: digitsEnd = afterPos + 1
                    Do While digitsEnd <= Len(p.text) And IsDigitChar(Mid$(p.text, digitsEnd, 1))
                        digitsEnd = digitsEnd + 1
                    Loop
                    If digitsEnd > afterPos + 1 Then
                        Dim full As String: full = Mid$(p.text, pos, digitsEnd - pos)
                        Dim expected As String: expected = CapitalizeFirstLetterVn(keyword) & Mid$(full, Len(keyword) + 1)
                        Result.Add MakeFindingInput(CLng(p.Index), "'" & full & "' n" & ChrW(&HEA) & "n vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '" & expected & _
                            "'. Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng h" & ChrW(&H1EE3) & "p t" & ChrW(&HEA) & "n " & ChrW(&H111) & ChrW(&H1A1) & "n v" & ChrW(&H1ECB) & " h" & ChrW(&HE0) & "nh ch" & ChrW(&HED) & "nh c" & ChrW(&H1EA5) & "u t" & ChrW(&H1EA1) & "o v" & ChrW(&H1EDB) & "i ch" & ChrW(&H1EEF) & " s" & ChrW(&H1ED1) & " th" & ChrW(&HEC) & " vi" & ChrW(&H1EBF) & "t hoa c" & ChrW(&H1EA3) & " danh t" & ChrW(&H1EEB) & " chung.", _
                            , CLng(pos - 1), full, expected)
                    End If
                End If
                searchFrom = pos + 1
            Loop
        Next kw
    Next p
    Set CheckAdministrativeUnitNumeralCase = Result
End Function

Private Function SpecialCapitalizationEntriesByMarker(ByVal marker As String) As Collection
    Dim Result As New Collection
    Dim entries As Object: Set entries = RuleLoader.GetSpecialCapitalizations()("entries")
    Dim e As Variant
    For Each e In entries
        If InStr(1, CStr(e("citation")), marker, vbBinaryCompare) > 0 Then Result.Add e
    Next e
    Set SpecialCapitalizationEntriesByMarker = Result
End Function

' Dung chung cho ba quy tac tra viet-hoa-dac-biet.json theo muc Phu luc II: ND30-PL2-M3-K1C (Muc
' III -- dia ly), ND30-PL2-M4-K1B (Muc IV -- co quan), ND30-PL2-M5-K1 (Muc V -- Nhan dan/ Nha
' nuoc). extraSentence: cau bo sung rieng cho M5-K1 (rong = khong co).
Private Function SpecialPhraseCheck(ByVal context As Object, ByVal marker As String, _
        ByVal extraSentence As String) As Object
    Dim entries As Collection: Set entries = SpecialCapitalizationEntriesByMarker(marker)
    If entries.count = 0 Then
        Set SpecialPhraseCheck = Nothing
        Exit Function
    End If
    Dim paragraphs As Collection: Set paragraphs = ScannableParagraphs(context)
    If paragraphs.count = 0 Then
        Set SpecialPhraseCheck = Nothing
        Exit Function
    End If

    Dim Result As New Collection
    Dim entry As Variant
    For Each entry In entries
        Dim phrase As String: phrase = CStr(entry("phrase"))
        Dim p As ParagraphSnapshot
        For Each p In paragraphs
            Dim matches As Collection: Set matches = FindWholePhraseCI(p.text, phrase)
            Dim m As Variant
            For Each m In matches
                Dim actual As String: actual = CStr(m("Actual"))
                If actual <> phrase Then
                    Dim msg As String
                    msg = "'" & actual & "' ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '" & phrase & "'." & extraSentence
                    Result.Add MakeFindingInput(CLng(p.Index), msg, , CLng(m("Index")), actual, phrase)
                End If
            Next m
        Next p
    Next entry
    Set SpecialPhraseCheck = Result
End Function

Public Function CheckSpecialGeographicCapitalization(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckSpecialGeographicCapitalization")
    Set CheckSpecialGeographicCapitalization = SpecialPhraseCheck(context, "M" & ChrW(&H1EE5) & "c III", "")
End Function

Public Function CheckSpecialOrganCapitalization(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckSpecialOrganCapitalization")
    Set CheckSpecialOrganCapitalization = SpecialPhraseCheck(context, "M" & ChrW(&H1EE5) & "c IV", "")
End Function

' "Bo qua (xoa code) ve kiem tra viet hoa chu Nha nuoc va Nhan dan") - qua nhieu bao sai trong
' thuc te.

' ND30-PL2-M3-K1D -- partial, autoFixable: danh tu chung chi dia hinh da thanh ten rieng (dia-
' danh-dia-hinh.json), viet hoa TAT CA.
Public Function CheckTerrainPlaceNameCapitalization(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckTerrainPlaceNameCapitalization")
    Dim places As Object: Set places = RuleLoader.GetTerrainPlaceNames()("places")
    If places.count = 0 Then
        Set CheckTerrainPlaceNameCapitalization = Nothing
        Exit Function
    End If
    Dim paragraphs As Collection: Set paragraphs = ScannableParagraphs(context)
    If paragraphs.count = 0 Then
        Set CheckTerrainPlaceNameCapitalization = Nothing
        Exit Function
    End If

    Dim Result As New Collection
    Dim place As Variant
    For Each place In places
        Dim placeStr As String: placeStr = CStr(place)
        Dim p As ParagraphSnapshot
        For Each p In paragraphs
            Dim matches As Collection: Set matches = FindWholePhraseCI(p.text, placeStr)
            Dim m As Variant
            For Each m In matches
                Dim actual As String: actual = CStr(m("Actual"))
                If actual <> placeStr Then
                    Result.Add MakeFindingInput(CLng(p.Index), "'" & actual & "' n" & ChrW(&HEA) & "n vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '" & placeStr & _
                        "'. Danh t" & ChrW(&H1EEB) & " chung ch" & ChrW(&H1EC9) & " " & ChrW(&H111) & ChrW(&H1ECB) & "a h" & ChrW(&HEC) & "nh " & ChrW(&H111) & ChrW(&HE3) & " tr" & ChrW(&H1EDF) & " th" & ChrW(&HE0) & "nh t" & ChrW(&HEA) & "n ri" & ChrW(&HEA) & "ng th" & ChrW(&HEC) & " vi" & ChrW(&H1EBF) & "t hoa t" & ChrW(&H1EA5) & "t c" & ChrW(&H1EA3) & ".", , _
                        CLng(m("Index")), actual, placeStr)
                End If
            Next m
        Next p
    Next place
    Set CheckTerrainPlaceNameCapitalization = Result
End Function

' ND30-PL2-M3-K1E -- partial, autoFixable: ten vung, mien, khu vuc (vung-mien.json).
Public Function CheckRegionNameCapitalization(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckRegionNameCapitalization")
    Dim regions As Object: Set regions = RuleLoader.GetRegionNames()("regions")
    If regions.count = 0 Then
        Set CheckRegionNameCapitalization = Nothing
        Exit Function
    End If
    Dim paragraphs As Collection: Set paragraphs = ScannableParagraphs(context)
    If paragraphs.count = 0 Then
        Set CheckRegionNameCapitalization = Nothing
        Exit Function
    End If

    Dim Result As New Collection
    Dim region As Variant
    For Each region In regions
        Dim regionStr As String: regionStr = CStr(region)
        Dim p As ParagraphSnapshot
        For Each p In paragraphs
            Dim matches As Collection: Set matches = FindWholePhraseCI(p.text, regionStr)
            Dim m As Variant
            For Each m In matches
                Dim actual As String: actual = CStr(m("Actual"))
                If actual <> regionStr Then
                    Result.Add MakeFindingInput(CLng(p.Index), "'" & actual & "' n" & ChrW(&HEA) & "n vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '" & regionStr & _
                        "'. Vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u c" & ChrW(&H1EE7) & "a t" & ChrW(&H1EA5) & "t c" & ChrW(&H1EA3) & " c" & ChrW(&HE1) & "c " & ChrW(&HE2) & "m ti" & ChrW(&H1EBF) & "t t" & ChrW(&H1EA1) & "o th" & ChrW(&HE0) & "nh t" & ChrW(&HEA) & "n g" & ChrW(&H1ECD) & "i.", , _
                        CLng(m("Index")), actual, regionStr)
                End If
            Next m
        Next p
    Next region
    Set CheckRegionNameCapitalization = Result
End Function

' ND30-PL2-M4-K1A -- warnOnly: ten co quan pho bien (ten-co-quan-pho-bien.json) -- CANH BAO, KHONG
' autoFixable.
Public Function CheckCommonOrganNameWarn(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckCommonOrganNameWarn")
    Dim organs As Object: Set organs = RuleLoader.GetCommonOrganNames()("organs")
    If organs.count = 0 Then
        Set CheckCommonOrganNameWarn = Nothing
        Exit Function
    End If
    Dim paragraphs As Collection: Set paragraphs = ScannableParagraphs(context)
    If paragraphs.count = 0 Then
        Set CheckCommonOrganNameWarn = Nothing
        Exit Function
    End If

    Dim Result As New Collection
    Dim organ As Variant
    For Each organ In organs
        Dim organStr As String: organStr = CStr(organ)
        Dim p As ParagraphSnapshot
        For Each p In paragraphs
            Dim matches As Collection: Set matches = FindWholePhraseCI(p.text, organStr)
            Dim m As Variant
            For Each m In matches
                Dim actual As String: actual = CStr(m("Actual"))
                If actual <> organStr Then
                    Result.Add MakeFindingInput(CLng(p.Index), "'" & actual & "' " & ChrW(&H2014) & " vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u c" & ChrW(&H1EE7) & "a " & _
                        "c" & ChrW(&HE1) & "c t" & ChrW(&H1EEB) & ", c" & ChrW(&H1EE5) & "m t" & ChrW(&H1EEB) & " ch" & ChrW(&H1EC9) & " lo" & ChrW(&H1EA1) & "i h" & ChrW(&HEC) & "nh c" & ChrW(&H1A1) & " quan t" & ChrW(&H1ED5) & " ch" & ChrW(&H1EE9) & "c v" & ChrW(&HE0) & " ch" & ChrW(&H1EE9) & "c n" & ChrW(&H103) & "ng, l" & ChrW(&H129) & "nh v" & ChrW(&H1EF1) & "c ho" & ChrW(&H1EA1) & "t " & ChrW(&H111) & ChrW(&H1ED9) & "ng. " & _
                        ChrW(&H110) & ChrW(&H1ED1) & "i chi" & ChrW(&H1EBF) & "u danh s" & ChrW(&HE1) & "ch c" & ChrW(&H1A1) & " quan ph" & ChrW(&H1ED5) & " bi" & ChrW(&H1EBF) & "n: c" & ChrW(&HF3) & " th" & ChrW(&H1EC3) & " " & ChrW(&H111) & ChrW(&HFA) & "ng l" & ChrW(&HE0) & " '" & organStr & "'.", , _
                        CLng(m("Index")), actual, organStr)
                End If
            Next m
        Next p
    Next organ
    Set CheckCommonOrganNameWarn = Result
End Function

' ND30-PL2-M5-K5 -- partial, autoFixable: ten ngay le, ngay ky niem (ngay-le.json).
Public Function CheckHolidayNameCapitalization(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckHolidayNameCapitalization")
    Dim entries As Object: Set entries = RuleLoader.GetHolidays()("entries")
    If entries.count = 0 Then
        Set CheckHolidayNameCapitalization = Nothing
        Exit Function
    End If
    Dim paragraphs As Collection: Set paragraphs = ScannableParagraphs(context)
    If paragraphs.count = 0 Then
        Set CheckHolidayNameCapitalization = Nothing
        Exit Function
    End If

    Dim Result As New Collection
    Dim holiday As Variant
    For Each holiday In entries
        Dim holidayName As String: holidayName = CStr(holiday("name"))
        Dim p As ParagraphSnapshot
        For Each p In paragraphs
            Dim matches As Collection: Set matches = FindWholePhraseCI(p.text, holidayName)
            Dim m As Variant
            For Each m In matches
                Dim actual As String: actual = CStr(m("Actual"))
                If actual <> holidayName Then
                    Result.Add MakeFindingInput(CLng(p.Index), "'" & actual & "' n" & ChrW(&HEA) & "n vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '" & holidayName & _
                        "'. Vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u c" & ChrW(&H1EE7) & "a " & ChrW(&HE2) & "m ti" & ChrW(&H1EBF) & "t t" & ChrW(&H1EA1) & "o th" & ChrW(&HE0) & "nh t" & ChrW(&HEA) & "n g" & ChrW(&H1ECD) & "i ng" & ChrW(&HE0) & "y l" & ChrW(&H1EC5) & ", ng" & ChrW(&HE0) & "y k" & ChrW(&H1EF7) & " ni" & ChrW(&H1EC7) & "m.", , _
                        CLng(m("Index")), actual, holidayName)
                End If
            Next m
        Next p
    Next holiday
    Set CheckHolidayNameCapitalization = Result
End Function

' ND30-PL2-M5-K8A -- partial, autoFixable: ten nam am lich (nam-am-lich.json, 60 nam can chi).
Public Function CheckLunarYearCapitalization(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckLunarYearCapitalization")
    Dim years As Object: Set years = RuleLoader.GetLunarYears()("years")
    If years.count = 0 Then
        Set CheckLunarYearCapitalization = Nothing
        Exit Function
    End If
    Dim paragraphs As Collection: Set paragraphs = ScannableParagraphs(context)
    If paragraphs.count = 0 Then
        Set CheckLunarYearCapitalization = Nothing
        Exit Function
    End If

    Dim Result As New Collection
    Dim year As Variant
    For Each year In years
        Dim yearStr As String: yearStr = CStr(year)
        Dim p As ParagraphSnapshot
        For Each p In paragraphs
            Dim matches As Collection: Set matches = FindWholePhraseCI(p.text, yearStr)
            Dim m As Variant
            For Each m In matches
                Dim actual As String: actual = CStr(m("Actual"))
                If actual <> yearStr Then
                    Result.Add MakeFindingInput(CLng(p.Index), "'" & actual & "' ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '" & yearStr & _
                        "'. Vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u c" & ChrW(&H1EE7) & "a t" & ChrW(&H1EA5) & "t c" & ChrW(&H1EA3) & " c" & ChrW(&HE1) & "c " & ChrW(&HE2) & "m ti" & ChrW(&H1EBF) & "t t" & ChrW(&H1EA1) & "o th" & ChrW(&HE0) & "nh t" & ChrW(&HEA) & "n g" & ChrW(&H1ECD) & "i.", , _
                        CLng(m("Index")), actual, yearStr)
                End If
            Next m
        Next p
    Next year
    Set CheckLunarYearCapitalization = Result
End Function

' Khop "dieu" VIET THUONG theo sau boi so La Ma/A Rap (khong lookbehind -- rui ro khop giua mot tu
' dai hon chap nhan duoc, cung muc do voi cac pattern khac cua PHAN 4 khong dung \b).
' "khong kiem tra, khong xu ly tat ca nhung van de lien quan den phan, chuong, muc, tieu muc - bo
' hoan toan khoi code"): BO han phan/chuong/muc/tieu muc khoi danh sach nay - CHI CON "dieu".
' Khoan/Diem (clausePattern/ pointPattern o CheckArticleClauseCapitalization duoi day) KHONG bi
' anh huong.
Private Function CitationLowercaseLevelKeywords() As Collection
    Dim c As New Collection
    c.Add Utils.ToLowerVn(VnDieu())
    Set CitationLowercaseLevelKeywords = c
End Function

' ND30-PL2-M5-K7 -- full, autoFixable: vien dan phan, chuong, muc, tieu muc, dieu, khoan, diem.
Public Function CheckArticleClauseCapitalization(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckArticleClauseCapitalization")
    Dim paragraphs As Collection: Set paragraphs = ScannableParagraphs(context)
    If paragraphs.count = 0 Then
        Set CheckArticleClauseCapitalization = Nothing
        Exit Function
    End If

    Dim lowerPattern As String
    lowerPattern = "(" & JoinCollection(CitationLowercaseLevelKeywords(), "|") & ")\s+([IVXLCDM]+|\d+)"
    Dim clausePattern As String: clausePattern = VnKhoan() & "\s+\d+"

    ' Lop ky tu "chu cai tieng Viet" tra THANG tu spec("vietnameseAlphabet") -- CAM so sanh chuoi
    ' mac dinh cua nen tang (dung tien le BuildPointPattern o PHAN 4).
    Dim spec As Object: Set spec = context("Spec")
    Dim charClass As String: charClass = ""
    Dim va As Variant
    For Each va In spec("vietnameseAlphabet")
        charClass = charClass & EscapeRegExpLiteral(CStr(va))
    Next va
    Dim pointPattern As String: pointPattern = VnDiem() & "\s+[" & charClass & "]\)"

    Dim dieuFirstChar As String: dieuFirstChar = left$(VnDieu(), 1)
    Dim articleMentionPattern As String
    articleMentionPattern = "[" & dieuFirstChar & Utils.ToLowerVn(dieuFirstChar) & "]" & Mid$(VnDieu(), 2) & "\s+\d+"

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        Dim text As String: text = p.text

        Dim matches As Object: Set matches = RegexAllMatches(lowerPattern, text)
        Dim m As Object
        For Each m In matches
            Dim keyword As String: keyword = CStr(m.SubMatches(0))
            Dim full As String: full = CStr(m.value)
            Dim expected As String: expected = CapitalizeFirstLetterVn(keyword) & Mid$(full, Len(keyword) + 1)
            Result.Add MakeFindingInput(CLng(p.Index), "'" & full & "' ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '" & expected & "'. " & _
                "Vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u c" & ChrW(&H1EE7) & "a " & ChrW(&H111) & "i" & ChrW(&H1EC1) & "u.", , CLng(m.firstIndex), full, expected)
        Next m

        If RegexTestCC(articleMentionPattern, text) Then
            Dim clauseMatches As Object: Set clauseMatches = RegexAllMatches(clausePattern, text)
            Dim cm As Object
            For Each cm In clauseMatches
                Dim cFull As String: cFull = CStr(cm.value)
                Dim cExpected As String: cExpected = Utils.ToLowerVn(left$(cFull, 1)) & Mid$(cFull, 2)
                Result.Add MakeFindingInput(CLng(p.Index), "'" & cFull & "' ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '" & cExpected & "'. " & _
                    "Vi" & ChrW(&H1EBF) & "t th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng kho" & ChrW(&H1EA3) & "n, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m khi vi" & ChrW(&H1EC7) & "n d" & ChrW(&H1EAB) & "n trong ng" & ChrW(&H1EEF) & " c" & ChrW(&H1EA3) & "nh c" & ChrW(&HF3) & " " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "u " & ChrW(&H111) & ChrW(&H1EE9) & "ng c" & ChrW(&HF9) & "ng.", , CLng(cm.firstIndex), _
                    cFull, cExpected)
            Next cm
            Dim pointMatches As Object: Set pointMatches = RegexAllMatches(pointPattern, text)
            Dim pm As Object
            For Each pm In pointMatches
                Dim pFull As String: pFull = CStr(pm.value)
                Dim pExpected As String: pExpected = Utils.ToLowerVn(left$(pFull, 1)) & Mid$(pFull, 2)
                Result.Add MakeFindingInput(CLng(p.Index), "'" & pFull & "' ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '" & pExpected & "'. " & _
                    "Vi" & ChrW(&H1EBF) & "t th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng kho" & ChrW(&H1EA3) & "n, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m khi vi" & ChrW(&H1EC7) & "n d" & ChrW(&H1EAB) & "n trong ng" & ChrW(&H1EEF) & " c" & ChrW(&H1EA3) & "nh c" & ChrW(&HF3) & " " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "u " & ChrW(&H111) & ChrW(&H1EE9) & "ng c" & ChrW(&HF9) & "ng.", , CLng(pm.firstIndex), _
                    pFull, pExpected)
            Next pm
        End If
    Next p
    Set CheckArticleClauseCapitalization = Result
End Function

Public Function CheckCapitalizationNotDetectable(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckCapitalizationNotDetectable")
    Set CheckCapitalizationNotDetectable = Nothing
End Function

' ----------------------------------------------------------------------------
' Nhom spelling -- QD1989-* (dau thanh, i/y) va LOCAL-TYPO-* (thong le soan thao).
' ----------------------------------------------------------------------------

Private Function IsUGroupToneKey(ByVal key As String) As Boolean
    Dim yLetters As String
    yLetters = "y" & ChrW(&H1EF3) & ChrW(&HFD) & ChrW(&H1EF7) & ChrW(&H1EF9) & ChrW(&H1EF5)
    Dim i As Long
    For i = 1 To Len(key)
        If InStr(yLetters, Mid$(key, i, 1)) > 0 Then
            IsUGroupToneKey = True
            Exit Function
        End If
    Next i
    IsUGroupToneKey = False
End Function

Private Function CountToneOccurrences(ByVal paragraphs As Collection, ByVal mapping As Object, _
        ByVal conditions As Object, ByRef firstParagraphIndex As Long) As Long
    Dim total As Long: total = 0
    Dim oConsonants As String: oConsonants = CStr(conditions("precedingConsonantsForOGroup"))
    Dim uConsonants As String: uConsonants = CStr(conditions("precedingConsonantsForUGroup"))

    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        Dim text As String: text = p.text
        Dim keyVar As Variant
        For Each keyVar In mapping.Keys
            Dim key As String: key = CStr(keyVar)
            ' Khoa "$comment" trong file JSON khong phai to hop dau thanh - bo qua.
            If left$(key, 1) = "$" Then GoTo ContinueKey
            Dim consonants As String
            If IsUGroupToneKey(key) Then consonants = uConsonants Else consonants = oConsonants

            Dim searchFrom As Long: searchFrom = 1
            Do
                Dim pos As Long: pos = InStr(searchFrom, text, key, vbBinaryCompare)
                If pos = 0 Then Exit Do
                Dim beforeOk As Boolean: beforeOk = False
                If pos > 1 Then
                    Dim beforeCh As String: beforeCh = Mid$(text, pos - 1, 1)
                    beforeOk = (InStr(consonants, beforeCh) > 0)
                End If
                Dim afterPos As Long: afterPos = pos + Len(key)
                Dim afterOk As Boolean: afterOk = True
                If afterPos <= Len(text) Then afterOk = Not IsLetterChar(Mid$(text, afterPos, 1))
                If beforeOk And afterOk Then
                    total = total + 1
                    If firstParagraphIndex = -1 Then firstParagraphIndex = p.Index
                End If
                searchFrom = pos + 1
            Loop
ContinueKey:
        Next keyVar
    Next p
    CountToneOccurrences = total
End Function

' "Viec bo dau khong duoc tinh la loi, do do bo han cac comment 'Vi tri dau thanh chua theo
' chuan'") - ham CheckToneMarkStyle cu da xoa. CHI kieu dat dau KHONG THONG NHAT trong CUNG mot
' van ban moi la loi thuc su, xem CheckToneMarkMix duoi day (giu nguyen, khong doi).

Public Function CheckToneMarkMix(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckToneMarkMix")
    Dim paragraphs As Collection: Set paragraphs = ScannableParagraphs(context)
    If paragraphs.count = 0 Then
        Set CheckToneMarkMix = Nothing
        Exit Function
    End If
    Dim mapping As Object: Set mapping = RuleLoader.GetToneMapping()
    Dim firstIndex As Long: firstIndex = -1
    Dim countA As Long ' kieu oa, uy (khong chuan) -- khoa mapToMainVowel
    Dim countB As Long ' kieu oa, uy (chuan) -- khoa mapToFirstVowel
    countA = CountToneOccurrences(paragraphs, mapping("mapToMainVowel"), mapping("conditions"), firstIndex)
    countB = CountToneOccurrences(paragraphs, mapping("mapToFirstVowel"), mapping("conditions"), firstIndex)

    Dim Result As New Collection
    If countA = 0 Or countB = 0 Then
        Set CheckToneMarkMix = Result
        Exit Function
    End If
    Result.Add MakeFindingInput(Null, "Ph" & ChrW(&HE1) & "t hi" & ChrW(&H1EC7) & "n " & countA & " ch" & ChrW(&H1ED7) & " ki" & ChrW(&H1EC3) & "u '" & ChrW(&HF2) & "a, " & ChrW(&HFA) & "y' v" & ChrW(&HE0) & " " & countB & _
        " ch" & ChrW(&H1ED7) & " ki" & ChrW(&H1EC3) & "u 'o" & ChrW(&HE0) & ", u" & ChrW(&HFD) & "' trong c" & ChrW(&HF9) & "ng v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n.", CLng(countA + countB))
    Set CheckToneMarkMix = Result
End Function

Private Function IyExcludedPhrases() As Collection
    Dim Result As New Collection
    Dim ex As Object: Set ex = RuleLoader.GetIyMapping()("excludeAbsolute")
    Dim groupKey As Variant
    For Each groupKey In Array("documentTypeNames", "typeAbbreviations", "nd30Terminology", "nationalTitle")
        Dim terms As Object: Set terms = ex(CStr(groupKey))("terms")
        Dim t As Variant
        For Each t In terms
            Result.Add CStr(t)
        Next t
    Next groupKey
    Set IyExcludedPhrases = Result
End Function

' Thay moi doan khop phrase (nguyen cum, khong phan biet hoa/thuong) trong text bang khoang trang
' cung do dai -- dung de LOAI TRU cac cum co dinh truoc khi do cap i/y.
Private Function MaskExcludedPhrases(ByVal text As String, ByVal phrases As Collection) As String
    Dim masked As String: masked = text
    Dim phrase As Variant
    For Each phrase In phrases
        Dim phraseStr As String: phraseStr = CStr(phrase)
        Dim matches As Collection: Set matches = FindWholePhraseCI(masked, phraseStr)
        Dim m As Variant
        For Each m In matches
            Dim idx As Long: idx = CLng(m("Index")) + 1 ' ve 1-based
            masked = left$(masked, idx - 1) & String(Len(phraseStr), " ") & Mid$(masked, idx + Len(phraseStr))
        Next m
    Next phrase
    MaskExcludedPhrases = masked
End Function

' excludeAbsolute.properNouns.skipCapitalizedMidSentence -- chu cai dau viet hoa VA khong o dau
' doan/dau cau (ky tu lien truoc, bo qua khoang trang, khong phai dau ket cau va khong phai dau
' doan) thi coi la danh tu rieng, loai tru.
Private Function IsMidSentenceCapitalized(ByVal text As String, ByVal indexZeroBased As Long, ByVal actual As String) As Boolean
    If Not IsUppercaseLetterChar(left$(actual, 1)) Then
        IsMidSentenceCapitalized = False
        Exit Function
    End If
    Dim i As Long: i = indexZeroBased ' 1-based vi tri ky tu ngay TRUOC actual
    Do While i >= 1
        If Mid$(text, i, 1) <> " " Then Exit Do
        i = i - 1
    Loop
    If i < 1 Then
        IsMidSentenceCapitalized = False
        Exit Function
    End If
    Dim ch As String: ch = Mid$(text, i, 1)
    IsMidSentenceCapitalized = (ch <> "." And ch <> "!" And ch <> "?")
End Function

' QD1989-IY-MIX -- full: van ban dung lan 'i' va 'y' cho cung mot tu.
Public Function CheckIyMix(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckIyMix")
    Dim mapping As Object: Set mapping = RuleLoader.GetIyMapping()
    Dim paragraphs As Collection: Set paragraphs = IyScannableParagraphs(context)
    If paragraphs.count = 0 Then
        Set CheckIyMix = Nothing
        Exit Function
    End If
    Dim excludedPhrases As Collection: Set excludedPhrases = IyExcludedPhrases()

    Dim Result As New Collection
    Dim pairKey As Variant
    For Each pairKey In mapping("pairs").Keys
        Dim formI As String: formI = CStr(pairKey)
        Dim formY As String: formY = CStr(mapping("pairs")(pairKey))

        Dim foundI As Boolean: foundI = False
        Dim foundIParagraph As Long, foundICharOffset As Long
        Dim foundY As Boolean: foundY = False
        Dim foundYParagraph As Long, foundYCharOffset As Long

        Dim p As ParagraphSnapshot
        For Each p In paragraphs
            Dim masked As String: masked = MaskExcludedPhrases(p.text, excludedPhrases)
            If Not foundI Then
                Dim matchesI As Collection: Set matchesI = FindWholePhraseCI(masked, formI)
                Dim mi As Variant
                For Each mi In matchesI
                    Dim actualI As String: actualI = CStr(mi("Actual"))
                    Dim idxI As Long: idxI = CLng(mi("Index"))
                    If left$(Utils.ToLowerVn(actualI), 2) <> "qu" Then
                        If Not IsMidSentenceCapitalized(p.text, idxI, actualI) Then
                            foundI = True
                            foundIParagraph = p.Index
                            foundICharOffset = idxI
                            Exit For
                        End If
                    End If
                Next mi
            End If
            If Not foundY Then
                Dim matchesY As Collection: Set matchesY = FindWholePhraseCI(masked, formY)
                Dim myv As Variant
                For Each myv In matchesY
                    Dim actualY As String: actualY = CStr(myv("Actual"))
                    Dim idxY As Long: idxY = CLng(myv("Index"))
                    If left$(Utils.ToLowerVn(actualY), 2) <> "qu" Then
                        If Not IsMidSentenceCapitalized(p.text, idxY, actualY) Then
                            foundY = True
                            foundYParagraph = p.Index
                            foundYCharOffset = idxY
                            Exit For
                        End If
                    End If
                Next myv
            End If
            If foundI And foundY Then Exit For
        Next p

        If foundI And foundY Then
            ' ParagraphIndex Null (khong con neo vao doan tim thay "i" dau tien) - nhan dinh TOAN
            ' VAN BAN ("dung lan i/y"), FindingTierAggregator.bas gop vao MOT comment duy nhat dau
            ' tai lieu (cung ly do voi CheckToneMarkMix o tren).
            Result.Add MakeFindingInput(Null, "Ph" & ChrW(&HE1) & "t hi" & ChrW(&H1EC7) & "n c" & ChrW(&H1EA3) & " '" & formI & "' v" & ChrW(&HE0) & " '" & formY & _
                "' trong c" & ChrW(&HF9) & "ng v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n.")
        End If
    Next pairKey
    Set CheckIyMix = Result
End Function

' LOCAL-TYPO-SPACE -- khoang trang thua (lien tiep, dau doan, cuoi doan). Ca ba truong hop nay
' da chuyen sang tu dong xoa o EdgeWhitespaceTrimmer/MultiSpaceCollapser (goi truoc buoc kiem
' tra, tu FindingReporter.RunCheckAndReport), nen ham nay tra Nothing mai mai. Ma quy tac giu
' nguyen trong registry de khong phai sua quy-tac-kiem-tra.json.
Public Function CheckExtraSpaceTypo(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckExtraSpaceTypo")
    Set CheckExtraSpaceTypo = Nothing
End Function

' LOCAL-TYPO-PUNCT -- full, autoFixable: khoang trang truoc dau cau.
Public Function CheckPunctuationSpacingTypo(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckPunctuationSpacingTypo")
    Dim paragraphs As Collection: Set paragraphs = ScannableParagraphs(context)
    If paragraphs.count = 0 Then
        Set CheckPunctuationSpacingTypo = Nothing
        Exit Function
    End If

    Dim Result As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        Dim matches As Object: Set matches = RegexAllMatches("[ \t]+([,.;:!?])", p.text)
        Dim m As Object
        For Each m In matches
            Dim punct As String: punct = CStr(m.SubMatches(0))
            Dim full As String: full = CStr(m.value)
            Result.Add MakeFindingInput(CLng(p.Index), "'" & full & "' n" & ChrW(&HEA) & "n vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '" & punct & "'. Tr" & ChrW(&H1B0) & ChrW(&H1EDB) & "c d" & ChrW(&H1EA5) & "u " & _
                "c" & ChrW(&HE2) & "u kh" & ChrW(&HF4) & "ng c" & ChrW(&HF3) & " d" & ChrW(&H1EA5) & "u c" & ChrW(&HE1) & "ch.", , CLng(m.firstIndex), full, punct)
        Next m
    Next p
    Set CheckPunctuationSpacingTypo = Result
End Function

Private Function HiddenCharLabel(ByVal ch As String) As String
    HiddenCharLabel = "U+" & Right$("0000" & Hex$(AscW(ch)), 4)
End Function

' LOCAL-TYPO-HIDDEN -- full, autoFixable: ky tu an (zero-width space, ZWNJ, ZWJ, BOM, khoang trang
' dac biet...) doc tu unicode-to-nfc.json.
Public Function CheckHiddenCharacters(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckHiddenCharacters")
    Dim paragraphs As Collection: Set paragraphs = ScannableParagraphs(context)
    If paragraphs.count = 0 Then
        Set CheckHiddenCharacters = Nothing
        Exit Function
    End If
    Dim hiddenChars As Object: Set hiddenChars = RuleLoader.GetUnicodeToNfc()("hiddenChars")
    Dim allChars As New Collection
    Dim v As Variant
    For Each v In hiddenChars("remove")
        allChars.Add CStr(v)
    Next v
    For Each v In hiddenChars("replaceWithSpace")
        allChars.Add CStr(v)
    Next v
    If allChars.count = 0 Then
        Set CheckHiddenCharacters = New Collection
        Exit Function
    End If

    Dim countByChar As Object: Set countByChar = Utils.NewDictionary()
    Dim firstIndexByChar As Object: Set firstIndexByChar = Utils.NewDictionary()
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        Dim ch As Variant
        For Each ch In allChars
            Dim chStr As String: chStr = CStr(ch)
            Dim searchFrom As Long: searchFrom = 1
            Do
                Dim pos As Long: pos = InStr(searchFrom, p.text, chStr, vbBinaryCompare)
                If pos = 0 Then Exit Do
                If countByChar.Exists(chStr) Then
                    countByChar(chStr) = countByChar(chStr) + 1
                Else
                    countByChar(chStr) = 1
                    firstIndexByChar(chStr) = p.Index
                End If
                searchFrom = pos + 1
            Loop
        Next ch
    Next p

    Dim Result As New Collection
    Dim key As Variant
    For Each key In countByChar.Keys
        Result.Add MakeFindingInput(CLng(firstIndexByChar(key)), "Ph" & ChrW(&HE1) & "t hi" & ChrW(&H1EC7) & "n " & countByChar(key) & _
            " k" & ChrW(&HFD) & " t" & ChrW(&H1EF1) & " " & ChrW(&H1EA9) & "n lo" & ChrW(&H1EA1) & "i " & HiddenCharLabel(CStr(key)) & ".", CLng(countByChar(key)))
    Next key
    Set CheckHiddenCharacters = Result
End Function

' Ap khuon hoa/thuong cua source len target -- tu-dien-chinh-ta.json khai preserveCase: true.
Private Function ApplyCasePatternVn(ByVal source As String, ByVal target As String) As String
    If Len(source) = 0 Then
        ApplyCasePatternVn = target
        Exit Function
    End If
    If Utils.ToUpperVn(source) = source And Utils.ToLowerVn(source) <> source Then
        ApplyCasePatternVn = Utils.ToUpperVn(target)
    ElseIf IsUppercaseLetterChar(left$(source, 1)) Then
        ApplyCasePatternVn = CapitalizeFirstLetterVn(target)
    Else
        ApplyCasePatternVn = target
    End If
End Function

' LOCAL-TYPO-DICT -- partial, autoFixable: tu dien loi chinh ta pho bien (tu-dien-chinh- ta.json),
' co xet protectedForms (dang viet DUNG, khong bao gio de nghi sua).
Public Function CheckDictionaryTypo(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckDictionaryTypo")
    Dim dict As Object: Set dict = RuleLoader.GetTypoDictionary()
    Dim paragraphs As Collection: Set paragraphs = ScannableParagraphs(context)
    If paragraphs.count = 0 Then
        Set CheckDictionaryTypo = Nothing
        Exit Function
    End If
    Dim preserveCase As Boolean: preserveCase = CBool(dict("preserveCase"))

    Dim protectedSet As Object: Set protectedSet = Utils.NewDictionary()
    Dim pf As Variant
    For Each pf In dict("protectedForms")
        protectedSet(Utils.ToLowerVn(CStr(pf))) = True
    Next pf

    Dim Result As New Collection
    Dim wrongKey As Variant
    For Each wrongKey In dict("corrections").Keys
        Dim wrong As String: wrong = CStr(wrongKey)
        Dim correct As String: correct = CStr(dict("corrections")(wrongKey))
        Dim p As ParagraphSnapshot
        For Each p In paragraphs
            Dim matches As Collection: Set matches = FindWholePhraseCI(p.text, wrong)
            Dim m As Variant
            For Each m In matches
                Dim actual As String: actual = CStr(m("Actual"))
                If Not protectedSet.Exists(Utils.ToLowerVn(actual)) Then
                    Dim suggested As String
                    If preserveCase Then suggested = ApplyCasePatternVn(actual, correct) Else suggested = correct
                    Result.Add MakeFindingInput(CLng(p.Index), "'" & actual & "' " & ChrW(&H2014) & " d" & ChrW(&H1EA1) & "ng " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng d" & ChrW(&HF9) & "ng l" & ChrW(&HE0) & " '" & _
                        suggested & "'.", , CLng(m("Index")), actual, suggested)
                End If
            Next m
        Next p
    Next wrongKey
    Set CheckDictionaryTypo = Result
End Function

' LOCAL-TYPO-TELEX -- partial, autoFixable: mau telex chua duoc bo go chuyen (telex-
' whitelist.json) -- CHI khop dung cac mau tuong minh trong danh sach trang.
Public Function CheckTelexLeftoverTypo(ByVal context As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComplianceChecker.CheckTelexLeftoverTypo")
    Dim paragraphs As Collection: Set paragraphs = ScannableParagraphs(context)
    If paragraphs.count = 0 Then
        Set CheckTelexLeftoverTypo = Nothing
        Exit Function
    End If

    Dim Result As New Collection
    Dim telexEntry As Variant
    For Each telexEntry In RuleLoader.GetTelexWhitelist()("patterns")
        Dim pattern As String: pattern = CStr(telexEntry("pattern"))
        Dim replacement As String: replacement = CStr(telexEntry("replacement"))
        Dim p As ParagraphSnapshot
        For Each p In paragraphs
            Dim matches As Object: Set matches = RegexAllMatches(pattern, p.text)
            Dim m As Object
            For Each m In matches
                Dim full As String: full = CStr(m.value)
                Result.Add MakeFindingInput(CLng(p.Index), "'" & full & "' c" & ChrW(&HF3) & " th" & ChrW(&H1EC3) & " l" & ChrW(&HE0) & " '" & replacement & _
                    "' do b" & ChrW(&H1ED9) & " g" & ChrW(&HF5) & " ch" & ChrW(&H1B0) & "a chuy" & ChrW(&H1EC3) & "n xong.", , CLng(m.firstIndex), full, replacement)
            Next m
        Next p
    Next telexEntry
    Set CheckTelexLeftoverTypo = Result
End Function

' ----------------------------------------------------------------------------
' Dang ky -- goi tu EnsureRegistryInitialized, ghi de placeholder cua dung 33 ma tren.
' ----------------------------------------------------------------------------

Private Sub RegisterCapitalizationSpellingChecks()
    mRegistry("ND30-PL1-M2-K4-CASE") = "ComplianceChecker.CheckPlaceNameLetterCase"

    mRegistry("ND30-PL2-M1") = "ComplianceChecker.CheckSentenceCapitalization"
    mRegistry("ND30-PL2-M2-K1") = "ComplianceChecker.CheckPersonNameCapitalizationWarn"
    mRegistry("ND30-PL2-M2-K2") = "ComplianceChecker.CheckCapitalizationNotDetectable"
    mRegistry("ND30-PL2-M3-K1A") = "ComplianceChecker.CheckAdministrativeUnitNameWarn"
    mRegistry("ND30-PL2-M3-K1B") = "ComplianceChecker.CheckAdministrativeUnitNumeralCase"
    mRegistry("ND30-PL2-M3-K1C") = "ComplianceChecker.CheckSpecialGeographicCapitalization"
    mRegistry("ND30-PL2-M3-K1D") = "ComplianceChecker.CheckTerrainPlaceNameCapitalization"
    mRegistry("ND30-PL2-M3-K1E") = "ComplianceChecker.CheckRegionNameCapitalization"
    mRegistry("ND30-PL2-M3-K2") = "ComplianceChecker.CheckCapitalizationNotDetectable"
    mRegistry("ND30-PL2-M4-K1A") = "ComplianceChecker.CheckCommonOrganNameWarn"
    mRegistry("ND30-PL2-M4-K1B") = "ComplianceChecker.CheckSpecialOrganCapitalization"
    mRegistry("ND30-PL2-M4-K2") = "ComplianceChecker.CheckCapitalizationNotDetectable"
    mRegistry("ND30-PL2-M5-K2") = "ComplianceChecker.CheckCapitalizationNotDetectable"
    mRegistry("ND30-PL2-M5-K3") = "ComplianceChecker.CheckCapitalizationNotDetectable"
    mRegistry("ND30-PL2-M5-K4") = "ComplianceChecker.CheckCapitalizationNotDetectable"
    mRegistry("ND30-PL2-M5-K5") = "ComplianceChecker.CheckHolidayNameCapitalization"
    mRegistry("ND30-PL2-M5-K6") = "ComplianceChecker.CheckCapitalizationNotDetectable"
    mRegistry("ND30-PL2-M5-K7") = "ComplianceChecker.CheckArticleClauseCapitalization"
    mRegistry("ND30-PL2-M5-K8A") = "ComplianceChecker.CheckLunarYearCapitalization"
    mRegistry("ND30-PL2-M5-K8B") = "ComplianceChecker.CheckCapitalizationNotDetectable"
    mRegistry("ND30-PL2-M5-K8C") = "ComplianceChecker.CheckCapitalizationNotDetectable"
    mRegistry("ND30-PL2-M5-K9") = "ComplianceChecker.CheckCapitalizationNotDetectable"
    mRegistry("ND30-PL2-M5-K10") = "ComplianceChecker.CheckCapitalizationNotDetectable"

    ' "Viec bo dau khong duoc tinh la loi") - chi kieu dat dau KHONG THONG NHAT trong CUNG mot van
    ' ban moi la loi (QD1989-D8-TONE-MIX, giu nguyen). Ham CheckToneMarkStyle da xoa.
    mRegistry("QD1989-D8-TONE-MIX") = "ComplianceChecker.CheckToneMarkMix"
    mRegistry("QD1989-IY-MIX") = "ComplianceChecker.CheckIyMix"
    mRegistry("LOCAL-TYPO-SPACE") = "ComplianceChecker.CheckExtraSpaceTypo"
    mRegistry("LOCAL-TYPO-PUNCT") = "ComplianceChecker.CheckPunctuationSpacingTypo"
    mRegistry("LOCAL-TYPO-HIDDEN") = "ComplianceChecker.CheckHiddenCharacters"
    mRegistry("LOCAL-TYPO-DICT") = "ComplianceChecker.CheckDictionaryTypo"
    mRegistry("LOCAL-TYPO-TELEX") = "ComplianceChecker.CheckTelexLeftoverTypo"
End Sub
