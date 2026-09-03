Attribute VB_Name = "ComponentDetector"
'==============================================================
' ComponentDetector â€” nhan dien vai tro the thuc cua TUNG DOAN.
' Moi dau hieu nam trong shared/rules/dau-hieu-nhan-dien.json qua RuleLoader.GetComponentSignals,
' KHONG hard-code o day (CLAUDE.md muc 3.1).
' Chi doc, khong sua tai lieu. Dung AscW/ChrW, khong dung Asc/Chr.
'==============================================================
Option Explicit

Private mNationalTitleStripMap As Object
Private mNationalTitleStripMapReady As Boolean

' ============================================================================
' DetectComponents â€” diem vao duy nhat
' ============================================================================

' snapshot: Dictionary tra ve tu DocumentSnapshot.CaptureDocument (khoa "Paragraphs").
' documentType: "congVan" | "coTenLoai" | "khongXacDinh" | "toTrinh" | "conLai". regime: "ND30" |
' "VIETTEL" | "DANG". VIETTEL dung chung bo dau hieu voi ND30.
' Tra Dictionary hai khoa: "LayoutMap" -> Dictionary(Long ParagraphIndex -> String Role), chi cac
' doan da gan chac. "Diagnostics" -> Collection cua ComponentRole, moi lan khop ke ca do tin cay
' thap.
' Thu tu uu tien lay tu SignalOrder(regime) â€” vua cho vi tri dependency dung thu tu (organName can
' codeNumberNotation da co, subject can typeName da co...), vua la co che loai tru duy nhat cho
' organName trong bo cuc dang bang.
Public Function DetectComponents(ByVal snapshot As Object, ByVal documentType As String, _
        Optional ByVal regime As String = "ND30") As Object
    On Error GoTo ErrHandler

    Dim paragraphs As Collection
    Set paragraphs = snapshot("Paragraphs")

    Dim signalsRoot As Object
    Set signalsRoot = RuleLoader.GetComponentSignals()("signals")

    Dim headerLastIndex As Long
    headerLastIndex = DocumentSnapshot.HeaderWindowLastIndex(paragraphs, RuleLoader.GetHeaderWindowChars())

    Dim layoutMap As Object
    Set layoutMap = Utils.NewDictionary()
    Dim diagnostics As New Collection
    Dim assignedSet As Object
    Set assignedSet = Utils.NewDictionary()

    Dim order As Variant
    order = SignalOrder(regime)

    Dim i As Long
    Dim signalKey As String
    Dim signal As Object
    Dim matches As Collection
    Dim m As Variant
    Dim role As String
    Dim confidence As String
    Dim rec As ComponentRole

    For i = LBound(order) To UBound(order)
        signalKey = CStr(order(i))
        If Not signalsRoot.Exists(signalKey) Then GoTo ContinueSignal
        If ShouldSkipForDocumentType(signalKey, documentType) Then GoTo ContinueSignal

        ' Mot dau hieu loi (vi du du lieu JSON bat thuong) khong duoc lam hong ca luot nhan dien -
        ' ghi log va bo qua RIENG dau hieu do, cac dau hieu khac van chay (T-71).
        On Error GoTo SignalError
        Set signal = signalsRoot(signalKey)
        Set matches = MatchSignal(signal, paragraphs, layoutMap, assignedSet, regime)
        Set matches = ApplyScopeFilters(signal, matches, layoutMap, headerLastIndex)

        role = CStr(signal("role"))
        confidence = CStr(signal("confidence"))

        For Each m In matches
            Set rec = New ComponentRole
            rec.paragraphIndex = CLng(m)
            rec.role = role
            rec.confidence = confidence
            rec.signalKey = signalKey
            diagnostics.Add rec

            assignedSet.Add CLng(m), True
            ' "low" chi de chan doan, khong gan chac (ADR-003). Van danh dau assignedSet de khong
            ' doan nao bi hai dau hieu tranh nhau.
            If confidence <> "low" Then layoutMap.Add CLng(m), role
        Next m
        On Error GoTo 0
        GoTo ContinueSignal
SignalError:
        DebugTrace.LogErr "ComponentDetector.DetectComponents", _
            "Dau hieu '" & signalKey & "' loi khi nhan dien - bo qua dau hieu nay", Err.number, Err.description
        Err.Clear
        On Error GoTo 0
ContinueSignal:
    Next i

    DebugTrace.LogLayoutMap "ComponentDetector.DetectComponents(" & documentType & "," & regime & ")", paragraphs, layoutMap

    Dim Result As Object
    Set Result = Utils.NewDictionary()
    Result.Add "LayoutMap", layoutMap
    Result.Add "Diagnostics", diagnostics
    Set DetectComponents = Result
    Exit Function
ErrHandler:
    Err.Raise Err.number, "ComponentDetector.DetectComponents", Err.description
End Function

Private Function SignalOrder(ByVal regime As String) As Variant
    If regime = "DANG" Then
        SignalOrder = Array("partyHeader", "codeNumberNotation", "starSeparator", _
            "placeAndIssuedDate", "subjectOfficialLetter", _
            "appendixLabel", "appendixTitle", "appendixReference", _
            "recipientSalutation", "recipientSalutationInline", "recipientSalutationList", _
            "recipientLabel", "recipientListClosing", "recipientListAfterLabel", _
            "typeName", "subject", "organName", "superiorOrganName", _
            "legalBasis", "signerAuthority", "signerAuthorityTitle")
    Else
        SignalOrder = Array("nationalTitle", "nationalMotto", "codeNumberNotation", _
            "placeAndIssuedDate", "subjectOfficialLetter", _
            "appendixLabel", "appendixTitle", "appendixReference", _
            "recipientSalutation", "recipientSalutationInline", "recipientSalutationList", _
            "recipientLabel", "recipientListClosing", "recipientListAfterLabel", _
            "typeName", "subject", "organName", "superiorOrganName", _
            "legalBasis", "signerAuthority", "signerAuthorityTitle")
    End If
End Function

' Cong van KHONG co doan ten loai va nguoc lai â€” bo qua nhanh khong the co. "khongXacDinh" chay CA
' HAI nhanh (khong chac thi khong loai tru â€” CLAUDE.md muc 5).
Private Function ShouldSkipForDocumentType(ByVal signalKey As String, ByVal documentType As String) As Boolean
    If (documentType = "coTenLoai" Or documentType = "toTrinh" Or documentType = "conLai") _
            And signalKey = "subjectOfficialLetter" Then
        ShouldSkipForDocumentType = True
    ElseIf documentType = "congVan" And (signalKey = "typeName" Or signalKey = "subject") Then
        ShouldSkipForDocumentType = True
    Else
        ShouldSkipForDocumentType = False
    End If
End Function

' ============================================================================
' pham vi 2000 ky tu dau, "chi lay ket qua dau tien", "chi lay ket qua cuoi cung cua van ban
' (khong tinh phu luc)".
' ============================================================================

Private Function ApplyScopeFilters(ByVal signal As Object, ByVal matches As Collection, _
        ByVal layoutMap As Object, ByVal headerLastIndex As Long) As Collection
    Dim headerOnly As Boolean: headerOnly = False
    If signal.Exists("headerWindowOnly") Then headerOnly = CBool(signal("headerWindowOnly"))

    Dim stopIndex As Long: stopIndex = -1
    If signal.Exists("stopAtRole") Then
        Dim stopAt As Variant
        stopAt = FirstIndexWithRole(layoutMap, CStr(signal("stopAtRole")))
        If Not IsNull(stopAt) Then stopIndex = CLng(stopAt)
    End If

    Dim kept As New Collection
    Dim m As Variant
    Dim idx As Long
    For Each m In matches
        idx = CLng(m)
        If headerOnly Then
            If idx > headerLastIndex Then GoTo ContinueMatch
        End If
        If stopIndex >= 0 Then
            If idx >= stopIndex Then GoTo ContinueMatch
        End If
        kept.Add idx
ContinueMatch:
    Next m

    Dim onlyFirst As Boolean: onlyFirst = False
    If signal.Exists("firstMatchOnly") Then onlyFirst = CBool(signal("firstMatchOnly"))
    Dim onlyLast As Boolean: onlyLast = False
    If signal.Exists("lastMatchOnly") Then onlyLast = CBool(signal("lastMatchOnly"))

    If kept.count <= 1 Then
        Set ApplyScopeFilters = kept
        Exit Function
    End If
    If Not (onlyFirst Or onlyLast) Then
        Set ApplyScopeFilters = kept
        Exit Function
    End If

    Dim best As Long: best = CLng(kept(1))
    Dim k As Variant
    For Each k In kept
        If onlyFirst Then
            If CLng(k) < best Then best = CLng(k)
        Else
            If CLng(k) > best Then best = CLng(k)
        End If
    Next k

    Dim oneOnly As New Collection
    oneOnly.Add best
    Set ApplyScopeFilters = oneOnly
End Function

' ============================================================================
' Khop mot dau hieu â€” moi "method" mot ham rieng, tra Collection cac Long ParagraphIndex khop.
' ============================================================================

Private Function MatchSignal(ByVal signal As Object, ByVal paragraphs As Collection, _
        ByVal layoutMap As Object, ByVal assignedSet As Object, ByVal regime As String) As Collection
    Select Case CStr(signal("method"))
        Case "normalizedContains"
            Set MatchSignal = MatchNormalizedContains(signal, paragraphs, assignedSet)
        Case "normalizedFuzzyContains"
            Set MatchSignal = MatchNormalizedFuzzyContains(signal, paragraphs, assignedSet)
        Case "normalizedFuzzyContainsAfterRole"
            Set MatchSignal = MatchNormalizedFuzzyContainsAfterRole(signal, paragraphs, layoutMap, assignedSet)
        Case "regex"
            Set MatchSignal = MatchRegex(signal, paragraphs, assignedSet, regime)
        Case "regexAfterRole"
            Set MatchSignal = MatchRegexAfterRole(signal, paragraphs, layoutMap, assignedSet, regime)
        Case "regexBeforeRole"
            Set MatchSignal = MatchRegexBeforeRole(signal, paragraphs, layoutMap, assignedSet, regime)
        Case "positionalAfterRole"
            Set MatchSignal = MatchPositionalAfterRole(signal, paragraphs, layoutMap, assignedSet)
        Case "positionalAfterRoleStyled"
            Set MatchSignal = MatchPositionalAfterRoleStyled(signal, paragraphs, layoutMap, assignedSet)
        Case "contiguousRunAfterRole"
            Set MatchSignal = MatchContiguousRunAfterRole(signal, paragraphs, layoutMap, assignedSet, regime)
        Case "typeNameDictionary"
            Set MatchSignal = MatchTypeNameDictionary(paragraphs, assignedSet, regime)
        Case "allCapsBeforeRole"
            Set MatchSignal = MatchAllCapsBeforeRole(signal, paragraphs, layoutMap, assignedSet)
        Case "immediatelyBeforeRole"
            Set MatchSignal = MatchImmediatelyBeforeRole(signal, paragraphs, layoutMap, assignedSet)
        Case "anchoredContiguousRun"
            Set MatchSignal = MatchLegalBasis(signal, paragraphs, layoutMap, assignedSet, regime)
        Case Else
            Set MatchSignal = New Collection
    End Select
End Function

' "regexByRegime" (Dictionary khoa ND30/VIETTEL/DANG) uu tien hon "regex" (chuoi don). Tra chuoi
' RONG neu khong co ca hai truong.
Private Function ResolveSignalRegex(ByVal signal As Object, ByVal regime As String) As String
    If signal.Exists("regexByRegime") Then
        Dim byRegime As Object: Set byRegime = signal("regexByRegime")
        If byRegime.Exists(regime) Then
            ResolveSignalRegex = CStr(byRegime(regime))
        ElseIf byRegime.Exists("ND30") Then
            ResolveSignalRegex = CStr(byRegime("ND30"))
        End If
    ElseIf signal.Exists("regex") Then
        ResolveSignalRegex = CStr(signal("regex"))
    End If
End Function

' Nhu ResolveSignalRegex nhung cho MOT truong bat ky (khong chi "regex"/"regexByRegime") â€” dung
' cho cac truong phu nhu "stopIfMatchesRegex"/"stopIfMatchesRegexByRegime" cua subject.
Private Function ResolveNamedRegex(ByVal signal As Object, ByVal baseFieldName As String, _
        ByVal regime As String) As String
    Dim byRegimeField As String
    byRegimeField = baseFieldName & "ByRegime"
    If signal.Exists(byRegimeField) Then
        Dim byRegime As Object: Set byRegime = signal(byRegimeField)
        If byRegime.Exists(regime) Then
            ResolveNamedRegex = CStr(byRegime(regime))
        ElseIf byRegime.Exists("ND30") Then
            ResolveNamedRegex = CStr(byRegime("ND30"))
        End If
    ElseIf signal.Exists(baseFieldName) Then
        ResolveNamedRegex = CStr(signal(baseFieldName))
    End If
End Function

' Khop CHUA (InStr) sau khi chuan hoa â€” Quoc hieu co the nam chung mot doan Word voi Tieu ngu hoac
' gach ngang trang tri qua xuong dong thu cong.
Private Function MatchNormalizedContains(ByVal signal As Object, ByVal paragraphs As Collection, _
        ByVal assignedSet As Object) As Collection
    Dim Result As New Collection
    If Not signal.Exists("normalizedTarget") Then
        Set MatchNormalizedContains = Result
        Exit Function
    End If

    Dim target As String
    target = CStr(signal("normalizedTarget"))

    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If Not assignedSet.Exists(p.Index) Then
            If InStr(NormalizeForNationalTitle(p.text), target) > 0 Then Result.Add p.Index
        End If
    Next p
    Set MatchNormalizedContains = Result
End Function

' Nhu MatchNormalizedContains nhung cho phep lech toi da "maxEditDistance" ky tu chen/xoa/thay
' (Levenshtein) o BAT KY vi tri nao trong cum, khong chi khop CHUA nguyen van â€” dung cho
' nationalTitle: tai lieu that hay go thieu/sai mot vai ky tu cua Quoc hieu.
Private Function MatchNormalizedFuzzyContains(ByVal signal As Object, ByVal paragraphs As Collection, _
        ByVal assignedSet As Object) As Collection
    Dim Result As New Collection
    If Not signal.Exists("normalizedTarget") Or Not signal.Exists("maxEditDistance") Then
        Set MatchNormalizedFuzzyContains = Result
        Exit Function
    End If

    Dim target As String
    target = CStr(signal("normalizedTarget"))
    Dim maxEdits As Long
    maxEdits = CLng(signal("maxEditDistance"))

    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If Not assignedSet.Exists(p.Index) Then
            If FuzzySubstringEditDistance(NormalizedSearchWindow(signal, p.text), target) <= maxEdits Then
                Result.Add p.Index
            End If
        End If
    Next p
    Set MatchNormalizedFuzzyContains = Result
End Function

' Chuan hoa mot doan van cho khop gan dung: bo dau (ke ca "Ä�"/"Ä‘", qua Utils.ToUnaccented â€” KHAC
' NormalizeForNationalTitle, ham do khong bo duoc "Ä�"), bo khoang trang, viet hoa. "maxSearchChars"
' (neu khai bao trong signal): CHI xet N ky tu DAU cua doan van da chuan hoa â€” dung cho nhan/label
' dung o DAU doan (vi du "NÆ¡i nháº­n:"), tranh khop nham mot cau than bai vo tinh nhac toi cum tu do
' o giua doan.
Private Function NormalizedSearchWindow(ByVal signal As Object, ByVal text As String) As String
    Dim normalized As String
    normalized = Utils.ToUpperVn(Replace$(Replace$(Utils.ToUnaccented(text), " ", ""), vbTab, ""))
    If signal.Exists("maxSearchChars") Then
        Dim maxChars As Long: maxChars = CLng(signal("maxSearchChars"))
        If Len(normalized) > maxChars Then normalized = left$(normalized, maxChars)
    End If
    NormalizedSearchWindow = normalized
End Function

' Nhu MatchNormalizedFuzzyContains nhung ket hop VI TRI kieu MatchRegexAfterRole (doan ngay sau
' afterRole, "skipBlankParagraphs" bo qua doan RONG xen giua) â€” dung cho nationalMotto: vua can
' dung vi tri (ngay sau Quoc hieu) vua can khop gan dung noi dung (Tieu ngu bi go sai/thieu).
Private Function MatchNormalizedFuzzyContainsAfterRole(ByVal signal As Object, ByVal paragraphs As Collection, _
        ByVal layoutMap As Object, ByVal assignedSet As Object) As Collection
    Dim Result As New Collection
    If Not signal.Exists("afterRole") Or Not signal.Exists("normalizedTarget") _
            Or Not signal.Exists("maxEditDistance") Then
        Set MatchNormalizedFuzzyContainsAfterRole = Result
        Exit Function
    End If

    Dim target As String: target = CStr(signal("normalizedTarget"))
    Dim maxEdits As Long: maxEdits = CLng(signal("maxEditDistance"))
    Dim skipBlanks As Boolean: skipBlanks = False
    If signal.Exists("skipBlankParagraphs") Then skipBlanks = CBool(signal("skipBlankParagraphs"))

    Dim anchor As Variant
    anchor = FirstIndexWithRole(layoutMap, CStr(signal("afterRole")))
    If IsNull(anchor) Then
        Set MatchNormalizedFuzzyContainsAfterRole = Result
        Exit Function
    End If

    Dim cur As Long
    cur = CLng(anchor) + 1
    Dim p As ParagraphSnapshot
    Do
        If assignedSet.Exists(cur) Then Exit Do
        Set p = ParagraphAtIndex(paragraphs, cur)
        If p Is Nothing Then Exit Do
        If skipBlanks And Len(Trim$(p.text)) = 0 Then
            cur = cur + 1
        Else
            If FuzzySubstringEditDistance(NormalizedSearchWindow(signal, p.text), target) <= maxEdits Then
                Result.Add cur
            End If
            Exit Do
        End If
    Loop

    Set MatchNormalizedFuzzyContainsAfterRole = Result
End Function

' Khoang cach Levenshtein NHO NHAT giua "target" va MOT doan con bat ky cua "text" (khop dau/cuoi
' tu do â€” thuat toan Sellers). Dung cho MatchNormalizedFuzzyContains. O(Len(text)*Len(target)),
' khong dang ngai voi do dai mot doan van/mot cum dau hieu (vai chuc ky tu).
Private Function FuzzySubstringEditDistance(ByVal text As String, ByVal target As String) As Long
    Dim n As Long, m As Long
    n = Len(text)
    m = Len(target)
    If m = 0 Then
        FuzzySubstringEditDistance = 0
        Exit Function
    End If
    If n = 0 Then
        FuzzySubstringEditDistance = m
        Exit Function
    End If

    ' prev/cur: hang thu i cua bang quy hoach dong (i = so ky tu DAU cua target da xet).
    ' prev(j) khoi tao = 0 voi moi j â€” "khop dau tu do": bat dau khop target o BAT KY vi tri nao
    ' trong text khong mat chi phi.
    Dim prev() As Long, cur() As Long
    ReDim prev(0 To n)
    ReDim cur(0 To n)
    Dim j As Long
    For j = 0 To n
        prev(j) = 0
    Next j

    Dim i As Long
    Dim tch As String, xch As String
    Dim costSub As Long, costDel As Long, costIns As Long
    Dim best As Long
    For i = 1 To m
        cur(0) = i
        tch = Mid$(target, i, 1)
        For j = 1 To n
            xch = Mid$(text, j, 1)
            If tch = xch Then
                costSub = prev(j - 1)
            Else
                costSub = prev(j - 1) + 1
            End If
            costDel = prev(j) + 1
            costIns = cur(j - 1) + 1
            cur(j) = costSub
            If costDel < cur(j) Then cur(j) = costDel
            If costIns < cur(j) Then cur(j) = costIns
        Next j
        For j = 0 To n
            prev(j) = cur(j)
        Next j
    Next i

    ' "khop cuoi tu do": ket qua la GTNN tren ca hang cuoi (target co the khop xong o BAT KY vi
    ' tri nao trong text, khong bat buoc phai chay het toi cuoi text).
    best = prev(0)
    For j = 1 To n
        If prev(j) < best Then best = prev(j)
    Next j

    FuzzySubstringEditDistance = best
End Function

Private Function MatchRegex(ByVal signal As Object, ByVal paragraphs As Collection, _
        ByVal assignedSet As Object, ByVal regime As String) As Collection
    Dim Result As New Collection
    Dim pattern As String
    pattern = ResolveSignalRegex(signal, regime)
    If Len(pattern) = 0 Then
        Set MatchRegex = Result
        Exit Function
    End If

    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If Not assignedSet.Exists(p.Index) Then
            If RegexTest(pattern, p.text) Then Result.Add p.Index
        End If
    Next p
    Set MatchRegex = Result
End Function

Private Function MatchRegexAfterRole(ByVal signal As Object, ByVal paragraphs As Collection, _
        ByVal layoutMap As Object, ByVal assignedSet As Object, ByVal regime As String) As Collection
    Dim Result As New Collection
    Dim pattern As String
    pattern = ResolveSignalRegex(signal, regime)
    If Len(pattern) = 0 Or Not signal.Exists("afterRole") Then
        Set MatchRegexAfterRole = Result
        Exit Function
    End If

    Dim afterRole As String
    afterRole = CStr(signal("afterRole"))

    ' "skipBlankParagraphs": bo qua cac doan Word RONG xen giua afterRole va doan can kiem tra
    ' (tai lieu that hay chen dong trong de tao khoang cach trong o bang, vi du giua Quoc hieu
    ' va Tieu ngu) â€” tim doan KHONG RONG dau tien sau afterRole roi moi kiem tra regex, thay vi
    ' doi hoi dung doan ke tiep (index+1).
    Dim skipBlanks As Boolean: skipBlanks = False
    If signal.Exists("skipBlankParagraphs") Then skipBlanks = CBool(signal("skipBlankParagraphs"))

    If skipBlanks Then
        Dim anchor As Variant
        anchor = FirstIndexWithRole(layoutMap, afterRole)
        If IsNull(anchor) Then
            Set MatchRegexAfterRole = Result
            Exit Function
        End If
        Dim cur As Long
        cur = CLng(anchor) + 1
        Dim pSkip As ParagraphSnapshot
        Do
            If assignedSet.Exists(cur) Then Exit Do
            Set pSkip = ParagraphAtIndex(paragraphs, cur)
            If pSkip Is Nothing Then Exit Do
            If Len(Trim$(pSkip.text)) = 0 Then
                cur = cur + 1
            Else
                If RegexTest(pattern, pSkip.text) Then Result.Add cur
                Exit Do
            End If
        Loop
        Set MatchRegexAfterRole = Result
        Exit Function
    End If

    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If assignedSet.Exists(p.Index) Then GoTo ContinueLoop
        If Not layoutMap.Exists(p.Index - 1) Then GoTo ContinueLoop
        If layoutMap(p.Index - 1) <> afterRole Then GoTo ContinueLoop
        If RegexTest(pattern, p.text) Then Result.Add p.Index
ContinueLoop:
    Next p
    Set MatchRegexAfterRole = Result
End Function

' Doi xung MatchRegexAfterRole â€” doan LIEN KE NGAY TRUOC mot doan mang vai tro beforeRole.
Private Function MatchRegexBeforeRole(ByVal signal As Object, ByVal paragraphs As Collection, _
        ByVal layoutMap As Object, ByVal assignedSet As Object, ByVal regime As String) As Collection
    Dim Result As New Collection
    Dim pattern As String
    pattern = ResolveSignalRegex(signal, regime)
    If Len(pattern) = 0 Or Not signal.Exists("beforeRole") Then
        Set MatchRegexBeforeRole = Result
        Exit Function
    End If

    Dim beforeRole As String
    beforeRole = CStr(signal("beforeRole"))

    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If assignedSet.Exists(p.Index) Then GoTo ContinueLoop
        If Not layoutMap.Exists(p.Index + 1) Then GoTo ContinueLoop
        If layoutMap(p.Index + 1) <> beforeRole Then GoTo ContinueLoop
        If RegexTest(pattern, p.text) Then Result.Add p.Index
ContinueLoop:
    Next p
    Set MatchRegexBeforeRole = Result
End Function

' Khoi LIEN TUC cac doan khop regex, bat dau ngay sau doan cuoi cung mang vai tro afterRole. Dung
' cho "recipientSalutationList" (danh sach cac noi kinh gui). Khong khai bao regex/regexByRegime
' nhung co "requireNonEmpty": true thi doi dieu kien dung tu "khop regex" sang "khong rong VA
' chua duoc gan vai tro khac" â€” dung cho "subject": trich yeu ngoai doi co the tach thanh nhieu
' doan Word that (Enter thay vi Shift+Enter), khong co dau hieu noi dung rieng de viet regex.
Private Function MatchContiguousRunAfterRole(ByVal signal As Object, ByVal paragraphs As Collection, _
        ByVal layoutMap As Object, ByVal assignedSet As Object, ByVal regime As String) As Collection
    Dim Result As New Collection
    If Not signal.Exists("afterRole") Then
        Set MatchContiguousRunAfterRole = Result
        Exit Function
    End If

    Dim pattern As String
    pattern = ResolveSignalRegex(signal, regime)

    Dim requireNonEmptyRun As Boolean: requireNonEmptyRun = False
    If signal.Exists("requireNonEmpty") Then requireNonEmptyRun = CBool(signal("requireNonEmpty"))

    If Len(pattern) = 0 And Not requireNonEmptyRun Then
        Set MatchContiguousRunAfterRole = Result
        Exit Function
    End If

    ' "stopIfMatchesRegex"/"...ByRegime": dung het khoi NGAY CA khi doan van khong rong, neu doan
    ' do doc nhu phan mo dau cua mot thanh phan khac dung sau nhung XU LY MUON HON trong
    ' SignalOrder (vi du legalBasis/"CÄƒn cá»©..." dung ngay sau subject KHONG qua dong trong â€” luc
    ' subject chay, legalBasis CHUA duoc gan nen assignedSet chua loai duoc doan do).
    Dim stopPattern As String
    stopPattern = ResolveNamedRegex(signal, "stopIfMatchesRegex", regime)

    Dim anchor As Variant
    anchor = LastIndexWithRole(layoutMap, CStr(signal("afterRole")))
    If IsNull(anchor) Then
        Set MatchContiguousRunAfterRole = Result
        Exit Function
    End If

    Dim cur As Long
    cur = CLng(anchor) + 1
    Dim p As ParagraphSnapshot
    Do
        If assignedSet.Exists(cur) Then Exit Do
        Set p = ParagraphAtIndex(paragraphs, cur)
        If p Is Nothing Then Exit Do
        If Len(stopPattern) > 0 Then
            If RegexTest(stopPattern, p.text) Then Exit Do
        End If
        If Len(pattern) > 0 Then
            If Not RegexTest(pattern, p.text) Then Exit Do
        Else
            If Len(Trim$(p.text)) = 0 Then Exit Do
        End If
        Result.Add cur
        cur = cur + 1
    Loop

    Set MatchContiguousRunAfterRole = Result
End Function

' "legalBasis": khoi LIEN TUC, neo vao doan NGAY SAU doan cuoi cung mang mot trong cac anchorRoles
' (thu theo dung thu tu khai bao). Mot doan "Can cu" dung le loi giua than bai KHONG duoc gan vai
' tro nay.
Private Function MatchLegalBasis(ByVal signal As Object, ByVal paragraphs As Collection, _
        ByVal layoutMap As Object, ByVal assignedSet As Object, ByVal regime As String) As Collection
    Dim Result As New Collection
    Dim pattern As String
    pattern = ResolveSignalRegex(signal, regime)
    If Len(pattern) = 0 Or Not signal.Exists("anchorRoles") Then
        Set MatchLegalBasis = Result
        Exit Function
    End If

    Dim anchorIndex As Long
    anchorIndex = -1

    Dim anchorRole As Variant
    Dim baseIndex As Variant
    Dim candidateIndex As Long
    Dim candidateParagraph As ParagraphSnapshot
    For Each anchorRole In signal("anchorRoles")
        baseIndex = LastIndexWithRole(layoutMap, CStr(anchorRole))
        If Not IsNull(baseIndex) Then
            candidateIndex = CLng(baseIndex) + 1
            If Not assignedSet.Exists(candidateIndex) Then
                Set candidateParagraph = ParagraphAtIndex(paragraphs, candidateIndex)
                If Not candidateParagraph Is Nothing Then
                    If RegexTest(pattern, candidateParagraph.text) Then
                        anchorIndex = candidateIndex
                        Exit For
                    End If
                End If
            End If
        End If
    Next anchorRole

    If anchorIndex = -1 Then
        Set MatchLegalBasis = Result
        Exit Function
    End If

    Result.Add anchorIndex

    Dim cur As Long
    cur = anchorIndex
    Dim nextIndex As Long
    Dim nextParagraph As ParagraphSnapshot
    Do
        nextIndex = cur + 1
        If assignedSet.Exists(nextIndex) Then Exit Do
        Set nextParagraph = ParagraphAtIndex(paragraphs, nextIndex)
        If nextParagraph Is Nothing Then Exit Do
        If Not RegexTest(pattern, nextParagraph.text) Then Exit Do
        Result.Add nextIndex
        cur = nextIndex
    Loop

    Set MatchLegalBasis = Result
End Function

Private Function MatchPositionalAfterRole(ByVal signal As Object, ByVal paragraphs As Collection, _
        ByVal layoutMap As Object, ByVal assignedSet As Object) As Collection
    Dim Result As New Collection
    If Not signal.Exists("afterRole") Then
        Set MatchPositionalAfterRole = Result
        Exit Function
    End If

    Dim anchor As Variant
    anchor = FirstIndexWithRole(layoutMap, CStr(signal("afterRole")))
    If IsNull(anchor) Then
        Set MatchPositionalAfterRole = Result
        Exit Function
    End If
    Dim anchorIndex As Long
    anchorIndex = CLng(anchor)

    ' "uptoRole": chan TREN â€” khong lay doan nao sau dong "Luu...". Khong tim thay dong do thi giu
    ' hanh vi khong chan tren (khong chac thi khong sua).
    Dim hasUpperBound As Boolean: hasUpperBound = False
    Dim upperBoundIndex As Long
    If signal.Exists("uptoRole") Then
        Dim upto As Variant
        upto = LastIndexWithRole(layoutMap, CStr(signal("uptoRole")))
        If Not IsNull(upto) Then
            hasUpperBound = True
            upperBoundIndex = CLng(upto)
        End If
    End If

    ' "skipIfAnchorMatchesRegex": nguoi soan gop ca "Noi nhan:" + danh sach + dong "Luu..." vao
    ' MOT doan Word qua xuong dong thu cong â€” khi do khong con gi de lay THEM.
    If signal.Exists("skipIfAnchorMatchesRegex") Then
        Dim anchorParagraph As ParagraphSnapshot
        Set anchorParagraph = ParagraphAtIndex(paragraphs, anchorIndex)
        If Not anchorParagraph Is Nothing Then
            If RegexTest(CStr(signal("skipIfAnchorMatchesRegex")), anchorParagraph.text) Then
                Set MatchPositionalAfterRole = Result
                Exit Function
            End If
        End If
    End If

    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If p.Index > anchorIndex And Not assignedSet.Exists(p.Index) Then
            If Not hasUpperBound Or p.Index <= upperBoundIndex Then Result.Add p.Index
        End If
    Next p
    Set MatchPositionalAfterRole = Result
End Function

Private Function MatchPositionalAfterRoleStyled(ByVal signal As Object, ByVal paragraphs As Collection, _
        ByVal layoutMap As Object, ByVal assignedSet As Object) As Collection
    Dim Result As New Collection
    If Not signal.Exists("afterRole") Then
        Set MatchPositionalAfterRoleStyled = Result
        Exit Function
    End If

    Dim afterRole As String
    afterRole = CStr(signal("afterRole"))

    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If assignedSet.Exists(p.Index) Then GoTo ContinueLoop
        If Not layoutMap.Exists(p.Index - 1) Then GoTo ContinueLoop
        If layoutMap(p.Index - 1) <> afterRole Then GoTo ContinueLoop
        If MatchesStyleConstraints(signal, p) Then Result.Add p.Index
ContinueLoop:
    Next p
    Set MatchPositionalAfterRoleStyled = Result
End Function

Private Function MatchTypeNameDictionary(ByVal paragraphs As Collection, ByVal assignedSet As Object, _
        ByVal regime As String) As Collection
    Dim Result As New Collection
    Dim match As ParagraphSnapshot
    Set match = DocumentTypeDetector.FindTypeNameParagraph(paragraphs, regime)
    If Not match Is Nothing Then
        If Not assignedSet.Exists(match.Index) Then Result.Add match.Index
    End If
    Set MatchTypeNameDictionary = Result
End Function

' Doan GAN NHAT (chi so lon nhat) truoc beforeRole, khong rong, chua duoc gan vai tro khac â€” xem
' "organName" trong dau-hieu-nhan-dien.json.
Private Function MatchAllCapsBeforeRole(ByVal signal As Object, ByVal paragraphs As Collection, _
        ByVal layoutMap As Object, ByVal assignedSet As Object) As Collection
    Dim Result As New Collection
    If Not signal.Exists("beforeRole") Then
        Set MatchAllCapsBeforeRole = Result
        Exit Function
    End If

    Dim anchor As Variant
    anchor = FirstIndexWithRole(layoutMap, CStr(signal("beforeRole")))
    If IsNull(anchor) Then
        Set MatchAllCapsBeforeRole = Result
        Exit Function
    End If
    Dim anchorIndex As Long
    anchorIndex = CLng(anchor)

    Dim candidate As ParagraphSnapshot
    Set candidate = Nothing
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If p.Index >= anchorIndex Then Exit For
        If assignedSet.Exists(p.Index) Then GoTo ContinueLoop
        If Not MatchesStyleConstraints(signal, p) Then GoTo ContinueLoop
        Set candidate = p
ContinueLoop:
    Next p

    If Not candidate Is Nothing Then Result.Add candidate.Index
    Set MatchAllCapsBeforeRole = Result
End Function

' Doan NGAY TRUOC beforeRole (index - 1) â€” "superiorOrganName". Doan rong hoac da duoc gan vai tro
' khac se khong khop: day chinh la cach loai truong hop khoi ten co quan CHI CO MOT DONG (khi do
' doan index-1 la Quoc hieu/Tieu ngu da gan, hoac khong ton tai).
Private Function MatchImmediatelyBeforeRole(ByVal signal As Object, ByVal paragraphs As Collection, _
        ByVal layoutMap As Object, ByVal assignedSet As Object) As Collection
    Dim Result As New Collection
    If Not signal.Exists("beforeRole") Then
        Set MatchImmediatelyBeforeRole = Result
        Exit Function
    End If

    Dim anchor As Variant
    anchor = FirstIndexWithRole(layoutMap, CStr(signal("beforeRole")))
    If IsNull(anchor) Then
        Set MatchImmediatelyBeforeRole = Result
        Exit Function
    End If

    Dim candidateIndex As Long
    candidateIndex = CLng(anchor) - 1
    If assignedSet.Exists(candidateIndex) Then
        Set MatchImmediatelyBeforeRole = Result
        Exit Function
    End If

    Dim found As ParagraphSnapshot
    Set found = ParagraphAtIndex(paragraphs, candidateIndex)
    If Not found Is Nothing Then
        If MatchesStyleConstraints(signal, found) Then Result.Add found.Index
    End If
    Set MatchImmediatelyBeforeRole = Result
End Function

' ============================================================================
' Tien ich dung chung
' ============================================================================

Private Function FirstIndexWithRole(ByVal layoutMap As Object, ByVal role As String) As Variant
    Dim best As Variant
    best = Null

    Dim key As Variant
    For Each key In layoutMap.Keys
        If layoutMap(key) = role Then
            If IsNull(best) Then
                best = key
            ElseIf CLng(key) < CLng(best) Then
                best = key
            End If
        End If
    Next key

    FirstIndexWithRole = best
End Function

Private Function LastIndexWithRole(ByVal layoutMap As Object, ByVal role As String) As Variant
    Dim best As Variant
    best = Null

    Dim key As Variant
    For Each key In layoutMap.Keys
        If layoutMap(key) = role Then
            If IsNull(best) Then
                best = key
            ElseIf CLng(key) > CLng(best) Then
                best = key
            End If
        End If
    Next key

    LastIndexWithRole = best
End Function

Private Function ParagraphAtIndex(ByVal paragraphs As Collection, ByVal idx As Long) As ParagraphSnapshot
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If p.Index = idx Then
            Set ParagraphAtIndex = p
            Exit Function
        End If
    Next p
    Set ParagraphAtIndex = Nothing
End Function

Private Function MatchesStyleConstraints(ByVal signal As Object, ByVal p As ParagraphSnapshot) As Boolean
    If signal.Exists("requireNonEmpty") Then
        If CBool(signal("requireNonEmpty")) And Len(Trim$(p.text)) = 0 Then
            MatchesStyleConstraints = False
            Exit Function
        End If
    End If
    If signal.Exists("requireAllCaps") Then
        If CBool(signal("requireAllCaps")) And Not p.AllCaps Then
            MatchesStyleConstraints = False
            Exit Function
        End If
    End If
    ' "excludeIfMatchesRegex": loai doan khoi ung vien neu noi dung doc nhu mot thanh phan KHAC
    ' (vi du organName khong duoc nhan doan doc nhu Tieu ngu) â€” dung khi thanh phan do bi go
    ' sai/thieu nen chua duoc gan vai tro rieng, van con "trong" luc dau hieu nay chay.
    If signal.Exists("excludeIfMatchesRegex") Then
        If RegexTest(CStr(signal("excludeIfMatchesRegex")), p.text) Then
            MatchesStyleConstraints = False
            Exit Function
        End If
    End If
    MatchesStyleConstraints = True
End Function

' ============================================================================
' Chuan hoa cho nationalTitle/partyHeader â€” "bo dau cach va dau". VBA khong co ham decompose NFD
' san: DAO NGUOC bang RuleLoader.GetUnicodeToNfc ("combiningToPrecomposed") de anh xa ky tu dung
' san ve chu goc trong MOT lan tra.
' ============================================================================

Private Sub EnsureNationalTitleStripMap()
    If mNationalTitleStripMapReady Then Exit Sub

    Set mNationalTitleStripMap = Utils.NewDictionary()

    Dim mapDict As Object
    Set mapDict = RuleLoader.GetUnicodeToNfc()("combiningToPrecomposed")

    Dim key As Variant
    Dim precomposed As String, baseChar As String
    For Each key In mapDict.Keys
        precomposed = CStr(mapDict(key))
        baseChar = left$(CStr(key), 1)
        If Not mNationalTitleStripMap.Exists(precomposed) Then
            mNationalTitleStripMap.Add precomposed, baseChar
        End If
    Next key

    mNationalTitleStripMapReady = True
End Sub

' Cong khai: RegimeDetector.bas dung lai CHINH ham nay, tranh hai noi tu chuan hoa theo hai cach
' khac nhau roi lech ket qua.
Public Function NormalizeForNationalTitle(ByVal text As String) As String
    EnsureNationalTitleStripMap

    Dim mapped As String
    mapped = ""

    Dim i As Long
    Dim ch As String
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If mNationalTitleStripMap.Exists(ch) Then
            mapped = mapped & mNationalTitleStripMap(ch)
        Else
            mapped = mapped & ch
        End If
    Next i

    Dim noSpace As String
    noSpace = Replace$(mapped, " ", "")
    noSpace = Replace$(noSpace, vbTab, "")

    NormalizeForNationalTitle = Utils.ToUpperVn(noSpace)
End Function

' Loi tren MOT doan khong duoc lam hong ca luot nhan dien thanh phan the thuc - ghi log va coi
' nhu KHONG khop, di tiep doan/dau hieu khac. Cung nguyen tac voi DocumentSnapshot.CaptureImages
' va DocumentTypeDetector.RegexTest (T-71).
Private Function RegexTest(ByVal pattern As String, ByVal s As String) As Boolean
    On Error GoTo ErrHandler
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern
    regex.IgnoreCase = False
    regex.Global = False
    RegexTest = regex.test(s)
    Exit Function
ErrHandler:
    DebugTrace.LogErr "ComponentDetector.RegexTest", "Loi khi kiem tra mau '" & pattern & _
        "' - coi nhu khong khop", Err.number, Err.description
    RegexTest = False
End Function
