Attribute VB_Name = "IyNormalizer"
Option Explicit

' ============================================================================
' Ky tu â€” dong bo voi ComplianceChecker.IsLetterChar / ToneNormalizer.IsLetterChar (moi module tu
' boc rieng, xem ghi chu dau ComplianceChecker.bas PHAN 5).
' ============================================================================

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

Private Function IsDigitChar(ByVal ch As String) As Boolean
    If ch = "" Then
        IsDigitChar = False
        Exit Function
    End If
    Dim code As Long: code = AscW(ch)
    IsDigitChar = (code >= 48 And code <= 57)
End Function

Private Function IsWordOrDigitChar(ByVal ch As String) As Boolean
    IsWordOrDigitChar = IsLetterChar(ch) Or IsDigitChar(ch)
End Function

Private Function IsWhitespaceChar(ByVal ch As String) As Boolean
    If ch = "" Then
        IsWhitespaceChar = False
        Exit Function
    End If
    Select Case ch
        Case " ", vbTab, vbCr, vbLf
            IsWhitespaceChar = True
        Case Else
            IsWhitespaceChar = False
    End Select
End Function

Private Function IsSentenceEndPunctuation(ByVal ch As String) As Boolean
    IsSentenceEndPunctuation = (ch = "." Or ch = "!" Or ch = "?")
End Function

Private Function IsCapitalized(ByVal s As String) As Boolean
    If Len(s) = 0 Then
        IsCapitalized = False
        Exit Function
    End If
    Dim first As String: first = left$(s, 1)
    IsCapitalized = (first = Utils.ToUpperVn(first)) And (first <> Utils.ToLowerVn(first))
End Function

Private Function IsAllUpper(ByVal s As String) As Boolean
    If Len(s) = 0 Then
        IsAllUpper = False
        Exit Function
    End If
    IsAllUpper = (s = Utils.ToUpperVn(s)) And (s <> Utils.ToLowerVn(s))
End Function

' ============================================================================
' Tach tu â€” cung dang Start(1-based)/Length/Text nhu ToneNormalizer.TokenizeWords
' ============================================================================

Private Function TokenizeWords(ByVal text As String) As Collection
    Dim Result As New Collection
    Dim n As Long: n = Len(text)
    Dim i As Long: i = 1
    Do While i <= n
        If IsLetterChar(Mid$(text, i, 1)) Then
            Dim startPos As Long: startPos = i
            Do While i <= n
                If Not IsLetterChar(Mid$(text, i, 1)) Then Exit Do
                i = i + 1
            Loop
            Dim token As Object: Set token = Utils.NewDictionary()
            token("Start") = startPos
            token("Length") = i - startPos
            token("Text") = Mid$(text, startPos, i - startPos)
            Result.Add token
        Else
            i = i + 1
        End If
    Loop
    Set TokenizeWords = Result
End Function

' Tu o vi tri wordStart (1-based) co phai dau cau khong â€” cung thuat toan
' ToneNormalizer.IsSentenceStart (lui qua khoang trang; het chuoi hoac gap dau ket cau thi coi la
' dau cau).
Private Function IsSentenceStart(ByVal text As String, ByVal wordStart As Long) As Boolean
    Dim i As Long: i = wordStart - 1
    Do While i >= 1
        If Not IsWhitespaceChar(Mid$(text, i, 1)) Then Exit Do
        i = i - 1
    Loop
    If i < 1 Then
        IsSentenceStart = True
    Else
        IsSentenceStart = IsSentenceEndPunctuation(Mid$(text, i, 1))
    End If
End Function

Private Function BuildWordMap(ByVal mapping As Object, ByVal direction As String) As Object
    Dim Result As Object: Set Result = Utils.NewDictionary()
    Dim pairs As Object: Set pairs = mapping("pairs")

    Dim iForm As Variant
    For Each iForm In pairs.Keys
        Dim yForm As String: yForm = CStr(pairs(iForm))
        Dim iWords() As String: iWords = Split(CStr(iForm), " ")
        Dim yWords() As String: yWords = Split(yForm, " ")
        If (UBound(iWords) - LBound(iWords)) <> (UBound(yWords) - LBound(yWords)) Then
            GoTo ContinueEntry ' an toan â€” khong nen xay ra, xem ghi chu dau file
        End If

        Dim idx As Long
        For idx = LBound(iWords) To UBound(iWords)
            If iWords(idx) <> yWords(idx) Then
                If direction = "toY" Then
                    Result(iWords(idx)) = yWords(idx)
                Else
                    Result(yWords(idx)) = iWords(idx)
                End If
            End If
        Next idx
ContinueEntry:
    Next iForm

    Set BuildWordMap = Result
End Function

' ============================================================================
' Danh sach loai tru tuyet doi theo cum co dinh (documentTypeNames, typeAbbreviations,
' nd30Terminology, nationalTitle) â€” docs/rules/07-chinh-ta-qd1989.md muc 3.2
' ============================================================================

' Tim moi vi tri khop NGUYEN CUM phrase trong text, khong phan biet hoa/thuong, bien an toan hai
' dau (IsWordOrDigitChar) â€” cung thuat toan findPhraseSpans ben TS. Tra Collection cua Dictionary
' {"Start" (Long, 1-based), "Length" (Long)}.
Private Function FindPhraseSpans(ByVal text As String, ByVal phrase As String) As Collection
    Dim Result As New Collection
    If Len(phrase) = 0 Then
        Set FindPhraseSpans = Result
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
            item("Start") = foundPos
            item("Length") = Len(phrase)
            Result.Add item
        End If
        fromPos = foundPos + 1
    Loop
    Set FindPhraseSpans = Result
End Function

Private Function CollectProtectedSpans(ByVal text As String, ByVal mapping As Object) As Collection
    Dim Result As New Collection
    Dim ex As Object: Set ex = mapping("excludeAbsolute")

    Dim groupKey As Variant
    For Each groupKey In Array("documentTypeNames", "typeAbbreviations", "nd30Terminology", "nationalTitle")
        Dim terms As Collection: Set terms = ex(CStr(groupKey))("terms")
        Dim term As Variant
        For Each term In terms
            Dim spans As Collection: Set spans = FindPhraseSpans(text, CStr(term))
            Dim s As Variant
            For Each s In spans
                Result.Add s
            Next s
        Next term
    Next groupKey

    Set CollectProtectedSpans = Result
End Function

Private Function IsInsideAnySpan(ByVal token As Object, ByVal spans As Collection) As Boolean
    Dim tokenStart As Long: tokenStart = token("Start")
    Dim tokenEnd As Long: tokenEnd = token("Start") + token("Length") ' vi tri ngay sau token
    Dim s As Variant
    For Each s In spans
        Dim spanStart As Long: spanStart = s("Start")
        Dim spanEnd As Long: spanEnd = s("Start") + s("Length")
        If tokenStart < spanEnd And tokenEnd > spanStart Then
            IsInsideAnySpan = True
            Exit Function
        End If
    Next s
    IsInsideAnySpan = False
End Function

' ============================================================================
' Diem vao cong khai
' ============================================================================

' Quyet dinh tu tai token co nen chuyen hay khong. Tra True va dien convertedOut (giu hoa/thuong
' goc) neu chuyen; tra False (convertedOut khong dung) neu giu nguyen. Thu tu kiem â€” dung "Rang
' buoc" cua:
' 1. Tien to "qu" â€” kiem TRUOC ca khi tra bang cap tu, chan ca truong hop chua liet ke.
' 2. Tra bang cap tu â€” khong co trong bang thi giu nguyen.
' 3. Cum co dinh loai tru tuyet doi â€” tu nam trong mot cum da khop thi giu nguyen.
' 4. Danh tu rieng viet hoa giua cau (khong o dau cau) â€” giu nguyen.
Private Function DecideConversion(ByVal text As String, ByVal token As Object, ByVal wordMap As Object, ByVal mapping As Object, ByVal protectedSpans As Collection, ByRef convertedOut As String) As Boolean
    Dim word As String: word = token("Text")
    Dim lower As String: lower = Utils.ToLowerVn(word)

    Dim ex As Object: Set ex = mapping("excludeAbsolute")

    If CBool(ex("startsWithQu")("enabled")) And left$(lower, 2) = "qu" Then
        DecideConversion = False
        Exit Function
    End If

    If Not wordMap.Exists(lower) Then
        DecideConversion = False
        Exit Function
    End If
    Dim mapped As String: mapped = CStr(wordMap(lower))

    If IsInsideAnySpan(token, protectedSpans) Then
        DecideConversion = False
        Exit Function
    End If

    If CBool(ex("properNouns")("skipCapitalizedMidSentence")) Then
        If IsCapitalized(word) And Not IsSentenceStart(text, token("Start")) Then
            DecideConversion = False
            Exit Function
        End If
    End If

    If IsAllUpper(word) Then
        convertedOut = Utils.ToUpperVn(mapped)
    ElseIf IsCapitalized(word) Then
        convertedOut = Utils.ToUpperVn(left$(mapped, 1)) & Mid$(mapped, 2)
    Else
        convertedOut = mapped
    End If
    DecideConversion = True
End Function

' Chuyen toan bo cap i/y trong text sang direction ("toI" hoac "toY"), co du bon lop loai tru o
' DecideConversion â€” KHONG dung vao doan mang vai tro trong GetExcludedComponentRoles (noi goi
' phai tu loc truoc, xem dau file). Idempotent: goi hai lan lien tiep cung chieu cho ket qua giong
' lan dau, vi sau lan dau khong con tu nao khop CHIEU DO trong bang nua. Tra Dictionary {"text":
' <chuoi da chuyen>, "changedCount": <so tu da doi>}.
Public Function NormalizeIy(ByVal text As String, ByVal direction As String) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("IyNormalizer.NormalizeIy")
    On Error GoTo ErrHandler

    Dim mapping As Object: Set mapping = RuleLoader.GetIyMapping()
    Dim wordMap As Object: Set wordMap = BuildWordMap(mapping, direction)
    Dim protectedSpans As Collection: Set protectedSpans = CollectProtectedSpans(text, mapping)
    Dim tokens As Collection: Set tokens = TokenizeWords(text)

    Dim Result As String: Result = ""
    Dim cursor As Long: cursor = 1
    Dim changedCount As Long: changedCount = 0

    Dim t As Variant
    For Each t In tokens
        Result = Result & Mid$(text, cursor, t("Start") - cursor)
        Dim convertedText As String
        If DecideConversion(text, t, wordMap, mapping, protectedSpans, convertedText) Then
            Result = Result & convertedText
            changedCount = changedCount + 1
        Else
            Result = Result & t("Text")
        End If
        cursor = t("Start") + t("Length")
    Next t
    Result = Result & Mid$(text, cursor)

    Dim out As Object: Set out = Utils.NewDictionary()
    out("text") = Result
    out("changedCount") = changedCount
    Set NormalizeIy = out
    Exit Function
ErrHandler:
    Err.Raise Err.number, "IyNormalizer.NormalizeIy", Err.description
End Function

Private Function AliasComponentRole(ByVal rawRole As String) As String
    If rawRole = "placeName" Then
        AliasComponentRole = "placeAndIssuedDate"
    Else
        AliasComponentRole = rawRole
    End If
End Function

' Danh sach vai tro the thuc KHONG duoc dung toi (Quoc hieu, Tieu ngu, ten co quan, ho ten nguoi
' ky, cum dia danh, so ky hieu van ban) â€” RibbonCallbacks doc de bo qua doan truoc khi goi
' NormalizeIy, xem dau file.
Public Function GetExcludedComponentRoles() As Collection
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("IyNormalizer.GetExcludedComponentRoles")
    On Error GoTo ErrHandler
    Dim Result As New Collection
    Dim rawRole As Variant
    For Each rawRole In RuleLoader.GetIyMapping()("excludeAbsolute")("componentRoles")("roles")
        Result.Add AliasComponentRole(CStr(rawRole))
    Next rawRole
    Set GetExcludedComponentRoles = Result
    Exit Function
ErrHandler:
    Err.Raise Err.number, "IyNormalizer.GetExcludedComponentRoles", Err.description
End Function

' ============================================================================
' ApplyIyStyle â€” nut 6.3/6.4, ap NormalizeIy cho TOAN VAN BAN qua Word Object Model
' ============================================================================

' Ap NormalizeIy cho TUNG DOAN cua ActiveDocument, bo qua doan mang vai tro trong
' GetExcludedComponentRoles â€” cau truc doc/ghi giong het ToneNormalizer.ApplyToneStyle, xem chu
' thich tai do. Nut loai B: bam la ap thang, khong xac nhan tung cho (CLAUDE.md muc 2.2).
Public Sub ApplyIyStyle(ByVal direction As String, ByVal opName As String)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("IyNormalizer.ApplyIyStyle")
    On Error GoTo ErrHandler
    Utils.BeginOperation opName

    Dim snapshot As Object: Set snapshot = DocumentSnapshot.CaptureDocument()
    Dim typeResult As Object: Set typeResult = DocumentTypeDetector.DetectDocumentType(snapshot)
    Dim detection As Object: Set detection = ComponentDetector.DetectComponents(snapshot, CStr(typeResult("Type")))
    Dim layoutMap As Object: Set layoutMap = detection("LayoutMap")

    Dim excludedRoles As Object: Set excludedRoles = Utils.NewDictionary()
    Dim r As Variant
    For Each r In GetExcludedComponentRoles()
        excludedRoles(CStr(r)) = True
    Next r

    Dim indexMap As Object: Set indexMap = DocumentSnapshot.BuildSnapshotIndexMap()

    Dim changedCount As Long: changedCount = 0
    Dim para As Variant
    For Each para In snapshot("Paragraphs")
        Dim idx As Long: idx = para.Index

        If layoutMap.Exists(idx) Then
            If excludedRoles.Exists(CStr(layoutMap(idx))) Then GoTo ContinueParagraph
        End If

        Dim normalized As Object: Set normalized = NormalizeIy(para.text, direction)
        Dim thisChanged As Long: thisChanged = CLng(normalized("changedCount"))
        If thisChanged > 0 And indexMap.Exists(idx) Then
            Dim rng As word.Range
            Set rng = ActiveDocument.paragraphs(CLng(indexMap(idx))).Range
            rng.MoveEnd wdCharacter, -1 ' bo dau ket doan (paragraph mark) khoi vung se ghi de
            rng.text = CStr(normalized("text"))
            changedCount = changedCount + thisChanged
        End If
ContinueParagraph:
    Next para

    Utils.EndOperation changedCount, False
    Exit Sub
ErrHandler:
    Utils.AbortOperation Err.description
End Sub
