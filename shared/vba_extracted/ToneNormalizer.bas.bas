Attribute VB_Name = "ToneNormalizer"
Option Explicit

Private mPlaceNameSet As Object ' Scripting.Dictionary {ten thuong da ha hoa -> True}, nap 1 lan

' ============================================================================
' Ky tu â€” dong bo voi ComplianceChecker.IsLetterChar (moi module tu boc rieng, xem ghi chu dau
' ComplianceChecker.bas PHAN 5).
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

' Chu dau tien cua s la chu hoa (bao gom ca truong hop toan chu hoa).
Private Function IsCapitalized(ByVal s As String) As Boolean
    If Len(s) = 0 Then
        IsCapitalized = False
        Exit Function
    End If
    Dim first As String: first = left$(s, 1)
    IsCapitalized = (first = Utils.ToUpperVn(first)) And (first <> Utils.ToLowerVn(first))
End Function

' Toan bo s la chu hoa (va co it nhat mot chu cai doi duoc hoa/thuong).
Private Function IsAllUpper(ByVal s As String) As Boolean
    If Len(s) = 0 Then
        IsAllUpper = False
        Exit Function
    End If
    IsAllUpper = (s = Utils.ToUpperVn(s)) And (s <> Utils.ToLowerVn(s))
End Function

' ============================================================================
' Phan tich chuoi pattern trong dau-thanh-hai-kieu.json/regex â€” VAN LA DU LIEU TU JSON, chi doc
' cau truc bang tay thay vi chay lookbehind (xem ghi chu dau file). Dang pattern:
' "(?<=[CONSONANTS])(combo1|combo2|...)(?![\p{L}])"
' ============================================================================

Private Sub ParseToneRegexPattern(ByVal pattern As String, ByRef consonantsOut As String, ByRef combosOut As Collection)
    Set combosOut = New Collection

    Dim bracketOpen As Long: bracketOpen = InStr(1, pattern, "[")
    Dim bracketClose As Long: bracketClose = InStr(bracketOpen + 1, pattern, "]")
    consonantsOut = Mid$(pattern, bracketOpen + 1, bracketClose - bracketOpen - 1)

    Dim groupOpen As Long: groupOpen = InStr(bracketClose + 1, pattern, "(")
    Dim groupClose As Long: groupClose = InStr(groupOpen + 1, pattern, ")")
    Dim combosStr As String: combosStr = Mid$(pattern, groupOpen + 1, groupClose - groupOpen - 1)

    Dim parts() As String: parts = Split(combosStr, "|")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        combosOut.Add parts(i)
    Next i
End Sub

' ============================================================================
' Tim vi tri khop to hop â€” thay the cho regex+lookbehind (xem ghi chu dau file).
' ============================================================================

Private Function FindGroupMatches(ByVal text As String, ByVal consonants As String, ByVal combos As Collection) As Collection
    Dim Result As New Collection
    Dim n As Long: n = Len(text)
    Dim lowerText As String: lowerText = Utils.ToLowerVn(text)

    Dim i As Long
    i = 1
    Do While i <= n
        Dim matchedLen As Long: matchedLen = 0
        Dim combo As Variant
        For Each combo In combos
            Dim comboLen As Long: comboLen = Len(CStr(combo))
            If i + comboLen - 1 <= n Then
                If Mid$(lowerText, i, comboLen) = CStr(combo) Then
                    ' Dieu kien 1: to hop phai nam CUOI AM TIET â€” ngay sau khong con chu cai.
                    Dim afterCh As String: afterCh = ""
                    If i + comboLen <= n Then afterCh = Mid$(text, i + comboLen, 1)
                    If Not IsLetterChar(afterCh) Then
                        ' Dieu kien 2: phu am dung truoc phai thuoc lop cho phep (da loai 'q' khoi
                        ' lop nay tu phia JSON cho nhom u).
                        If i > 1 Then
                            Dim beforeLower As String: beforeLower = Mid$(lowerText, i - 1, 1)
                            If InStr(1, consonants, beforeLower, vbBinaryCompare) > 0 Then
                                matchedLen = comboLen
                            End If
                        End If
                    End If
                End If
            End If
            If matchedLen > 0 Then Exit For
        Next combo

        If matchedLen > 0 Then
            Dim item As Object: Set item = Utils.NewDictionary()
            item("Start") = i
            item("Length") = matchedLen
            Result.Add item
            i = i + matchedLen
        Else
            i = i + 1
        End If
    Loop

    Set FindGroupMatches = Result
End Function

' Gop ket qua cua oGroup + uGroup mot chieu, sap theo vi tri (FindGroupMatches da tra ve theo thu
' tu tang dan trong tung nhom, hai nhom khong giao ky tu nen chi can gop hai Collection bang mot
' lan duyet-hop-nhat don gian).
Private Function CollectDirectionMatches(ByVal text As String, ByVal oConsonants As String, ByVal oCombos As Collection, ByVal uConsonants As String, ByVal uCombos As Collection) As Collection
    Dim oMatches As Collection: Set oMatches = FindGroupMatches(text, oConsonants, oCombos)
    Dim uMatches As Collection: Set uMatches = FindGroupMatches(text, uConsonants, uCombos)

    Dim merged As New Collection
    Dim m As Variant
    For Each m In oMatches
        merged.Add m
    Next m
    For Each m In uMatches
        merged.Add m
    Next m

    ' Sap xep don gian (chen tuyen tinh) â€” so luong to hop trong mot doan van rat nho, khong can
    ' thuat toan phuc tap.
    Dim Result As New Collection
    Do While merged.count > 0
        Dim bestIdx As Long: bestIdx = 1
        Dim bestStart As Long: bestStart = merged(1)("Start")
        Dim k As Long
        For k = 2 To merged.count
            If merged(k)("Start") < bestStart Then
                bestStart = merged(k)("Start")
                bestIdx = k
            End If
        Next k
        Result.Add merged(bestIdx)
        merged.Remove bestIdx
    Loop

    Set CollectDirectionMatches = Result
End Function

' Ap bang anh xa cho DUNG chuoi to hop da khop, giu nguyen hoa/thuong cua to hop goc.
Private Function MapCombo(ByVal matchedOriginalCase As String, ByVal mapDict As Object) As String
    Dim lower As String: lower = Utils.ToLowerVn(matchedOriginalCase)
    If Not mapDict.Exists(lower) Then
        MapCombo = matchedOriginalCase ' khong nen xay ra neu FindGroupMatches khop dung combos
        Exit Function
    End If
    Dim mapped As String: mapped = CStr(mapDict(lower))
    If IsAllUpper(matchedOriginalCase) Then
        MapCombo = Utils.ToUpperVn(mapped)
    Else
        MapCombo = mapped
    End If
End Function

' ============================================================================
' Tach tu va tim vi tri â€” dung chung cho quy tac chinh lan tra cum dia danh
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

' Chi so 1-based trong tokens chua vi tri pos, hoac 0 neu khong tim thay.
Private Function FindTokenIndexAt(ByVal tokens As Collection, ByVal pos As Long) As Long
    Dim i As Long
    For i = 1 To tokens.count
        Dim t As Object: Set t = tokens(i)
        If pos >= t("Start") And pos < t("Start") + t("Length") Then
            FindTokenIndexAt = i
            Exit Function
        End If
    Next i
    FindTokenIndexAt = 0
End Function

' Tu o vi tri wordStart (1-based) co phai dau cau khong: lui qua khoang trang, het chuoi (dau
' doan) hoac ky tu lien truoc la dau ket cau thi coi la dau cau.
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

Private Function IsAdjacentBySingleSpace(ByVal text As String, ByVal a As Object, ByVal b As Object) As Boolean
    Dim aEnd As Long: aEnd = a("Start") + a("Length") ' vi tri ngay sau token a
    IsAdjacentBySingleSpace = (b("Start") = aEnd + 1) And (Mid$(text, aEnd, 1) = " ")
End Function

' Cum chu hoa lien ke (cach nhau dung mot dau cach) chua tokens(wordIndex) â€” tra (dau, cuoi) la
' chi so 1-based TRONG tokens.
Private Sub FindCapitalizedCluster(ByVal text As String, ByVal tokens As Collection, ByVal wordIndex As Long, ByRef firstOut As Long, ByRef lastOut As Long)
    Dim first As Long: first = wordIndex
    Dim last As Long: last = wordIndex

    Do While first > 1
        If Not IsAdjacentBySingleSpace(text, tokens(first - 1), tokens(first)) Then Exit Do
        If Not IsCapitalized(tokens(first - 1)("Text")) Then Exit Do
        first = first - 1
    Loop

    Do While last < tokens.count
        If Not IsAdjacentBySingleSpace(text, tokens(last), tokens(last + 1)) Then Exit Do
        If Not IsCapitalized(tokens(last + 1)("Text")) Then Exit Do
        last = last + 1
    Loop

    firstOut = first
    lastOut = last
End Sub

' Ap chuyen doi TOAN CUM sang kieu "toFirstVowel" de tra danh sach dia danh â€” bat ke nut nao dang
' duoc bam, vi dia-danh-viet-nam.json luu o dang do (xem $placeListNote trong JSON).
Private Function ConvertPhraseToFirstVowelStyle(ByVal phrase As String, ByVal mapping As Object) As String
    Dim regexGroup As Object: Set regexGroup = mapping("regex")("toFirstVowel")
    Dim oConsonants As String, oCombos As Collection
    Dim uConsonants As String, uCombos As Collection
    ParseToneRegexPattern CStr(regexGroup("oGroup")), oConsonants, oCombos
    ParseToneRegexPattern CStr(regexGroup("uGroup")), uConsonants, uCombos

    Dim matches As Collection
    Set matches = CollectDirectionMatches(phrase, oConsonants, oCombos, uConsonants, uCombos)

    Dim mapDict As Object: Set mapDict = mapping("mapToFirstVowel")
    Dim Result As String: Result = ""
    Dim cursor As Long: cursor = 1
    Dim m As Variant
    For Each m In matches
        Result = Result & Mid$(phrase, cursor, m("Start") - cursor)
        Result = Result & MapCombo(Mid$(phrase, m("Start"), m("Length")), mapDict)
        cursor = m("Start") + m("Length")
    Next m
    Result = Result & Mid$(phrase, cursor)
    ConvertPhraseToFirstVowelStyle = Result
End Function

' ============================================================================
' Loai tru ten rieng â€” docs/rules/04-loi-go-may.md muc 6.4
' ============================================================================

Private Function GetPlaceNameSet() As Object
    If mPlaceNameSet Is Nothing Then
        Set mPlaceNameSet = Utils.NewDictionary()
        Dim p As Variant
        For Each p In RuleLoader.GetPlaceNames()("places")
            mPlaceNameSet(Utils.ToLowerVn(CStr(p))) = True
        Next p
    End If
    Set GetPlaceNameSet = mPlaceNameSet
End Function

' Quyet dinh mot to hop da khop dieu kien 6.3 (o vi tri matchStart, 1-based) co nen chuyen hay
' khong, theo thuat toan muc 6.4: chu thuong -> chuyen; hoa dau cau -> chuyen; hoa giua cau -> tra
' cum (dai nhat chua, thu toi ngan nhat) trong danh sach dia danh, co match thi chuyen tru khi
' surnameGuard chan (tu lien truoc cum khop la ho nguoi Viet pho bien).
Private Function DecideShouldConvert(ByVal text As String, ByVal tokens As Collection, ByVal matchStart As Long, ByVal mapping As Object) As Boolean
    Dim wordIndex As Long: wordIndex = FindTokenIndexAt(tokens, matchStart)
    If wordIndex = 0 Then
        DecideShouldConvert = True ' khong nen xay ra â€” an toan thi van ap quy tac chung
        Exit Function
    End If

    Dim word As String: word = tokens(wordIndex)("Text")
    Dim rules As Object: Set rules = mapping("properNounHandling")

    If Not IsCapitalized(word) Then
        DecideShouldConvert = (CStr(rules("lowercase")) = "convert")
        Exit Function
    End If

    If IsSentenceStart(text, tokens(wordIndex)("Start")) Then
        DecideShouldConvert = (CStr(rules("capitalizedAtSentenceStart")) = "convert")
        Exit Function
    End If

    ' Viet hoa giua cau â€” tra cum trong danh sach dia danh (checkPlaceNameList).
    Dim placeNames As Object: Set placeNames = GetPlaceNameSet()
    Dim clusterFirst As Long, clusterLast As Long
    FindCapitalizedCluster text, tokens, wordIndex, clusterFirst, clusterLast

    Dim length As Long
    For length = (clusterLast - clusterFirst + 1) To 1 Step -1
        Dim startIdx As Long
        For startIdx = clusterFirst To clusterLast - length + 1
            Dim endIdx As Long: endIdx = startIdx + length - 1
            If wordIndex >= startIdx And wordIndex <= endIdx Then ' cum thu phai chua tu dang xet
                Dim phraseStart As Long: phraseStart = tokens(startIdx)("Start")
                Dim phraseEnd As Long: phraseEnd = tokens(endIdx)("Start") + tokens(endIdx)("Length") - 1
                Dim phrase As String: phrase = Mid$(text, phraseStart, phraseEnd - phraseStart + 1)
                Dim normalizedPhrase As String: normalizedPhrase = ConvertPhraseToFirstVowelStyle(phrase, mapping)

                If placeNames.Exists(Utils.ToLowerVn(normalizedPhrase)) Then
                    Dim guard As Object: Set guard = rules("surnameGuard")
                    If CBool(guard("enabled")) And startIdx > 1 Then
                        Dim precedingWord As String: precedingWord = tokens(startIdx - 1)("Text")
                        Dim surname As Variant
                        For Each surname In guard("surnames")
                            If precedingWord = CStr(surname) Then
                                DecideShouldConvert = False ' ho nguoi Viet pho bien -> ten nguoi
                                Exit Function
                            End If
                        Next surname
                    End If
                    DecideShouldConvert = True ' khop dia danh, khong bi surnameGuard chan
                    Exit Function
                End If
            End If
        Next startIdx
    Next length

    DecideShouldConvert = False ' khong khop dia danh nao -> coi la ten nguoi, giu nguyen
End Function

' ============================================================================
' wordInitialForms â€” to hop o dau tu, khong co phu am dung truoc (docs/rules/04 muc 6.3, ghi
' chu "Khoang trong da biet"; xac nhan CO xu ly theo A4.1, docs/process/cau-hoi-con-mo.md)
' ============================================================================

Private Function ApplyWordInitialForms(ByVal text As String, ByVal mapping As Object, ByVal direction As String) As Object
    Dim wif As Object: Set wif = mapping("wordInitialForms")
    Dim Result As Object: Set Result = Utils.NewDictionary()

    If Not CBool(wif("enabled")) Then
        Result("text") = text
        Result("changedCount") = 0
        Set ApplyWordInitialForms = Result
        Exit Function
    End If

    Dim dict As Object
    If direction = "toMainVowel" Then
        Set dict = wif("toMainVowel")
    Else
        Set dict = wif("toFirstVowel")
    End If

    Dim tokens As Collection: Set tokens = TokenizeWords(text)
    Dim newText As String: newText = ""
    Dim cursor As Long: cursor = 1
    Dim changedCount As Long: changedCount = 0

    Dim i As Long
    For i = 1 To tokens.count
        Dim t As Object: Set t = tokens(i)
        Dim word As String: word = t("Text")
        Dim lower As String: lower = Utils.ToLowerVn(word)

        newText = newText & Mid$(text, cursor, t("Start") - cursor)
        If dict.Exists(lower) Then
            Dim mapped As String: mapped = CStr(dict(lower))
            changedCount = changedCount + 1
            If IsAllUpper(word) Then
                newText = newText & Utils.ToUpperVn(mapped)
            ElseIf IsCapitalized(word) Then
                newText = newText & Utils.ToUpperVn(left$(mapped, 1)) & Mid$(mapped, 2)
            Else
                newText = newText & mapped
            End If
        Else
            newText = newText & word
        End If
        cursor = t("Start") + t("Length")
    Next i
    newText = newText & Mid$(text, cursor)

    Result("text") = newText
    Result("changedCount") = changedCount
    Set ApplyWordInitialForms = Result
End Function

' ============================================================================
' Diem vao cong khai
' ============================================================================

' Chuyen toan bo to hop dau thanh trong text sang direction ("toMainVowel" hoac "toFirstVowel"),
' co loai tru ten rieng â€” KHONG dung vao doan mang vai tro trong GetExcludedComponentRoles (noi
' goi phai tu loc truoc, xem dau file). Idempotent: goi hai lan lien tiep cung chieu cho ket qua
' giong lan dau, vi sau lan dau khong con to hop nao khop CHIEU DO nua. Tra Dictionary {"text":
' <chuoi da chuyen>, "changedCount": <so to hop da doi>}.
Public Function NormalizeTone(ByVal text As String, ByVal direction As String) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ToneNormalizer.NormalizeTone")
    On Error GoTo ErrHandler

    Dim mapping As Object: Set mapping = RuleLoader.GetToneMapping()
    Dim mapDict As Object
    If direction = "toMainVowel" Then
        Set mapDict = mapping("mapToMainVowel")
    Else
        Set mapDict = mapping("mapToFirstVowel")
    End If

    Dim regexGroup As Object: Set regexGroup = mapping("regex")(direction)
    Dim oConsonants As String, oCombos As Collection
    Dim uConsonants As String, uCombos As Collection
    ParseToneRegexPattern CStr(regexGroup("oGroup")), oConsonants, oCombos
    ParseToneRegexPattern CStr(regexGroup("uGroup")), uConsonants, uCombos

    Dim matches As Collection
    Set matches = CollectDirectionMatches(text, oConsonants, oCombos, uConsonants, uCombos)

    Dim tokens As Collection: Set tokens = TokenizeWords(text)

    Dim Result As String: Result = ""
    Dim cursor As Long: cursor = 1
    Dim changedCount As Long: changedCount = 0

    Dim m As Variant
    For Each m In matches
        Dim matchStart As Long: matchStart = m("Start")
        Dim matchLen As Long: matchLen = m("Length")
        Dim shouldConvert As Boolean
        shouldConvert = DecideShouldConvert(text, tokens, matchStart, mapping)

        Result = Result & Mid$(text, cursor, matchStart - cursor)
        Dim matchedOriginal As String: matchedOriginal = Mid$(text, matchStart, matchLen)
        If shouldConvert Then
            Result = Result & MapCombo(matchedOriginal, mapDict)
            changedCount = changedCount + 1
        Else
            Result = Result & matchedOriginal
        End If
        cursor = matchStart + matchLen
    Next m
    Result = Result & Mid$(text, cursor)

    Dim wifResult As Object: Set wifResult = ApplyWordInitialForms(Result, mapping, direction)

    Dim out As Object: Set out = Utils.NewDictionary()
    out("text") = wifResult("text")
    out("changedCount") = changedCount + CLng(wifResult("changedCount"))
    Set NormalizeTone = out
    Exit Function
ErrHandler:
    Err.Raise Err.number, "ToneNormalizer.NormalizeTone", Err.description
End Function

' Danh sach vai tro the thuc KHONG duoc dung toi (Quoc hieu, Tieu ngu, ten co quan, ho ten nguoi
' ky) â€” RibbonCallbacks doc de bo qua doan truoc khi goi NormalizeTone, xem dau file.
Public Function GetExcludedComponentRoles() As Collection
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ToneNormalizer.GetExcludedComponentRoles")
    On Error GoTo ErrHandler
    Set GetExcludedComponentRoles = RuleLoader.GetToneMapping()("excludeComponentRoles")("roles")
    Exit Function
ErrHandler:
    Err.Raise Err.number, "ToneNormalizer.GetExcludedComponentRoles", Err.description
End Function

' ============================================================================
' ApplyToneStyle â€” nut 6.1/6.2, ap NormalizeTone cho TOAN VAN BAN qua Word Object Model
' ============================================================================

Public Sub ApplyToneStyle(ByVal direction As String, ByVal opName As String)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ToneNormalizer.ApplyToneStyle")
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

        Dim normalized As Object: Set normalized = NormalizeTone(para.text, direction)
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
