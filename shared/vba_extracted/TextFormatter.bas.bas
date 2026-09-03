Attribute VB_Name = "TextFormatter"
'==============================================================
' TextFormatter â€” Phong, co chu, kieu chu, mau chu ap le ngoai bo Styles, va ba nut gian chu
' (â—€ / â¬¤ / â–¶ tren ribbon; khong tach thanh mot CharSpacingFormatter rieng).
' Cac ham nhan THANG mot Word.Range do tang goi truyen vao - KHONG tu quyet dinh pham vi
' Chon/Toan bo, khac ba thu tuc gian chu (tu goi Utils.GetSelectionOnly). Moi ham tu boc
' Utils.BeginOperation/EndOperation - KHONG duoc long nhau (hai bien module-level
' mCurrentOpName/mOpStartTime cua Utils.bas khong tai nhap).
' Moi thu tuc cong khai bat dau bang On Error GoTo ErrHandler. Dung AscW/ChrW, khong dung Asc/Chr
' (phu thuoc code page).
'==============================================================
Option Explicit

' ============================================================================
' ApplyBlackColor -- mau chu ve den, FR-FMT-05
' ============================================================================

' Vai tro "phan tiep theo" khi mot doan Word gop chung hai vai tro qua xuong dong thu cong
' (Shift+Enter).
Private Function ContinuationRoleAfterBreakSize(ByVal role As String) As String
    Select Case role
        Case "recipientLabel": ContinuationRoleAfterBreakSize = "recipientList"
        Case "signerAuthority": ContinuationRoleAfterBreakSize = "signerAuthorityTitle"
        Case "nationalTitle": ContinuationRoleAfterBreakSize = "nationalMotto"
        Case Else: ContinuationRoleAfterBreakSize = ""
    End Select
End Function

' Vai tro cua phan noi dung nam CUNG DONG, ngay sau dau hai cham â€” hien chi co "Kinh gui:"/"Kinh
' trinh:" co noi dung viet lien sau.
Private Function InlineContentRole(ByVal role As String) As String
    Select Case role
        Case "recipientSalutationInline": InlineContentRole = "recipientSalutationInlineContent"
        Case Else: InlineContentRole = ""
    End Select
End Function

' Vi tri (1-based) ky tu dau tien SAU dau hai cham dau tien, bo qua khoang trang. 0 neu khong co.
Private Function InlineContentStartPos(ByVal text As String) As Long
    Dim posColon As Long: posColon = InStr(text, ":")
    If posColon = 0 Then Exit Function

    Dim i As Long: i = posColon + 1
    Dim ch As String
    Do While i <= Len(text)
        ch = Mid$(text, i, 1)
        If ch <> " " And ch <> vbTab Then Exit Do
        i = i + 1
    Loop
    If i > Len(text) Then Exit Function
    InlineContentStartPos = i
End Function

' Vi tri xuong dong thu cong (Chr(11)) NGAY SAU noi dung "nhan" cua doan â€” bo qua khoang trang/
' xuong dong THUA o dau doan truoc (co tai lieu that co Shift+Enter thua truoc ca chu "Noi nhan:",
' tim Chr(11) dau tien mot cach ngay tho se tach trung ngay dau doan).
Private Function FindContinuationBreakPos(ByVal text As String) As Long
    Dim searchFrom As Long: searchFrom = 1
    Dim ch As String
    Do While searchFrom <= Len(text)
        ch = Mid$(text, searchFrom, 1)
        If ch <> Chr(11) And ch <> " " And ch <> vbTab Then Exit Do
        searchFrom = searchFrom + 1
    Loop
    FindContinuationBreakPos = InStr(searchFrom, text, Chr(11))
End Function

Private Function AlignmentValue(ByVal alignment As String) As WdParagraphAlignment
    Select Case alignment
        Case "left": AlignmentValue = wdAlignParagraphLeft
        Case "center": AlignmentValue = wdAlignParagraphCenter
        Case "right": AlignmentValue = wdAlignParagraphRight
        Case "justify": AlignmentValue = wdAlignParagraphJustify
        Case Else: AlignmentValue = wdAlignParagraphLeft
    End Select
End Function

' Dam/nghieng/gach chan/in hoa. Nhan Range (khong phai Paragraph) de dung duoc ca cho phan tiep
' noi nam trong CUNG mot doan.
Private Sub ApplyRoleCharStyle(ByVal rng As word.Range, ByVal spec As Object)
    If spec.Exists("style") Then
        Dim st As Object: Set st = spec("style")
        If st.Exists("bold") Then rng.Font.bold = CBool(st("bold"))
        If st.Exists("italic") Then rng.Font.Italic = CBool(st("italic"))
        If st.Exists("underline") Then
            If CBool(st("underline")) Then
                rng.Font.Underline = wdUnderlineSingle
            Else
                rng.Font.Underline = wdUnderlineNone
            End If
        End If
    End If
    If spec.Exists("letterCase") Then rng.Font.AllCaps = (CStr(spec("letterCase")) = "upper")
End Sub

Private Sub ApplyLineSpacing(ByVal p As word.paragraph, ByVal spec As Object, ByVal lineSpacing As Object)
    Dim zone As String: zone = "fixed"
    If spec.Exists("lineSpacingZone") Then zone = CStr(spec("lineSpacingZone"))

    If zone = "body" Then
        If CStr(lineSpacing("rule")) = "exactly" Then
            Dim pt As Double: pt = CDbl(lineSpacing("exactlyPt"))
            If pt > 0 Then
                p.LineSpacingRule = wdLineSpaceExactly
                p.lineSpacing = pt
                Exit Sub
            End If
        End If
    End If
    p.LineSpacingRule = wdLineSpaceSingle
End Sub

' Canh le + lui dau dong + khoang cach truoc doan + gian dong.
Private Sub ApplyRoleParagraphLayout(ByVal rng As word.Range, ByVal spec As Object, ByVal lineSpacing As Object)
    Dim p As word.paragraph
    For Each p In rng.paragraphs
        If spec.Exists("alignment") Then p.alignment = AlignmentValue(CStr(spec("alignment")))
        If spec.Exists("firstLineIndentCm") Then
            p.LeftIndent = 0
            p.FirstLineIndent = Utils.CmToPoint(CDbl(spec("firstLineIndentCm")))
        End If
        If spec.Exists("spaceBeforePt") Then p.SpaceBefore = CDbl(spec("spaceBeforePt"))
        ApplyLineSpacing p, spec, lineSpacing
    Next p
End Sub

' Times New Roman cho MOI story (than bai, dau/chan trang, hop van ban, chu thich...). Gan ca bon
' "khe" phong chu vi Word giu rieng tung khe cho tung bang ma.
Private Sub ApplyFontNameToStoryChain(ByVal firstStory As word.Range, ByVal fontName As String)
    Dim cur As word.Range
    Set cur = firstStory
    Do While Not cur Is Nothing
        On Error Resume Next
        cur.Font.name = fontName
        cur.Font.NameAscii = fontName
        cur.Font.NameOther = fontName
        cur.Font.NameFarEast = fontName
        cur.Font.NameBi = fontName
        On Error GoTo 0
        Set cur = cur.NextStoryRange
    Loop
End Sub

Private Sub ApplyFontNameWholeDocument()
    Dim fontName As String
    fontName = CStr(RuleLoader.GetFormatSpec()("font")("name"))

    Dim story As word.Range
    For Each story In ActiveDocument.StoryRanges
        ApplyFontNameToStoryChain story, fontName
    Next story
End Sub

' Ten loai van ban (noi dung doan mang vai tro "typeName") â€” can cho ghi de
' "styleByDocumentTypeGroup" cua Viettel. Chuoi rong neu khong nhan dien duoc.
Private Function TypeNameTextFrom(ByVal layoutMap As Object, ByVal indexMap As Object) As String
    Dim key As Variant
    For Each key In layoutMap.Keys
        If CStr(layoutMap(key)) = "typeName" Then
            If indexMap.Exists(CLng(key)) Then
                Dim raw As String
                On Error Resume Next
                raw = ActiveDocument.paragraphs(CLng(indexMap(CLng(key)))).Range.text
                On Error GoTo 0
                raw = Replace$(raw, Chr(13), "")
                raw = Replace$(raw, Chr(7), "")
                TypeNameTextFrom = Trim$(raw)
            End If
            Exit Function
        End If
    Next key
End Function

Public Sub ApplyFontSizeWholeDocument(ByVal sizeSetKey As String, Optional ByVal regime As String = "ND30")
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("TextFormatter.ApplyFontSizeWholeDocument")
    On Error GoTo ErrHandler
    Dim opName As String
    opName = ChrW(&H1EE8) & "ng d" & ChrW(&H1EE5) & "ng c" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & _
        " to" & ChrW(&HE0) & "n v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n"
    Utils.BeginOperation opName

    DebugTrace.Log "TextFormatter.ApplyFontSizeWholeDocument", "Bat dau, sizeSetKey=" & sizeSetKey & " regime=" & regime

    Dim sizeSet As Object: Set sizeSet = RuleLoader.GetEffectiveFontSizeSet(regime, sizeSetKey)
    Dim lineSpacing As Object: Set lineSpacing = RuleLoader.GetRegimeLineSpacing(regime, sizeSetKey)
    Dim bodySize As Double: bodySize = CDbl(sizeSet("bodyText"))

    ApplyFontNameWholeDocument

    Dim layoutMap As Object: Set layoutMap = ComponentFormatter.DetectLayoutMap(regime)
    ' Chi so ban chup KHONG trung chi so ActiveDocument.Paragraphs tren tai lieu co bang â€” phai
    ' dich qua indexMap, khong tu "+1" (xem DocumentSnapshot.BuildSnapshotIndexMap).
    Dim indexMap As Object: Set indexMap = DocumentSnapshot.BuildSnapshotIndexMap()
    Dim docTypeName As String: docTypeName = TypeNameTextFrom(layoutMap, indexMap)

    Dim rolefulWordParaIndices As Object: Set rolefulWordParaIndices = Utils.NewDictionary()
    Dim keyPre As Variant
    For Each keyPre In layoutMap.Keys
        If indexMap.Exists(CLng(keyPre)) Then rolefulWordParaIndices(CLng(indexMap(CLng(keyPre)))) = True
    Next keyPre

    Dim bodySpec As Object
    Set bodySpec = RuleLoader.GetEffectiveComponentSpec(regime, "bodyText", docTypeName)

    Dim appliedCount As Long: appliedCount = 0
    Dim skippedCount As Long: skippedCount = 0
    Dim p As word.paragraph
    Dim wordParaIdx As Long: wordParaIdx = 0
    For Each p In ActiveDocument.paragraphs
        wordParaIdx = wordParaIdx + 1
        On Error GoTo BodyItemError
        p.Range.Font.size = bodySize
        ' Doan trong bang giu nguyen bo cuc (thut le/gian dong cua bang do nguoi soan quyet dinh),
        ' chi nhan co chu va phong chu.
        If Not rolefulWordParaIndices.Exists(wordParaIdx) And Not p.Range.Information(wdWithInTable) Then
            ' KHONG dung toi Bold: doan tieu de Dieu/Khoan nguoi dung tu bold thu cong cung roi
            ' vao nhanh nay (ComponentDetector khong gan vai tro cho cap cau truc) â€” ep Bold=False
            ' se xoa mat dinh dang nguoi dung da lam. Nghieng thi an toan (ND 30 Muc II khoan 6e:
            ' than bai kieu chu dung).
            p.Range.Font.Italic = False
            p.Range.Font.Underline = wdUnderlineNone
            If bodySpec.Exists("alignment") Then
                If p.alignment = wdAlignParagraphLeft Or p.alignment = wdAlignParagraphRight Then
                    p.alignment = AlignmentValue(CStr(bodySpec("alignment")))
                End If
            End If
            If bodySpec.Exists("firstLineIndentCm") Then
                p.LeftIndent = 0
                p.FirstLineIndent = Utils.CmToPoint(CDbl(bodySpec("firstLineIndentCm")))
            End If
            If bodySpec.Exists("spaceBeforePt") Then p.SpaceBefore = CDbl(bodySpec("spaceBeforePt"))
            ApplyLineSpacing p, bodySpec, lineSpacing
        End If
        appliedCount = appliedCount + 1
        GoTo BodyItemDone
BodyItemError:
        skippedCount = skippedCount + 1
        DebugTrace.LogErr "TextFormatter.ApplyFontSizeWholeDocument", "buoc bodyText, doan #" & CStr(wordParaIdx), _
            Err.number, Err.description
        Resume BodyItemDone
BodyItemDone:
        On Error GoTo ErrHandler
    Next p

    Dim key As Variant
    Dim role As String
    For Each key In layoutMap.Keys
        On Error GoTo RoleItemError
        role = CStr(layoutMap(key))
        If role <> "unknown" Then
            If sizeSet.Exists(role) And indexMap.Exists(CLng(key)) Then
                Dim rng As word.Range
                Set rng = ActiveDocument.paragraphs(CLng(indexMap(CLng(key)))).Range
                Dim roleSpec As Object
                Set roleSpec = RuleLoader.GetEffectiveComponentSpec(regime, role, docTypeName)

                rng.Font.size = CDbl(sizeSet(role))
                ApplyRoleCharStyle rng, roleSpec
                ApplyRoleParagraphLayout rng, roleSpec, lineSpacing
                appliedCount = appliedCount + 1

                ' Hai vai tro gop CHUNG mot doan Word qua xuong dong thu cong: phan sau dau
                ' Chr(11) mang vai tro "tiep noi", phai ap rieng (chay SAU nen ghi de dung).
                Dim contRole As String: contRole = ContinuationRoleAfterBreakSize(role)
                If Len(contRole) > 0 And sizeSet.Exists(contRole) Then
                    Dim breakPos As Long: breakPos = FindContinuationBreakPos(rng.text)
                    If breakPos > 0 Then
                        Dim tailRng As word.Range: Set tailRng = rng.Duplicate
                        tailRng.SetRange rng.Start + breakPos, rng.End
                        tailRng.Font.size = CDbl(sizeSet(contRole))
                        ApplyRoleCharStyle tailRng, RuleLoader.GetEffectiveComponentSpec(regime, contRole, docTypeName)
                    End If
                End If

                ' Noi dung viet lien sau dau hai cham cua "Kinh gui:" â€” CUNG doan, tach theo ky
                ' tu.
                Dim inlineRole As String: inlineRole = InlineContentRole(role)
                If Len(inlineRole) > 0 And sizeSet.Exists(inlineRole) Then
                    Dim startPos As Long: startPos = InlineContentStartPos(rng.text)
                    If startPos > 0 Then
                        Dim inlineRng As word.Range: Set inlineRng = rng.Duplicate
                        inlineRng.SetRange rng.Start + startPos - 1, rng.End
                        inlineRng.Font.size = CDbl(sizeSet(inlineRole))
                        ApplyRoleCharStyle inlineRng, RuleLoader.GetEffectiveComponentSpec(regime, inlineRole, docTypeName)
                    End If
                End If
            End If
        End If
        GoTo RoleItemDone
RoleItemError:
        skippedCount = skippedCount + 1
        DebugTrace.LogErr "TextFormatter.ApplyFontSizeWholeDocument", _
            "buoc theo vai tro, doan #" & CStr(key) & " [" & role & "]", Err.number, Err.description
        Resume RoleItemDone
RoleItemDone:
        On Error GoTo ErrHandler
    Next key

    DebugTrace.Log "TextFormatter.ApplyFontSizeWholeDocument", _
        "Hoan tat: applied=" & CStr(appliedCount) & " skipped=" & CStr(skippedCount)
    Utils.EndOperation appliedCount, (skippedCount > 0)
    Exit Sub
ErrHandler:
    DebugTrace.LogErr "TextFormatter.ApplyFontSizeWholeDocument", "loi ngoai vong lap (khong tiep tuc duoc)", _
        Err.number, Err.description
    Utils.AbortOperation Err.description
End Sub

Private Function NoSelectionMessage() As String
    NoSelectionMessage = "Ch" & ChrW(&H1B0) & "a b" & ChrW(&HF4) & "i " & ChrW(&H111) & "en ph" & _
        ChrW(&H1EA7) & "n v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n c" & ChrW(&H1EA7) & "n gi" & _
        ChrW(&HE3) & "n/co ch" & ChrW(&H1EEF) & ". H" & ChrW(&HE3) & "y ch" & ChrW(&H1ECD) & "n " & _
        "(b" & ChrW(&HF4) & "i " & ChrW(&H111) & "en) " & ChrW(&H111) & "o" & ChrW(&H1EA1) & "n c" & _
        ChrW(&H1EA7) & "n " & ChrW(&HE1) & "p r" & ChrW(&H1ED3) & "i b" & ChrW(&H1EA5) & "m l" & _
        ChrW(&H1EA1) & "i n" & ChrW(&HFA) & "t ""<"" ""o"" "">""."
End Function

' Nut 4.1 "<" - giam mot buoc (co chu lai, condensed).
Public Sub CondenseSpacing()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("TextFormatter.CondenseSpacing")
    On Error GoTo ErrHandler
    StepSpacing -1
    Exit Sub
ErrHandler:
    Utils.AbortOperation Err.description
End Sub

' Nut 4.2 "o" - ve 0pt (normal), khong phu thuoc gia tri hien tai.
Public Sub ResetSpacing()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("TextFormatter.ResetSpacing")
    On Error GoTo ErrHandler

    Dim sel As word.Range
    Set sel = Utils.GetSelectionOnly()
    If sel Is Nothing Then
        MsgBoxW.Show NoSelectionMessage(), vbExclamation, "Chu" & ChrW(&H1EA9) & "n h" & ChrW(&HF3) & _
            "a th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
        Exit Sub
    End If

    Utils.BeginOperation "V" & ChrW(&H1EC1) & " Normal"
    sel.Font.Spacing = 0
    Utils.EndOperation 1, False
    Exit Sub
ErrHandler:
    Utils.AbortOperation Err.description
End Sub

' Nut 4.3 ">" - tang mot buoc (gian chu ra, expanded).
Public Sub ExpandSpacing()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("TextFormatter.ExpandSpacing")
    On Error GoTo ErrHandler
    StepSpacing 1
    Exit Sub
ErrHandler:
    Utils.AbortOperation Err.description
End Sub

' Cong/tru mot buoc vao khoang cach ky tu hien tai cua vung chon, kep trong [minPt, maxPt]
' signDirection: -1 = condense, +1 = expand.
' Selection.Font.Spacing tra ve 9999999 (wdUndefined) khi vung chon co nhieu muc gian chu khac
' nhau - PHAI bat truong hop nay va lay 0pt (normal) lam moc, khong duoc cong thang delta vao
' 9999999 (se sinh gia tri rac) - dieu 4 cua.
Private Sub StepSpacing(ByVal signDirection As Integer)
    Dim sel As word.Range
    Set sel = Utils.GetSelectionOnly()
    If sel Is Nothing Then
        MsgBoxW.Show NoSelectionMessage(), vbExclamation, "Chu" & ChrW(&H1EA9) & "n h" & ChrW(&HF3) & _
            "a th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
        Exit Sub
    End If

    Dim spec As Object
    Set spec = RuleLoader.GetFormatSpec()("charSpacing")
    Dim stepPt As Double, minPt As Double, maxPt As Double
    stepPt = CDbl(spec("stepPt"))
    minPt = CDbl(spec("minPt"))
    maxPt = CDbl(spec("maxPt"))

    Dim opName As String
    opName = IIf(signDirection < 0, "Co ch" & ChrW(&H1EEF) & " l" & ChrW(&H1EA1) & "i", _
        "Gi" & ChrW(&HE3) & "n ch" & ChrW(&H1EEF) & " ra")
    Utils.BeginOperation opName

    Dim currentSpacingPt As Double
    Dim rawSpacing As Variant
    rawSpacing = sel.Font.Spacing
    If rawSpacing = wdUndefined Then
        currentSpacingPt = 0
    Else
        currentSpacingPt = CDbl(rawSpacing)
    End If

    Dim newSpacingPt As Double
    newSpacingPt = currentSpacingPt + signDirection * stepPt
    If newSpacingPt < minPt Then newSpacingPt = minPt
    If newSpacingPt > maxPt Then newSpacingPt = maxPt

    sel.Font.Spacing = newSpacingPt
    Utils.EndOperation 1, False
End Sub
