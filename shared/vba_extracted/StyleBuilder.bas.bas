Attribute VB_Name = "StyleBuilder"
Option Explicit

Private Const SIZE_TOKEN_PLACEHOLDER As String = "${SZ_MAIN}"

' ============================================================================
' BuildStyles - nut 2.2, goi SAU khi frmFontSizeDialog (P8) tra ve lua chon
' ============================================================================

Public Function BuildStyles(ByVal sizeSetKey As String) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("StyleBuilder.BuildStyles")
    On Error GoTo ErrHandler

    Dim opName As String
    opName = "D" & ChrW(&H1EF1) & "ng b" & ChrW(&H1ED9) & " Styles"
    Utils.BeginOperation opName

    Dim sheet As Object
    Set sheet = RuleLoader.GetStyleSheet()

    If Not sheet("sizeToken")("bySet").Exists(sizeSetKey) Then
        Err.Raise vbObjectError + 520, "StyleBuilder.BuildStyles", _
            "Bo co chu """ & sizeSetKey & """ khong co trong bo-styles.json/sizeToken.bySet."
    End If
    Dim sizeInfo As Object
    Set sizeInfo = sheet("sizeToken")("bySet")(sizeSetKey)

    Dim docDefaults As Object
    Set docDefaults = sheet("docDefaults")

    ' styleId -> ten Word that (vd "Heading1" -> "heading 1") - dung cho basedOn/next.
    Dim styleNameById As Object
    Set styleNameById = Utils.NewDictionary()
    Dim def As Variant
    For Each def In sheet("styles")
        styleNameById(CStr(def("styleId"))) = CStr(def("name"))
    Next def

    Dim styleCount As Long
    styleCount = 0
    For Each def In sheet("styles")
        Dim wordStyle As word.Style
        Set wordStyle = GetOrCreateStyle(CStr(def("name")), ToWordStyleType(CStr(def("type"))))
        ApplyStyleDefinition wordStyle, def, docDefaults, sizeInfo, styleNameById
        styleCount = styleCount + 1
    Next def

    ' Dong bo 8 style *Char lien ket (Heading1Char...Heading6Char, TitleChar, SubtitleChar) - CHI
    ' khi da ton tai san (Word tu tao cap link cho style built-in nay), KHONG tu goi Styles.Add
    ' cho chung - giu dung 19 style, doi chieu ghi chu SyncLinkedCharStyles.
    SyncLinkedCharStyles sheet, docDefaults, sizeInfo

    Dim Result As Object
    Set Result = Utils.NewDictionary()
    Result("styleCount") = styleCount
    Result("sizeSetKey") = sizeSetKey

    Utils.EndOperation styleCount, False
    Set BuildStyles = Result
    Exit Function
ErrHandler:
    Utils.AbortOperation Err.description
    Set BuildStyles = Nothing
End Function

' Tra ve style da co (theo ten Word that) hoac tao moi neu chua co - khong bao gio tao style thu
' 20, chi tao dung 19 style khai trong bo-styles.json/styles.
Private Function GetOrCreateStyle(ByVal styleName As String, ByVal wordStyleType As Long) As word.Style
    On Error GoTo ErrHandler
    Dim st As word.Style
    Set st = Nothing
    On Error Resume Next
    Set st = ActiveDocument.Styles(styleName)
    On Error GoTo ErrHandler
    If st Is Nothing Then
        Set st = ActiveDocument.Styles.Add(name:=styleName, Type:=wordStyleType)
    End If
    Set GetOrCreateStyle = st
    Exit Function
ErrHandler:
    Err.Raise Err.number, "StyleBuilder.GetOrCreateStyle", Err.description
End Function

Private Function ToWordStyleType(ByVal styleType As String) As Long
    Select Case styleType
        Case "paragraph"
            ToWordStyleType = wdStyleTypeParagraph
        Case "character"
            ToWordStyleType = wdStyleTypeCharacter
        Case "table"
            ToWordStyleType = wdStyleTypeTable
        Case Else
            Err.Raise vbObjectError + 521, "StyleBuilder.ToWordStyleType", _
                "Kieu style """ & styleType & """ chua duoc ho tro."
    End Select
End Function

' ============================================================================
' Ap mot dinh nghia style (mot phan tu cua bo-styles.json/styles) len mot Word.Style da co
' (existing hoac vua Styles.Add)
' ============================================================================

Private Sub ApplyStyleDefinition(ByVal wordStyle As word.Style, ByVal def As Object, _
        ByVal docDefaults As Object, ByVal sizeInfo As Object, ByVal styleNameById As Object)
    If def.Exists("basedOn") Then
        wordStyle.BaseStyle = ResolveStyleName(CStr(def("basedOn")), styleNameById)
    End If
    If def.Exists("next") Then
        wordStyle.NextParagraphStyle = ResolveStyleName(CStr(def("next")), styleNameById)
    End If
    If CLng(def("uiPriority")) > 0 Then wordStyle.Priority = CLng(def("uiPriority"))
    If def.Exists("qFormat") Then wordStyle.QuickStyle = CBool(def("qFormat"))
    If def.Exists("unhideWhenUsed") Then wordStyle.UnhideWhenUsed = CBool(def("unhideWhenUsed"))
    If def.Exists("semiHidden") Then wordStyle.Visibility = Not CBool(def("semiHidden"))
    If def.Exists("autoRedefine") Then wordStyle.AutomaticallyUpdate = CBool(def("autoRedefine"))

    If def.Exists("paragraphFormat") Then
        ApplyParagraphFormat wordStyle.ParagraphFormat, _
            ResolveParagraphFormat(def("paragraphFormat"), docDefaults)
    End If
    If def.Exists("runFormat") Then
        ApplyRunFormat wordStyle.Font, ResolveRunFormat(def("runFormat"), docDefaults, sizeInfo)
        ' languageId (vi-VN) - phan con lai cua docDefaults.runFormat.lang, ghi tuong minh vi
        ' docDefaults khong ke thua duoc (xem dau file). Thuoc tinh nay nam tren CHINH Word.Style,
        ' khong phai Word.Style.Font - xem ghi chu ApplyRunFormat.
        wordStyle.LanguageID = wdVietnamese
    End If
    If def.Exists("tableFormat") Then
        ApplyTableFormat wordStyle, def("tableFormat"), docDefaults, sizeInfo
    End If
End Sub

Private Function ResolveBuiltinStyleName(ByVal styleId As String) As String
    Select Case styleId
        Case "TableNormal"
            ResolveBuiltinStyleName = "Table Normal"
        Case "DefaultParagraphFont"
            ResolveBuiltinStyleName = "Default Paragraph Font"
        Case Else
            ResolveBuiltinStyleName = styleId
    End Select
End Function

Private Function ResolveStyleName(ByVal styleIdOrName As String, ByVal styleNameById As Object) As String
    If styleNameById.Exists(styleIdOrName) Then
        ResolveStyleName = CStr(styleNameById(styleIdOrName))
    Else
        ResolveStyleName = ResolveBuiltinStyleName(styleIdOrName)
    End If
End Function

' ============================================================================
' paragraphFormat - hop nhat docDefaults.paragraphFormat (nen) voi style.paragraphFormat (ghi de),
' roi ap len Word.ParagraphFormat. Moi gia tri do trong bo-styles.json la TWIP - Utils.TwipToPoint
' quy doi sang point (don vi Word.ParagraphFormat dung).
' ============================================================================

Private Function ResolveParagraphFormat(ByVal partial As Object, ByVal docDefaults As Object) As Object
    Dim baseSpacing As Object
    Set baseSpacing = docDefaults("paragraphFormat")("spacing")

    Dim beforeTwip As Double, afterTwip As Double, lineTwip As Double, lineRule As String
    beforeTwip = CDbl(baseSpacing("before"))
    afterTwip = CDbl(baseSpacing("after"))
    lineTwip = CDbl(baseSpacing("line"))
    lineRule = CStr(baseSpacing("lineRule"))

    If partial.Exists("spacing") Then
        Dim s2 As Object
        Set s2 = partial("spacing")
        If s2.Exists("before") Then beforeTwip = CDbl(s2("before"))
        If s2.Exists("after") Then afterTwip = CDbl(s2("after"))
        If s2.Exists("line") Then lineTwip = CDbl(s2("line"))
        If s2.Exists("lineRule") Then lineRule = CStr(s2("lineRule"))
    End If

    ' Toan bo 19 style deu dung lineRule "auto" - chan som neu bo-styles.json them gia tri khac ma
    ' ham nay chua xu ly, thay vi ap sai lang le.
    If lineRule <> "auto" Then
        Err.Raise vbObjectError + 522, "StyleBuilder.ResolveParagraphFormat", _
            "lineRule """ & lineRule & """ chua duoc ho tro (hien chi xu ly ""auto"")."
    End If

    Dim resolved As Object
    Set resolved = Utils.NewDictionary()

    If partial.Exists("widowControl") Then
        resolved("widowControl") = CBool(partial("widowControl"))
    Else
        resolved("widowControl") = CBool(docDefaults("paragraphFormat")("widowControl"))
    End If
    If partial.Exists("keepNext") Then resolved("keepNext") = CBool(partial("keepNext"))
    If partial.Exists("keepLines") Then resolved("keepLines") = CBool(partial("keepLines"))
    If partial.Exists("jc") Then resolved("jc") = CStr(partial("jc"))
    If partial.Exists("outlineLvl") Then resolved("outlineLvl") = CLng(partial("outlineLvl"))

    resolved("spacingBeforeTwip") = beforeTwip
    resolved("spacingAfterTwip") = afterTwip
    resolved("spacingLineTwip") = lineTwip

    Dim firstLineTwip As Double
    firstLineTwip = 0
    If partial.Exists("indent") Then
        Dim ind As Object
        Set ind = partial("indent")
        If ind.Exists("firstLine") Then firstLineTwip = CDbl(ind("firstLine"))
        If ind.Exists("left") Then resolved("leftIndentTwip") = CDbl(ind("left"))
        If ind.Exists("right") Then resolved("rightIndentTwip") = CDbl(ind("right"))
    End If
    resolved("firstLineIndentTwip") = firstLineTwip

    If partial.Exists("tabs") Then Set resolved("tabs") = partial("tabs")

    Set ResolveParagraphFormat = resolved
End Function

Private Function MapJcToAlignment(ByVal jc As String) As Long
    Select Case jc
        Case "both"
            MapJcToAlignment = wdAlignParagraphJustify
        Case "center"
            MapJcToAlignment = wdAlignParagraphCenter
        Case "left"
            MapJcToAlignment = wdAlignParagraphLeft
        Case "right"
            MapJcToAlignment = wdAlignParagraphRight
        Case Else
            Err.Raise vbObjectError + 523, "StyleBuilder.MapJcToAlignment", _
                "Gia tri jc """ & jc & """ chua co anh xa sang WdParagraphAlignment."
    End Select
End Function

' outlineLvl trong bo-styles.json la 0-based (0..9, khop OOXML w:outlineLvl); WdOutlineLevel la
' 1-based, va wdOutlineLevelBodyText (9 trong bo-styles.json) khong phai cap outline nao.
Private Function MapOutlineLevel(ByVal outlineLvl As Long) As Long
    If outlineLvl = 9 Then
        MapOutlineLevel = wdOutlineLevelBodyText
    ElseIf outlineLvl >= 0 And outlineLvl <= 8 Then
        MapOutlineLevel = outlineLvl + 1
    Else
        Err.Raise vbObjectError + 524, "StyleBuilder.MapOutlineLevel", _
            "outlineLvl " & outlineLvl & " nam ngoai dai 0-9 hop le."
    End If
End Function

Private Function MapTabAlignment(ByVal val As String) As Long
    Select Case val
        Case "right"
            MapTabAlignment = wdAlignTabRight
        Case "left"
            MapTabAlignment = wdAlignTabLeft
        Case "center"
            MapTabAlignment = wdAlignTabCenter
        Case Else
            Err.Raise vbObjectError + 525, "StyleBuilder.MapTabAlignment", _
                "Gia tri tab val """ & val & """ chua co anh xa sang WdTabAlignment."
    End Select
End Function

Private Function MapTabLeader(ByVal leader As String) As Long
    Select Case leader
        Case "dot"
            MapTabLeader = wdTabLeaderDots
        Case "none"
            MapTabLeader = wdTabLeaderSpaces
        Case Else
            Err.Raise vbObjectError + 526, "StyleBuilder.MapTabLeader", _
                "Gia tri tab leader """ & leader & """ chua co anh xa sang WdTabLeader."
    End Select
End Function

Private Sub ApplyParagraphFormat(ByVal target As word.ParagraphFormat, ByVal resolved As Object)
    target.WidowControl = CBool(resolved("widowControl"))
    If resolved.Exists("keepNext") Then target.KeepWithNext = CBool(resolved("keepNext"))
    If resolved.Exists("keepLines") Then target.KeepTogether = CBool(resolved("keepLines"))
    If resolved.Exists("jc") Then target.alignment = MapJcToAlignment(CStr(resolved("jc")))
    If resolved.Exists("outlineLvl") Then target.OutlineLevel = MapOutlineLevel(CLng(resolved("outlineLvl")))

    target.SpaceBefore = Utils.TwipToPoint(CDbl(resolved("spacingBeforeTwip")))
    target.SpaceAfter = Utils.TwipToPoint(CDbl(resolved("spacingAfterTwip")))
    ' Quy uoc CUA Word: voi LineSpacingRule=wdLineSpaceMultiple, LineSpacing tinh bang POINT theo
    ' cong thuc 12pt = mot dong don ("Single") - khop twipToPoint vi toan bo 19 style dung
    ' line=240 twip (don dong).
    target.LineSpacingRule = wdLineSpaceMultiple
    target.lineSpacing = Utils.TwipToPoint(CDbl(resolved("spacingLineTwip")))
    target.FirstLineIndent = Utils.TwipToPoint(CDbl(resolved("firstLineIndentTwip")))
    If resolved.Exists("leftIndentTwip") Then _
        target.LeftIndent = Utils.TwipToPoint(CDbl(resolved("leftIndentTwip")))
    If resolved.Exists("rightIndentTwip") Then _
        target.RightIndent = Utils.TwipToPoint(CDbl(resolved("rightIndentTwip")))

    ' Xoa het tab cu truoc khi ghi lai, de chay lai BuildStyles nhieu lan khong cong don tab stop.
    target.TabStops.ClearAll
    If resolved.Exists("tabs") Then
        Dim tabDef As Variant
        For Each tabDef In resolved("tabs")
            target.TabStops.Add Position:=Utils.TwipToPoint(CDbl(tabDef("pos"))), _
                alignment:=MapTabAlignment(CStr(tabDef("val"))), _
                leader:=MapTabLeader(CStr(tabDef("leader")))
        Next tabDef
    End If
End Sub

' ============================================================================
' runFormat - hop nhat docDefaults.runFormat (nen) voi style.runFormat (ghi de). Dung chung cho
' font cua style chinh, font cua firstRow (ConditionalStyle) va font cua style *Char lien ket.
' ============================================================================

Private Function ResolveRunFormat(ByVal partial As Object, ByVal docDefaults As Object, _
        ByVal sizeInfo As Object) As Object
    Dim resolved As Object
    Set resolved = Utils.NewDictionary()

    If (Not partial Is Nothing) And partial.Exists("font") Then
        resolved("font") = CStr(partial("font"))
    Else
        resolved("font") = CStr(docDefaults("runFormat")("font"))
    End If

    If (Not partial Is Nothing) And partial.Exists("b") Then
        resolved("bold") = CBool(partial("b"))
    Else
        resolved("bold") = False
    End If

    Dim rawSize As Variant
    If (Not partial Is Nothing) And partial.Exists("sz") Then
        rawSize = partial("sz")
    Else
        rawSize = docDefaults("runFormat")("sz")
    End If
    resolved("sizePoint") = ResolveSizeToken(rawSize, sizeInfo) / 2

    Set ResolveRunFormat = resolved
End Function

Private Function ResolveSizeToken(ByVal rawSize As Variant, ByVal sizeInfo As Object) As Double
    If VarType(rawSize) = vbString Then
        If CStr(rawSize) = SIZE_TOKEN_PLACEHOLDER Then
            ResolveSizeToken = CDbl(sizeInfo("halfPoint"))
        Else
            Err.Raise vbObjectError + 527, "StyleBuilder.ResolveSizeToken", _
                "Gia tri co chu """ & rawSize & """ khong phai so va khong phai token."
        End If
    Else
        ResolveSizeToken = CDbl(rawSize)
    End If
End Function

Private Sub ApplyRunFormat(ByVal target As word.Font, ByVal resolved As Object)
    target.name = CStr(resolved("font"))
    target.size = CSng(resolved("sizePoint"))
    target.bold = CBool(resolved("bold"))
    target.Kerning = 0
    target.Ligatures = wdLigaturesNone
    target.Color = wdColorAutomatic
End Sub

' ============================================================================
' tableFormat - chi style TableGrid dung. Vien qua Word.Style.Table.Borders, dinh dang hang dau
' qua Word.Style.Table.Condition(wdFirstRow).
' ============================================================================

Private Function MapBorderLocation(ByVal key As String) As Long
    Select Case key
        Case "top"
            MapBorderLocation = wdBorderTop
        Case "left"
            MapBorderLocation = wdBorderLeft
        Case "bottom"
            MapBorderLocation = wdBorderBottom
        Case "right"
            MapBorderLocation = wdBorderRight
        Case "insideH"
            MapBorderLocation = wdBorderHorizontal
        Case "insideV"
            MapBorderLocation = wdBorderVertical
        Case Else
            Err.Raise vbObjectError + 528, "StyleBuilder.MapBorderLocation", _
                "Vi tri vien """ & key & """ chua co anh xa sang WdBorderType."
    End Select
End Function

Private Function MapBorderLineStyle(ByVal val As String) As Long
    Select Case val
        Case "single"
            MapBorderLineStyle = wdLineStyleSingle
        Case Else
            Err.Raise vbObjectError + 529, "StyleBuilder.MapBorderLineStyle", _
                "Kieu vien """ & val & """ chua co anh xa sang WdLineStyle."
    End Select
End Function

' sz cua vien bang la eighths-of-point (OOXML w:sz cho border, KHAC w:sz co chu la half-point).
' Gia tri WdLineWidth trung khop TRUC TIEP voi sz eighths-of-point (vd wdLineWidth050pt = 4),
' nhung van anh xa tuong minh qua Select Case de bao loi ro rang neu bo-styles.json them gia tri
' chua xu ly, thay vi gan bua so nguyen.
Private Function MapBorderLineWidth(ByVal eighthPoint As Long) As Long
    Select Case eighthPoint
        Case 2
            MapBorderLineWidth = wdLineWidth025pt
        Case 4
            MapBorderLineWidth = wdLineWidth050pt
        Case 6
            MapBorderLineWidth = wdLineWidth075pt
        Case 8
            MapBorderLineWidth = wdLineWidth100pt
        Case 12
            MapBorderLineWidth = wdLineWidth150pt
        Case 18
            MapBorderLineWidth = wdLineWidth225pt
        Case 24
            MapBorderLineWidth = wdLineWidth300pt
        Case 36
            MapBorderLineWidth = wdLineWidth450pt
        Case 48
            MapBorderLineWidth = wdLineWidth600pt
        Case Else
            Err.Raise vbObjectError + 530, "StyleBuilder.MapBorderLineWidth", _
                "Do day vien " & eighthPoint & " (eighths-of-point) chua co anh xa."
    End Select
End Function

Private Sub ApplyTableFormat(ByVal wordStyle As word.Style, ByVal tableFormat As Object, _
        ByVal docDefaults As Object, ByVal sizeInfo As Object)
    If tableFormat.Exists("borders") Then
        Dim borders As Object
        Set borders = tableFormat("borders")
        Dim key As Variant
        For Each key In borders.Keys
            Dim spec As Object
            Set spec = borders(key)
            Dim bdTable As Object
            Set bdTable = wordStyle.table.borders(MapBorderLocation(CStr(key)))
            bdTable.LineStyle = MapBorderLineStyle(CStr(spec("val")))
            bdTable.LineWidth = MapBorderLineWidth(CLng(spec("sz")))
            ' spec("color") luon la "auto" trong bo-styles.json - khong dat, xem ghi chu
            ' ApplyRunFormat.
        Next key
    End If

    If tableFormat.Exists("firstRow") Then
        Dim firstRow As Object
        Set firstRow = tableFormat("firstRow")
        Dim cond As Object
        Set cond = wordStyle.table.Condition(wdFirstRow)
        If firstRow.Exists("runFormat") Then
            ApplyRunFormat cond.Font, ResolveRunFormat(firstRow("runFormat"), docDefaults, sizeInfo)
        End If
        If firstRow.Exists("paragraphFormat") Then
            Dim fpf As Object
            Set fpf = firstRow("paragraphFormat")
            If fpf.Exists("jc") Then cond.ParagraphFormat.alignment = MapJcToAlignment(CStr(fpf("jc")))
        End If
    End If
End Sub

' ============================================================================
' Dong bo 8 style *Char lien ket (Heading1Char...Heading6Char, TitleChar, SubtitleChar)
' ============================================================================

' Ten Word that cua character style lien ket - quy uoc built-in cua Word: "<ten> Char".
Private Function BuildLinkedCharStyleName(ByVal paragraphStyleName As String) As String
    BuildLinkedCharStyleName = paragraphStyleName & " Char"
End Function

Private Sub SyncLinkedCharStyles(ByVal sheet As Object, ByVal docDefaults As Object, ByVal sizeInfo As Object)
    On Error GoTo ErrHandler
    Dim def As Variant
    For Each def In sheet("styles")
        If def.Exists("link") And def.Exists("runFormat") Then
            Dim charName As String
            charName = BuildLinkedCharStyleName(CStr(def("name")))

            Dim st As word.Style
            Set st = Nothing
            On Error Resume Next
            Set st = ActiveDocument.Styles(charName)
            On Error GoTo ErrHandler

            If Not st Is Nothing Then
                ApplyRunFormat st.Font, ResolveRunFormat(def("runFormat"), docDefaults, sizeInfo)
                st.LanguageID = wdVietnamese
            End If
        End If
    Next def
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "StyleBuilder.SyncLinkedCharStyles", Err.description
End Sub
