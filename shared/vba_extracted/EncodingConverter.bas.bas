Attribute VB_Name = "EncodingConverter"
Option Explicit

' ============================================================================
' Khai bao cap module â€” PHAI dung TRUOC toan bo Sub/Function (quy tac VBA; dat xen giua hai
' Function khien "Debug > Compile Project" bao loi, xem ghi chu dau Utils.bas/SafetyGuard.bas).
' ============================================================================

Public Const VNI_CONVERSION_ENABLED As Boolean = False

' Byte thap nhat ma TCVN3/VNI that su dung de ma hoa ky tu co dau â€” duoi nguong nay luon la ASCII
' thuong (chu khong dau, so, khoang trang, dau cau), khong phai loi anh xa.
Private Const SPECIAL_BYTE_MIN As Long = &H80

' Bang chuan hoa NFC dung trong ComposeWithMarks â€” nap mot lan qua EnsureNfcMaps.
Private mCombiningToPrecomposed As Object
Private mPrecomposedToCombining As Object

' Chuoi hien thi tieng Viet. VBA khong cho goi ChrW trong bieu thuc Const nen phai dung bien cap
' module dien qua EnsureTexts (cung cach frmWarning.frm lam). se ra soat lai toan bo chuoi giao
' dien.
Private mTextsReady As Boolean
Private STORY_MAIN As String
Private STORY_HEADER As String
Private STORY_FOOTER As String
Private STORY_FOOTNOTE As String
Private STORY_ENDNOTE As String
Private STORY_TEXTBOX As String
Private PROGRESS_PREFIX As String
Private ERR_PARTIAL_DONE As String
Private ERR_PARTIAL_NONE As String
Private ERR_PARTIAL_FAILED As String

Private Sub EnsureTexts()
    If mTextsReady Then Exit Sub

    STORY_MAIN = "th" & ChrW(&HE2) & "n b" & ChrW(&HE0) & "i"

    STORY_HEADER = ChrW(&H111) & ChrW(&H1EA7) & "u trang"

    STORY_FOOTER = "ch" & ChrW(&HE2) & "n trang"

    STORY_FOOTNOTE = "ch" & ChrW(&HFA) & " th" & ChrW(&HED) & "ch cu"
    STORY_FOOTNOTE = STORY_FOOTNOTE & ChrW(&H1ED1) & "i trang"

    STORY_ENDNOTE = "ch" & ChrW(&HFA) & " th" & ChrW(&HED) & "ch cu"
    STORY_ENDNOTE = STORY_ENDNOTE & ChrW(&H1ED1) & "i v" & ChrW(&H103) & "n b" & ChrW(&H1EA3)
    STORY_ENDNOTE = STORY_ENDNOTE & "n"

    STORY_TEXTBOX = "h" & ChrW(&H1ED9) & "p v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n"

    PROGRESS_PREFIX = ChrW(&H110) & "ang chuy" & ChrW(&H1EC3) & "n " & ChrW(&H111)
    PROGRESS_PREFIX = PROGRESS_PREFIX & ChrW(&H1ED5) & "i b" & ChrW(&H1EA3) & "ng m" & ChrW(&HE3)
    PROGRESS_PREFIX = PROGRESS_PREFIX & ": "

    ERR_PARTIAL_DONE = ChrW(&H110) & ChrW(&HE3) & " chuy" & ChrW(&H1EC3) & "n xong: "

    ERR_PARTIAL_NONE = "Ch" & ChrW(&H1B0) & "a ph" & ChrW(&H1EA7) & "n n"
    ERR_PARTIAL_NONE = ERR_PARTIAL_NONE & ChrW(&HE0) & "o " & ChrW(&H111) & ChrW(&H1B0) & ChrW(&H1EE3)
    ERR_PARTIAL_NONE = ERR_PARTIAL_NONE & "c chuy" & ChrW(&H1EC3) & "n."

    ERR_PARTIAL_FAILED = ChrW(&H110) & "ang d" & ChrW(&H1EDF) & " ph" & ChrW(&H1EA7)
    ERR_PARTIAL_FAILED = ERR_PARTIAL_FAILED & "n: "

    mTextsReady = True
End Sub

Public Function ClassifyFontName(ByVal fontName As String) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("EncodingConverter.ClassifyFontName")
    On Error GoTo ErrHandler

    Dim Result As Object
    Set Result = Utils.NewDictionary()

    Dim trimmed As String
    trimmed = Trim$(fontName)

    If Len(trimmed) = 0 Then
        ' Run khong khai bao font rieng (ke thua mac dinh cua style) - coi la unicode, khong doan
        ' mo theo tang style/docDefaults day du.
        Result("encoding") = "unicode"
        Result("legacyPattern") = Null
        Set ClassifyFontName = Result
        Exit Function
    End If

    Dim patterns As Object
    Set patterns = RuleLoader.GetFontPatterns()("patterns")

    If InCollectionOfStrings(RuleLoader.GetFontPatterns()("excludeFromUpperMatch"), trimmed) Then
        Result("encoding") = "tcvn3"
        Result("legacyPattern") = "tcvn3Lower"
        Set ClassifyFontName = Result
        Exit Function
    End If

    If RegexTest(CStr(patterns("tcvn3Upper")("regex")), trimmed) Then
        Result("encoding") = "tcvn3"
        Result("legacyPattern") = "tcvn3Upper"
        Set ClassifyFontName = Result
        Exit Function
    End If

    If RegexTest(CStr(patterns("tcvn3Lower")("regex")), trimmed) Then
        Result("encoding") = "tcvn3"
        Result("legacyPattern") = "tcvn3Lower"
        Set ClassifyFontName = Result
        Exit Function
    End If

    If RegexTest(CStr(patterns("vni")("regex")), trimmed) Then
        Result("encoding") = "vni"
        Result("legacyPattern") = "vni"
        Set ClassifyFontName = Result
        Exit Function
    End If

    Result("encoding") = "unicode"
    Result("legacyPattern") = Null
    Set ClassifyFontName = Result
    Exit Function
ErrHandler:
    Err.Raise Err.number, "EncodingConverter.ClassifyFontName", Err.description
End Function

' Kiem tra regex ECMAScript bang VBScript.RegExp - engine duy nhat co san trong VBA khong can tham
' chieu thu vien ngoai, va co ho tro lookahead "(?!...)" (can cho mau tcvn3Lower).
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
    Err.Raise Err.number, "EncodingConverter.RegexTest", Err.description
End Function

Private Function InCollectionOfStrings(ByVal coll As Collection, ByVal s As String) As Boolean
    Dim item As Variant
    For Each item In coll
        If CStr(item) = s Then
            InCollectionOfStrings = True
            Exit Function
        End If
    Next item
    InCollectionOfStrings = False
End Function

' ============================================================================
' AggregateRuns â€” gop danh sach run tho (da co ten phong + so ky tu) thanh ket qua nhan dien
' toan tai lieu.
' - Bo qua run rong (charCount = 0).
' - nonUnicodeCount chi cong don run KHONG PHAI unicode.
' - encoding tong hop CHI xet run khong phai unicode: khong co -> "unicode"; dung mot loai -> dung
'   loai do; tu hai loai tro len (vd.VnTime lan VNI-Times) -> "mixed". Run unicode xen giua (vd
'   so/ky hieu van dung font Unicode ngay trong van ban TCVN3) KHONG keo ve "mixed".
' rawRuns: Collection cac Dictionary {"paragraphIndex": Variant, "fontName": String, "charCount":
' Long} - dau vao THUAN TUY, khong tu di quet tai lieu (DetectEncoding o duoi moi la ham quet tai
' lieu that).
' ============================================================================
Public Function AggregateRuns(ByVal rawRuns As Collection) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("EncodingConverter.AggregateRuns")
    On Error GoTo ErrHandler

    Dim classifiedRuns As New Collection
    Dim nonUnicodeCount As Long
    Dim hasTcvn3 As Boolean, hasVni As Boolean

    Dim raw As Variant
    For Each raw In rawRuns
        If CLng(raw("charCount")) > 0 Then
            Dim cls As Object
            Set cls = ClassifyFontName(CStr(raw("fontName")))

            Dim outRun As Object
            Set outRun = Utils.NewDictionary()
            outRun("paragraphIndex") = raw("paragraphIndex")
            outRun("fontName") = raw("fontName")
            outRun("charCount") = raw("charCount")
            outRun("encoding") = cls("encoding")
            outRun("legacyPattern") = cls("legacyPattern")
            classifiedRuns.Add outRun

            If cls("encoding") <> "unicode" Then
                nonUnicodeCount = nonUnicodeCount + CLng(raw("charCount"))
                If cls("encoding") = "tcvn3" Then hasTcvn3 = True
                If cls("encoding") = "vni" Then hasVni = True
            End If
        End If
    Next raw

    Dim docEncoding As String
    If hasTcvn3 And hasVni Then
        docEncoding = "mixed"
    ElseIf hasTcvn3 Then
        docEncoding = "tcvn3"
    ElseIf hasVni Then
        docEncoding = "vni"
    Else
        docEncoding = "unicode"
    End If

    Dim Result As Object
    Set Result = Utils.NewDictionary()
    Result("encoding") = docEncoding
    Set Result("runs") = classifiedRuns
    Result("nonUnicodeCount") = nonUnicodeCount
    Set AggregateRuns = Result
    Exit Function
ErrHandler:
    Err.Raise Err.number, "EncodingConverter.AggregateRuns", Err.description
End Function

' ============================================================================
' DetectEncoding â€” quet TOAN TAI LIEU (FR-ENC-01/06): than bai + bang bieu (doan trong bang
' nam san trong Paragraphs cua story chinh, khong can xu ly rieng), dau trang, chan trang, chu
' thich (footnote/endnote). Chi DOC, khong sua gi tai lieu - khong rui ro Undo.
' ============================================================================
Public Function DetectEncoding() As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("EncodingConverter.DetectEncoding")
    On Error GoTo ErrHandler

    Dim allRuns As New Collection
    Dim story As Range
    Dim cur As Range

    For Each story In ActiveDocument.StoryRanges
        Set cur = story
        Do While Not cur Is Nothing
            If ShouldScanStoryType(cur.StoryType) Then
                CollectRunsFromStory cur, allRuns
            End If
            Set cur = cur.NextStoryRange
        Loop
    Next story

    Set DetectEncoding = AggregateRuns(allRuns)
    Exit Function
ErrHandler:
    Err.Raise Err.number, "EncodingConverter.DetectEncoding", Err.description
End Function

Private Function ShouldScanStoryType(ByVal st As WdStoryType) As Boolean
    Select Case st
        Case wdMainTextStory, wdFootnotesStory, wdEndnotesStory, _
             wdPrimaryHeaderStory, wdPrimaryFooterStory, _
             wdEvenPagesHeaderStory, wdEvenPagesFooterStory, _
             wdFirstPageHeaderStory, wdFirstPageFooterStory, _
             wdTextFrameStory
            ShouldScanStoryType = True
        Case Else
            ShouldScanStoryType = False
    End Select
End Function

' Gom run cua mot story vao allRuns. Story chinh (wdMainTextStory) duyet tung Paragraph de gan
' paragraphIndex 0-based dung quy uoc toan du an; cac story khac (dau/chan trang, chu thich) gom
' ca story thanh mot lan quet voi paragraphIndex = Null (khong co chi so doan trong than bai de
' tro toi) - khoi luong chu o day thuong nho, chap nhan duoc ve hieu nang.
Private Sub CollectRunsFromStory(ByVal storyRange As Range, ByVal allRuns As Collection)
    On Error GoTo ErrHandler

    If storyRange.StoryType = wdMainTextStory Then
        Dim p As paragraph
        Dim idx As Long
        idx = 0
        For Each p In storyRange.paragraphs
            Dim pr As Range
            Set pr = p.Range.Duplicate
            TrimTrailingParaMark pr

            Dim runsInPara As Collection
            Set runsInPara = CollectFontRuns(pr)

            Dim r As Variant
            For Each r In runsInPara
                r("paragraphIndex") = idx
                allRuns.Add r
            Next r

            idx = idx + 1
        Next p
    Else
        Dim sr As Range
        Set sr = storyRange.Duplicate
        TrimTrailingParaMark sr

        Dim runsInStory As Collection
        Set runsInStory = CollectFontRuns(sr)

        Dim rr As Variant
        For Each rr In runsInStory
            rr("paragraphIndex") = Null
            allRuns.Add rr
        Next rr
    End If
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "EncodingConverter.CollectRunsFromStory", Err.description
End Sub

' Bo dau doan (Chr(13)) o cuoi Range neu co - tranh no lan vao charCount va bi tinh font rieng cua
' ky tu dau doan.
Private Sub TrimTrailingParaMark(ByVal rng As Range)
    If Len(rng.text) > 0 Then
        If Right$(rng.text, 1) = vbCr Then
            rng.MoveEnd wdCharacter, -1
        End If
    End If
End Sub

' ============================================================================
' CollectFontRuns â€” doc run (ten phong + so ky tu) cua MOT Range, khong tu chia theo doan.
' rng.Font.Name tra "" (chuoi rong, KHONG phai Null) khi vung co nhieu phong - luc do chia nho
' xuong tung ky tu (Characters), gop lien tiep cung phong thanh mot run. Chi xay ra o pham vi hep
' (mot doan, hoac mot story dau/chan trang/chu thich nho) nen chap nhan duoc ve hieu nang - KHONG
' bao gio chay tren toan bo than bai cung luc.
' ============================================================================
Private Function CollectFontRuns(ByVal srcRange As Range) As Collection
    On Error GoTo ErrHandler
    Dim Result As New Collection

    Dim rng As Range
    Set rng = srcRange.Duplicate

    If Len(rng.text) = 0 Then
        Set CollectFontRuns = Result
        Exit Function
    End If

    Dim fName As String
    fName = rng.Font.name

    If fName <> "" Then
        AddFontRunItem Result, fName, Len(rng.text)
        Set CollectFontRuns = Result
        Exit Function
    End If

    Dim n As Long
    n = rng.Characters.count

    Dim curFont As String, curCount As Long, i As Long
    curCount = 0
    For i = 1 To n
        Dim chFont As String
        chFont = rng.Characters(i).Font.name
        If curCount = 0 Then
            curFont = chFont
            curCount = 1
        ElseIf chFont = curFont Then
            curCount = curCount + 1
        Else
            AddFontRunItem Result, curFont, curCount
            curFont = chFont
            curCount = 1
        End If
    Next i
    If curCount > 0 Then AddFontRunItem Result, curFont, curCount

    Set CollectFontRuns = Result
    Exit Function
ErrHandler:
    Err.Raise Err.number, "EncodingConverter.CollectFontRuns", Err.description
End Function

Private Sub AddFontRunItem(ByVal target As Collection, ByVal fontName As String, ByVal charCount As Long)
    Dim item As Object
    Set item = Utils.NewDictionary()
    item("fontName") = fontName
    item("charCount") = charCount
    target.Add item
End Sub

' Phong dich sau khi chuyen â€” Phu luc I Muc I khoan 4 (FR-ENC-07). Doc tu RuleData.FONT_NAME (sinh
' tu shared/rules/thong-so-the-thuc.json/font.name) chu KHONG viet hang so tai cho, CLAUDE.md muc
' 3.1.
Private Function TargetFontName() As String
    TargetFontName = RuleData.FONT_NAME
End Function

' AscW tra Integer CO DAU (-32768..32767) - ma > 0x7FFF ve am. Cong 65536 de lay lai gia tri khong
' dau. Dung AscW, cam Asc (phu thuoc ma trang) - CLAUDE.md muc 5.
Private Function CodePointOf(ByVal ch As String) As Long
    Dim c As Long
    c = AscW(ch)
    If c < 0 Then c = c + 65536
    CodePointOf = c
End Function

Private Sub EnsureNfcMaps()
    If Not mCombiningToPrecomposed Is Nothing Then Exit Sub

    Set mCombiningToPrecomposed = RuleLoader.GetUnicodeToNfc()("combiningToPrecomposed")

    Set mPrecomposedToCombining = Utils.NewDictionary()
    Dim k As Variant
    For Each k In mCombiningToPrecomposed.Keys
        Dim v As String
        v = CStr(mCombiningToPrecomposed(k))
        If Not mPrecomposedToCombining.Exists(v) Then
            mPrecomposedToCombining(v) = CStr(k)
        End If
    Next k
End Sub

Private Function LookupPrecomposed(ByVal baseChar As String, ByVal marks As String) As String
    Dim key As String
    key = baseChar & marks
    If mCombiningToPrecomposed.Exists(key) Then
        LookupPrecomposed = CStr(mCombiningToPrecomposed(key))
    Else
        LookupPrecomposed = ""
    End If
End Function

' prev la mot "manh" da dung o vi tri lien truoc - gan nhu luon dung mot ky tu (gia tri cua bang
' direct deu la mot ky tu). Chi ghep vao KY TU CUOI cua prev, giu nguyen phan dau, de truong hop
' hiem (mot lan ghep truoc do that bai va de lai <goc>+<dau>) khong lam hong ca manh.
Private Function ComposeWithMarks(ByVal prev As String, ByVal marks As String) As String
    On Error GoTo ErrHandler
    EnsureNfcMaps

    Dim baseChar As String, leading As String
    baseChar = Right$(prev, 1)
    leading = left$(prev, Len(prev) - 1)

    Dim composed As String
    composed = LookupPrecomposed(baseChar, marks)
    If Len(composed) > 0 Then
        ComposeWithMarks = leading & composed
        Exit Function
    End If

    If mPrecomposedToCombining.Exists(baseChar) Then
        Dim nfdBase As String
        nfdBase = CStr(mPrecomposedToCombining(baseChar))

        Dim letterPart As String, baseMarks As String
        letterPart = left$(nfdBase, 1)
        baseMarks = Mid$(nfdBase, 2)

        composed = LookupPrecomposed(letterPart, baseMarks & marks)
        If Len(composed) > 0 Then
            ComposeWithMarks = leading & composed
            Exit Function
        End If

        composed = LookupPrecomposed(letterPart, marks & baseMarks)
        If Len(composed) > 0 Then
            ComposeWithMarks = leading & composed
            Exit Function
        End If
    End If

    ComposeWithMarks = prev & marks
    Exit Function
ErrHandler:
    Err.Raise Err.number, "EncodingConverter.ComposeWithMarks", Err.description
End Function

' ============================================================================
' ApplyEncodingTable â€” ap MOT bang anh xa (da quy ve {direct, combining}) len mot chuoi, theo
' dung bon buoc cua $thuatToan trong bang-ma-vni.json:
' 1. Duyet chuoi tu trai sang phai, xet tung ky tu mot.
' 2. Khop trong "direct" -> thay bang chuoi tuong ung. chuoi dai nam o GIA TRI, khong nam o khoa.)
' 3. Khop trong "combining" -> noi chuoi to hop dau vao SAU ky tu vua dung o vi tri lien truoc,
'   chuan hoa NFC, thay cho CA CAP hai ky tu.
' 4. Khong khop gi, hoac khop "combining" nhung khong co ky tu lien truoc -> GIU NGUYEN, khong
'   doan; chi tinh vao unmappedCount neu ma >= 0x80 (ASCII thuong khong tinh).
' Tra Dictionary {"text": <chuoi ket qua>, "unmappedCount": <so ky tu giu nguyen>}
' ============================================================================
Public Function ApplyEncodingTable(ByVal text As String, ByVal table As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("EncodingConverter.ApplyEncodingTable")
    On Error GoTo ErrHandler

    Dim directMap As Object, combiningMap As Object
    Set directMap = table("direct")
    Set combiningMap = table("combining")

    Dim n As Long
    n = Len(text)

    ' Moi vi tri sinh toi da MOT phan tu trong pieces (ket hop dau chi GOP vao phan tu da co,
    ' khong them phan tu moi) â€” kich thuoc n la chan tren an toan, giong cach
    ' UnicodeNormalizer.NormalizeNfc dung. O chua dung giu vbNullString, Join bo qua nen khong can
    ' cat.
    Dim pieces() As String
    If n > 0 Then ReDim pieces(0 To n - 1)
    Dim pieceCount As Long
    pieceCount = 0

    Dim unmappedCount As Long
    unmappedCount = 0

    Dim i As Long
    For i = 1 To n
        Dim ch As String
        ch = Mid$(text, i, 1)

        Dim code As Long
        code = CodePointOf(ch)

        If code > &HFF Then
            ' Ngoai dai 0x00-0xFF: chac chan khong phai byte bang ma cu (TCVN3/VNI chi dung 8-bit)
            ' - giu nguyen, khong xet bang, khong tinh unmapped.
            pieces(pieceCount) = ch
            pieceCount = pieceCount + 1
        Else
            Dim key As String
            key = Right$("0" & Hex$(code), 2)

            If directMap.Exists(key) Then
                pieces(pieceCount) = CStr(directMap(key))
                pieceCount = pieceCount + 1
            ElseIf combiningMap.Exists(key) Then
                If pieceCount > 0 Then
                    pieces(pieceCount - 1) = ComposeWithMarks(pieces(pieceCount - 1), CStr(combiningMap(key)))
                Else
                    ' Buoc 4: dau thanh dung dau chuoi, khong co ky tu de ghep - giu nguyen, ghi
                    ' nhat ky (luon la byte >= 0x80 vi combining chi chua khoa trong dai do).
                    pieces(pieceCount) = ch
                    pieceCount = pieceCount + 1
                    unmappedCount = unmappedCount + 1
                End If
            Else
                pieces(pieceCount) = ch
                pieceCount = pieceCount + 1
                If code >= SPECIAL_BYTE_MIN Then unmappedCount = unmappedCount + 1
            End If
        End If
    Next i

    Dim Result As Object
    Set Result = Utils.NewDictionary()
    If n > 0 Then
        Result("text") = Join(pieces, vbNullString)
    Else
        Result("text") = text
    End If
    Result("unmappedCount") = unmappedCount
    Set ApplyEncodingTable = Result
    Exit Function
ErrHandler:
    Err.Raise Err.number, "EncodingConverter.ApplyEncodingTable", Err.description
End Function

' Chuyen mot chuoi theo DUNG bang ung voi legacyPattern (tra qua RuleLoader.GetEncodingTable -
' nguon chan ly shared/rules/bang-ma-*.json). Bang luon co san tu; nhanh "khong co bang" chi la
' phong ve - tra nguyen van khong doi, an toan hon nem loi giua chung mot thao tac loai B.
Public Function ConvertRunText(ByVal text As String, ByVal legacyPattern As String) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("EncodingConverter.ConvertRunText")
    On Error GoTo ErrHandler

    Dim table As Object
    Set table = RuleLoader.GetEncodingTable(legacyPattern)

    If table Is Nothing Then
        Dim unchanged As Object
        Set unchanged = Utils.NewDictionary()
        unchanged("text") = text
        unchanged("unmappedCount") = 0
        Set ConvertRunText = unchanged
        Exit Function
    End If

    Set ConvertRunText = ApplyEncodingTable(text, table)
    Exit Function
ErrHandler:
    Err.Raise Err.number, "EncodingConverter.ConvertRunText", Err.description
End Function

' ============================================================================
' ConvertRun â€” phan loai font truoc (ClassifyFontName), roi:
' - "unicode" -> giu nguyen (FR-ENC-09, da dung chuan), skipped = "alreadyUnicode".
' - "vni" ma VNI_CONVERSION_ENABLED = False -> giu nguyen, skipped = "vniDisabled" (coi nhu "bang
'   ma khong ho tro" theo dung tinh than FR-ENC-08, KHONG phai loi cua nguoi dung).
' - con lai -> tra bang theo legacyPattern, doi phong dich.
' Tra Dictionary {"encoding", "text", "newFontName", "unmappedCount", "skipped"}. newFontName la
' chuoi rong khi giu nguyen phong goc; skipped la chuoi rong khi da chuyen that.
' ============================================================================
Public Function ConvertRun(ByVal fontName As String, ByVal text As String) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("EncodingConverter.ConvertRun")
    On Error GoTo ErrHandler

    Dim cls As Object
    Set cls = ClassifyFontName(fontName)

    Dim Result As Object
    Set Result = Utils.NewDictionary()
    Result("encoding") = cls("encoding")
    Result("text") = text
    Result("newFontName") = ""
    Result("unmappedCount") = 0
    Result("skipped") = ""

    If cls("encoding") = "unicode" Then
        Result("skipped") = "alreadyUnicode"
        Set ConvertRun = Result
        Exit Function
    End If

    If cls("encoding") = "vni" And Not VNI_CONVERSION_ENABLED Then
        Result("skipped") = "vniDisabled"
        Set ConvertRun = Result
        Exit Function
    End If

    Dim converted As Object
    Set converted = ConvertRunText(text, CStr(cls("legacyPattern")))
    Result("text") = converted("text")
    Result("unmappedCount") = converted("unmappedCount")
    Result("newFontName") = TargetFontName()

    Set ConvertRun = Result
    Exit Function
ErrHandler:
    Err.Raise Err.number, "EncodingConverter.ConvertRun", Err.description
End Function

' ============================================================================
' ConvertToUnicode â€” nut 1.2, thao tac RUI RO CAO NHAT cua nhom 1 (ADR-007). CHI duoc goi SAU
' KHI nguoi dung da xac nhan canh bao P4 (frmWarning.ShowHighRisk) - xem
' RibbonCallbacks.OnChuyenDoiUnicode. Ham nay KHONG tu hien canh bao.
' Hai pha tach bach: PHA 1 â€” GOM (chi doc): duyet StoryRanges, cat tung story thanh cac "run
' chuyen doi duoc" (cung phong, khong chua ky tu dieu khien), giu lai Range + van ban + ten phong.
' PHA 2 â€” GHI: voi tung story, duyet run theo THU TU NGUOC roi ghi Range.Text va Range.Font.Name.
' Vi sao thu tu nguoc: chuoi ket qua thuong NGAN HON chuoi goc (moi cap <chu cai>+<dau thanh> gop
' lai con mot ky tu), nen moi lan ghi lam lech vi tri cua phan phia sau. Di tu cuoi ve dau thi cac
' run chua xu ly deu nam TRUOC cho vua sua, vi tri khong doi. Gom truoc - ghi sau cung tranh viec
' sua tai lieu ngay giua mot vong For Each tren Paragraphs.
' FR-SAF-07 (loi giua chung): VBA khong co giao dich. Bu lai bang hai thu, dung nhu "Rang buoc"
' cua cho phep: (a) toan bo boc trong Utils.BeginOperation nen mot lan Ctrl+Z hoan tac tron thao
' tac, (b) chuyen XONG HAN tung story roi moi sang story ke tiep, va ghi ro story nao da xong
' trong thong bao loi.
' ============================================================================
Public Function ConvertToUnicode() As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("EncodingConverter.ConvertToUnicode")
    On Error GoTo ErrHandler
    EnsureTexts

    Dim summary As Object
    Set summary = Utils.NewDictionary()
    summary("convertedRunCount") = 0
    summary("unmappedCharCount") = 0
    summary("skippedVniRunCount") = 0
    summary("completedStories") = ""
    summary("failed") = False

    ' Mo bao thao tac TRUOC ca pha gom: ScreenUpdating tat som giup pha gom (doc Range.Text/
    ' Range.Font.Name hang nghin lan) nhanh han han, va ErrHandler duoi day luon co mot bao thao
    ' tac dang mo de dong lai cho dung.
    Utils.BeginOperation SafetyGuard.HIGH_RISK_ENCODING_CONVERSION

    Dim plans As Collection
    Set plans = CollectStoryPlans()

    Dim totalRuns As Long
    Dim p As Variant
    For Each p In plans
        totalRuns = totalRuns + p("runs").count
    Next p

    Dim processedRuns As Long
    processedRuns = 0

    Dim plan As Variant
    For Each plan In plans
        ConvertOneStory plan, summary, processedRuns, totalRuns
        summary("completedStories") = AppendStoryLabel(CStr(summary("completedStories")), CStr(plan("label")))
    Next plan

    Application.StatusBar = False
    Utils.EndOperation CLng(summary("convertedRunCount")), _
        (CLng(summary("unmappedCharCount")) > 0 Or CLng(summary("skippedVniRunCount")) > 0)

    Set ConvertToUnicode = summary
    Exit Function

ErrHandler:
    Application.StatusBar = False
    summary("failed") = True
    Utils.AbortOperation BuildPartialFailureMessage(summary, Err.description)
    Set ConvertToUnicode = summary
End Function

' Ghep them mot ten story vao danh sach da xong, ngan cach bang dau phay.
Private Function AppendStoryLabel(ByVal current As String, ByVal LABEL As String) As String
    If Len(current) = 0 Then
        AppendStoryLabel = LABEL
    Else
        AppendStoryLabel = current & ", " & LABEL
    End If
End Function

' Thong bao loi giua chung â€” noi ro phan nao da chuyen xong de nguoi dung biet tai lieu dang o
' trang thai nao truoc khi quyet dinh Ctrl+Z (FR-SAF-07).
Private Function BuildPartialFailureMessage(ByVal summary As Object, ByVal errDescription As String) As String
    Dim done As String
    done = CStr(summary("completedStories"))

    If Len(done) = 0 Then
        BuildPartialFailureMessage = ERR_PARTIAL_NONE & vbCrLf & errDescription
    Else
        BuildPartialFailureMessage = ERR_PARTIAL_DONE & done & "." & vbCrLf & errDescription
    End If
End Function

' ============================================================================
' PHA 1 â€” GOM. Chi doc, khong sua gi tai lieu.
' ============================================================================

Private Function CollectStoryPlans() As Collection
    On Error GoTo ErrHandler
    EnsureTexts

    Dim plans As New Collection
    Dim story As Range
    Dim cur As Range

    For Each story In ActiveDocument.StoryRanges
        Set cur = story
        Do While Not cur Is Nothing
            If ShouldScanStoryType(cur.StoryType) Then
                Dim runs As Collection
                Set runs = New Collection
                CollectConvertibleRunsFromStory cur, runs

                If runs.count > 0 Then
                    Dim plan As Object
                    Set plan = Utils.NewDictionary()
                    plan("label") = StoryTypeLabel(cur.StoryType)
                    Set plan("storyRange") = cur.Duplicate
                    Set plan("runs") = runs
                    plans.Add plan
                End If
            End If
            Set cur = cur.NextStoryRange
        Loop
    Next story

    Set CollectStoryPlans = plans
    Exit Function
ErrHandler:
    Err.Raise Err.number, "EncodingConverter.CollectStoryPlans", Err.description
End Function

' Ten tieng Viet cua tung loai story, dung trong thong bao loi giua chung.
Private Function StoryTypeLabel(ByVal st As WdStoryType) As String
    EnsureTexts
    Select Case st
        Case wdMainTextStory
            StoryTypeLabel = STORY_MAIN
        Case wdPrimaryHeaderStory, wdEvenPagesHeaderStory, wdFirstPageHeaderStory
            StoryTypeLabel = STORY_HEADER
        Case wdPrimaryFooterStory, wdEvenPagesFooterStory, wdFirstPageFooterStory
            StoryTypeLabel = STORY_FOOTER
        Case wdFootnotesStory
            StoryTypeLabel = STORY_FOOTNOTE
        Case wdTextFrameStory
            StoryTypeLabel = STORY_TEXTBOX
        Case Else
            StoryTypeLabel = STORY_ENDNOTE
    End Select
End Function

' Story chinh duyet tung Paragraph (de gan paragraphIndex 0-based dung quy uoc toan du an, va de
' moi lan doc Range.Text/Range.Font.Name chi tren mot doan); cac story khac gom ca story thanh mot
' lan quet voi paragraphIndex = Null - khoi luong chu o do nho.
Private Sub CollectConvertibleRunsFromStory(ByVal storyRange As Range, ByVal target As Collection)
    On Error GoTo ErrHandler

    If storyRange.StoryType = wdMainTextStory Then
        Dim p As paragraph
        Dim idx As Long
        idx = 0
        For Each p In storyRange.paragraphs
            CollectConvertibleRunsFromRange p.Range, target, idx
            idx = idx + 1
        Next p
    Else
        CollectConvertibleRunsFromRange storyRange, target, Null
    End If
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "EncodingConverter.CollectConvertibleRunsFromStory", Err.description
End Sub

' Ky tu KHONG duoc dua vao mot run chuyen doi duoc. Moi ma < 32 deu la ky tu dieu khien cua Word
' chu khong phai chu: dau doan (13), dau ket o/ket hang bang (7), ngat dong (11), ngat trang/cot
' (12, 14), anh chen dong (1), tham chieu chu thich (2), dau phan cach truong (19, 20, 21), tab
' (9). Gan Range.Text de len cac ky tu nay se XOA anh, xoa truong, hoac lam hong cau truc bang -
' nen chung duoc dung lam RANH GIOI cat run va bi bo ra ngoai moi run.
Private Function IsControlChar(ByVal ch As String) As Boolean
    IsControlChar = (CodePointOf(ch) < 32)
End Function

' Cat mot Range thanh cac doan lien tuc: cung ten phong VA khong chua ky tu dieu khien.
' rng.Font.Name tra "" (chuoi rong, KHONG phai Null) khi vung co nhieu phong - chi luc do moi phai
' hoi phong tung ky tu, duong cham nhat. Tai lieu TCVN3 that gan nhu luon dong nhat phong trong
' mot doan, nen duong nhanh (mot lan doc Font.Name cho ca doan) la duong chay chinh.
Private Sub CollectConvertibleRunsFromRange(ByVal srcRange As Range, ByVal target As Collection, ByVal paragraphIndex As Variant)
    On Error GoTo ErrHandler

    Dim rng As Range
    Set rng = srcRange.Duplicate

    Dim s As String
    s = rng.text
    Dim n As Long
    n = Len(s)
    If n = 0 Then Exit Sub

    Dim fonts() As String
    ReDim fonts(1 To n)

    Dim uniformFont As String
    uniformFont = rng.Font.name

    Dim i As Long
    If uniformFont <> "" Then
        For i = 1 To n
            fonts(i) = uniformFont
        Next i
    Else
        ' Duong cham: phai hoi phong tung ky tu. Chi chay khi vung that su co nhieu phong. Chi so
        ' cua Characters phai khop 1-1 voi chi so trong chuoi Range.Text - lech thi bo qua ca
        ' vung, KHONG doan (CLAUDE.md muc 5 "khong chac thi khong sua").
        If rng.Characters.count <> n Then Exit Sub
        For i = 1 To n
            If IsControlChar(Mid$(s, i, 1)) Then
                fonts(i) = ""
            Else
                fonts(i) = rng.Characters(i).Font.name
            End If
        Next i
    End If

    Dim segStart As Long
    segStart = 0

    For i = 1 To n + 1
        Dim mustBreak As Boolean
        If i > n Then
            mustBreak = True
        ElseIf IsControlChar(Mid$(s, i, 1)) Then
            mustBreak = True
        ElseIf segStart > 0 Then
            mustBreak = (fonts(i) <> fonts(segStart))
        Else
            mustBreak = False
        End If

        If mustBreak And segStart > 0 Then
            AddConvertibleRun rng, segStart, i - 1, fonts(segStart), Mid$(s, segStart, i - segStart), target, paragraphIndex
            segStart = 0
        End If

        If i <= n Then
            If Not IsControlChar(Mid$(s, i, 1)) Then
                If segStart = 0 Then segStart = i
            End If
        End If
    Next i
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "EncodingConverter.CollectConvertibleRunsFromRange", Err.description
End Sub

' startOffset/endOffset la vi tri 1-based trong chuoi baseRange.Text; luu ra vi tri TUYET DOI
' trong story (startPos/endPos) de pha 2 dung lai Range. Quy uoc mot ky tu = mot don vi vi tri cua
' Range dung cho moi thu Word tra trong Range.Text (ke ca anh chen dong hay tham chieu chu thich -
' deu dem la mot); du vay pha 2 van doi chieu lai Range.Text truoc khi ghi, nen mot sai lech vi
' tri se dan den BO QUA run do chu khong ghi nham cho.
Private Sub AddConvertibleRun(ByVal baseRange As Range, ByVal startOffset As Long, ByVal endOffset As Long, _
        ByVal fontName As String, ByVal text As String, ByVal target As Collection, ByVal paragraphIndex As Variant)

    Dim item As Object
    Set item = Utils.NewDictionary()
    item("startPos") = baseRange.Start + startOffset - 1
    item("endPos") = baseRange.Start + endOffset
    item("fontName") = fontName
    item("text") = text
    item("paragraphIndex") = paragraphIndex
    target.Add item
End Sub

' ============================================================================
' PHA 2 â€” GHI. Duyet run theo THU TU NGUOC (xem giai thich o ConvertToUnicode).
' ============================================================================

Private Sub ConvertOneStory(ByVal plan As Object, ByVal summary As Object, ByRef processedRuns As Long, ByVal totalRuns As Long)
    On Error GoTo ErrHandler

    Dim runs As Collection
    Set runs = plan("runs")

    Dim storyRange As Range
    Set storyRange = plan("storyRange")

    Dim lastPercent As Long
    lastPercent = -1

    Dim i As Long
    For i = runs.count To 1 Step -1
        ApplyRunConversion storyRange, runs(i), summary

        processedRuns = processedRuns + 1

        ' Lam tron xuong boi cua 10 phan tram. Dau ngoac quanh phep chia nguyen la BAT BUOC: trong
        ' VBA "*" co do uu tien CAO HON "\", nen a \ b * 10 se bi hieu la a \ (b * 10).
        Dim percent As Long
        If totalRuns > 0 Then percent = ((processedRuns * 10&) \ totalRuns) * 10
        If percent <> lastPercent Then
            Application.StatusBar = PROGRESS_PREFIX & CStr(percent) & "%"
            lastPercent = percent
        End If
    Next i
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "EncodingConverter.ConvertOneStory (" & ERR_PARTIAL_FAILED & plan("label") & ")", Err.description
End Sub

Private Sub ApplyRunConversion(ByVal storyRange As Range, ByVal item As Object, ByVal summary As Object)
    Dim converted As Object
    Set converted = ConvertRun(CStr(item("fontName")), CStr(item("text")))

    If CStr(converted("skipped")) = "vniDisabled" Then
        summary("skippedVniRunCount") = CLng(summary("skippedVniRunCount")) + 1
        Exit Sub
    End If
    If Len(CStr(converted("skipped"))) > 0 Then Exit Sub

    ' Dung lai Range NGAY TRUOC khi ghi, khong giu san tu pha 1 - xem ghi chu hieu nang o
    ' CollectStoryPlans. Duplicate roi SetRange giu nguyen story cua storyRange (dau trang, chan
    ' trang, chu thich deu co he vi tri rieng, khong dung chung voi than bai).
    Dim segRange As Range
    Set segRange = storyRange.Duplicate
    segRange.SetRange CLng(item("startPos")), CLng(item("endPos"))

    ' Cong chan an toan: van ban thuc te tai vi tri nay phai DUNG BANG van ban da gom o pha 1.
    ' Lech nghia la vi tri da xe dich (gia thiet "mot ky tu = mot don vi vi tri" khong dung o cho
    ' nay) - bo qua run, KHONG ghi de nham cho. Doi chieu quy uoc "khong chac thi khong sua"
    ' (CLAUDE.md muc 5).
    If segRange.text <> CStr(item("text")) Then Exit Sub

    segRange.text = CStr(converted("text"))

    Dim newFontName As String: newFontName = CStr(converted("newFontName"))
    With segRange.Font
        .name = newFontName
        On Error Resume Next
        .NameAscii = newFontName
        .NameFarEast = newFontName
        .NameOther = newFontName
        .NameBi = newFontName
        On Error GoTo 0
    End With

    summary("convertedRunCount") = CLng(summary("convertedRunCount")) + 1
    summary("unmappedCharCount") = CLng(summary("unmappedCharCount")) + CLng(converted("unmappedCount"))
End Sub
