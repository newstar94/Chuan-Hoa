Attribute VB_Name = "DocumentSnapshot"
'==============================================================
' DocumentSnapshot
' HIEU NANG â€” diem chet nguoi cua ban Legacy: goi Range.Information cho tung doan tren tai lieu
' dai la RAT CHAM neu lap lai truy cap Paragraphs(i) nhieu lan. Bat buoc:
' - Application.ScreenUpdating = False suot qua trinh, bat lai truoc khi thoat KE CA khi loi.
' - Doc Paragraphs(i).Range MOT LAN, gan vao bien, roi lay nhieu thuoc tinh tu do (KHONG goi lai
'   ActiveDocument.Paragraphs(i) nhieu lan trong cung mot vong lap).
' - Khong co bat ky loi goi Word nao khac trong vong lap ngoai nhung gi that su can.
' CACHE TRONG PHIEN KIEM TRA: CaptureDocument KHONG tu cache â€” noi goi (ComplianceChecker, tro di)
' chiu trach nhiem goi MOT LAN cho moi lan bam "Kiem tra" va dung lai ket qua cho moi module,
' khong de tung module tu chup lai.
'==============================================================
Option Explicit

Private mLastCaptureDurationMs As Double
Private mHasCaptured As Boolean

' ============================================================================
' Diem vao duy nhat
' ============================================================================

Public Function CaptureDocument() As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("DocumentSnapshot.CaptureDocument")
    On Error GoTo ErrHandler
    Dim startTime As Single
    startTime = Timer

    Application.ScreenUpdating = False

    Dim paragraphs As Collection
    Set paragraphs = CaptureParagraphs()

    Dim Result As Object
    Set Result = Utils.NewDictionary()
    Result.Add "Paragraphs", paragraphs
    Result.Add "Sections", CaptureSections()
    Result.Add "Tables", CaptureTables(paragraphs)
    Result.Add "Images", CaptureImages()
    Result.Add "FileName", GetDocumentFileNameOrEmpty()
    Result.Add "IsDocx", IsCurrentDocumentDocx()

    Application.ScreenUpdating = True

    ' Timer chay lai tu 0 luc nua dem â€” hieu la am neu chup vat qua moc do (rat hiem, chi anh
    ' huong so lieu do, khong anh huong ket qua chup).
    mLastCaptureDurationMs = (Timer - startTime) * 1000
    mHasCaptured = True

    Set CaptureDocument = Result
    Exit Function
ErrHandler:
    Application.ScreenUpdating = True
    Err.Raise Err.number, "DocumentSnapshot.CaptureDocument", Err.description
End Function

' Ban CHUP NHE - CHI doan van, khong Sections/Tables/Images. Dung cho nhung noi CHI can noi
' dung/dinh dang doan (vi du BlankFieldSpacer chi do khoang trang trong text) - tranh di qua
' CaptureImages mot cach khong can thiet ( chan doan "Kiem tra chinh ta chay rat lau": tren mot
' tai lieu that co InlineShape bi "Object has been deleted", MOI lan CaptureImages cham vao anh do
' ton them 9-16 giay Word tu do truoc khi loi noi len - goi CaptureDocument HAI LAN moi luot "Kiem
' tra" (mot lan trong BlankFieldSpacer, mot lan trong BuildCheckContext) nhan doi chi phi nay. Ham
' nay cat bo MOT trong hai lan, giu nguyen lan con lai trong BuildCheckContext (noi THAT SU can du
' lieu Images/Tables/Sections).
Public Function CaptureParagraphsOnly() As Collection
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("DocumentSnapshot.CaptureParagraphsOnly")
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Set CaptureParagraphsOnly = CaptureParagraphs()
    Application.ScreenUpdating = True
    Exit Function
ErrHandler:
    Application.ScreenUpdating = True
    Err.Raise Err.number, "DocumentSnapshot.CaptureParagraphsOnly", Err.description
End Function

' Nhu CaptureParagraphsOnly, nhung goi trong Dictionary khoa "Paragraphs" - dung cho noi CAN goi
' DocumentTypeDetector.DetectDocumentType/ComponentDetector.DetectComponents (doi hoi tham so
' snapshot dang Dictionary) nhung KHONG can Sections/Tables/Images: ComponentFormatter.DetectLayoutMap,
' EdgeWhitespaceTrimmer.TrimEdgeWhitespace, MultiSpaceCollapser.CollapseDoubleSpaces - ca ba CHI
' doc snapshot("Paragraphs"), tung goi CaptureDocument() day du truoc day nen moi lan "Kiem tra"
' cham vao anh loi (T-71) NHIEU LAN thay vi mot (BuildCheckContext, noi THAT SU can Images/Tables/
' Sections).
Public Function CaptureParagraphsOnlySnapshot() As Object
    Dim Result As Object: Set Result = Utils.NewDictionary()
    Result.Add "Paragraphs", CaptureParagraphsOnly()
    Set CaptureParagraphsOnlySnapshot = Result
End Function

' ============================================================================
' ============================================================================

' Noi cac doan DAU TAI LIEU cho toi khi du maxChars ky tu â€” tin hieu BO TRO cho
' RegimeDetector.DetectRegime, bat duoc truong hop cum tu bi tach lech qua nhieu doan.
Public Function BuildHeaderWindow(ByVal paragraphs As Collection, Optional ByVal maxChars As Long = 2000) As String
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("DocumentSnapshot.BuildHeaderWindow")
    Dim buf As String
    buf = ""
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        buf = buf & p.text & vbLf
        If Len(buf) >= maxChars Then Exit For
    Next p
    BuildHeaderWindow = buf
End Function

' Chi so doan CUOI CUNG con nam trong pham vi maxChars dau tai lieu â€” cung quy uoc tich luy voi
' BuildHeaderWindow (doan lam day bo dem VAN duoc tinh la trong pham vi). ComponentDetector dung
' de thuc thi co "headerWindowOnly" cua dau hieu nhan dien. Tai lieu rong tra -1.
Public Function HeaderWindowLastIndex(ByVal paragraphs As Collection, Optional ByVal maxChars As Long = 2000) As Long
    Dim last As Long
    last = -1
    Dim used As Long
    used = 0
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        last = p.Index
        used = used + Len(p.text) + 1
        If used >= maxChars Then Exit For
    Next p
    HeaderWindowLastIndex = last
End Function

' ============================================================================
' Doan van
' ============================================================================

Private Function CaptureParagraphs() As Collection
    Dim Result As New Collection
    Dim total As Long
    total = ActiveDocument.paragraphs.count

    Dim i As Long
    Dim rng As word.Range
    Dim rowMarkProbe As word.Range
    Dim isRowMark As Boolean
    Dim fmt As word.ParagraphFormat
    Dim text As String
    Dim snap As ParagraphSnapshot

    ' Chi so doan cua ban chup (0-based) - KHONG dung i-1 vi vong lap co bo qua dau ket dong cua
    ' bang (xem duoi), nen chi so Word va chi so ban chup lech nhau.
    Dim snapIndex As Long
    snapIndex = 0

    For i = 1 To total
        ' BUOC 1 - lay Range va xac dinh co phai dau ket dong cua bang khong. Boc rieng khoi BUOC
        ' 2: day la PHAN DUY NHAT DocumentSnapshot.BuildSnapshotIndexMap cung lam - hai ham PHAI
        ' ra cung quyet dinh "bo qua hay dem" cho cung mot i, neu khong chi so ban chup
        ' (snapIndex) se lech nhau giua hai noi goi (xem canh bao dau ham do).
        On Error GoTo Step1Error
        Set rng = ActiveDocument.paragraphs(i).Range

        ' BO QUA dau ket DONG cua bang (end-of-row mark). Dau ket O (end-of-cell) VAN giu
        ' PHAI thu tren ban sao DA THU GON VE DAU:
        Set rowMarkProbe = rng.Duplicate
        rowMarkProbe.Collapse wdCollapseStart
        isRowMark = rowMarkProbe.Information(wdAtEndOfRowMarker)
        On Error GoTo 0
        If isRowMark Then GoTo ContinueParagraph

        ' BUOC 2 - doc tung thuoc tinh cua doan. dung mot ban chup RONG/an toan thay the va DI
        ' TIEP - dung khuon mau da co san o TextFormatter.ApplyFontSizeWholeDocument. Khac BUOC 1:
        ' o day VAN dem vao snapIndex (BuildSnapshotIndexMap khong doc cac thuoc tinh nay nen
        ' khong the gap loi cung diem - dem binh thuong de hai ham khop chi so).
        On Error GoTo Step2Error
        Set fmt = rng.ParagraphFormat
        text = StripParagraphMark(RemoveInlineShapeChars(rng))

        Set snap = New ParagraphSnapshot
        snap.Index = snapIndex
        snap.text = text
        snap.styleName = MixedToEmptyString(GetStyleNameSafe(rng))
        snap.alignment = ToParagraphAlignment(fmt.alignment)
        If Len(text) = 0 Then
            snap.fontName = ""
            snap.FontSizePt = 0
            snap.bold = False
            snap.Italic = False
            snap.FontColor = ""
        Else
            snap.fontName = MixedToEmptyString(rng.Font.name)
            snap.FontSizePt = MixedToZero(rng.Font.size)
            snap.bold = MixedToNullBoolean(rng.Font.bold)
            snap.Italic = MixedToNullBoolean(rng.Font.Italic)
            snap.FontColor = MixedToHexColor(rng.Font.Color)
        End If
        ' Theo NOI DUNG, khong doc Font.AllCaps â€” xem ghi chu dau ParagraphSnapshot.cls.
        snap.AllCaps = IsAllCapsContent(text)
        snap.FirstLineIndentPt = fmt.FirstLineIndent
        snap.SpaceBeforePt = fmt.SpaceBefore
        snap.SpaceAfterPt = fmt.SpaceAfter
        snap.lineSpacing = fmt.lineSpacing
        snap.isInTable = CBool(rng.Information(wdWithInTable))
        snap.sectionIndex = CLng(rng.Information(wdActiveEndSectionNumber)) - 1
        snap.PageNumber = CLng(rng.Information(wdActiveEndPageNumber))
        On Error GoTo 0
        GoTo Step2Done

Step1Error:
        DebugTrace.LogErr "DocumentSnapshot.CaptureParagraphs", _
            "Doan Word i=" & CStr(i) & " loi ngay khi lay Range/kiem tra dau ket dong - bo qua han doan nay", _
            Err.number, Err.description
        Err.Clear
        On Error GoTo 0
        GoTo ContinueParagraph

Step2Error:
        DebugTrace.LogErr "DocumentSnapshot.CaptureParagraphs", _
            "Doan Word i=" & CStr(i) & " (snapIndex=" & CStr(snapIndex) & ") loi khi doc thuoc tinh - dung ban chup rong thay the", _
            Err.number, Err.description
        Err.Clear
        On Error GoTo 0
        Set snap = New ParagraphSnapshot
        snap.Index = snapIndex
        snap.text = ""
        snap.styleName = ""
        snap.alignment = "unknown"
        snap.fontName = ""
        snap.FontSizePt = 0
        snap.FontColor = ""
        snap.AllCaps = False
        snap.FirstLineIndentPt = 0
        snap.SpaceBeforePt = 0
        snap.SpaceAfterPt = 0
        snap.lineSpacing = 0
        snap.isInTable = False
        snap.sectionIndex = 0
        snap.PageNumber = 0
Step2Done:
        Result.Add snap
        snapIndex = snapIndex + 1
ContinueParagraph:
    Next i

    Set CaptureParagraphs = Result
End Function

' ============================================================================
' cac nut "Sua" ghi nham vao sai doan tren van ban that co bang, vi du bang bo cuc "Noi nhan"/chu
' ky dat canh nhau): ComponentFormatter.bas va TextFormatter.ApplyFontSizeWholeDocument truoc day
' tu SUY snapIndex + 1 = chi so Word that â€” SAI tren MOI tai lieu co bang, vi CaptureParagraphs
' (tren) BO QUA dau ket dong cua bang (end-of-row mark) khi danh snapIndex nhung
' ActiveDocument.Paragraphs KHONG bo qua, nen hai chi so LECH nhau dung bang so dau ket dong da
' gap TRUOC doan can tim. Ham nay la NOI DUY NHAT duoc phep dich chi so â€” MOI cho khac can ghi vao
' tai lieu tu mot chi so lay ra tu ParagraphSnapshot/LayoutMap PHAI di qua day, KHONG tu +1.
' ============================================================================

' Dung MOT LAN cho ca vong lap sua nhieu doan (hieu nang - xem canh bao dau file "diem chet nguoi"
' - lap lai goi ham don-le se O(n) MOI lan tra cuu). Tra Dictionary Long(snapIndex 0-based) ->
' Long(chi so 1-based that cua ActiveDocument.Paragraphs).
Public Function BuildSnapshotIndexMap() As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("DocumentSnapshot.BuildSnapshotIndexMap")
    Dim Result As Object: Set Result = Utils.NewDictionary()

    Dim total As Long: total = ActiveDocument.paragraphs.count
    Dim i As Long
    Dim rng As word.Range
    Dim rowMarkProbe As word.Range
    Dim isRowMark As Boolean
    Dim snapIndex As Long: snapIndex = 0

    For i = 1 To total
        ' PHAI ra CUNG quyet dinh "bo qua hay dem" nhu Buoc 1 cua CaptureParagraphs o tren cho
        ' cung mot i - xem canh bao dau ham do. Loi o day thi BO QUA HAN, khong dem, giong het
        ' nhanh Step1Error ben do.
        On Error GoTo ItemError
        Set rng = ActiveDocument.paragraphs(i).Range
        Set rowMarkProbe = rng.Duplicate
        rowMarkProbe.Collapse wdCollapseStart
        isRowMark = rowMarkProbe.Information(wdAtEndOfRowMarker)
        On Error GoTo 0
        If isRowMark Then GoTo ContinueScan

        Result.Add snapIndex, i
        snapIndex = snapIndex + 1
        GoTo ContinueScan
ItemError:
        DebugTrace.LogErr "DocumentSnapshot.BuildSnapshotIndexMap", _
            "Doan Word i=" & CStr(i) & " loi ngay khi lay Range/kiem tra dau ket dong - bo qua han doan nay", _
            Err.number, Err.description
        Err.Clear
        On Error GoTo 0
ContinueScan:
    Next i

    Set BuildSnapshotIndexMap = Result
End Function

' Paragraph.Style co the la Style object hoac (hiem) nem loi tren mot so tai lieu bi hong lien ket
' style - boc rieng de mot doan loi khong lam hong ca lan chup.
Private Function GetStyleNameSafe(ByVal rng As word.Range) As String
    On Error Resume Next
    GetStyleNameSafe = rng.Style.NameLocal
    On Error GoTo 0
End Function

' Dau ket doan o CUOI chuoi: CR/LF, va U+0007 (ky tu ket o ma Range.Text tra ve cho doan cuoi
' trong mot o bang). Cat o day de hai ban cho ra cung mot chuoi Van ban cua doan sau khi BO cac ky
' tu dai dien cho anh inline. Word chen mot ky tu vao Range.Text cho moi InlineShape (thuong la
' U+0001, nhung da gap ca ky tu khac tuy tai lieu - xem bo fixture bien-hop-canh-05-anh, Word 2016
' tra ve U+002F), nen KHONG duoc loc theo ma ky tu ma phai loc theo VI TRI that cua tung
' InlineShape.
Private Function RemoveInlineShapeChars(ByVal rng As word.Range) As String
    Dim text As String
    text = rng.text
    If rng.InlineShapes.count = 0 Then
        RemoveInlineShapeChars = text
        Exit Function
    End If

    ' Xoa tu CUOI ve DAU de vi tri cua cac ky tu chua xoa khong bi doi.
    Dim offsets As Collection
    Set offsets = New Collection
    Dim k As Long
    For k = 1 To rng.InlineShapes.count
        Dim offset As Long
        offset = rng.InlineShapes(k).Range.Start - rng.Start + 1 ' ve 1-based trong text
        If offset >= 1 And offset <= Len(text) Then
            ' Chen vao DAU de danh sach xep giam dan theo vi tri. Collection.Add voi Before:=1 bao
            ' loi khi danh sach dang rong, nen muc dau tien phai them binh thuong.
            If offsets.count = 0 Then
                offsets.Add offset
            Else
                offsets.Add offset, , 1
            End If
        End If
    Next k

    Dim j As Long
    For j = 1 To offsets.count
        Dim pos As Long
        pos = CLng(offsets(j))
        text = left$(text, pos - 1) & Mid$(text, pos + 1)
    Next j

    RemoveInlineShapeChars = text
End Function

Private Function StripParagraphMark(ByVal text As String) As String
    Dim Result As String
    Result = text

    Do While Len(Result) > 0 And _
            (Right$(Result, 1) = vbCr Or Right$(Result, 1) = vbLf Or _
             Right$(Result, 1) = ChrW(7))
        Result = left$(Result, Len(Result) - 1)
    Loop

    ' U+0001 (neo doi tuong inline - anh), U+000C (ngat trang/ngat section), U+000E (ngat cot).
    Result = Replace$(Result, ChrW(1), "")
    Result = Replace$(Result, ChrW(12), "")
    Result = Replace$(Result, ChrW(14), "")

    StripParagraphMark = Result
End Function

' Co it nhat mot chu cai va toan bo chu hoa theo Utils.ToUpperVn (UCase$, Unicode dung cho tieng
' Viet co dau)
Private Function IsAllCapsContent(ByVal text As String) As Boolean
    If Not HasLetter(text) Then
        IsAllCapsContent = False
        Exit Function
    End If
    IsAllCapsContent = (text = Utils.ToUpperVn(text))
End Function

Private Function HasLetter(ByVal text As String) As Boolean
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If UCase$(ch) <> LCase$(ch) Then
            HasLetter = True
            Exit Function
        End If
    Next i
    HasLetter = False
End Function

' Bon bien the canh deu gop ve "justify" vi ND 30 chi phan biet canh deu hay khong.
Private Function ToParagraphAlignment(ByVal alignment As Long) As String
    Select Case alignment
        Case wdAlignParagraphLeft
            ToParagraphAlignment = "left"
        Case wdAlignParagraphCenter
            ToParagraphAlignment = "center"
        Case wdAlignParagraphRight
            ToParagraphAlignment = "right"
        Case wdAlignParagraphJustify, wdAlignParagraphJustifyMed, _
             wdAlignParagraphJustifyHi, wdAlignParagraphJustifyLow
            ToParagraphAlignment = "justify"
        Case Else
            ToParagraphAlignment = "unknown"
    End Select
End Function

' ============================================================================
' Section
' ============================================================================

Private Function CaptureSections() As Collection
    Dim Result As New Collection
    Dim total As Long
    total = ActiveDocument.sections.count

    Dim i As Long
    Dim ps As word.PageSetup
    Dim item As Object

    ' mot section loi khong duoc lam hong ca lan chup, bo qua section do va di tiep thay vi nem
    ' loi len tan CaptureDocument.
    For i = 1 To total
        On Error GoTo SectionError
        Set ps = ActiveDocument.sections(i).PageSetup
        Set item = Utils.NewDictionary()
        item.Add "Index", i - 1
        ' Ban Legacy toan quyen Word Object Model nen PageSetup luon doc duoc
        item.Add "PageSetupAvailable", True
        item.Add "PageWidthPt", ps.PageWidth
        item.Add "PageHeightPt", ps.PageHeight
        item.Add "TopMarginPt", ps.TopMargin
        item.Add "BottomMarginPt", ps.BottomMargin
        item.Add "LeftMarginPt", ps.LeftMargin
        item.Add "RightMarginPt", ps.RightMargin
        item.Add "Orientation", IIf(ps.Orientation = wdOrientPortrait, "portrait", "landscape")
        item.Add "DifferentFirstPageHeaderFooter", CBool(ps.DifferentFirstPageHeaderFooter)
        On Error GoTo 0

        Result.Add item
        GoTo NextSection
SectionError:
        DebugTrace.LogErr "DocumentSnapshot.CaptureSections", _
            "Section i=" & CStr(i) & " loi khi doc - bo qua section nay", Err.number, Err.description
        Err.Clear
        On Error GoTo 0
NextSection:
    Next i

    Set CaptureSections = Result
End Function

' ============================================================================
' Bang â€” chi bang cap ngoai cung, doi chieu Document.Tables (khong tra bang long) ben VBA va
' ============================================================================

Private Function CaptureTables(ByVal paragraphs As Collection) As Collection
    Dim Result As New Collection
    Dim runs As Collection
    Set runs = MatchTableParagraphRuns(paragraphs, ActiveDocument.Tables.count)

    Dim i As Long
    Dim t As word.table
    Dim item As Object
    Dim run As Object

    ' Boc chong crash tung bang - cung nguyen tac voi CaptureSections o tren.
    For i = 1 To ActiveDocument.Tables.count
        On Error GoTo TableError
        Set t = ActiveDocument.Tables(i)

        Set item = Utils.NewDictionary()
        item.Add "Index", i - 1

        If i <= runs.count Then
            Set run = runs(i)
            item.Add "FirstParagraphIndex", run("Start")
            item.Add "ParagraphCount", run("Count")
        Else
            ' So dai lech so bang - tra -1/0 thay vi ghep sai, dung nguyen tac "khong chac thi
            ' khong sua" (CLAUDE.md muc 5).
            item.Add "FirstParagraphIndex", -1
            item.Add "ParagraphCount", 0
        End If

        item.Add "RowCount", MixedToZero(t.Rows.count)
        item.Add "HeaderRowCount", CountHeaderRows(t)
        item.Add "Alignment", ToParagraphAlignment(SafeTableAlignment(t))
        item.Add "IsUniform", CBool(t.Uniform)
        item.Add "StyleName", MixedToEmptyString(GetTableStyleNameSafe(t))
        On Error GoTo 0

        Result.Add item
        GoTo NextTable
TableError:
        DebugTrace.LogErr "DocumentSnapshot.CaptureTables", _
            "Bang i=" & CStr(i) & " loi khi doc - bo qua bang nay", Err.number, Err.description
        Err.Clear
        On Error GoTo 0
NextTable:
    Next i

    Set CaptureTables = Result
End Function

' Table.Rows(1).HeadingFormat khong dai dien cho ca bang - dem so dong lien tiep tu dau bang co
' HeadingFormat = True, doi chieu y nghia headerRowCount la so dong LAP LAI o dau moi trang.
Private Function CountHeaderRows(ByVal t As word.table) As Long
    On Error GoTo ErrHandler
    Dim count As Long
    Dim r As word.row
    count = 0

    For Each r In t.Rows
        If r.HeadingFormat Then
            count = count + 1
        Else
            Exit For
        End If
    Next r

    CountHeaderRows = count
    Exit Function
ErrHandler:
    CountHeaderRows = 0
End Function

' Table.Rows.Alignment (WdRowAlignment: wdAlignRowLeft=0/Center=1/Right=2 - KHONG co Justify, bang
' khong canh deu) trung gia tri so voi ba truong hop dau cua WdParagraphAlignment nen dung chung
' duoc ToParagraphAlignment ben duoi.
Private Function SafeTableAlignment(ByVal t As word.table) As Long
    On Error GoTo ErrHandler
    SafeTableAlignment = t.Rows.alignment
    Exit Function
ErrHandler:
    SafeTableAlignment = -1
End Function

Private Function GetTableStyleNameSafe(ByVal t As word.table) As String
    On Error Resume Next
    GetTableStyleNameSafe = t.Style.NameLocal
    On Error GoTo 0
End Function

' Tra Collection RONG (thay vi ghep lech) khi so dai khac so bang.
Private Function MatchTableParagraphRuns(ByVal paragraphs As Collection, ByVal tableCount As Long) As Collection
    Dim runs As New Collection
    Dim startIndex As Long
    startIndex = -1

    Dim i As Long
    Dim total As Long
    Dim p As ParagraphSnapshot
    total = paragraphs.count

    For i = 1 To total
        Set p = paragraphs(i)

        If p.isInTable And startIndex < 0 Then
            startIndex = p.Index
        ElseIf Not p.isInTable And startIndex >= 0 Then
            AddRun runs, startIndex, p.Index - startIndex
            startIndex = -1
        End If
    Next i

    If startIndex >= 0 Then
        AddRun runs, startIndex, total - startIndex
    End If

    If runs.count = tableCount Then
        Set MatchTableParagraphRuns = runs
    Else
        Set MatchTableParagraphRuns = New Collection
    End If
End Function

Private Sub AddRun(ByVal runs As Collection, ByVal startIndex As Long, ByVal count As Long)
    Dim run As Object
    Set run = Utils.NewDictionary()
    run.Add "Start", startIndex
    run.Add "Count", count
    runs.Add run
End Sub

' ============================================================================
' Anh
' ============================================================================

Private Function CaptureImages() As Collection
    Dim Result As New Collection
    Dim total As Long
    total = ActiveDocument.InlineShapes.count

    Dim i As Long
    Dim shp As word.InlineShape
    Dim item As Object

    ' Boc chong crash tung anh - cung nguyen tac voi CaptureSections/CaptureTables o tren.
    For i = 1 To total
        On Error GoTo ImageError
        Set shp = ActiveDocument.InlineShapes(i)
        Set item = Utils.NewDictionary()
        item.Add "Index", i - 1
        item.Add "ParagraphIndex", FindParagraphIndexForRange(shp.Range)
        item.Add "WidthPt", shp.width
        item.Add "HeightPt", shp.Height
        item.Add "LockAspectRatio", CBool(shp.LockAspectRatio)
        item.Add "AltTextDescription", MixedToEmptyString(shp.AlternativeText)
        On Error GoTo 0

        Result.Add item
        GoTo NextImage
ImageError:
        DebugTrace.LogErr "DocumentSnapshot.CaptureImages", _
            "Anh i=" & CStr(i) & " loi khi doc - bo qua anh nay", Err.number, Err.description
        Err.Clear
        On Error GoTo 0
NextImage:
    Next i

    Set CaptureImages = Result
End Function

' Ghep anh voi doan chua no.
Private Function FindParagraphIndexForRange(ByVal shapeRange As word.Range) As Long
    On Error GoTo ErrHandler
    Dim containingPara As word.paragraph
    Set containingPara = shapeRange.paragraphs(1)

    FindParagraphIndexForRange = ActiveDocument.Range(0, containingPara.Range.Start).paragraphs.count
    Exit Function
ErrHandler:
    FindParagraphIndexForRange = -1
End Function

' ============================================================================
' Ten file va.docx
' ============================================================================

' "" khi tai lieu chua luu lan nao (ActiveDocument.Path rong)
Private Function GetDocumentFileNameOrEmpty() As String
    If Len(ActiveDocument.Path) = 0 Then
        GetDocumentFileNameOrEmpty = ""
    Else
        GetDocumentFileNameOrEmpty = ActiveDocument.name
    End If
End Function

Private Function IsCurrentDocumentDocx() As Boolean
    Dim fileName As String
    fileName = GetDocumentFileNameOrEmpty()
    If Len(fileName) = 0 Then
        IsCurrentDocumentDocx = False
        Exit Function
    End If

    Dim dotPos As Long
    dotPos = InStrRev(fileName, ".")
    If dotPos = 0 Then
        IsCurrentDocumentDocx = False
        Exit Function
    End If

    IsCurrentDocumentDocx = (LCase$(Mid$(fileName, dotPos + 1)) = "docx")
End Function

' ============================================================================
' Tham so khai bao Variant vi Range.Font.Name/Size/Bold/Italic co the tra ve wdUndefined (9999999)
' khi vung co nhieu gia tri.
' ============================================================================

Private Function MixedToEmptyString(ByVal value As Variant) As String
    If IsNull(value) Then
        MixedToEmptyString = ""
    Else
        MixedToEmptyString = CStr(value)
    End If
End Function

Private Function MixedToZero(ByVal value As Variant) As Double
    If IsNull(value) Then
        MixedToZero = 0
    ElseIf CDbl(value) = wdUndefined Then
        MixedToZero = 0
    Else
        MixedToZero = CDbl(value)
    End If
End Function

' Bold/Italic ben Word.Font tra ve Long (True=-1/False=0/wdUndefined=9999999) chu khong phai
' Boolean thuan - phai so voi wdUndefined truoc khi ep kieu Boolean.
Private Function MixedToNullBoolean(ByVal value As Variant) As Variant
    If IsNull(value) Then
        MixedToNullBoolean = Null
    ElseIf CLng(value) = wdUndefined Then
        MixedToNullBoolean = Null
    Else
        MixedToNullBoolean = CBool(value)
    End If
End Function

' Word.Font.Color la mot WdColor (Long), dong goi RGB kieu 0x00BBGGRR (COLORREF cua Windows - byte
' thap nhat la R). wdColorAutomatic (-16777216) nghia la "khong dat mau tuong minh" (Word tu ve
' den tren nen sang) wdUndefined (9999999) la gia tri hon hop (nhieu mau RGB khac nhau trong mot
' doan) -> "" theo quy uoc chung cua ca hai ban.
' Doi chieu OOXML goc (word/document.xml giai nen thu cong) xac nhan CA 213 lan xuat hien
' "w:color" trong toan van ban DEU la "000000" (kem w:themeColor="text1") - khong he co mot mau do
' nao trong file that su. Nguyen nhan: cac doan nay co doan van ban ghi mau qua
' w:themeColor="text1" (tham chieu theme) trong khi dau doan (paragraph mark, w:pPr/w:rPr) hoac
' mot phan khac trong cung doan KHONG mang cung kieu tham chieu (VI DU: RSID/thoi diem sinh doan
' vet dinh dang cu qua accept track changes) - hai ben CUNG la mau den khi HIEN THI nhung Word coi
' la "khong the quy ve MOT gia tri Long duy nhat theo kieu RGB thuan", tra ve MOT SENTINEL RIENG
' cho truong hop "hon hop lien quan theme" nay (0xDD00FFFF), KHAC voi wdUndefined (9999999) ma
' Word dung cho truong hop hai mau RGB THUAN khac nhau ro rang. MOI gia tri am KHAC
' wdColorAutomatic deu la co che noi bo cua Word (tham chieu theme/sentinel khac), KHONG PHAI RGB
' van (RGB(r,g,b) hop le luon >= 0, toi da 16777215) - quy ve "" (hon hop/khong xac dinh) GIONG
' wdUndefined, AN TOAN HON la bia ra mot "mau" khong co that. Danh doi da CAN NHAC: mau nen qua
' tham chieu theme KHAC "den" (vi du chu do qua Accent 2) co the CUNG roi vao nhanh nay va bi bo
' qua thay vi bi bao sai - chap nhan duoc vi (a) nguoi dung thuong chon mau qua bang mau tieu
' chuan (luu RGB truc tiep, khong qua theme, VAN bi bat binh thuong), (b) uu tien loai het BAO SAI
' dang xay ra that (ADR-003, "khong chac thi khong sua" - tot hon la bo sot mot truong hop hiem
' con hon la bao sai hang loat).
Private Function MixedToHexColor(ByVal value As Variant) As String
    If IsNull(value) Then
        MixedToHexColor = ""
        Exit Function
    End If

    Dim rgbValue As Long
    rgbValue = CLng(value)

    If rgbValue = wdColorAutomatic Then
        MixedToHexColor = "#000000"
        Exit Function
    End If

    If rgbValue = wdUndefined Then
        MixedToHexColor = ""
        Exit Function
    End If

    If rgbValue < 0 Then
        ' Am nhung KHAC wdColorAutomatic - sentinel noi bo khac cua Word (xem ghi chu o tren),
        ' khong phai RGB van - quy ve hon hop/khong xac dinh, KHONG giai ma.
        MixedToHexColor = ""
        Exit Function
    End If

    Dim r As Long, g As Long, b As Long
    r = rgbValue Mod 256
    g = (rgbValue \ 256) Mod 256
    b = (rgbValue \ 65536) Mod 256

    MixedToHexColor = "#" & Right$("0" & Hex$(r), 2) & Right$("0" & Hex$(g), 2) & Right$("0" & Hex$(b), 2)
End Function

' Khi xong: ), commit ": DocumentSnapshot ban Legacy".

' Bo sung: them FontColor (ParagraphSnapshot.cls) va MixedToHexColor o tren de
' ND30-PL1-M1-K4-COLOR (ComplianceChecker.bas) kiem duoc â€” luc dau chua chup mau chu.
