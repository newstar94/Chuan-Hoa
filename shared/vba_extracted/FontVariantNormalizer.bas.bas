Attribute VB_Name = "FontVariantNormalizer"
Option Explicit

Private Const BASE_FONT_NAME As String = "Times New Roman"

' Bien module TAM dung de FixRangeOrRuns/ApplyRunIfVariant (goi qua nhieu tang, VBA khong cho
' truyen tham so mac dinh qua Function ho tro de quy sau nhu vay gon) truy cap duoc bang bien the
' hien hanh MA KHONG PHAI truyen qua tung tham so - gan dau NormalizeFontVariants, xoa cuoi ham do
' (khong ro ri sang lan goi khac). PHAI dat O DAY, dau file, TRUOC toan bo Sub/Function - quy tac
' VBA (xem ghi chu tuong tu o dau Utils.bas).
Private variants_ByModule As Object

' Tach hau to bien the tu mot ten font, vi du "Times New Roman Bold" -> "Bold", "Times New
' Roman,Bold Italic" -> "Bold Italic". Tra chuoi rong neu fontName KHONG bat dau bang "Times New
' Roman" theo sau boi mot dau phan cach/hau to (tuc CHINH LA "Times New Roman" nguyen ven, hoac
' mot font khac han - ca hai deu KHONG can sua).
Private Function VariantSuffix(ByVal fontName As String) As String
    If Len(fontName) <= Len(BASE_FONT_NAME) Then Exit Function
    If LCase$(left$(fontName, Len(BASE_FONT_NAME))) <> LCase$(BASE_FONT_NAME) Then Exit Function

    Dim suffix As String
    suffix = Mid$(fontName, Len(BASE_FONT_NAME) + 1)

    ' Bo dau phan cach o dau hau to (dau cach/phay/gach ngang) - "Times New RomanBold" (dinh lien,
    ' khong dau phan cach) KHONG duoc coi la bien the, tranh khop nham mot ten font khac hoan toan
    ' vo tinh co tien to trung (vi du mot font hu cau "Times New RomanBoldFace").
    If Len(suffix) = 0 Then Exit Function
    Select Case left$(suffix, 1)
        Case " ", ",", "-"
            suffix = Mid$(suffix, 2)
        Case Else
            Exit Function
    End Select

    VariantSuffix = Trim$(suffix)
End Function

' day la TOAN BO cach ky hieu kieu chu ma mot font.TTF rieng biet co the mang trong ten (chi co
' bon to hop kieu chu: thuong, dam, nghieng, dam+nghieng - "thuong" thi khong co hau to nen khong
' toi day).
Private Function IsRecognizedVariantSuffix(ByVal suffix As String) As Boolean
    Dim s As String
    s = LCase$(Replace$(Replace$(suffix, " ", ""), vbTab, ""))
    Select Case s
        Case "bold", "italic", "bolditalic", "italicbold"
            IsRecognizedVariantSuffix = True
    End Select
End Function

Private Function SuffixIsBold(ByVal suffix As String) As Boolean
    SuffixIsBold = (InStr(1, suffix, "bold", vbTextCompare) > 0)
End Function

Private Function SuffixIsItalic(ByVal suffix As String) As Boolean
    SuffixIsItalic = (InStr(1, suffix, "italic", vbTextCompare) > 0)
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

' Gom TAP HOP CAC TEN FONT PHAN BIET dang xuat hien trong mot story vao "found" (key = ten font,
' value = True). Uu tien doc Font.Name cua CA story mot lan (nhanh, dung khi story chi mot font
' duy nhat); story tra ve "" (chuoi rong) nghia la NHIEU font tron lan - luc do moi chia nho theo
' TUNG DOAN (van con nhanh hon nhieu so voi chia toi tung ky tu, vi da so doan trong mot van ban
' hanh chinh chi dung MOT font); mot doan van con tron font (hiem) thi moi chia toi tung ky tu -
' dung DUNG mot lan cho ca tai lieu, giong "canh bao" da ghi o EncodingConverter.CollectFontRuns
' ve chi phi vong lap ky tu, nhung o day chi anh huong so LUONG font phan biet tim duoc.
Private Sub CollectDistinctFontNames(ByVal storyRange As Range, ByVal found As Object)
    On Error Resume Next
    Dim wholeName As String
    wholeName = storyRange.Font.name
    If wholeName <> "" Then
        found(wholeName) = True
        Exit Sub
    End If

    If storyRange.StoryType = wdMainTextStory Then
        Dim p As paragraph
        For Each p In storyRange.paragraphs
            CollectFromRange p.Range, found
        Next p
    Else
        CollectFromRange storyRange, found
    End If
    On Error GoTo 0
End Sub

Private Sub CollectFromRange(ByVal rng As Range, ByVal found As Object)
    On Error Resume Next
    Dim r As Range: Set r = rng.Duplicate
    If Len(r.text) = 0 Then Exit Sub
    If Right$(r.text, 1) = vbCr Then r.MoveEnd wdCharacter, -1
    If Len(r.text) = 0 Then Exit Sub

    Dim fName As String
    fName = r.Font.name
    If fName <> "" Then
        found(fName) = True
        Exit Sub
    End If

    Dim n As Long: n = r.Characters.count
    Dim i As Long
    For i = 1 To n
        Dim chFont As String
        chFont = r.Characters(i).Font.name
        If chFont <> "" Then found(chFont) = True
    Next i
End Sub

' Neu tenFont khop mot bien the DA NHAN DIEN trong "variants" (Dictionary ten font -> Dictionary
' {"Bold","Italic"}), tra True kem gan boldOut/italicOut. Ham tra ket qua qua tham so ByRef vi VBA
' khong co kieu Tuple/Nullable gon.
Private Function TryGetVariant(ByVal fontName As String, ByVal variants As Object, _
        ByRef boldOut As Boolean, ByRef italicOut As Boolean) As Boolean
    If Not variants.Exists(fontName) Then Exit Function
    Dim v As Object: Set v = variants(fontName)
    boldOut = CBool(v("Bold"))
    italicOut = CBool(v("Italic"))
    TryGetVariant = True
End Function

' AP DUNG TRUC TIEP len mot Range: neu Font.Name (CA range, khong tron) khop mot bien the, gan
' Font.Name/Bold/Italic thang tren CHINH Range do - AN TOAN vi day la gan thuoc tinh don gian,
' KHONG qua co che Find noi tai cua Word (nguyen nhan crash da xac dinh - xem dau file). Tra True
' neu co sua.
Private Function ApplyIfVariant(ByVal rng As Range, ByVal variants As Object) As Boolean
    On Error Resume Next
    Dim fName As String: fName = rng.Font.name
    If fName = "" Then Exit Function

    Dim vBold As Boolean, vItalic As Boolean
    If TryGetVariant(fName, variants, vBold, vItalic) Then
        rng.Font.name = BASE_FONT_NAME
        rng.Font.bold = vBold
        rng.Font.Italic = vItalic
        ApplyIfVariant = True
    End If
End Function

' uu tien duong nhanh (ca story mang MOT font duy nhat, khop bien the thi sua thang len CA STORY,
' mot lan gan thuoc tinh - cuc nhanh); khong duoc thi (story tron font) duyet tung DOAN (voi
' wdMainTextStory) hoac ca story nhu mot khoi (story khac); moi Range van tron font thi moi chia
' toi tung ky tu, GOM CAC KY TU LIEN TIEP CUNG FONT thanh MOT run roi ap dung thang len Range cua
' run do (khong ap tung ky tu rieng le - giam so lan gan thuoc tinh). Tra so LAN sua (so Range da
' doi, dung cho log/nhat ky, khong phai so bien the).
Private Function FixStory(ByVal storyRange As Range) As Long
    On Error Resume Next
    Dim fixedHere As Long: fixedHere = 0

    If ApplyIfVariant(storyRange, variants_ByModule) Then
        FixStory = 1
        Exit Function
    End If
    If storyRange.Font.name <> "" Then Exit Function ' font dong nhat nhung KHONG phai bien the

    If storyRange.StoryType = wdMainTextStory Then
        Dim p As paragraph
        For Each p In storyRange.paragraphs
            fixedHere = fixedHere + FixRangeOrRuns(p.Range)
        Next p
    Else
        fixedHere = fixedHere + FixRangeOrRuns(storyRange)
    End If

    FixStory = fixedHere
End Function

' Ap dung cho MOT Range (mot doan, hoac mot story khong phai than bai): thu ca Range truoc (nhanh
' nhat), khong duoc thi chia toi tung ky tu VA GOM cac ky tu LIEN TIEP mang CUNG mot ten font
' thanh tung "run" - CHI ap dung Font.Name/Bold/Italic MOT LAN cho moi run (khong phai moi ky tu),
' giam dang ke so lan gan thuoc tinh so voi ap tung ky tu rieng le.
Private Function FixRangeOrRuns(ByVal rng As Range) As Long
    On Error Resume Next
    Dim r As Range: Set r = rng.Duplicate
    If Len(r.text) = 0 Then Exit Function
    If Right$(r.text, 1) = vbCr Then r.MoveEnd wdCharacter, -1
    If Len(r.text) = 0 Then Exit Function

    If ApplyIfVariant(r, variants_ByModule) Then
        FixRangeOrRuns = 1
        Exit Function
    End If
    If r.Font.name <> "" Then Exit Function ' dong nhat nhung khong phai bien the

    Dim fixedHere As Long: fixedHere = 0
    Dim n As Long: n = r.Characters.count
    Dim runStart As Long: runStart = 0
    Dim runFont As String: runFont = ""
    Dim i As Long
    For i = 1 To n
        Dim chFont As String: chFont = r.Characters(i).Font.name
        If chFont <> runFont Then
            If runStart > 0 And runFont <> "" Then
                fixedHere = fixedHere + ApplyRunIfVariant(r, runStart, i - 1)
            End If
            runStart = i
            runFont = chFont
        End If
    Next i
    If runStart > 0 And runFont <> "" Then
        fixedHere = fixedHere + ApplyRunIfVariant(r, runStart, n)
    End If
    FixRangeOrRuns = fixedHere
End Function

' Cat Range con tu Range cha theo chi so KY TU 1-based [startIdx, endIdx] (Characters(startIdx)
' den Characters(endIdx)), roi ap dung neu font cua run do khop mot bien the.
Private Function ApplyRunIfVariant(ByVal parentRange As Range, ByVal startIdx As Long, ByVal endIdx As Long) As Long
    On Error Resume Next
    Dim runRange As Range
    Set runRange = parentRange.Characters(startIdx)
    runRange.SetRange runRange.Start, parentRange.Characters(endIdx).End
    If ApplyIfVariant(runRange, variants_ByModule) Then ApplyRunIfVariant = 1
End Function

' Diem vao duy nhat -- FindingReporter.RunCheckAndReport goi TU DONG truoc moi lan "Kiem tra",
' CUNG buoc voi EdgeWhitespaceTrimmer/MultiSpaceCollapser. Tra so LAN sua (so Range da doi font)
' -- dung cho nhat ky thao tac.
Public Function NormalizeFontVariants() As Long
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("FontVariantNormalizer.NormalizeFontVariants")
    On Error GoTo ErrHandler
    Dim t0 As Double: t0 = Timer
    DebugTrace.Log "FontVariantNormalizer.NormalizeFontVariants", "Bat dau"

    ' Buoc 1 (CHI DOC): gom ten font phan biet dang dung trong tai lieu, loc ra cac ten khop dung
    ' mau "Times New Roman" + hau to Bold/Italic hop le.
    Dim allNames As Object: Set allNames = Utils.NewDictionary()
    Dim story As Range, cur As Range
    For Each story In ActiveDocument.StoryRanges
        Set cur = story
        Do While Not cur Is Nothing
            If ShouldScanStoryType(cur.StoryType) Then CollectDistinctFontNames cur, allNames
            Set cur = cur.NextStoryRange
        Loop
    Next story
    DebugTrace.Log "FontVariantNormalizer.NormalizeFontVariants", "Gom ten font xong, " & _
        Format$(Timer - t0, "0.00") & "s, " & CStr(allNames.count) & " ten font phan biet"

    ' variants: Dictionary ten font bien the -> Dictionary {"Bold","Italic"} - tra cuu O(1) khi ap
    ' dung, thay vi Collection duyet tuan tu nhu ban cu (Find/Replace khong can tra cuu nhanh vi
    ' moi bien the chi Execute mot lan; ban gan truc tiep nay tra cuu CHO TUNG doan/run nen can
    ' Dictionary).
    Dim variants As Object: Set variants = Utils.NewDictionary()
    Dim nameKey As Variant
    For Each nameKey In allNames.Keys
        Dim suffix As String: suffix = VariantSuffix(CStr(nameKey))
        If Len(suffix) > 0 Then
            If IsRecognizedVariantSuffix(suffix) Then
                Dim v As Object: Set v = Utils.NewDictionary()
                v("Bold") = SuffixIsBold(suffix)
                v("Italic") = SuffixIsItalic(suffix)
                Set variants(CStr(nameKey)) = v
            End If
        End If
    Next nameKey

    If variants.count = 0 Then
        DebugTrace.Log "FontVariantNormalizer.NormalizeFontVariants", "Khong co font bien the - thoat som"
        NormalizeFontVariants = 0
        Exit Function
    End If

    Dim opName As String
    opName = "Chu" & ChrW(&H1EA9) & "n h" & ChrW(&HF3) & "a font bi" & ChrW(&H1EBF) & "n th" & ChrW(&H1EC3)
    Dim opStarted As Boolean
    Utils.BeginOperation opName
    opStarted = True

    ' Buoc 2 (GHI): MOT LUOT duy nhat qua toan bo story (khac ban cu - lap rieng MOI bien the LA
    ' MOT luot rieng qua ca tai lieu; nay xu ly TAT CA bien the CUNG LUC trong MOT luot, vi gan
    ' truc tiep tra cuu Dictionary theo ten font gap phai, khong can biet truoc dang tim ten nao).
    Set variants_ByModule = variants
    Dim fixedCount As Long: fixedCount = 0
    Dim s2 As Range, c2 As Range
    For Each s2 In ActiveDocument.StoryRanges
        Set c2 = s2
        Do While Not c2 Is Nothing
            If ShouldScanStoryType(c2.StoryType) Then
                fixedCount = fixedCount + FixStory(c2)
            End If
            Set c2 = c2.NextStoryRange
        Loop
    Next s2
    Set variants_ByModule = Nothing
    DebugTrace.Log "FontVariantNormalizer.NormalizeFontVariants", "Da sua " & CStr(fixedCount) & _
        " vung, " & Format$(Timer - t0, "0.00") & "s"

    Utils.EndOperation fixedCount, False

    ' Buoc 3 (CHI DOC, tu xac nhan) - quet lai sau khi sua, canh bao qua log neu VAN CON ten bien
    ' the (chung to mot story nao do chua duoc quet toi) - giu tu, van huu ich du da doi ky thuat
    ' sua.
    Dim afterNames As Object: Set afterNames = Utils.NewDictionary()
    Dim story3 As Range, cur3 As Range
    For Each story3 In ActiveDocument.StoryRanges
        Set cur3 = story3
        Do While Not cur3 Is Nothing
            If ShouldScanStoryType(cur3.StoryType) Then CollectDistinctFontNames cur3, afterNames
            Set cur3 = cur3.NextStoryRange
        Loop
    Next story3
    Dim stillVariant As String: stillVariant = ""
    Dim afterKey As Variant
    For Each afterKey In afterNames.Keys
        Dim afterSuffix As String: afterSuffix = VariantSuffix(CStr(afterKey))
        If Len(afterSuffix) > 0 Then
            If IsRecognizedVariantSuffix(afterSuffix) Then
                stillVariant = stillVariant & """" & CStr(afterKey) & """ "
            End If
        End If
    Next afterKey
    If Len(stillVariant) > 0 Then
        DebugTrace.Log "FontVariantNormalizer.NormalizeFontVariants", _
            "CANH BAO: sau khi sua VAN CON ten font bien the: " & stillVariant & _
            "(co the nam trong mot story chua duoc quet)"
    End If

    DebugTrace.Log "FontVariantNormalizer.NormalizeFontVariants", "Hoan tat, " & _
        Format$(Timer - t0, "0.00") & "s, fixedCount=" & fixedCount
    NormalizeFontVariants = fixedCount
    Exit Function

ErrHandler:
    Set variants_ByModule = Nothing
    DebugTrace.LogErr "FontVariantNormalizer.NormalizeFontVariants", "loi giua chung, " & _
        Format$(Timer - t0, "0.00") & "s", Err.number, Err.description
    If opStarted Then
        Utils.AbortOperation Err.description
    Else
        Err.Raise Err.number, "FontVariantNormalizer.NormalizeFontVariants", Err.description
    End If
    NormalizeFontVariants = 0
End Function
