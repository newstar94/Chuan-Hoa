Attribute VB_Name = "PageNumberFormatter"
Option Explicit

' ============================================================================
' InsertPageNumbers - nut 2.3
' ============================================================================

' Chen field PAGE (chu so A Rap) vao header chinh cua MOI section, canh giua theo chieu ngang
' trong phan le tren, an so trang o trang 1 cua toan tai lieu. Chay lai nhieu lan khong sinh field
' trung lap - quet xoa moi field PAGE cu truoc khi chen lai. Tra ve Dictionary {sectionCount,
' pageNumberFontSizePt} hoac Nothing neu loi giua chung (da hien MsgBox qua Utils.AbortOperation).
Public Function InsertPageNumbers() As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("PageNumberFormatter.InsertPageNumbers")
    On Error GoTo ErrHandler

    Dim opName As String
    opName = "Ch" & ChrW(&H1EBF) & "n s" & ChrW(&H1ED1) & " trang"
    Utils.BeginOperation opName

    Dim pageNumberSizePt As Double
    pageNumberSizePt = ResolvePageNumberFontSize()

    Dim headerStyleName As String, pageNumberStyleName As String
    headerStyleName = LookupStyleName("Header")
    pageNumberStyleName = LookupStyleName("PageNumber")

    RemoveExistingPageFields

    Dim sectionCount As Long
    sectionCount = ActiveDocument.sections.count

    Dim i As Long
    For i = 1 To sectionCount
        ApplySectionPageNumber ActiveDocument.sections(i), (i = 1), pageNumberSizePt, _
            headerStyleName, pageNumberStyleName
    Next i

    Dim Result As Object
    Set Result = Utils.NewDictionary()
    Result("sectionCount") = sectionCount
    Result("pageNumberFontSizePt") = pageNumberSizePt

    Utils.EndOperation sectionCount, False
    Set InsertPageNumbers = Result
    Exit Function
ErrHandler:
    Utils.AbortOperation Err.description
    Set InsertPageNumbers = Nothing
End Function

' Ten Word that cua style trong bo-styles.json ung voi styleId - doc qua RuleLoader thay vi hard-
' code chuoi "header"/"page number" tai day (CLAUDE.md muc 3.1), dong bo tuyet doi voi
' StyleBuilder.bas - style nao StyleBuilder da dung, ham nay dung DUNG ten do.
Private Function LookupStyleName(ByVal styleId As String) As String
    On Error GoTo ErrHandler
    Dim sheet As Object
    Set sheet = RuleLoader.GetStyleSheet()
    Dim def As Variant
    For Each def In sheet("styles")
        If def("styleId") = styleId Then
            LookupStyleName = CStr(def("name"))
            Exit Function
        End If
    Next def
    Err.Raise vbObjectError + 543, "PageNumberFormatter.LookupStyleName", _
        "Khong tim thay styleId """ & styleId & """ trong bo-styles.json."
    Exit Function
ErrHandler:
    Err.Raise Err.number, "PageNumberFormatter.LookupStyleName", Err.description
End Function

' ============================================================================
' Suy ra co chu so trang tu co chu doan van "Normal" dau tien co noi dung
' ============================================================================

Private Function ResolvePageNumberFontSize() As Double
    On Error GoTo ErrHandler
    Dim p As word.paragraph
    Set p = FindNormalBodyParagraph()
    If p Is Nothing Then
        ' Chi khi THAT SU khong co doan nao ngoai bang chua noi dung (van ban trong hoan toan,
        ' hoac toan bo noi dung nam trong bang - hiem) - dung quy uoc "<=1" giong
        ' RibbonCallbacks.IsDocEmpty (Content.Text luon con it nhat mot Chr(13)).
        If Len(Trim$(ActiveDocument.Content.text)) <= 1 Then
            ' van ban TRUC SU trong thi dung co chu HIEN TAI cua style "Normal" hien co lam co chu
            ' so trang - khong tra fontSizeSets nao ca (khong co noi dung de doi chieu bo co chu
            ' 13/14/15).
            ResolvePageNumberFontSize = ActiveDocument.Styles(wdStyleNormal).Font.size
            Exit Function
        End If

        Err.Raise vbObjectError + 544, "PageNumberFormatter.ResolvePageNumberFontSize", _
            "Khong tim thay doan van nao co noi dung ngoai bang (toan bo noi dung nam trong " & _
            "bang). Hay go them noi dung than van ban ngoai bang truoc khi chen so trang."
    End If

    Dim bodySize As Variant
    bodySize = p.Range.Font.size
    If bodySize = wdUndefined Then
        Err.Raise vbObjectError + 545, "PageNumberFormatter.ResolvePageNumberFontSize", _
            "Khong xac dinh duoc co chu than van ban (doan van co co chu khong dong nhat " & _
            "ben trong). Hay dat co chu thong nhat (Co chu 13 hoac Co chu 14) cho than van " & _
            "ban truoc khi chen so trang."
    End If

    Dim fontSizeSets As Object
    Set fontSizeSets = RuleLoader.GetFormatSpec()("fontSizeSets")

    Dim sizeSetKey As String
    If CDbl(bodySize) = CDbl(fontSizeSets("set1")("bodyText")) Then
        sizeSetKey = "set1"
    ElseIf CDbl(bodySize) = CDbl(fontSizeSets("set2")("bodyText")) Then
        sizeSetKey = "set2"
    Else
        Err.Raise vbObjectError + 546, "PageNumberFormatter.ResolvePageNumberFontSize", _
            "Co chu than van ban hien tai (" & CStr(bodySize) & "pt) khong khop bo co chu " & _
            "nao theo ND 30 (Co chu 13 hoac Co chu 14). Hay dung nut ""Dung bo Styles"" " & _
            "truoc khi chen so trang."
    End If

    ResolvePageNumberFontSize = CDbl(fontSizeSets(sizeSetKey)("pageNumber"))
    Exit Function
ErrHandler:
    Err.Raise Err.number, "PageNumberFormatter.ResolvePageNumberFontSize", Err.description
End Function

Private Function FindNormalBodyParagraph() As word.paragraph
    Dim p As word.paragraph
    For Each p In ActiveDocument.paragraphs
        If Not p.Range.Information(wdWithInTable) Then
            ' Range.Text luon co it nhat mot ky tu danh dau doan (Chr 13) o cuoi - "> 1" de loai
            ' doan trong.
            If Len(Trim$(p.Range.text)) > 1 Then
                Set FindNormalBodyParagraph = p
                Exit Function
            End If
        End If
    Next p
    Set FindNormalBodyParagraph = Nothing
End Function

' ============================================================================
' Quet xoa field PAGE cu - dieu kien 3 cua: chay lai nhieu lan khong sinh field trung.
' ============================================================================

' Duyet toan bo Header/Footer (ca ba loai: Primary/FirstPage/EvenPages) cua MOI section, xoa moi
' field kieu wdFieldPage tim thay. Boc On Error Resume Next rieng cho tung HeaderFooter.Range -
' mot section dang LinkToPrevious=True tro toi noi dung CHIA SE cua section truoc, doc/xoa qua
' Range cua no van hop le nhung de an toan truoc truong hop bien chua luong truoc (vd tai lieu tu
' ban dung khac), khong de mot loi o day lam hong ca thao tac.
Private Sub RemoveExistingPageFields()
    Dim i As Long
    For i = 1 To ActiveDocument.sections.count
        RemovePageFieldsIn ActiveDocument.sections(i).Headers(wdHeaderFooterPrimary)
        RemovePageFieldsIn ActiveDocument.sections(i).Headers(wdHeaderFooterFirstPage)
        RemovePageFieldsIn ActiveDocument.sections(i).Headers(wdHeaderFooterEvenPages)
        RemovePageFieldsIn ActiveDocument.sections(i).Footers(wdHeaderFooterPrimary)
        RemovePageFieldsIn ActiveDocument.sections(i).Footers(wdHeaderFooterFirstPage)
        RemovePageFieldsIn ActiveDocument.sections(i).Footers(wdHeaderFooterEvenPages)
    Next i
End Sub

Private Sub RemovePageFieldsIn(ByVal hf As word.HeaderFooter)
    On Error Resume Next
    Dim j As Long
    For j = hf.Range.Fields.count To 1 Step -1
        If hf.Range.Fields(j).Type = wdFieldPage Then hf.Range.Fields(j).Delete
    Next j
    On Error GoTo 0
End Sub

' ============================================================================
' Ap dung cho mot section
' ============================================================================

Private Sub ApplySectionPageNumber(ByVal sect As word.section, ByVal isFirstSection As Boolean, _
        ByVal pageNumberSizePt As Double, ByVal headerStyleName As String, _
        ByVal pageNumberStyleName As String)
    sect.PageSetup.DifferentFirstPageHeaderFooter = isFirstSection

    Dim hdr As word.HeaderFooter
    Set hdr = sect.Headers(wdHeaderFooterPrimary)

    If isFirstSection Then
        ' Section dau khong co "section truoc" de lien ket - tu ghi noi dung.
        hdr.LinkToPrevious = False
        hdr.Range.Delete

        On Error Resume Next
        hdr.Range.Style = headerStyleName
        On Error GoTo 0

        Dim fld As word.Field
        Set fld = ActiveDocument.Fields.Add(hdr.Range, wdFieldPage)

        On Error Resume Next
        fld.Result.Style = pageNumberStyleName
        On Error GoTo 0

        ' Ap THANG phong/co chu cho field - khong phu thuoc style "page number" gan thanh cong hay
        ' khong (cung ly do voi ParagraphFormat duoi day).
        fld.Result.Font.name = CStr(RuleLoader.GetFormatSpec()("font")("name"))
        fld.Result.Font.size = pageNumberSizePt

        ' Dat SAU CUNG (sau Fields.Add) - day moi la thu THAT SU quyet dinh hinh thuc hien tren
        ' man hinh, khong bi ghi de boi buoc nao chay sau no nua.
        With hdr.Range.ParagraphFormat
            .alignment = wdAlignParagraphCenter
            .FirstLineIndent = 0
            .LeftIndent = 0
            .RightIndent = 0
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
    Else
        hdr.LinkToPrevious = True
    End If
End Sub
