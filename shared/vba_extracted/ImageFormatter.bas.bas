Attribute VB_Name = "ImageFormatter"
'==============================================================
' ImageFormatter â€” Can giua, bo lui dau dong, co be ngang vua vung trinh bay - nut 5.5
' O day xu ly ca hai loai anh:
' - Anh inline (ActiveDocument.InlineShapes): can giua doan, bo lui dau dong, co be ngang.
' - Anh noi (ActiveDocument.Shapes, Type = msoPicture): co be ngang + can giua theo vung trinh
'   bay, dung wdShapeCenter/RelativeHorizontalPosition (khong co doan/Alignment vi anh noi khong
'   "nam trong" mot doan theo nghia canh le).
' Xac dinh section chua anh qua Range.Sections(1).Index (anh inline) hoac Anchor.Sections(1).Index
' (anh noi). Hieu nang: boc trong Utils.BeginOperation/EndOperation de tat ScreenUpdating - tai
' lieu nhieu anh khong tat se rat cham. Moi thu tuc cong khai bat dau bang On Error GoTo
' ErrHandler. Dung AscW/ChrW, khong dung Asc/Chr (phu thuoc code page) - CLAUDE.md muc 5.
'==============================================================
Option Explicit

' Nut 5.5 "Chuan hoa anh" - xu ly CA anh inline lan anh noi, tra ve tong so anh da xu ly (khong
' tinh anh trong bang - bo qua co y, khong phai loi).
Public Function NormalizeImages() As Long
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ImageFormatter.NormalizeImages")
    On Error GoTo ErrHandler

    Dim opName As String
    opName = "Chu" & ChrW(&H1EA9) & "n h" & ChrW(&HF3) & "a " & ChrW(&H1EA3) & "nh"
    Utils.BeginOperation opName

    Dim inlineProcessed As Long
    Dim inlineSkippedInTable As Long
    NormalizeInlineImages inlineProcessed, inlineSkippedInTable

    Dim floatingProcessed As Long
    Dim floatingSkippedInTable As Long
    NormalizeFloatingImages floatingProcessed, floatingSkippedInTable

    Dim totalProcessed As Long
    totalProcessed = inlineProcessed + floatingProcessed

    Utils.EndOperation totalProcessed, False
    NormalizeImages = totalProcessed
    Exit Function
ErrHandler:
    Utils.AbortOperation Err.description
    NormalizeImages = -1
End Function

Private Sub NormalizeInlineImages(ByRef processedCount As Long, ByRef skippedInTableCount As Long)
    processedCount = 0
    skippedInTableCount = 0

    Dim shp As word.InlineShape
    Dim sectionIndex As Long
    Dim contentWidthPt As Double

    For Each shp In ActiveDocument.InlineShapes
        If shp.Range.Information(wdWithInTable) Then
            skippedInTableCount = skippedInTableCount + 1
        Else
            sectionIndex = shp.Range.sections(1).Index
            contentWidthPt = ContentWidthPointsForSection(sectionIndex)

            shp.Range.ParagraphFormat.alignment = wdAlignParagraphCenter
            shp.Range.ParagraphFormat.FirstLineIndent = 0
            shp.Range.ParagraphFormat.LeftIndent = 0

            ' Co be ngang vua vung trinh bay, giu ty le - anh nho hon thi giu nguyen, khong phong
            ' to. LockAspectRatio PHAI dat TRUOC khi doi Width.
            If shp.width > contentWidthPt Then
                shp.LockAspectRatio = msoTrue
                shp.width = contentWidthPt
            End If

            processedCount = processedCount + 1
        End If
    Next shp
End Sub

' Duyet ActiveDocument.Shapes, chi xu ly Type = msoPicture (bo qua duong ke, text box, WordArt...
' - ngoai pham vi nut 5.5).
Private Sub NormalizeFloatingImages(ByRef processedCount As Long, ByRef skippedInTableCount As Long)
    processedCount = 0
    skippedInTableCount = 0

    Dim shp As word.Shape
    Dim sectionIndex As Long
    Dim contentWidthPt As Double

    For Each shp In ActiveDocument.Shapes
        If shp.Type = msoPicture Then
            If shp.anchor.Information(wdWithInTable) Then
                ' Dong bo anh inline - anh noi neo trong bang cung bo qua, khong dung tay vao.
                skippedInTableCount = skippedInTableCount + 1
            Else
                sectionIndex = shp.anchor.sections(1).Index
                contentWidthPt = ContentWidthPointsForSection(sectionIndex)

                If shp.width > contentWidthPt Then
                    shp.LockAspectRatio = msoTrue
                    shp.width = contentWidthPt
                End If

                ' Can giua anh noi theo vung trinh bay - khong co "doan chua anh" nhu inline nen
                ' dung wdShapeCenter (canh giua tuyet doi so voi RelativeHorizontalPosition da
                ' chon, o day la trang giay) thay vi ParagraphFormat.Alignment.
                shp.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
                shp.left = wdShapeCenter

                processedCount = processedCount + 1
            End If
        End If
    Next shp
End Sub

' Be rong vung trinh bay (point) cua section thu sectionIndex - quy doi tu
' PageFormatter.GetContentWidthMm (mm) bang cung he so units.mmToPoint da dung o
' PageFormatter/StyleBuilder, khong hard-code (CLAUDE.md muc 3.1).
Private Function ContentWidthPointsForSection(ByVal sectionIndex As Long) As Double
    On Error GoTo ErrHandler
    Dim factor As Double
    factor = RuleLoader.GetFormatSpec()("units")("mmToPoint")
    ContentWidthPointsForSection = PageFormatter.GetContentWidthMm(sectionIndex) * factor
    Exit Function
ErrHandler:
    Err.Raise Err.number, "ImageFormatter.ContentWidthPointsForSection", Err.description
End Function
