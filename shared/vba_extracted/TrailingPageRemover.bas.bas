Attribute VB_Name = "TrailingPageRemover"
'==============================================================
' TrailingPageRemover â€” Xoa trang trang thua sinh ra khi bang sat mep duoi trang cuoi - nut 6.5
' Ban Legacy co ComputeStatistics(wdStatisticPages) va Range.Information(wdActiveEndPageNumber)
' nen do truc tiep duoc:
' Moi thu tuc cong khai bat dau bang On Error GoTo ErrHandler. Dung AscW/ChrW, khong dung Asc/Chr
' (phu thuoc code page) - CLAUDE.md muc 5.
'==============================================================
Option Explicit

Private Const MIN_FONT_SIZE_PT As Single = 1

' Gian dong nho nhat, dat theo LineSpacingRule = wdLineSpaceExactly de gia tri point ap dung DUNG
' NGHIA (khac wdLineSpaceMultiple noi LineSpacing la % - xem ghi chu cua StyleBuilder.bas).
Private Const MIN_LINE_SPACING_PT As Single = 1

' ============================================================================
' Nut 6.5 "Xoa trang thua o cuoi". Tra ve: 1 neu da xu ly va xac nhan bot dung mot trang, 0 neu
' khong co trang thua (bao MsgBox, khong dong gi) hoac xu ly xong nhung khong bot duoc trang (hoan
' tac, bao MsgBox), -1 neu loi.
' ============================================================================
Public Function RemoveTrailingPage() As Long
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("TrailingPageRemover.RemoveTrailingPage")
    On Error GoTo ErrHandler

    Dim lastPara As word.paragraph
    Set lastPara = ActiveDocument.paragraphs.last

    ' Dieu 1 (Viec phai lam ) - doan cuoi rong. Doi chieu Len(Trim$) <= 1 vi Range.Text cua doan
    ' luon co it nhat ky tu xuong dong wdParagraphMark (do dai 1).
    If Len(Trim$(lastPara.Range.text)) > 1 Then
        ShowNoTrailingPageMessage
        RemoveTrailingPage = 0
        Exit Function
    End If

    ' Doan cuoi phai o CAP BODY (khong nam trong bang) - dung API goi y thu hai cua
    ' ("Paragraphs.Last.Range.Information(wdWithInTable) = False ket hop kiem bang lien truoc").
    If lastPara.Range.Information(wdWithInTable) Then
        ShowNoTrailingPageMessage
        RemoveTrailingPage = 0
        Exit Function
    End If

    ' Khong co noi dung nao truoc doan cuoi (doan cuoi la doan duy nhat cua tai lieu) - khong co
    ' gi de kiem tra "bang lien truoc".
    If lastPara.Range.Start = 0 Then
        ShowNoTrailingPageMessage
        RemoveTrailingPage = 0
        Exit Function
    End If

    ' Dieu 2 - phan tu ngay truoc doan cuoi la bang: ky tu ngay truoc diem bat dau doan cuoi nam
    ' TRONG bang (wdWithInTable) tuc la bang ket thuc lien ke, khong co noi dung nao khac xen giua
    Dim precedingPos As word.Range
    Set precedingPos = ActiveDocument.Range(lastPara.Range.Start - 1, lastPara.Range.Start)
    If Not precedingPos.Information(wdWithInTable) Then
        ShowNoTrailingPageMessage
        RemoveTrailingPage = 0
        Exit Function
    End If

    ' Dieu 3 - trang cuoi trong: doan cuoi phai dang o trang CUOI CUNG cua tai lieu (kiem tra
    ' phong thu, luon dung trong thuc te) VA phan noi dung ngay truoc no (cuoi bang) phai nam o
    ' mot trang SOM HON - neu khong, trang cuoi cung con chua noi dung khac ngoai doan rong nay,
    ' khong phai tinh huong trang thua do bang sinh ra.
    Dim totalPagesBefore As Long
    totalPagesBefore = ActiveDocument.ComputeStatistics(wdStatisticPages)

    If lastPara.Range.Information(wdActiveEndPageNumber) <> totalPagesBefore Then
        ShowNoTrailingPageMessage
        RemoveTrailingPage = 0
        Exit Function
    End If
    If precedingPos.Information(wdActiveEndPageNumber) >= totalPagesBefore Then
        ShowNoTrailingPageMessage
        RemoveTrailingPage = 0
        Exit Function
    End If

    ' ----- Du ba dieu kien - tien hanh xu ly
    Dim opName As String
    opName = "X" & ChrW(&HF3) & "a trang th" & ChrW(&H1EEB) & "a " & ChrW(&H1EDF) & " cu" & _
        ChrW(&H1ED1) & "i"
    Utils.BeginOperation opName

    ' Luu lai gia tri goc de hoan tac neu do trang sau khi sua van khong bot (khong dua vao
    ' Application.Undo o day - UndoRecord tuy chinh dang mo, goi Undo giua chung khong dam bao
    ' hanh vi dung; ghi lai va gan tra truc tiep an toan hon).
    Dim originalFontSize As Single
    Dim originalSpaceBefore As Single
    Dim originalSpaceAfter As Single
    Dim originalLineSpacingRule As WdLineSpacing
    Dim originalLineSpacing As Single
    originalFontSize = lastPara.Range.Font.size
    originalSpaceBefore = lastPara.SpaceBefore
    originalSpaceAfter = lastPara.SpaceAfter
    originalLineSpacingRule = lastPara.LineSpacingRule
    originalLineSpacing = lastPara.lineSpacing

    ApplyMinimalFormat lastPara

    Dim totalPagesAfter As Long
    totalPagesAfter = ActiveDocument.ComputeStatistics(wdStatisticPages)

    If totalPagesAfter = totalPagesBefore - 1 Then
        Utils.EndOperation 1, False
        RemoveTrailingPage = 1
    Else
        ' Khong bot dung mot trang - hoan tac ve gia tri goc, ghi log co canh bao, bao nguoi dung
        ' "Khong xu ly duoc".
        lastPara.Range.Font.size = originalFontSize
        lastPara.SpaceBefore = originalSpaceBefore
        lastPara.SpaceAfter = originalSpaceAfter
        lastPara.LineSpacingRule = originalLineSpacingRule
        lastPara.lineSpacing = originalLineSpacing

        Utils.EndOperation 0, True
        ShowCannotProcessMessage
        RemoveTrailingPage = 0
    End If

    Exit Function
ErrHandler:
    Utils.AbortOperation Err.description
    RemoveTrailingPage = -1
End Function

Private Sub ApplyMinimalFormat(ByVal target As word.paragraph)
    target.Range.Font.size = MIN_FONT_SIZE_PT
    target.SpaceBefore = 0
    target.SpaceAfter = 0
    target.LineSpacingRule = wdLineSpaceExactly
    target.lineSpacing = MIN_LINE_SPACING_PT
End Sub

' Thong bao khi tai lieu khong co trang thua o cuoi (thieu mot trong ba dieu kien nhan dien) - bao
' ro, khong dong gi.
Private Sub ShowNoTrailingPageMessage()
    Dim msg As String
    msg = "T" & ChrW(&HE0) & "i li" & ChrW(&H1EC7) & "u kh" & ChrW(&HF4) & "ng c" & _
        ChrW(&HF3) & " trang th" & ChrW(&H1EEB) & "a " & ChrW(&H1EDF) & " cu" & _
        ChrW(&H1ED1) & "i."
    MsgBoxW.Show msg, vbInformation, "Chu" & ChrW(&H1EA9) & "n h" & ChrW(&HF3) & "a th" & _
        ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
End Sub

' Thong bao khi da thu nho dinh dang doan cuoi nhung do lai van khong bot duoc mot trang (da hoan
' tac ve nguyen trang) - dieu 3 cua.
Private Sub ShowCannotProcessMessage()
    Dim msg As String
    msg = "Kh" & ChrW(&HF4) & "ng x" & ChrW(&H1EED) & " l" & ChrW(&HFD) & " " & _
        ChrW(&H111) & ChrW(&H1B0) & ChrW(&H1EE3) & "c. Vui l" & ChrW(&HF2) & "ng ki" & _
        ChrW(&H1EC3) & "m tra th" & ChrW(&H1EE7) & " c" & ChrW(&HF4) & "ng."
    MsgBoxW.Show msg, vbExclamation, "Chu" & ChrW(&H1EA9) & "n h" & ChrW(&HF3) & "a th" & _
        ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
End Sub
