Attribute VB_Name = "TableFormatter"
'==============================================================
' TableFormatter â€” AutoFit, can giua bang, lap dong tieu de, canh doc trong o - nut 5.1-5.4
' Table.AutoFitBehavior wdAutoFitWindow tu tinh be rong dung bang vung trinh bay CUA SECTION CHUA
' BANG (khong phai toan cuc)
' Bang long nhau: ActiveDocument.Tables duyet CA bang long ben trong lan bang ngoai cung trong
' Word Object Model - khac gia dinh ban dau cua dac ta. Nut 5.3/5.4 khong can loc vi pham vi da la
' "bang dang chon" (Selection.Tables(1))
' Moi thu tuc cong khai bat dau bang On Error GoTo ErrHandler. Dung AscW/ChrW, khong dung Asc/Chr
' (phu thuoc code page) - CLAUDE.md muc 5.
'==============================================================
Option Explicit

' ============================================================================
' Thong bao dung chung
' ============================================================================

' Khong co bang nao chua con tro - dung cho ca nut 5.3 va 5.4.
Private Function NoTableSelectedMessage() As String
    ' "Con trá»�" sai, phai la "Con trá»�" - ChrW(&H1ED0) la "á»�" (hoa), dung nham vi tri cua
    ' ChrW(&H1ECF) la "á»�" (thuong, dung).
    NoTableSelectedMessage = "Con tr" & ChrW(&H1ECF) & " hi" & ChrW(&H1EC7) & "n kh" & _
        ChrW(&HF4) & "ng n" & ChrW(&H1EB1) & "m trong b" & ChrW(&H1EA3) & "ng n" & _
        ChrW(&HE0) & "o. H" & ChrW(&HE3) & "y " & ChrW(&H111) & ChrW(&H1EB7) & "t " & _
        "con tr" & ChrW(&H1ECF) & " v" & ChrW(&HE0) & "o m" & ChrW(&H1ED9) & "t " & _
        ChrW(&HF4) & " b" & ChrW(&H1EA5) & "t k" & ChrW(&H1EF3) & " trong b" & _
        ChrW(&H1EA3) & "ng c" & ChrW(&H1EA7) & "n canh d" & ChrW(&H1ECD) & "c r" & _
        ChrW(&H1ED3) & "i b" & ChrW(&H1EA5) & "m l" & ChrW(&H1EA1) & "i n" & ChrW(&HFA) & "t."
End Function

' ============================================================================
' Nut 5.1 va 5.2 - duyet MOI bang NGOAI CUNG cua toan tai lieu (khong gioi han vung chon).
' ============================================================================

' Nut 5.1 "Chuan hoa bang" - AutoFit to Window (tu tinh be rong theo vung trinh bay cua section
' chua bang) + canh giua CA BANG so voi cot trang (Rows.Alignment - KHONG phai canh trong o, dieu
' 1 cua ).
Public Function NormalizeTables() As Long
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("TableFormatter.NormalizeTables")
    On Error GoTo ErrHandler

    Dim opName As String
    opName = "Chu" & ChrW(&H1EA9) & "n h" & ChrW(&HF3) & "a b" & ChrW(&H1EA3) & "ng"
    Utils.BeginOperation opName

    Dim t As word.table
    Dim tableCount As Long
    tableCount = 0
    For Each t In ActiveDocument.Tables
        If t.NestingLevel = 1 Then
            t.AutoFitBehavior wdAutoFitWindow
            t.Rows.alignment = wdAlignRowCenter
            tableCount = tableCount + 1
        End If
    Next t

    Utils.EndOperation tableCount, False
    NormalizeTables = tableCount
    Exit Function
ErrHandler:
    Utils.AbortOperation Err.description
    NormalizeTables = -1
End Function

' Nut 5.2 "Lap dong tieu de" - bat Repeat Header Rows cho hang dau. HeadingFormat = True chi NANG
' len, khong ha xuong: neu nguoi dung da tu dat nhieu hon 1 hang tieu de lap lai (hang 2, 3...
' cung HeadingFormat = True), dat lai True cho hang 1 khong dong den cac hang do
Public Function RepeatHeaderRows() As Long
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("TableFormatter.RepeatHeaderRows")
    On Error GoTo ErrHandler

    Dim opName As String
    opName = "L" & ChrW(&H1EB7) & "p d" & ChrW(&HF2) & "ng ti" & ChrW(&HEA) & "u " & _
        ChrW(&H111) & ChrW(&H1EC1)
    Utils.BeginOperation opName

    Dim t As word.table
    Dim tableCount As Long
    tableCount = 0
    For Each t In ActiveDocument.Tables
        If t.NestingLevel = 1 Then
            t.Rows(1).HeadingFormat = True
            tableCount = tableCount + 1
        End If
    Next t

    Utils.EndOperation tableCount, False
    RepeatHeaderRows = tableCount
    Exit Function
ErrHandler:
    Utils.AbortOperation Err.description
    RepeatHeaderRows = -1
End Function

' ============================================================================
' Nut 5.3 va 5.4 - CHI bang chua con tro (Selection.Tables(1)), khong bang nao khac.
' ============================================================================

' Nut 5.3 "Can dinh o" - canh doc moi o cua bang dang chon ve Top.
Public Sub AlignCellsTop()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("TableFormatter.AlignCellsTop")
    On Error GoTo ErrHandler
    ApplySelectedTableVerticalAlignment wdCellAlignVerticalTop, "C" & ChrW(&H103) & "n " & _
        ChrW(&H111) & ChrW(&H1EC9) & "nh " & ChrW(&HF4)
    Exit Sub
ErrHandler:
    Utils.AbortOperation Err.description
End Sub

' Nut 5.4 "Can giua o" - canh doc moi o cua bang dang chon ve Center.
Public Sub AlignCellsCenter()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("TableFormatter.AlignCellsCenter")
    On Error GoTo ErrHandler
    ApplySelectedTableVerticalAlignment wdCellAlignVerticalCenter, "C" & ChrW(&H103) & "n gi" & _
        ChrW(&H1EEF) & "a " & ChrW(&HF4)
    Exit Sub
ErrHandler:
    Utils.AbortOperation Err.description
End Sub

Private Sub ApplySelectedTableVerticalAlignment(ByVal cellAlign As WdCellVerticalAlignment, _
        ByVal opName As String)
    If Selection.Tables.count = 0 Then
        MsgBoxW.Show NoTableSelectedMessage(), vbExclamation, "Chu" & ChrW(&H1EA9) & "n h" & _
            ChrW(&HF3) & "a th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
        Exit Sub
    End If

    Utils.BeginOperation opName

    Dim tbl As word.table
    Set tbl = Selection.Tables(1)

    Dim cel As word.Cell
    Dim cellCount As Long
    cellCount = 0
    For Each cel In tbl.Range.Cells
        cel.VerticalAlignment = cellAlign
        cellCount = cellCount + 1
    Next cel

    Utils.EndOperation cellCount, False
End Sub
