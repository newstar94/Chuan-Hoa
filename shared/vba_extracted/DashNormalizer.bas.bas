Attribute VB_Name = "DashNormalizer"
Option Explicit

Private Function EnDash() As String
    EnDash = ChrW(&H2013)
End Function

Private Function EmDash() As String
    EmDash = ChrW(&H2014)
End Function

' Tim moi vi tri (0-based) cua En Dash/Em Dash trong text -- tra ve Collection cac Long, TANG DAN
' theo vi tri xuat hien (doi chieu GIAM DAN o noi goi truoc khi ap dung).
Private Function FindDashPositions(ByVal text As String) As Collection
    ' ten bien cuc bo "enDash"/"emDash" TRUNG (khong phan biet hoa/thuong) voi ten ham
    ' EnDash/EmDash cung module - VBA khong bien dich duoc, bao loi "Expected array" tai dong gan.
    ' Doi ten bien thanh "dashEn"/"dashEm" de tranh dung ten.
    Dim Result As New Collection
    Dim dashEn As String: dashEn = EnDash()
    Dim dashEm As String: dashEm = EmDash()
    Dim n As Long: n = Len(text)
    Dim i As Long
    For i = 1 To n
        Dim ch As String: ch = Mid$(text, i, 1)
        If ch = dashEn Or ch = dashEm Then Result.Add i - 1 ' quy ve 0-based
    Next i
    Set FindDashPositions = Result
End Function

' Diem vao duy nhat -- FindingReporter.RunCheckCore goi TU DONG truoc moi lan "Kiem tra". Tra so
' cho da thay - dung cho nhat ky thao tac.
Public Function NormalizeDashes() As Long
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("DashNormalizer.NormalizeDashes")
    Dim t0 As Double: t0 = Timer
    DebugTrace.Log "DashNormalizer.NormalizeDashes", "Bat dau"

    ' CHI can noi dung doan van - dung ban chup NHE, tranh cham anh loi khong can thiet (T-71,
    ' xem ghi chu dau DocumentSnapshot.CaptureParagraphsOnlySnapshot).
    On Error GoTo StepSnapshot
    Dim snapshot As Object: Set snapshot = DocumentSnapshot.CaptureParagraphsOnlySnapshot()
    DebugTrace.Log "DashNormalizer.NormalizeDashes", "CaptureParagraphsOnlySnapshot xong, " & Format$(Timer - t0, "0.00") & "s"

    On Error GoTo ErrHandler
    Dim targets As New Collection
    Dim p As ParagraphSnapshot
    For Each p In snapshot("Paragraphs")
        Dim positions As Collection: Set positions = FindDashPositions(p.text)
        If positions.count > 0 Then
            Dim item As Object: Set item = Utils.NewDictionary()
            item("Index") = p.Index
            Set item("Positions") = positions
            targets.Add item
        End If
    Next p

    If targets.count = 0 Then
        DebugTrace.Log "DashNormalizer.NormalizeDashes", "Khong co cho can thay - thoat som"
        NormalizeDashes = 0
        Exit Function
    End If

    Dim indexMap As Object: Set indexMap = DocumentSnapshot.BuildSnapshotIndexMap()
    DebugTrace.Log "DashNormalizer.NormalizeDashes", "BuildSnapshotIndexMap xong, " & Format$(Timer - t0, "0.00") & "s"

    Dim opName As String
    opName = "Thay d" & ChrW(&H1EA5) & "u g" & ChrW(&H1EA1) & "ch ngang b" & ChrW(&H1EB1) & "ng d" & _
        ChrW(&H1EA5) & "u g" & ChrW(&H1EA1) & "ch n" & ChrW(&H1ED1) & "i"
    Dim opStarted As Boolean
    Utils.BeginOperation opName
    opStarted = True

    Dim fixedCount As Long: fixedCount = 0
    Dim t As Variant
    For Each t In targets
        Dim idx As Long: idx = CLng(t("Index"))
        If indexMap.Exists(idx) Then
            Dim wordPara As word.paragraph: Set wordPara = ActiveDocument.paragraphs(CLng(indexMap(idx)))
            Dim rng As word.Range: Set rng = wordPara.Range
            rng.MoveEnd wdCharacter, -1 ' bo dau doan o cuoi Range
            Dim paraStart As Long: paraStart = rng.Start

            Dim applyPositions As Collection: Set applyPositions = t("Positions")
            Dim k As Long
            For k = applyPositions.count To 1 Step -1
                Dim pos As Long: pos = CLng(applyPositions(k))
                ActiveDocument.Range(paraStart + pos, paraStart + pos + 1).text = "-"
                fixedCount = fixedCount + 1
            Next k
        End If
    Next t

    Utils.EndOperation fixedCount, False
    DebugTrace.Log "DashNormalizer.NormalizeDashes", "Hoan tat, " & Format$(Timer - t0, "0.00") & "s, fixedCount=" & fixedCount
    NormalizeDashes = fixedCount
    Exit Function

StepSnapshot:
    DebugTrace.LogErr "DashNormalizer.NormalizeDashes", "[CaptureDocument]", Err.number, Err.description
    Err.Raise Err.number, "DashNormalizer.NormalizeDashes", "[CaptureDocument] " & Err.description
ErrHandler:
    DebugTrace.LogErr "DashNormalizer.NormalizeDashes", "loi giua chung, " & Format$(Timer - t0, "0.00") & "s", Err.number, Err.description
    If opStarted Then
        Utils.AbortOperation Err.description
    Else
        Err.Raise Err.number, "DashNormalizer.NormalizeDashes", Err.description
    End If
    NormalizeDashes = 0
End Function
