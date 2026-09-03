Attribute VB_Name = "EllipsisNormalizer"
'==============================================================
' "Chua thay the dau ba cham bang
' dau cham lung "). Loai B (CLAUDE.md muc 2.2: chuyen doi co hoc, khong nhap nhang) -- MOI cum
' "..." (BA dau cham lien tiep TRO LEN, khong phai HAI hay MOT) duoc thay THANG bang MOT ky tu dau
' cham lung U+2026 ("") -- khong co truong hop hop le nao ba dau cham lien tiep KHONG phai la dau
' cham lung go tay, nen KHONG can danh sach loai tru (khac tu dien chinh ta/quy tac viet hoa, Loai
' C).
' Cung khuon voi MultiSpaceCollapser.bas: chup snapshot, tim vi tri TRONG TUNG DOAN, ap dung GIAM
' DAN (day la phep THU GON - ba+ ky tu thay bang MOT, lam ngan doan van) de vi tri con lai (luon o
' TRUOC) khong bi lech khi mot cho o SAU da thay doi do dai.
' KHONG loai tru bang/vai tro nao (khac EdgeWhitespaceTrimmer/MultiSpaceCollapser) -- day la
' chuyen doi thuan tuy ky tu, ap dung DONG NHAT moi noi trong tai lieu (kha bang, letterhead, Noi
' nhan...), khong co ly do nghiep vu nao de loai tru mot vi tri cu the.
' Diem vao goi TU DONG tu FindingReporter.RunCheckCore, cung vi tri voi MultiSpaceCollapser/
' EdgeWhitespaceTrimmer/FontVariantNormalizer (truoc BuildCheckContext, "van ban phai on dinh
' truoc buoc nhan dien+kiem tra").
'==============================================================
Option Explicit

' Tim moi vi tri (0-based, chi so ky tu DAU cua cum) co BA dau cham lien tiep TRO LEN trong text
' -- tra ve Collection cua Dictionary {"Start" (0-based), "Len" (do dai cum)}, TANG DAN theo vi
' tri xuat hien (doi chieu GIAM DAN o noi goi truoc khi ap dung).
Private Function FindDotRuns(ByVal text As String) As Collection
    Dim Result As New Collection
    Dim n As Long: n = Len(text)
    Dim i As Long: i = 1
    Do While i <= n
        If Mid$(text, i, 1) = "." Then
            Dim runStart As Long: runStart = i
            Dim runLen As Long: runLen = 0
            Do While i <= n And Mid$(text, i, 1) = "."
                runLen = runLen + 1
                i = i + 1
            Loop
            If runLen >= 3 Then
                Dim item As Object: Set item = Utils.NewDictionary()
                item("Start") = runStart - 1 ' quy ve 0-based
                item("Len") = runLen
                Result.Add item
            End If
        Else
            i = i + 1
        End If
    Loop
    Set FindDotRuns = Result
End Function

' Diem vao duy nhat -- FindingReporter.RunCheckCore goi TU DONG truoc moi lan "Kiem tra". Tra so
' cho da thay - dung cho nhat ky thao tac.
Public Function NormalizeEllipsis() As Long
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("EllipsisNormalizer.NormalizeEllipsis")
    Dim t0 As Double: t0 = Timer
    DebugTrace.Log "EllipsisNormalizer.NormalizeEllipsis", "Bat dau"

    ' CHI can noi dung doan van - dung ban chup NHE, tranh cham anh loi khong can thiet (T-71,
    ' xem ghi chu dau DocumentSnapshot.CaptureParagraphsOnlySnapshot).
    On Error GoTo StepSnapshot
    Dim snapshot As Object: Set snapshot = DocumentSnapshot.CaptureParagraphsOnlySnapshot()
    DebugTrace.Log "EllipsisNormalizer.NormalizeEllipsis", "CaptureParagraphsOnlySnapshot xong, " & Format$(Timer - t0, "0.00") & "s"

    On Error GoTo ErrHandler
    Dim targets As New Collection
    Dim p As ParagraphSnapshot
    For Each p In snapshot("Paragraphs")
        Dim runs As Collection: Set runs = FindDotRuns(p.text)
        If runs.count > 0 Then
            Dim item As Object: Set item = Utils.NewDictionary()
            item("Index") = p.Index
            Set item("Runs") = runs
            targets.Add item
        End If
    Next p

    If targets.count = 0 Then
        DebugTrace.Log "EllipsisNormalizer.NormalizeEllipsis", "Khong co cho can thay - thoat som"
        NormalizeEllipsis = 0
        Exit Function
    End If

    Dim indexMap As Object: Set indexMap = DocumentSnapshot.BuildSnapshotIndexMap()
    DebugTrace.Log "EllipsisNormalizer.NormalizeEllipsis", "BuildSnapshotIndexMap xong, " & Format$(Timer - t0, "0.00") & "s"

    Dim opName As String
    opName = "Thay d" & ChrW(&H1EA5) & "u ba ch" & ChrW(&H1EA5) & "m b" & ChrW(&H1EB1) & "ng d" & _
        ChrW(&H1EA5) & "u ch" & ChrW(&H1EA5) & "m l" & ChrW(&H1EED) & "ng"
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

            ' Duyet GIAM DAN de vi tri con lai (o TRUOC) khong bi lech khi mot cum o SAU da rut
            ' ngan do dai.
            Dim applyRuns As Collection: Set applyRuns = t("Runs")
            Dim k As Long
            For k = applyRuns.count To 1 Step -1
                Dim one As Object: Set one = applyRuns(k)
                Dim runStart As Long: runStart = CLng(one("Start"))
                Dim runLen As Long: runLen = CLng(one("Len"))
                ActiveDocument.Range(paraStart + runStart, paraStart + runStart + runLen).text = ChrW(&H2026)
                fixedCount = fixedCount + 1
            Next k
        End If
    Next t

    Utils.EndOperation fixedCount, False
    DebugTrace.Log "EllipsisNormalizer.NormalizeEllipsis", "Hoan tat, " & Format$(Timer - t0, "0.00") & "s, fixedCount=" & fixedCount
    NormalizeEllipsis = fixedCount
    Exit Function

StepSnapshot:
    DebugTrace.LogErr "EllipsisNormalizer.NormalizeEllipsis", "[CaptureDocument]", Err.number, Err.description
    Err.Raise Err.number, "EllipsisNormalizer.NormalizeEllipsis", "[CaptureDocument] " & Err.description
ErrHandler:
    DebugTrace.LogErr "EllipsisNormalizer.NormalizeEllipsis", "loi giua chung, " & Format$(Timer - t0, "0.00") & "s", Err.number, Err.description
    If opStarted Then
        Utils.AbortOperation Err.description
    Else
        Err.Raise Err.number, "EllipsisNormalizer.NormalizeEllipsis", Err.description
    End If
    NormalizeEllipsis = 0
End Function
