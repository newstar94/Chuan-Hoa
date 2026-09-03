Attribute VB_Name = "EdgeWhitespaceTrimmer"
'==============================================================
' "Dau cach va ky tu tab thua o DAU doan: tu dong xoa" + "Cuoi doan co dau cach thua, nen xoa. Voi
' loi nay, hay tu dong xoa giup nguoi dung ma khong can thong bao, khong can danh dau" -- ca hai
' truong hop la Loai B (co hoc, khong nhap nhang: khong co truong hop nao dau cach/tab NGAY DAU
' hoac NGAY CUOI doan mang y nghia can giu), tu dong ap dung TOAN VAN BAN, KHONG con sinh
' Finding/Word Comment nua (xem ComplianceChecker.CheckExtraSpaceTypo, tra Nothing MAI MAI vi ba
' nhanh cu deu da chuyen sang co che khac -- xem ghi chu tai do). RIENG truong hop "nhieu dau
' cach/tab lien tiep O GIUA doan" GIU RIENG o module khac (MultiSpaceCollapser.bas) vi hanh vi
' khac han (dung sai theo SO LUONG, khong phai theo VI TRI).
' Diem vao goi TU DONG tu FindingReporter.RunCheckAndReport, NGAY SAU
' LineBreakNormalizer.NormalizeLineBreaks va TRUOC BuildCheckContext -- cung vi tri, cung ly do
' voi module do (doan phai "on dinh" truoc khi nhan dien vai tro/quet loi). Xoa dau cach mep doan
' KHONG lam doi so doan Word (khac viec tach doan cua LineBreakNormalizer) nen thu tu truoc/sau
' LineBreakNormalizer khong anh huong DUNG SAI.
' Pham vi loai tru GIONG HET ComplianceChecker.ScannableParagraphs (Private, khong goi truc tiep
' duoc tu module khac nen tinh lai o day): bo qua doan trong bang (co nut rieng "Xoa ky tu thua
' bang Excel" -- ExcelPasteCleaner.bas -- cho truong hop trong bang, PHAM VI RONG HON: ca NBSP,
' khong chi dau cach/tab) + cac vai tro "ten rieng" nhay cam (tu-dien-chinh-ta.json/skipContexts/
' componentRoles).
' Ky thuat sua truc tiep tren Word.Range: dung DO DAI van ban da chup (ParagraphSnapshot.Text) de
' tinh vi tri vung dau cach thua o HAI DAU, roi xoa qua Range tai VI TRI TUYET DOI trong tai lieu
' (paraStart + offset) -- xoa CUOI TRUOC, DAU SAU (xoa cuoi khong lam lech vi tri dau, nguoc lai
' thi co) -- cung ky thuat "tin tuong offset tu p.Text anh xa thang sang Range" ma
' LineBreakNormalizer.NormalizeOneParagraph da dung an toan.
' Moi thu tuc cong khai bat dau bang On Error GoTo ErrHandler.
'==============================================================
Option Explicit

Private Function SkipRoleSet() As Object
    Dim Result As Object: Set Result = Utils.NewDictionary()
    Dim r As Variant
    For Each r In RuleLoader.GetTypoDictionary()("skipContexts")("componentRoles")
        Result(CStr(r)) = True
    Next r
    Set SkipRoleSet = Result
End Function

Private Function IsSpaceOrTab(ByVal ch As String) As Boolean
    IsSpaceOrTab = (ch = " " Or ch = vbTab)
End Function

' Dem so ky tu dau cach/tab o CUOI text.
Private Function TrailingWhitespaceLength(ByVal text As String) As Long
    Dim n As Long: n = 0
    Dim i As Long: i = Len(text)
    Do While i > 0
        If IsSpaceOrTab(Mid$(text, i, 1)) Then
            n = n + 1
            i = i - 1
        Else
            Exit Do
        End If
    Loop
    TrailingWhitespaceLength = n
End Function

' Dem so ky tu dau cach/tab o DAU text, KHONG vuot qua phan da tinh la "cuoi" (tranh dem chong
' cheo khi CA doan chi toan khoang trang).
Private Function LeadingWhitespaceLength(ByVal text As String, ByVal capAt As Long) As Long
    Dim n As Long: n = 0
    Dim i As Long: i = 1
    Do While i <= capAt
        If IsSpaceOrTab(Mid$(text, i, 1)) Then
            n = n + 1
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    LeadingWhitespaceLength = n
End Function

' Diem vao duy nhat -- FindingReporter.RunCheckAndReport goi TU DONG truoc moi lan "Kiem tra". Tra
' so doan da xoa dau cach o (it nhat) mot mep -- dung cho nhat ky thao tac.
Public Function TrimEdgeWhitespace() As Long
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("EdgeWhitespaceTrimmer.TrimEdgeWhitespace")
    Dim t0 As Double: t0 = Timer
    DebugTrace.Log "EdgeWhitespaceTrimmer.TrimEdgeWhitespace", "Bat dau"

    ' CHI can noi dung doan van + DetectDocumentType/DetectComponents (ca hai cung CHI doc
    ' snapshot("Paragraphs")) - dung ban chup NHE, tranh cham anh loi khong can thiet (T-71, xem
    ' ghi chu dau DocumentSnapshot.CaptureParagraphsOnlySnapshot).
    On Error GoTo StepSnapshot
    Dim snapshot As Object: Set snapshot = DocumentSnapshot.CaptureParagraphsOnlySnapshot()
    DebugTrace.Log "EdgeWhitespaceTrimmer.TrimEdgeWhitespace", "CaptureParagraphsOnlySnapshot xong, " & Format$(Timer - t0, "0.00") & "s"

    On Error GoTo StepDetect
    Dim typeResult As Object: Set typeResult = DocumentTypeDetector.DetectDocumentType(snapshot)
    Dim componentsResult As Object
    Set componentsResult = ComponentDetector.DetectComponents(snapshot, CStr(typeResult("Type")))
    Dim layoutMap As Object: Set layoutMap = componentsResult("LayoutMap")
    DebugTrace.Log "EdgeWhitespaceTrimmer.TrimEdgeWhitespace", "DetectComponents xong, " & Format$(Timer - t0, "0.00") & "s"

    On Error GoTo ErrHandler
    Dim skipRoles As Object: Set skipRoles = SkipRoleSet()

    ' targets: Collection cua Dictionary {"Index", "LeadLen", "TrailLen", "TotalLen"}.
    Dim targets As New Collection
    Dim p As ParagraphSnapshot
    For Each p In snapshot("Paragraphs")
        If Not p.isInTable Then
            Dim skipThis As Boolean: skipThis = False
            If layoutMap.Exists(p.Index) Then
                If skipRoles.Exists(CStr(layoutMap(p.Index))) Then skipThis = True
            End If
            If Not skipThis Then
                Dim totalLen As Long: totalLen = Len(p.text)
                Dim trailLen As Long: trailLen = TrailingWhitespaceLength(p.text)
                Dim leadLen As Long: leadLen = LeadingWhitespaceLength(p.text, totalLen - trailLen)
                If trailLen > 0 Or leadLen > 0 Then
                    Dim item As Object: Set item = Utils.NewDictionary()
                    item("Index") = p.Index
                    item("LeadLen") = leadLen
                    item("TrailLen") = trailLen
                    item("TotalLen") = totalLen
                    targets.Add item
                End If
            End If
        End If
    Next p

    If targets.count = 0 Then
        DebugTrace.Log "EdgeWhitespaceTrimmer.TrimEdgeWhitespace", "Khong co doan can xoa - thoat som"
        TrimEdgeWhitespace = 0
        Exit Function
    End If

    Dim indexMap As Object: Set indexMap = DocumentSnapshot.BuildSnapshotIndexMap()
    DebugTrace.Log "EdgeWhitespaceTrimmer.TrimEdgeWhitespace", "BuildSnapshotIndexMap xong, " & Format$(Timer - t0, "0.00") & "s"

    Dim opName As String
    opName = "X" & ChrW(&HF3) & "a d" & ChrW(&H1EA5) & "u c" & ChrW(&HE1) & "ch/tab th" & _
        ChrW(&H1EEB) & "a hai mep " & ChrW(&H111) & "o" & ChrW(&H1EA1) & "n"
    Dim opStarted As Boolean
    Utils.BeginOperation opName
    opStarted = True

    ' Xoa mep khong lam doi so doan Word nen thu tu duyet cac DOAN khong quan trong -- moi doan
    ' truy cap qua Paragraphs(N) theo THU TU, khong theo vi tri ky tu tuyet doi. TRONG cung mot
    ' doan: xoa CUOI TRUOC, DAU SAU (xoa cuoi khong lam lech vi tri dau).
    Dim fixedCount As Long: fixedCount = 0
    Dim t As Variant
    For Each t In targets
        Dim idx As Long: idx = CLng(t("Index"))
        If indexMap.Exists(idx) Then
            Dim wordPara As word.paragraph: Set wordPara = ActiveDocument.paragraphs(CLng(indexMap(idx)))
            Dim rng As word.Range: Set rng = wordPara.Range
            rng.MoveEnd wdCharacter, -1 ' bo dau doan (1 ky tu Cr) o cuoi Range
            Dim paraStart As Long: paraStart = rng.Start
            Dim tLen As Long: tLen = CLng(t("TotalLen"))
            Dim trL As Long: trL = CLng(t("TrailLen"))
            Dim leL As Long: leL = CLng(t("LeadLen"))

            If trL > 0 Then
                ActiveDocument.Range(paraStart + tLen - trL, paraStart + tLen).text = ""
            End If
            If leL > 0 Then
                ActiveDocument.Range(paraStart, paraStart + leL).text = ""
            End If
            fixedCount = fixedCount + 1
        End If
    Next t

    Utils.EndOperation fixedCount, False
    DebugTrace.Log "EdgeWhitespaceTrimmer.TrimEdgeWhitespace", "Hoan tat, " & Format$(Timer - t0, "0.00") & "s, fixedCount=" & fixedCount
    TrimEdgeWhitespace = fixedCount
    Exit Function

StepSnapshot:
    DebugTrace.LogErr "EdgeWhitespaceTrimmer.TrimEdgeWhitespace", "[CaptureDocument]", Err.number, Err.description
    Err.Raise Err.number, "EdgeWhitespaceTrimmer.TrimEdgeWhitespace", "[CaptureDocument] " & Err.description
StepDetect:
    DebugTrace.LogErr "EdgeWhitespaceTrimmer.TrimEdgeWhitespace", "[DetectComponents]", Err.number, Err.description
    Err.Raise Err.number, "EdgeWhitespaceTrimmer.TrimEdgeWhitespace", "[DetectComponents] " & Err.description
ErrHandler:
    DebugTrace.LogErr "EdgeWhitespaceTrimmer.TrimEdgeWhitespace", "loi giua chung, " & Format$(Timer - t0, "0.00") & "s", Err.number, Err.description
    If opStarted Then
        Utils.AbortOperation Err.description
    Else
        Err.Raise Err.number, "EdgeWhitespaceTrimmer.TrimEdgeWhitespace", Err.description
    End If
    TrimEdgeWhitespace = 0
End Function
