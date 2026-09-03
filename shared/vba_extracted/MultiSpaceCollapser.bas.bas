Attribute VB_Name = "MultiSpaceCollapser"
'==============================================================
' "Nhieu dau cach lien tiep: Neu
' chi co 2 thi xoa bot 1, neu nhieu hon 2 thi giu nguyen". Doi tu Loai C (LOCAL-TYPO-SPACE, de
' nghi cho nguoi dung xac nhan) sang xu ly TU DONG hoan toan theo dung SO LUONG: DUNG 2 ky tu dau
' cach/tab lien tiep -> tu dong gop con 1 dau cach; NHIEU HON 2 -> khong dung toi, KHONG sinh
' Finding/comment nao (co the la dan trang thu cong co chu dich, khac han truong hop 2 ky tu gan
' chac la go nham do go hai lan phim cach). Xem ComplianceChecker.CheckExtraSpaceTypo (tra Nothing
' MAI MAI, ba nhanh cu deu da chuyen sang co che khac).
' Diem vao goi TU DONG tu FindingReporter.RunCheckAndReport, CUNG vi tri voi EdgeWhitespaceTrimmer
' (truoc BuildCheckContext) -- goi SAU EdgeWhitespaceTrimmer de khong xu ly trung vao khoang trang
' O MEP doan (da duoc xoa het truoc do, nen "[ \t]{2}" con lai trong pham vi module nay chac chan
' nam O GIUA doan).
' Pham vi loai tru GIONG HET ComplianceChecker.ScannableParagraphs/EdgeWhitespaceTrimmer: bo qua
' doan trong bang (co nut rieng "Xoa ky tu thua bang Excel") + cac vai tro ten rieng nhay cam.
' Ky thuat sua: MOI vi tri "dung 2 ky tu lien tiep" duoc XOA 1 ky tu (giu lai 1) -- day la PHEP
' XOA (giam do dai doan), nen PHAI duyet CAC VI TRI trong CUNG mot doan theo thu tu GIAM DAN
' (giong LineBreakNormalizer) de vi tri cac cho CHUA xu ly (luon dung TRUOC trong doan) khong bi
' lech khi xoa mot cho o SAU no. Giua CAC DOAN khac nhau thi khong quan trong thu tu (truy cap qua
' Paragraphs(N) theo thu tu doan, khong theo vi tri ky tu tuyet doi).
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

' Tim moi vi tri (0-based, chi so ky tu DAU cua cum) co DUNG 2 ky tu dau cach/tab lien tiep (khong
' phai 1, khong phai 3+) trong text -- tra ve Collection cac Long, TANG DAN theo vi tri xuat hien
' (se duoc doi chieu GIAM DAN o noi goi truoc khi ap dung, xem ghi chu dau file).
Private Function FindExactlyTwoRuns(ByVal text As String) As Collection
    Dim Result As New Collection
    Dim n As Long: n = Len(text)
    Dim i As Long: i = 1
    Do While i <= n
        If IsSpaceOrTab(Mid$(text, i, 1)) Then
            Dim runStart As Long: runStart = i
            Dim runLen As Long: runLen = 0
            Do While i <= n And IsSpaceOrTab(Mid$(text, i, 1))
                runLen = runLen + 1
                i = i + 1
            Loop
            If runLen = 2 Then Result.Add runStart - 1 ' quy ve 0-based
        Else
            i = i + 1
        End If
    Loop
    Set FindExactlyTwoRuns = Result
End Function

' Diem vao duy nhat -- FindingReporter.RunCheckAndReport goi TU DONG truoc moi lan "Kiem tra". Tra
' so cho da gop -- dung cho nhat ky thao tac.
Public Function CollapseDoubleSpaces() As Long
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("MultiSpaceCollapser.CollapseDoubleSpaces")
    Dim t0 As Double: t0 = Timer
    DebugTrace.Log "MultiSpaceCollapser.CollapseDoubleSpaces", "Bat dau"

    ' CHI can noi dung doan van + DetectDocumentType/DetectComponents (ca hai cung CHI doc
    ' snapshot("Paragraphs")) - dung ban chup NHE, tranh cham anh loi khong can thiet (T-71, xem
    ' ghi chu dau DocumentSnapshot.CaptureParagraphsOnlySnapshot).
    On Error GoTo StepSnapshot
    Dim snapshot As Object: Set snapshot = DocumentSnapshot.CaptureParagraphsOnlySnapshot()
    DebugTrace.Log "MultiSpaceCollapser.CollapseDoubleSpaces", "CaptureParagraphsOnlySnapshot xong, " & Format$(Timer - t0, "0.00") & "s"

    On Error GoTo StepDetect
    Dim typeResult As Object: Set typeResult = DocumentTypeDetector.DetectDocumentType(snapshot)
    Dim componentsResult As Object
    Set componentsResult = ComponentDetector.DetectComponents(snapshot, CStr(typeResult("Type")))
    Dim layoutMap As Object: Set layoutMap = componentsResult("LayoutMap")
    DebugTrace.Log "MultiSpaceCollapser.CollapseDoubleSpaces", "DetectComponents xong, " & Format$(Timer - t0, "0.00") & "s"

    On Error GoTo ErrHandler
    Dim skipRoles As Object: Set skipRoles = SkipRoleSet()

    ' targets: Collection cua Dictionary {"Index", "Positions" (Collection Long, 0-based, TANG
    ' DAN)}.
    Dim targets As New Collection
    Dim p As ParagraphSnapshot
    For Each p In snapshot("Paragraphs")
        If Not p.isInTable Then
            Dim skipThis As Boolean: skipThis = False
            If layoutMap.Exists(p.Index) Then
                If skipRoles.Exists(CStr(layoutMap(p.Index))) Then skipThis = True
            End If
            If Not skipThis Then
                Dim positions As Collection: Set positions = FindExactlyTwoRuns(p.text)
                If positions.count > 0 Then
                    Dim item As Object: Set item = Utils.NewDictionary()
                    item("Index") = p.Index
                    Set item("Positions") = positions
                    targets.Add item
                End If
            End If
        End If
    Next p

    If targets.count = 0 Then
        DebugTrace.Log "MultiSpaceCollapser.CollapseDoubleSpaces", "Khong co cho can gop - thoat som"
        CollapseDoubleSpaces = 0
        Exit Function
    End If

    Dim indexMap As Object: Set indexMap = DocumentSnapshot.BuildSnapshotIndexMap()
    DebugTrace.Log "MultiSpaceCollapser.CollapseDoubleSpaces", "BuildSnapshotIndexMap xong, " & Format$(Timer - t0, "0.00") & "s"

    Dim opName As String
    opName = "G" & ChrW(&HF2) & "p " & ChrW(&H111) & "up d" & ChrW(&H1EA5) & "u c" & ChrW(&HE1) & _
        "ch li" & ChrW(&HEA) & "n ti" & ChrW(&H1EBF) & "p"
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

            ' Duyet GIAM DAN (tu vi tri lon nhat/o SAU nhat truoc) de cac vi tri con lai (o TRUOC)
            ' khong bi lech khi mot cho o SAU bi xoa bot 1 ky tu.
            Dim applyPositions As Collection: Set applyPositions = t("Positions")
            Dim k As Long
            For k = applyPositions.count To 1 Step -1
                Dim pos As Long: pos = CLng(applyPositions(k)) ' 0-based, vi tri ky tu DAU cua cum 2
                ' Xoa KY TU THU HAI cua cum (giu lai ky tu dau, thuong la dau cach thuan) -- xoa
                ' [paraStart+pos+1, paraStart+pos+2).
                ActiveDocument.Range(paraStart + pos + 1, paraStart + pos + 2).text = ""
                fixedCount = fixedCount + 1
            Next k
        End If
    Next t

    Utils.EndOperation fixedCount, False
    DebugTrace.Log "MultiSpaceCollapser.CollapseDoubleSpaces", "Hoan tat, " & Format$(Timer - t0, "0.00") & "s, fixedCount=" & fixedCount
    CollapseDoubleSpaces = fixedCount
    Exit Function

StepSnapshot:
    DebugTrace.LogErr "MultiSpaceCollapser.CollapseDoubleSpaces", "[CaptureDocument]", Err.number, Err.description
    Err.Raise Err.number, "MultiSpaceCollapser.CollapseDoubleSpaces", "[CaptureDocument] " & Err.description
StepDetect:
    DebugTrace.LogErr "MultiSpaceCollapser.CollapseDoubleSpaces", "[DetectComponents]", Err.number, Err.description
    Err.Raise Err.number, "MultiSpaceCollapser.CollapseDoubleSpaces", "[DetectComponents] " & Err.description
ErrHandler:
    DebugTrace.LogErr "MultiSpaceCollapser.CollapseDoubleSpaces", "loi giua chung, " & Format$(Timer - t0, "0.00") & "s", Err.number, Err.description
    If opStarted Then
        Utils.AbortOperation Err.description
    Else
        Err.Raise Err.number, "MultiSpaceCollapser.CollapseDoubleSpaces", Err.description
    End If
    CollapseDoubleSpaces = 0
End Function
