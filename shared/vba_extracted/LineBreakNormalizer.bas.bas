Attribute VB_Name = "LineBreakNormalizer"
'==============================================================
' "Tu dong sua: Neu quoc hieu
' tieu ngu, co quan... danh sach noi nhan dung (SHIFT+ENTER) hoac cac line break nao chua chuan
' thi tu dong sua thanh dang line break chuan"). Loai B (chuyen doi co hoc, danh sach vai tro DAY
' DU o TARGET_ROLES duoi day, khong hoi lai tung cho â€” CLAUDE.md muc 2.2).
' "Toi chua tung yeu cau ban tao nut nay. NormalizeLineBreaks gio CHI con goi TU DUY NHAT MOT NOI:
' FindingReporter.RunCheckAndReport, NGAY TRUOC buoc nhan dien+kiem tra (am tham, khong MsgBox) â€”
' moi lan bam nut "Kiem tra" la mot lan tu dong chuan hoa xuong dong TRUOC KHI quet, dung nghia
' "hoan toan tu dong" nguoi dung yeu cau. Ham nay VAN TU BOC Utils.BeginOperation/ EndOperation
' rieng (mot dong nhat ky/mot buoc Undo rieng, TACH voi buoc "Kiem tra" ngay sau
' BOI CANH: ... da xac nhan nhieu lan â€” cac doan Quoc hieu/Tieu ngu/co quan/Noi nhan/nguoi ky
' trong van ban hanh chinh THAT thuong GOP nhieu "y" logic vao MOT doan Word DUY NHAT qua xuong
' dong thu cong (Chr(11), Shift+Enter) thay vi Enter that. Day la NGUON GOC cua hang loat loi
' (nhan nham recipientList/recipientLabel, gan sai co chu, chen nham comment) ma // phai vong qua
' bang FindContinuationBreakPos/ContinuationRoleAfterBreak o NHIEU noi (TextFormatter.bas,
' DebugAnnotator.bas, FindingAnnotator.bas). Module nay tan goc: TACH THAT cac doan nay thanh
' nhieu doan Word rieng biet ngay tu dau â€” sau khi chay xong, cac ky thuat "tim diem ngat" noi
' tren khong con gap doan gop nua (van GIU LAI code do, khong xoa â€” nguoi dung co the co van ban
' CHUA chay nut nay, hoac go them Shift+Enter sau).
' Pham vi: CHI cac doan van da duoc ComponentDetector gan MOT trong TARGET_ROLES duoi day â€” che
' ngu than bai (Dieu/Khoan/Diem, danh sach gach dau dong trong noi dung...) KHONG bi dung toi.
' Word tach doan dung vi tri, doan moi ke thua nguyen dinh dang cua doan goc, khong lam hong noi
' dung xung quanh). Day la thao tac soan thao BINH THUONG (tuong duong nguoi dung tu xoa
' Shift+Enter go lai Enter) â€” KHONG pha vo ngan xep Undo (khac OOXML/chuyen bang ma), nen KHONG
' can canh bao rui ro cao qua SafetyGuard (doi chieu ADR-007).
' Duyet TU CUOI TAI LIEU VE DAU (chi so LayoutMap giam dan) de vi tri cac doan CHUA xu ly khong bi
' lech khi tach doan lam tang tong so doan phia truoc no. Trong CUNG mot doan, xu ly Chr(11) cuoi
' cung TRUOC (cung ly do).
' Moi thu tuc cong khai bat dau bang On Error GoTo ErrHandler.
'==============================================================
Option Explicit

Private Function TargetRoles() As Variant
    TargetRoles = Array( _
        "nationalTitle", "nationalMotto", "organName", "superiorOrganName", _
        "recipientLabel", "recipientList", "signerAuthority", "signerAuthorityTitle", _
        "recipientSalutation", "recipientSalutationList", "recipientSalutationInline")
End Function

Private Function IsTargetRole(ByVal role As String) As Boolean
    Dim roles As Variant: roles = TargetRoles()
    Dim i As Long
    For i = LBound(roles) To UBound(roles)
        If roles(i) = role Then
            IsTargetRole = True
            Exit Function
        End If
    Next i
    IsTargetRole = False
End Function

' Diem vao duy nhat â€” FindingReporter.RunCheckAndReport goi TU DONG truoc moi lan "Kiem tra". Tra
' so dau ngat doan da tach â€” dung cho nhat ky thao tac (Utils.EndOperation).
Public Function NormalizeLineBreaks() As Long
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("LineBreakNormalizer.NormalizeLineBreaks")
    Dim t0 As Double: t0 = Timer
    DebugTrace.Log "LineBreakNormalizer.NormalizeLineBreaks", "Bat dau"

    On Error GoTo StepDetect
    Dim layoutMap As Object: Set layoutMap = ComponentFormatter.DetectLayoutMap()
    DebugTrace.Log "LineBreakNormalizer.NormalizeLineBreaks", "DetectLayoutMap xong, " & _
        Format$(Timer - t0, "0.00") & "s, " & layoutMap.count & " vai tro"

    On Error GoTo StepSort
    Dim indexes As Collection: Set indexes = SortedTargetIndexesDescending(layoutMap)
    DebugTrace.Log "LineBreakNormalizer.NormalizeLineBreaks", "SortedTargetIndexesDescending xong, " & _
        Format$(Timer - t0, "0.00") & "s, " & indexes.count & " doan muc tieu"

    On Error GoTo ErrHandler
    If indexes.count = 0 Then
        DebugTrace.Log "LineBreakNormalizer.NormalizeLineBreaks", "Khong co doan muc tieu - thoat som"
        NormalizeLineBreaks = 0
        Exit Function
    End If

    Dim indexMap As Object: Set indexMap = DocumentSnapshot.BuildSnapshotIndexMap()
    DebugTrace.Log "LineBreakNormalizer.NormalizeLineBreaks", "BuildSnapshotIndexMap xong, " & _
        Format$(Timer - t0, "0.00") & "s"

    Dim opName As String
    opName = "Chu" & ChrW(&H1EA9) & "n h" & ChrW(&HF3) & "a xu" & ChrW(&H1ED1) & "ng d" & ChrW(&HF2) & "ng"
    Dim opStarted As Boolean
    Utils.BeginOperation opName
    opStarted = True

    Dim fixedCount As Long: fixedCount = 0
    Dim idx As Variant
    Dim iterCount As Long: iterCount = 0
    For Each idx In indexes
        iterCount = iterCount + 1
        If indexMap.Exists(CLng(idx)) Then
            Dim p As word.paragraph
            Set p = ActiveDocument.paragraphs(CLng(indexMap(idx)))
            fixedCount = fixedCount + NormalizeOneParagraph(p.Range)
        End If
        DebugTrace.Log "LineBreakNormalizer.NormalizeLineBreaks", "  vong " & iterCount & "/" & _
            indexes.count & " (doan #" & CStr(idx) & ") xong, " & Format$(Timer - t0, "0.00") & "s"
    Next idx

    Utils.EndOperation fixedCount, False
    DebugTrace.Log "LineBreakNormalizer.NormalizeLineBreaks", "Hoan tat, " & _
        Format$(Timer - t0, "0.00") & "s, fixedCount=" & fixedCount
    NormalizeLineBreaks = fixedCount
    Exit Function
StepDetect:
    DebugTrace.LogErr "LineBreakNormalizer.NormalizeLineBreaks", "[DetectLayoutMap], " & _
        Format$(Timer - t0, "0.00") & "s", Err.number, Err.description
    Err.Raise Err.number, "LineBreakNormalizer.NormalizeLineBreaks", "[DetectLayoutMap] " & Err.description
StepSort:
    DebugTrace.LogErr "LineBreakNormalizer.NormalizeLineBreaks", "[SortedTargetIndexesDescending], " & _
        Format$(Timer - t0, "0.00") & "s", Err.number, Err.description
    Err.Raise Err.number, "LineBreakNormalizer.NormalizeLineBreaks", "[SortedTargetIndexesDescending] " & Err.description
ErrHandler:
    DebugTrace.LogErr "LineBreakNormalizer.NormalizeLineBreaks", "loi giua chung, " & _
        Format$(Timer - t0, "0.00") & "s", Err.number, Err.description
    If opStarted Then
        Utils.AbortOperation Err.description
    Else
        Err.Raise Err.number, "LineBreakNormalizer.NormalizeLineBreaks", Err.description
    End If
    NormalizeLineBreaks = 0
End Function

' Chi so doan (Long, GIAM DAN) mang MOT trong TARGET_ROLES â€” giam dan de tach doan (lam tang tong
' so doan phia SAU no trong tai lieu) khong lam lech vi tri cac doan CHUA xu ly con lai (luon nam
' TRUOC doan dang xu ly trong vong lap nay).
Private Function SortedTargetIndexesDescending(ByVal layoutMap As Object) As Collection
    Dim raw As New Collection
    Dim key As Variant
    For Each key In layoutMap.Keys
        If IsTargetRole(CStr(layoutMap(key))) Then raw.Add CLng(key)
    Next key

    Dim n As Long: n = raw.count
    Dim Result As New Collection
    If n = 0 Then
        Set SortedTargetIndexesDescending = Result
        Exit Function
    End If

    Dim arr() As Long
    ReDim arr(0 To n - 1)
    Dim i As Long: i = 0
    Dim v As Variant
    For Each v In raw
        arr(i) = CLng(v)
        i = i + 1
    Next v

    Dim j As Long, tmp As Long
    For i = 1 To n - 1
        tmp = arr(i)
        j = i - 1
        Do While j >= 0
            If arr(j) < tmp Then
                arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        arr(j + 1) = tmp
    Next i

    For i = 0 To n - 1
        Result.Add arr(i)
    Next i
    Set SortedTargetIndexesDescending = Result
End Function

Private Function NormalizeOneParagraph(ByVal paraRange As word.Range) As Long
    Dim text As String: text = paraRange.text
    Dim positions As New Collection
    Dim i As Long
    For i = 1 To Len(text)
        If Mid$(text, i, 1) = Chr(11) Then positions.Add i
    Next i
    If positions.count = 0 Then
        NormalizeOneParagraph = 0
        Exit Function
    End If

    Dim paraStart As Long: paraStart = paraRange.Start
    Dim k As Long
    For k = positions.count To 1 Step -1
        Dim pos As Long: pos = CLng(positions(k))
        Dim breakRng As word.Range
        Set breakRng = ActiveDocument.Range(paraStart + pos - 1, paraStart + pos)
        breakRng.text = vbCr
    Next k

    NormalizeOneParagraph = positions.count
End Function
