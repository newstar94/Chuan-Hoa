Attribute VB_Name = "DebugAnnotator"
'==============================================================
' DebugAnnotator â€” TINH NANG GO LOI, chen Word Comment tren MOI doan da duoc ComponentDetector
' "khi nhan dien ra doan van nao tuong ung voi thanh phan nao cua the thuc, hay chen Comment de
' toi kiem tra"). SE GIU LAI trong code khi phat hanh, chi TAT (ENABLED = False) â€” bat/tat MOT
' dong duy nhat bo duoi day, khong xoa code.
' Goi tu FindingReporter.BuildCheckContext (moi lan bam "Kiem tra") VA DataReader.RunCore (nut
' "Doc du lieu") â€” CA HAI la diem vao co chu dinh, KHONG goi tu ComponentDetector.DetectComponents
' (ham do bi goi RAT NHIEU LAN trong mot phien â€” moi luong "Sua"/nut "Co chu 13"/"14" deu tu nhan
' dien lai, xem DebugTrace.log thuc te â€” chen comment o do se chong chat lap lai vo han). MOI lan
' goi XOA HET comment danh dau CU (theo tien to COMMENT_MARKER) roi chen lai TU DAU â€” idempotent
' qua nhieu lan bam "Kiem tra" lien tiep, khong con lai rac cu.
' TAT ENABLED cho ban dung tiep theo ("o ban dung tiep theo: tam thoi tat chen comment debug") â€”
' chua xac dinh duoc day co phai nguyen nhan gay Not Responding hay khong, nhung la mot trong
' nhung nghi ngo hop ly (chen/xoa Word Comment qua COM la thao tac tuong doi nang, da ghi nhan)
' nen tat truoc de thu hep pham vi chan doan. Code GIU NGUYEN, chi doi hang so nay.
' "Báº¡n hĂ£y báº­t cháº¿ Ä‘á»™ debug gáº¯n comment nháº­n diá»‡n Ä‘á»ƒ tĂ´i kiá»ƒm tra" â€” dang QA truc tiep dia danh-
' ngay/chuc vu dong hai/than van ban).
' "HĂ£y táº¯t DebugAnnotator") - da doi chieu xong nhung dieu can doi chieu (dia danh-ngay/chuc vu
' dong hai/than van ban), khong con can hien Word Comment go loi tren tai lieu that nua.
Public Const ENABLED As Boolean = False
Private Const COMMENT_MARKER As String = "[DEBUG-CHUAN-HOA]"

Private Function ContinuationRoleAfterBreak(ByVal role As String) As String
    Select Case role
        Case "recipientLabel": ContinuationRoleAfterBreak = "recipientList"
        Case "signerAuthority": ContinuationRoleAfterBreak = "signerAuthorityTitle"
        Case "nationalTitle": ContinuationRoleAfterBreak = "nationalMotto"
        Case Else: ContinuationRoleAfterBreak = ""
    End Select
End Function

Private Function FindContinuationBreakPos(ByVal text As String) As Long
    Dim searchFrom As Long: searchFrom = 1
    Do While searchFrom <= Len(text)
        Dim ch As String: ch = Mid$(text, searchFrom, 1)
        If ch <> Chr(11) And ch <> " " And ch <> vbTab Then Exit Do
        searchFrom = searchFrom + 1
    Loop
    FindContinuationBreakPos = InStr(searchFrom, text, Chr(11))
End Function

Public Sub AnnotateLayoutMap(ByVal paragraphs As Collection, ByVal layoutMap As Object)
    On Error GoTo ErrHandler
    RemoveDebugComments
    If Not ENABLED Then Exit Sub

    Dim indexMap As Object: Set indexMap = DocumentSnapshot.BuildSnapshotIndexMap()

    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If layoutMap.Exists(p.Index) Then
            If indexMap.Exists(p.Index) Then
                Dim rng As word.Range
                Set rng = ActiveDocument.paragraphs(CLng(indexMap(p.Index))).Range
                Dim role As String: role = CStr(layoutMap(p.Index))

                ' Doan gop nhan + noi dung qua xuong dong thu cong trong CUNG mot doan Word - chen
                ' HAI comment rieng thay vi mot, de hien dung "phan nao la gi" (xem dau
                ' ContinuationRoleAfterBreak).
                Dim contRole As String: contRole = ContinuationRoleAfterBreak(role)
                Dim breakPos As Long: breakPos = 0
                If Len(contRole) > 0 Then breakPos = FindContinuationBreakPos(rng.text)

                If breakPos > 0 Then
                    Dim labelRng As word.Range: Set labelRng = rng.Duplicate
                    labelRng.SetRange rng.Start, rng.Start + breakPos - 1
                    On Error Resume Next
                    ActiveDocument.Comments.Add labelRng, COMMENT_MARKER & " role=" & role & _
                        " (doan #" & CStr(p.Index) & ")"
                    On Error GoTo ErrHandler

                    Dim contRng As word.Range: Set contRng = rng.Duplicate
                    contRng.SetRange rng.Start + breakPos, rng.End
                    On Error Resume Next
                    ActiveDocument.Comments.Add contRng, COMMENT_MARKER & " role=" & contRole & _
                        " (phan sau xuong dong thu cong, CUNG doan #" & CStr(p.Index) & " voi " & role & ")"
                    On Error GoTo ErrHandler
                Else
                    On Error Resume Next
                    ActiveDocument.Comments.Add rng, COMMENT_MARKER & " role=" & role & _
                        " (doan #" & CStr(p.Index) & ")"
                    On Error GoTo ErrHandler
                End If
            End If
        End If
    Next p
    Exit Sub
ErrHandler:
    ' Tinh nang phu (khong phai loi cua ban than "Kiem tra") - im lang bo qua, khong hien MsgBox
    ' chan luong quet chinh.
End Sub

' Xoa toan bo comment danh dau CU do chinh module nay chen truoc do (nhan dien qua COMMENT_MARKER
' o dau noi dung comment) - KHONG dung comment nao khac cua nguoi dung.
Private Sub RemoveDebugComments()
    On Error Resume Next
    Dim i As Long
    For i = ActiveDocument.Comments.count To 1 Step -1
        If left$(ActiveDocument.Comments(i).Range.text, Len(COMMENT_MARKER)) = COMMENT_MARKER Then
            ActiveDocument.Comments(i).Delete
        End If
    Next i
    On Error GoTo 0
End Sub

' Xoa tay tu Immediate Window (Ctrl+G: DebugAnnotator.ClearAll) khi muon don comment go loi ma
' khong quet lai.
Public Sub ClearAll()
    RemoveDebugComments
End Sub
