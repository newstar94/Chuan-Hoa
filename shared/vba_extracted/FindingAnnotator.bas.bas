Attribute VB_Name = "FindingAnnotator"
'==============================================================
' FindingAnnotator â€” bao ket qua kiem tra bang Word Comment chen thang vao tai lieu
' (highlight turquoise/do, da xoa) lam MOT co che DUY NHAT bao CA the thuc lan chinh ta: Word
' Comment chen dung vi tri, noi dung la f.Title/f.Message (dang "hien tai..., can...", xem
' MakeFindingInput trong ComplianceChecker.bas) cong f.Before/After khi co (quy tac chinh ta/ tu
' dien).
' TRUOC DAY chi co MOT marker duy nhat cho moi Finding. Nay co HAI marker rieng biet â€”
' MarkerTagTheThuc/MarkerTagChinhTa. "moi nut chi xoa comment mang tag cua chinh no" â€” xem
' ClearMarker duoi day (so khop CHINH XAC tien to marker, KHAC RemoveFindingComments/ClearAll cu
' van la "[KI"+"TRA" bao trum moi bien the, giu lai lam cong cu don SACH TOAN BO qua Immediate
' Window khi can).
' Marker moi KHONG con khop tien to catch-all cu "[KI"+"TRA trong 25 ky tu dau" ("[THá»‚
' THá»¨C]"/"[CHĂ�NH Táº¢]" khong bat dau bang "[KI") - da CAP NHAT IsOwnMarkerComment de nhan dien THEM
' ca hai marker moi tuong minh, dam bao ClearAll van don duoc CA marker cu (tai lieu da "Kiem tra"
' tu cac ban truoc) LAN marker moi trong cung mot lan quet.
' Goi tu FindingReporter.bas (RunFormatCheck/RunSpellingCheck), ngay sau
' ComplianceChecker.RunAllChecks. Moi lan goi XOA COMMENT MANG DUNG MARKER CUA CHINH NO (qua
' ClearMarker) roi chen lai tu dau - idempotent qua nhieu lan bam "Kiem tra" lien tiep, va KHONG
' dong cham comment cua nut kia (dung D-4).
' Comment do add-in chen mang ten tac gia "Add-in Chuáº©n hĂ³a thá»ƒ thá»©c" (CommentAuthorName), gan
' THANG vao Word.Comment.Author cua tung comment ngay sau Comments.Add - CO Y KHONG dung cach doi
' tam Application.UserName (pho bien tren forum VBA nhung la cai dat CAP MAY TINH, ton tai qua ca
' luc dong/mo lai Word; neu add-in crash giua luc chua kip khoi phuc se ghi de vinh vien ten nguoi
' dung that cua chu may).
'==============================================================
Option Explicit

' VBA KHONG cho goi ChrW trong bieu thuc khoi tao Const nen hai marker la BIEN module thuong, gan
' MOT LAN qua EnsureTexts - cung ky thuat RibbonCallbacks.bas/CheckGate.bas/frmWarning.frm.
Private mTextsReady As Boolean
Private MARKER_THE_THUC As String
Private MARKER_CHINH_TA As String

Private Sub EnsureTexts()
    If mTextsReady Then Exit Sub
    ' &H1EC2 la chu "E mu moc" HOA (khac &H1EC3 chu thuong - bay da vap phai o, canh giac lai o
    ' day).
    MARKER_THE_THUC = "[TH" & ChrW(&H1EC2) & " TH" & ChrW(&H1EE8) & "C]"
    MARKER_CHINH_TA = "[CH" & ChrW(&HCD) & "NH T" & ChrW(&H1EA2) & "]"
    mTextsReady = True
End Sub

' Marker cua nut "Kiem tra the thuc" â€” FindingTierAggregator.bas dung LAI de gan cung marker cho
' comment toan cuc/theo trang cua no khi chay tang the thuc.
Public Function MarkerTagTheThuc() As String
    EnsureTexts
    MarkerTagTheThuc = MARKER_THE_THUC
End Function

' Marker cua nut "Kiem tra chinh ta".
Public Function MarkerTagChinhTa() As String
    EnsureTexts
    MarkerTagChinhTa = MARKER_CHINH_TA
End Function

' Dung trong FindingReporter.RunCheckAndReport (doi tam Application. UserName truoc khi chen).
Public Function CommentAuthorName() As String
    CommentAuthorName = "Add-in Chu" & ChrW(&H1EA9) & "n h" & ChrW(&HF3) & "a th" & _
        ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
End Function

' findings: Collection cua Finding (KHONG loc truoc theo nhom â€” FindingReporter da loai bo phat
' hien thuoc nhom bi cong chan.docx/Unicode truoc khi goi vao day, xem ghi chu RunCheckAndReport).
' Loi o TUNG Finding rieng le KHONG duoc chan ca vong lap (mot doan tim khong ra khong duoc lam
' hong nhung doan khac) - Resume Next co chu dinh trong vong lap, cung mau voi
' SpellingHighlighter.ApplyHighlights (da xoa) truoc day. marker: MarkerTagTheThuc hoac
' MarkerTagChinhTa - BAT BUOC truyen tuong minh de tranh chen nham marker cua nut kia.
Public Sub AnnotateFindings(ByVal findings As Collection, ByVal marker As String)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("FindingAnnotator.AnnotateFindings")
    On Error GoTo ErrHandler
    EnsureTexts
    ' KHONG con tu xoa o day - FindingReporter.bas (RunFormatCheck/ RunSpellingCheck) goi
    ' FindingAnnotator.ClearMarker MOT LAN DUY NHAT truoc CA HAI buoc chen (Tier roi Local) CUA
    ' CHINH nut do, xem chu thich dau file.

    Dim indexMap As Object: Set indexMap = DocumentSnapshot.BuildSnapshotIndexMap()
    Dim f As Finding
    For Each f In findings
        On Error Resume Next
        AnnotateOneFinding f, indexMap, marker
        On Error GoTo ErrHandler
    Next f
    Exit Sub
ErrHandler:
    ' Tinh nang phu (khong phai loi cua ban than "Kiem tra") - im lang bo qua, khong hien MsgBox
    ' chan luong quet chinh (cung nguyen tac voi DebugAnnotator.AnnotateLayoutMap).
End Sub

Private Sub AnnotateOneFinding(ByVal f As Finding, ByVal indexMap As Object, ByVal marker As String)
    If IsNull(f.paragraphIndex) Then Exit Sub

    Dim paraIdx As Long: paraIdx = CLng(f.paragraphIndex)
    If Not indexMap.Exists(paraIdx) Then Exit Sub

    Dim p As word.paragraph
    Set p = ActiveDocument.paragraphs(CLng(indexMap(paraIdx)))
    Dim rng As word.Range
    Set rng = p.Range

    If Len(f.Before) > 0 Then
        Dim paraText As String: paraText = rng.text
        Dim searchFrom As Long: searchFrom = 1
        If Not IsEmpty(f.charOffset) And Not IsNull(f.charOffset) Then
            searchFrom = CLng(f.charOffset) + 1
        End If
        Dim matchPos As Long
        matchPos = InStr(searchFrom, paraText, f.Before)
        If matchPos = 0 Then matchPos = InStr(1, paraText, f.Before) ' CharOffset lech thi tim lai tu dau
        If matchPos > 0 Then
            Dim hitRng As word.Range
            Set hitRng = rng.Duplicate
            hitRng.SetRange rng.Start + matchPos - 1, rng.Start + matchPos - 1 + Len(f.Before)
            Set rng = hitRng
        End If
    End If

    Dim cmt As word.Comment
    Set cmt = ActiveDocument.Comments.Add(rng, marker & " " & BuildCommentText(f))
    On Error Resume Next
    cmt.Author = CommentAuthorName()
    On Error GoTo 0
End Sub

Private Function BuildCommentText(ByVal f As Finding) As String
    Dim txt As String
    txt = f.title & ": " & f.message
    If Len(f.Before) > 0 Then
        txt = txt & vbCrLf & "S" & ChrW(&H1EED) & "a: """ & f.Before & """ " & ChrW(&H2192) & " """ & f.After & """"
    End If
    BuildCommentText = txt
End Function

' Xoa toan bo comment danh dau CU do chinh module nay chen truoc do - khong dung comment nao khac
' cua nguoi dung (ke ca comment "[DEBUG-CHUAN-HOA]" cua DebugAnnotator.bas, tien to khac han - xem
' IsOwnMarkerComment).
' Nhan dien theo tien to on dinh "[KI" thay vi so sanh CHINH XAC tung phien ban marker: marker co
' dau tung sua chinh ta hai lan (bien the sai dau khong khop nua khi doi), so khop chinh xac se
' lam comment cu bi bo sot va chen chong lai. Nhan het moi bien the (ke ca loi chinh ta chua biet
' toi) tranh phai liet ke tay. Cung quet luon comment toan cuc/theo trang cua
' FindingTierAggregator.bas (IsLegacyTierComment) de don trong mot lan.
Private Sub RemoveFindingComments()
    On Error Resume Next
    Dim i As Long
    For i = ActiveDocument.Comments.count To 1 Step -1
        Dim txt As String: txt = ActiveDocument.Comments(i).Range.text
        If IsOwnMarkerComment(txt) Or FindingTierAggregator.IsLegacyTierComment(txt) Then
            ActiveDocument.Comments(i).Delete
        End If
    Next i
    On Error GoTo 0
End Sub

Private Function IsOwnMarkerComment(ByVal text As String) As Boolean
    EnsureTexts
    If left$(text, 3) = "[KI" And InStr(1, left$(text, 25), "TRA") > 0 Then
        IsOwnMarkerComment = True
        Exit Function
    End If
    If left$(text, Len(MARKER_THE_THUC)) = MARKER_THE_THUC Or _
       left$(text, Len(MARKER_CHINH_TA)) = MARKER_CHINH_TA Then
        IsOwnMarkerComment = True
        Exit Function
    End If
    IsOwnMarkerComment = False
End Function

Public Sub ClearMarker(ByVal marker As String)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("FindingAnnotator.ClearMarker")
    On Error Resume Next
    Dim i As Long
    For i = ActiveDocument.Comments.count To 1 Step -1
        Dim txt As String: txt = ActiveDocument.Comments(i).Range.text
        If left$(txt, Len(marker)) = marker Then
            ActiveDocument.Comments(i).Delete
        End If
    Next i
    On Error GoTo 0
End Sub

' Xoa tay tu Immediate Window (Ctrl+G: FindingAnnotator.ClearAll) khi muon don TOAN BO comment ket
' qua kiem tra (CA HAI marker + moi bien the cu) ma khong quet lai - KHAC ClearMarker (chi xoa
' dung mot marker, dung tu hai nut "Kiem tra" tren ribbon theo D-4).
Public Sub ClearAll()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("FindingAnnotator.ClearAll")
    EnsureTexts
    RemoveFindingComments
End Sub
