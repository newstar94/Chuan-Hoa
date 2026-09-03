Attribute VB_Name = "FindingTierAggregator"
'==============================================================
' khi nghe cau tra loi cua Claude ve cach xu ly, "hay lam luon item 3"): (a) TOAN CUC â€” mau chu
' khong dong nhat, cach dat dau khong thong nhat, cach dung i/y khong thong nhat, co chu chiem ty
' le cao nhat khong phai 13 hoac 14, phong chu chiem ty le cao nhat khong phai Times New Roman,
' mau chu chiem ty le cao nhat khong phai Automatic -> CHEN 1 COMMENT DUY NHAT vao vi tri DAU TIEN
' cua van ban. (b) THEO TUNG TRANG GIAY â€” kho giay, le trang (huong giay warnOnly cung xep vao
' day) -> CHEN 1 COMMENT cho MOI SECTION vi pham, o vi tri DAU TIEN cua section do. Word Section
' (kieu ngat "Next Page", pho bien nhat) LUON bat dau mot trang MOI - Sections(i).Range. Start
' CHINH LA "vi tri dau tien cua trang" khong can tu duyet so trang. (c) CUC BO â€” sai font/co
' chu/mau chu o mot vi tri nho (thieu so, khong phai da so - xem
' ComplianceChecker.GroupWrongParagraphsByValue), sai chinh ta -> GIU NGUYEN co che cu, mot
' Finding mot comment, chen dung vi tri (FindingAnnotator.bas, khong doi).
' THIET KE: KHONG them truong moi vao Finding.cls - phan loai tang DUA VAO QUY UOC CO SAN:
' - Nhom (a)/(b) deu dat ParagraphIndex = Null trong ham kiem (dung quy uoc "khong co vi tri doan
'   van cu the" von co san cho quy tac cap trang/toan van ban).
' - Phan biet (a) voi (b) qua RULECODE: ba ma "ND30-PL1-M1-K1/-K2/-K3" (kho giay/huong giay/ le
'   trang) la (b), CON LAI (ma ParagraphIndex=Null) la (a).
' - Rieng ND30-PL1-M1-K4-FONT/-COLOR: Null CHI khi gia tri sai chiem DA SO (>50% doan da quet) -
'   ComplianceChecker.GroupWrongParagraphsByValue tu quyet dinh; thieu so thi VAN la
'   ParagraphIndex that (tang cuc bo, khong dung toi module nay).
' - Nhom (b) doc SO SECTION tu CHINH message ("Section N..." - tien to CHUNG ca ba ham kiem kho
'   giay/huong giay/le trang) thay vi them truong rieng - tranh dong cham Finding.cls.
' Goi TU DUY NHAT MOT NOI: FindingReporter.RunCheckAndReport, TRUOC khi goi
' FindingAnnotator.AnnotateFindings â€” tach nhom (a)/(b) ra chen rieng, TRA VE phan con lai (nhom
' (c)) de luong cu xu ly y nguyen.
' comment toan cuc/theo trang chen o day TRUOC DAY khong mang marker nao â€” SUA: gan CUNG marker
' voi FindingAnnotator (xem InsertGlobalComment/InsertPageComments duoi day) va viec XOA CU chuyen
' ve DUY NHAT o FindingReporter.bas (goi FindingAnnotator.ClearMarker/ClearAll TRUOC CA module nay
' chay) â€” module nay KHONG con tu xoa gi truoc khi chen. IsLegacyTierComment duoi day phuc vu don
' rac MOT LAN cho comment CU tu truoc ban sua nay (chua tung mang marker nao).
' ApplyTieredFindings nay nhan marker LAM THAM SO (khong con goi thang FindingAnnotator.MarkerTag
' - ham do da XOA, thay bang MarkerTagTheThuc/MarkerTagChinhTa) de gan DUNG marker cua nut dang
' chay cho comment toan cuc/theo trang chen o day.
'==============================================================
Option Explicit

Private mTextsReady As Boolean
Private TEXT_GLOBAL_TITLE As String
Private TEXT_PAGE_TITLE_PREFIX As String
Private TEXT_FONTSIZE_DOMINANT_PREFIX As String
Private TEXT_FONTSIZE_DOMINANT_MID As String

Private Sub EnsureTexts()
    If mTextsReady Then Exit Sub
    ' "Loi toan cuc (anh huong ca van ban):" â€” tieu de gop dau comment toan cuc. thieu chu "c"
    ' cuoi ("cá»¥" thay vi "cá»¥c") - da them lai.
    TEXT_GLOBAL_TITLE = "L" & ChrW(&H1ED7) & "i to" & ChrW(&HE0) & "n c" & ChrW(&H1EE5) & "c" & _
        " (" & ChrW(&H1EA3) & "nh h" & ChrW(&H1B0) & ChrW(&H1EDF) & "ng c" & ChrW(&H1EA3) & " v" & _
        ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n):"
    ' "Loi ap dung cho ca Section N (moi trang trong section):"
    TEXT_PAGE_TITLE_PREFIX = "L" & ChrW(&H1ED7) & "i " & ChrW(&HE1) & "p d" & ChrW(&H1EE5) & "ng cho c" & _
        ChrW(&H1EA3) & " Section "
    ' "Co chu chiem ty le cao nhat la {n}pt, khong phai 13 hoac 14 - khong xac dinh duoc dang dung
    ' Co chu 13 hay Co chu 14."
    TEXT_FONTSIZE_DOMINANT_PREFIX = "C" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & " chi" & ChrW(&H1EBF) & _
        "m t" & ChrW(&H1EF7) & " l" & ChrW(&H1EC7) & " cao nh" & ChrW(&H1EA5) & "t l" & ChrW(&HE0) & " "
    TEXT_FONTSIZE_DOMINANT_MID = "pt, kh" & ChrW(&HF4) & "ng ph" & ChrW(&H1EA3) & "i 13 ho" & ChrW(&H1EB7) & _
        "c 14 " & ChrW(&H2014) & " kh" & ChrW(&HF4) & "ng x" & ChrW(&HE1) & "c " & ChrW(&H111) & _
        ChrW(&H1ECB) & "nh " & ChrW(&H111) & ChrW(&H1B0) & ChrW(&H1EE3) & "c " & ChrW(&H111) & "ang " & _
        ChrW(&H111) & ChrW(&HF9) & "ng C" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & " 13 hay C" & ChrW(&H1EE1) & _
        " ch" & ChrW(&H1EEF) & " 14."
    mTextsReady = True
End Sub

Public Function IsLegacyTierComment(ByVal text As String) As Boolean
    EnsureTexts
    IsLegacyTierComment = (left$(text, Len(TEXT_GLOBAL_TITLE)) = TEXT_GLOBAL_TITLE) Or _
        (left$(text, Len(TEXT_PAGE_TITLE_PREFIX)) = TEXT_PAGE_TITLE_PREFIX)
End Function

Private Function PageTierRuleCodes() As Variant
    PageTierRuleCodes = Array("ND30-PL1-M1-K1", "ND30-PL1-M1-K2", "ND30-PL1-M1-K3")
End Function

Private Function IsInArray(ByVal value As String, ByVal arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = value Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

Public Function ApplyTieredFindings(ByVal findings As Collection, ByVal context As Object, _
        ByVal marker As String, ByVal includeDominantFontSizeNote As Boolean) As Collection
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("FindingTierAggregator.ApplyTieredFindings")
    On Error GoTo ErrHandler
    EnsureTexts

    Dim localFindings As New Collection
    Dim globalLines As New Collection
    Dim pageFindingsBySection As Object: Set pageFindingsBySection = Utils.NewDictionary()
    Dim pageCodes As Variant: pageCodes = PageTierRuleCodes()

    Dim f As Finding
    For Each f In findings
        If IsNull(f.paragraphIndex) Then
            If IsInArray(f.ruleCode, pageCodes) Then
                Dim sectionIdx As Long: sectionIdx = ParseSectionNumber(f.message)
                If sectionIdx >= 0 Then
                    Dim key As String: key = CStr(sectionIdx)
                    Dim lines As Collection
                    If pageFindingsBySection.Exists(key) Then
                        Set lines = pageFindingsBySection(key)
                    Else
                        Set lines = New Collection
                        Set pageFindingsBySection(key) = lines
                    End If
                    lines.Add f.title & ": " & f.message
                Else
                    ' Khong tach duoc so section (khong nen xay ra voi ba ham kiem hien co) - danh
                    ' dau AN TOAN: coi nhu toan cuc thay vi mat han thong tin.
                    globalLines.Add f.title & ": " & f.message
                End If
            Else
                globalLines.Add f.title & ": " & f.message
            End If
        Else
            localFindings.Add f
        End If
    Next f

    ' Them nhan dinh "co chu chiem ty le cao nhat" â€” tinh RIENG tai day (khong qua
    ' ComplianceChecker/quy-tac-kiem-tra.json, tranh them ma quy tac moi cho mot phep tinh don
    ' gian), dua tren CHINH cac doan da duoc LayoutMap gan vai tro (dang tin cay, cung nguyen tac
    ' voi CheckFontSizeConsistency). CHI tinh khi chay tang "the thuc" (xem chu thich tham so
    ' includeDominantFontSizeNote o dau ham).
    If includeDominantFontSizeNote Then
        Dim dominantSizeNote As String
        dominantSizeNote = DominantFontSizeNoteIfInvalid(context)
        If Len(dominantSizeNote) > 0 Then globalLines.Add dominantSizeNote
    End If

    InsertGlobalComment globalLines, marker
    InsertPageComments pageFindingsBySection, marker

    Set ApplyTieredFindings = localFindings
    Exit Function
ErrHandler:
    ' Tinh nang phu (khong phai loi cua ban than "Kiem tra") - tra nguyen findings goc, KHONG chan
    ' luong quet chinh (cung nguyen tac voi FindingAnnotator.AnnotateFindings).
    Set ApplyTieredFindings = findings
End Function

' Tach so section tu dau chuoi message (tien to CHUNG "Section N" cua CA BA ham kiem kho giay/
' huong giay/le trang trong ComplianceChecker.bas - K1/K3 co dau hai cham theo sau, K2 thi khong,
' nen chi bat "Section (\d+)" o dau chuoi, khong doi hoi ky tu ke tiep). Tra -1 neu khong tach
' duoc (an toan, xem noi goi).
Private Function ParseSectionNumber(ByVal message As String) As Long
    On Error GoTo ErrHandler
    Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = "^Section (\d+)"
    regex.IgnoreCase = False
    Dim m As Object: Set m = regex.Execute(message)
    If m.count = 0 Then
        ParseSectionNumber = -1
    Else
        ' "Section N" trong message la 1-based (section("Index") + 1 luc dung MakeFindingInput) ->
        ' tra ve 0-based de khop ActiveDocument.Sections(idx + 1).
        ParseSectionNumber = CLng(m(0).SubMatches(0)) - 1
    End If
    Exit Function
ErrHandler:
    ParseSectionNumber = -1
End Function

Private Sub InsertGlobalComment(ByVal lines As Collection, ByVal marker As String)
    If lines.count = 0 Then Exit Sub
    On Error Resume Next
    Dim txt As String: txt = marker & " " & TEXT_GLOBAL_TITLE
    Dim v As Variant
    For Each v In lines
        txt = txt & vbCrLf & ChrW(&H2022) & " " & CStr(v)
    Next v
    Dim rng As word.Range
    Set rng = ActiveDocument.Range(0, 0)
    Dim cmt As word.Comment: Set cmt = ActiveDocument.Comments.Add(rng, txt)
    cmt.Author = FindingAnnotator.CommentAuthorName()
    On Error GoTo 0
End Sub

Private Sub InsertPageComments(ByVal findingsBySection As Object, ByVal marker As String)
    If findingsBySection.count = 0 Then Exit Sub
    Dim key As Variant
    For Each key In findingsBySection.Keys
        On Error Resume Next
        Dim sectionIdx As Long: sectionIdx = CLng(key)
        Dim wordSectionNumber As Long: wordSectionNumber = sectionIdx + 1
        If wordSectionNumber >= 1 And wordSectionNumber <= ActiveDocument.sections.count Then
            Dim lines As Collection: Set lines = findingsBySection(key)
            Dim txt As String
            txt = marker & " " & TEXT_PAGE_TITLE_PREFIX & CStr(wordSectionNumber) & ":"
            Dim v As Variant
            For Each v In lines
                txt = txt & vbCrLf & ChrW(&H2022) & " " & CStr(v)
            Next v

            Dim sectionRange As word.Range
            Set sectionRange = ActiveDocument.sections(wordSectionNumber).Range
            Dim anchorRange As word.Range
            Set anchorRange = sectionRange.Duplicate
            anchorRange.SetRange sectionRange.Start, sectionRange.Start
            Dim cmt As word.Comment: Set cmt = ActiveDocument.Comments.Add(anchorRange, txt)
            cmt.Author = FindingAnnotator.CommentAuthorName()
        End If
        On Error GoTo 0
    Next key
End Sub

Private Function DominantFontSizeNoteIfInvalid(ByVal context As Object) As String
    On Error GoTo ErrHandler
    Dim snapshot As Object: Set snapshot = context("Snapshot")
    Dim layoutMap As Object: Set layoutMap = context("LayoutMap")
    If snapshot Is Nothing Or layoutMap Is Nothing Then
        DominantFontSizeNoteIfInvalid = ""
        Exit Function
    End If

    Dim countBySize As Object: Set countBySize = Utils.NewDictionary()
    Dim p As ParagraphSnapshot
    For Each p In snapshot("Paragraphs")
        If layoutMap.Exists(p.Index) Then
            Dim sz As Variant: sz = p.FontSizePt
            If Not IsNull(sz) Then
                If CDbl(sz) > 0 Then
                    Dim key As String: key = CStr(CDbl(sz))
                    If countBySize.Exists(key) Then
                        countBySize(key) = countBySize(key) + 1
                    Else
                        countBySize(key) = 1
                    End If
                End If
            End If
        End If
    Next p

    If countBySize.count = 0 Then
        DominantFontSizeNoteIfInvalid = ""
        Exit Function
    End If

    Dim bestKey As String: bestKey = ""
    Dim bestCount As Long: bestCount = -1
    Dim k As Variant
    For Each k In countBySize.Keys
        If CLng(countBySize(k)) > bestCount Then
            bestCount = CLng(countBySize(k))
            bestKey = CStr(k)
        End If
    Next k

    Dim bestSize As Double: bestSize = CDbl(bestKey)
    If bestSize = 13 Or bestSize = 14 Then
        DominantFontSizeNoteIfInvalid = ""
    Else
        DominantFontSizeNoteIfInvalid = TEXT_FONTSIZE_DOMINANT_PREFIX & CStr(bestSize) & TEXT_FONTSIZE_DOMINANT_MID
    End If
    Exit Function
ErrHandler:
    DominantFontSizeNoteIfInvalid = ""
End Function
