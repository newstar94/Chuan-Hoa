Attribute VB_Name = "DocumentTypeDetector"
'==============================================================
' DocumentTypeDetector â€” Xac dinh loai van ban (cong van / co ten loai / khong xac dinh)
' Nhan dien BANG LUAT (regex + vi tri tuong doi), khong doan - xem ADR-003.
' Doc DocumentSnapshot qua Dictionary tra ve tu DocumentSnapshot.CaptureDocument â€” KHONG tu chup
' lai, noi goi (ComplianceChecker, tro di) chiu trach nhiem truyen snapshot vao. Moi thu tuc cong
' khai bat dau bang On Error GoTo ErrHandler. Dung AscW/ChrW, khong dung Asc/Chr (phu thuoc code
' page) â€” CLAUDE.md muc 5.
'==============================================================
Option Explicit

' ============================================================================
' Chuoi hien thi tieng Viet dung trong regex â€” VBA khong cho goi ChrW trong bieu thuc Const nen
' phai dien qua bien cap module + EnsureTexts (cung cach EncodingConverter.bas lam).
' ============================================================================

Private mTextsReady As Boolean

' O 5b â€” Trich yeu cong van: "V/v" (chap nhan "V/v.", "V/v:"). Khong dau, khong can ChrW.
Private Const VV_PATTERN As String = "^\s*V\/v\b"

' O 4 â€” Dia danh va thoi gian ban hanh, moc phan ranh "sau o 4" cho tim ten loai (muc 2). Lay
' nguyen van tu docs/rules/01-tham-so-the-thuc.md muc 5.
Private PLACE_AND_DATE_PATTERN As String

' Moc phan ranh "truoc phan noi dung": cong van mo dau bang "Kinh gui:", van ban co ten loai va
' cau truc Dieu/Khoan mo dau noi dung bang "Dieu 1."
Private RECIPIENT_PATTERN As String
Private ARTICLE_PATTERN As String

Private KINH_TRINH_PATTERN As String

Private KINH_GUI_LOOSE_PATTERN As String

Private Sub EnsureTexts()
    If mTextsReady Then Exit Sub

    ' ",\s*ngay\s+\d{1,2}\s+thang\s+\d{1,2}\s+nam\s+\d{4}"
    Dim ngay As String, thang As String, nam As String
    ngay = "ng" & ChrW(&HE0) & "y"
    thang = "th" & ChrW(&HE1) & "ng"
    nam = "n" & ChrW(&H103) & "m"
    PLACE_AND_DATE_PATTERN = ",\s*" & ngay & "\s+\d{1,2}\s+" & thang & "\s+\d{1,2}\s+" & nam & "\s+\d{4}"

    ' "^\s*Kinh\s+gui\s*:"
    Dim kinh As String, gui As String
    kinh = "K" & ChrW(&HED) & "nh"
    gui = "g" & ChrW(&H1EED) & "i"
    RECIPIENT_PATTERN = "^\s*" & kinh & "\s+" & gui & "\s*:"

    ' "^\s*Kinh\s+trinh\s*:"
    Dim trinh As String
    trinh = "tr" & ChrW(&HEC) & "nh"
    KINH_TRINH_PATTERN = "^\s*" & kinh & "\s+" & trinh & "\s*:"

    ' "^\s*Kinh\s+gui\b" - KHONG doi hoi dau hai cham, xem ghi chu tai khai bao bien.
    KINH_GUI_LOOSE_PATTERN = "^\s*" & kinh & "\s+" & gui & "\b"

    ' "^\s*Dieu\s+\d+\s*\."
    Dim dieu As String
    dieu = ChrW(&H110) & "i" & ChrW(&H1EC1) & "u"
    ARTICLE_PATTERN = "^\s*" & dieu & "\s+\d+\s*\."

    mTextsReady = True
End Sub

' ============================================================================
' DetectDocumentType â€” diem vao duy nhat
' ============================================================================

' snapshot: Dictionary tra ve tu DocumentSnapshot.CaptureDocument (khoa "Paragraphs"). regime:
' "ND30" | "VIETTEL" | "DANG". Che do DANG di theo nhanh rieng - xem DetectDocumentType_DANG. Tra
' ve Dictionary voi ba khoa: "Type" ("congVan"|"coTenLoai"|"khongXacDinh", hoac "toTrinh"|"conLai"
' khi regime = "DANG"), "Confidence" ("high"|"medium"|"low"), "EvidenceParagraphIndex" (Long, Null
' neu khong co).
' 1. Co doan khop o 5b ("V/v") VA co doan "Kinh gui:" (BAT KY dau trong tai lieu, khong bat buoc
'   cung doan) -> "congVan", tin cay cao. TRUOC DAY chi doi hoi "V/v" - THEM dieu kien "Kinh gui:"
'   de tang do chac chan, tranh nham mot cum "V/v" tinh co xuat hien o van ban khac loai.
' 2. Khong co (1), co doan toan hoa/canh giua khop danh muc ten loai (loc theo regime), nam SAU o
'   4 va TRUOC phan noi dung -> "coTenLoai", tin cay trung binh.
' 3. Khong dau hieu nao -> "khongXacDinh", khong doan (ADR-003).
' Quet thanh hai luot rieng (khong gop mot vong lap)
Public Function DetectDocumentType(ByVal snapshot As Object, Optional ByVal regime As String = "ND30") As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("DocumentTypeDetector.DetectDocumentType")
    On Error GoTo ErrHandler
    EnsureTexts

    If regime = "DANG" Then
        Set DetectDocumentType = DetectDocumentType_DANG(snapshot)
        Exit Function
    End If

    Dim paragraphs As Collection
    Set paragraphs = snapshot("Paragraphs")

    Dim vvParagraph As ParagraphSnapshot
    Set vvParagraph = FindVvParagraph(paragraphs)
    Dim kinhGuiParagraph As ParagraphSnapshot
    Set kinhGuiParagraph = FindMatchingParagraph(paragraphs, KINH_GUI_LOOSE_PATTERN)
    If Not vvParagraph Is Nothing And Not kinhGuiParagraph Is Nothing Then
        Set DetectDocumentType = BuildResult("congVan", "high", vvParagraph.Index)
        Exit Function
    End If

    Dim typeNameParagraph As ParagraphSnapshot
    Set typeNameParagraph = FindTypeNameParagraph(paragraphs, regime)
    If Not typeNameParagraph Is Nothing Then
        Set DetectDocumentType = BuildResult("coTenLoai", "medium", typeNameParagraph.Index)
        Exit Function
    End If

    Set DetectDocumentType = BuildResult("khongXacDinh", "low", Null)
    Exit Function
ErrHandler:
    Err.Raise Err.number, "DocumentTypeDetector.DetectDocumentType", Err.description
End Function

Private Function DetectDocumentType_DANG(ByVal snapshot As Object) As Object
    Dim paragraphs As Collection
    Set paragraphs = snapshot("Paragraphs")

    Dim vvParagraph As ParagraphSnapshot
    Set vvParagraph = FindVvParagraph(paragraphs)
    Dim kinhGuiParagraph As ParagraphSnapshot
    Set kinhGuiParagraph = FindMatchingParagraph(paragraphs, KINH_GUI_LOOSE_PATTERN)
    If Not vvParagraph Is Nothing And Not kinhGuiParagraph Is Nothing Then
        Set DetectDocumentType_DANG = BuildResult("congVan", "high", vvParagraph.Index)
        Exit Function
    End If

    Dim kinhTrinhParagraph As ParagraphSnapshot
    Set kinhTrinhParagraph = FindMatchingParagraph(paragraphs, KINH_TRINH_PATTERN)
    If Not kinhTrinhParagraph Is Nothing Then
        Set DetectDocumentType_DANG = BuildResult("toTrinh", "high", kinhTrinhParagraph.Index)
        Exit Function
    End If

    Dim typeNameParagraph As ParagraphSnapshot
    Set typeNameParagraph = FindTypeNameParagraph(paragraphs, "DANG")
    If Not typeNameParagraph Is Nothing Then
        If Utils.ToUpperVn(StripTrailingPunctuation(Trim$(typeNameParagraph.text))) = _
                "T" & ChrW(&H1EDD) & " TR" & ChrW(&HCC) & "NH" Then
            Set DetectDocumentType_DANG = BuildResult("toTrinh", "medium", typeNameParagraph.Index)
            Exit Function
        End If
    End If

    Set DetectDocumentType_DANG = BuildResult("conLai", "low", Null)
End Function

' Doan DAU TIEN khop pattern - tien ich dung chung cho FindVvParagraph/"Kinh gui"/"Kinh trinh".
Private Function FindMatchingParagraph(ByVal paragraphs As Collection, ByVal pattern As String) As ParagraphSnapshot
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If RegexTest(pattern, p.text) Then
            Set FindMatchingParagraph = p
            Exit Function
        End If
    Next p
    Set FindMatchingParagraph = Nothing
End Function

Private Function BuildResult(ByVal docType As String, ByVal confidence As String, _
        ByVal evidenceParagraphIndex As Variant) As Object
    Dim Result As Object
    Set Result = Utils.NewDictionary()
    Result.Add "Type", docType
    Result.Add "Confidence", confidence
    Result.Add "EvidenceParagraphIndex", evidenceParagraphIndex
    Set BuildResult = Result
End Function

' ============================================================================
' Buoc 1 â€” "V/v"
' ============================================================================

Private Function FindVvParagraph(ByVal paragraphs As Collection) As ParagraphSnapshot
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If RegexTest(VV_PATTERN, p.text) Then
            Set FindVvParagraph = p
            Exit Function
        End If
    Next p
    Set FindVvParagraph = Nothing
End Function

' ============================================================================
' Buoc 2 â€” Ten loai
' ============================================================================

Public Function FindTypeNameParagraph(ByVal paragraphs As Collection, _
        Optional ByVal regime As String = "ND30") As ParagraphSnapshot
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("DocumentTypeDetector.FindTypeNameParagraph")
    ' Ham nay gio la diem vao cong khai TU BEN NGOAI module (ComponentDetector.bas), khong con
    ' chac chan da di qua DetectDocumentType (noi goi EnsureTexts truoc) - phai tu dam bao
    ' PLACE_AND_DATE_PATTERN/RECIPIENT_PATTERN/ARTICLE_PATTERN da duoc dien. EnsureTexts tu thoat
    ' ngay neu da nap roi (mTextsReady), goi lai khong ton kem.
    EnsureTexts

    Dim startIndex As Long, endIndex As Long
    ResolveSearchWindow paragraphs, startIndex, endIndex

    Dim p As ParagraphSnapshot
    Dim candidate As String
    For Each p In paragraphs
        If p.Index >= startIndex And p.Index < endIndex Then
            If p.AllCaps And p.alignment = "center" Then
                candidate = StripTrailingPunctuation(Trim$(p.text))
                If Not FindBestFuzzyTypeNameEntry(candidate, regime) Is Nothing Then
                    Set FindTypeNameParagraph = p
                    Exit Function
                End If
            End If
        End If
    Next p

    Set FindTypeNameParagraph = Nothing
End Function

' Chuan hoa MOT nhan/ten loai de doi chieu gan dung: bo dau (ke ca "Ä�"), bo khoang trang, viet
' hoa. Dung Utils.ToUnaccented (khong dung ComponentDetector.NormalizeForNationalTitle â€” ham do
' KHONG bo duoc "Ä�"/"Ä‘", chi bo duoc cac dau to hop tu unicode-to-nfc.json).
Private Function NormalizeTypeNameCandidate(ByVal s As String) As String
    NormalizeTypeNameCandidate = Utils.ToUpperVn(Replace$(Utils.ToUnaccented(s), " ", ""))
End Function

' Nguong Levenshtein toi da chap nhan duoc, tinh theo do dai chuoi da chuan hoa â€” cang ngan cang
' it dung sai, tranh nham giua hai ten loai NGAN va gan giong nhau (vi du "Dá»± Ă¡n"/"Ä�á»� Ă¡n" chi
' cach nhau 1 ky tu nhung la hai loai van ban khac han nhau).
Private Function TypeNameEditThreshold(ByVal normalizedLen As Long) As Long
    TypeNameEditThreshold = normalizedLen \ 5
End Function

' Doi chieu GAN DUNG (Levenshtein, sau khi bo dau/khoang trang) voi tung ten loai trong tu dien
' chu-viet-tat-ten-loai.json â€” CHI nhan khi ket qua tot nhat la DUY NHAT (khong hoa/gan bang mot
' ten loai KHAC, vi du "ThĂ´ng bĂ¡o"/"ThĂ´ng cĂ¡o" chi cach nhau 1 ky tu) VA nam trong nguong
' TypeNameEditThreshold cua chinh ten loai do. Tra Nothing neu khong chac chan â€” "khong chac thi
' khong sua" (CLAUDE.md muc 5). Dung chung cho ca FindTypeNameParagraph va MatchedTypeNameOrdinal
' de hai noi khong lech nhau.
Private Function FindBestFuzzyTypeNameEntry(ByVal candidateText As String, ByVal regime As String) As Object
    Dim candidateNorm As String
    candidateNorm = NormalizeTypeNameCandidate(candidateText)
    If Len(candidateNorm) = 0 Then
        Set FindBestFuzzyTypeNameEntry = Nothing
        Exit Function
    End If

    Dim documentTypes As Collection
    Set documentTypes = RuleLoader.GetDocTypeAbbreviations()("documentTypes")

    Dim bestEntry As Object: Set bestEntry = Nothing
    Dim bestEntryNorm As String
    Dim bestDist As Long: bestDist = -1
    Dim tieCount As Long: tieCount = 0

    Dim entry As Object
    Dim entryNorm As String
    Dim d As Long
    For Each entry In documentTypes
        If EntryAppliesToRegime(entry, regime) Then
            entryNorm = NormalizeTypeNameCandidate(StripTrailingQualifier(CStr(entry("typeName"))))
            d = Utils.LevenshteinDistance(candidateNorm, entryNorm)
            If bestDist = -1 Or d < bestDist Then
                bestDist = d
                Set bestEntry = entry
                bestEntryNorm = entryNorm
                tieCount = 1
            ElseIf d = bestDist Then
                tieCount = tieCount + 1
            End If
        End If
    Next entry

    If bestEntry Is Nothing Then
        Set FindBestFuzzyTypeNameEntry = Nothing
        Exit Function
    End If
    If tieCount > 1 Then
        Set FindBestFuzzyTypeNameEntry = Nothing
        Exit Function
    End If
    If bestDist > TypeNameEditThreshold(Len(bestEntryNorm)) Then
        Set FindBestFuzzyTypeNameEntry = Nothing
        Exit Function
    End If

    Set FindBestFuzzyTypeNameEntry = bestEntry
End Function

Private Function EntryAppliesToRegime(ByVal entry As Object, ByVal regime As String) As Boolean
    If Not entry.Exists("regimes") Then
        EntryAppliesToRegime = True ' du phong an toan - xem ghi chu tren
        Exit Function
    End If
    Dim r As Variant
    For Each r In entry("regimes")
        If CStr(r) = regime Then
            EntryAppliesToRegime = True
            Exit Function
        End If
    Next r
    EntryAppliesToRegime = False
End Function

' Bo hau to phan loai kieu "(ca biet)" - shared/rules/chu-viet-tat-ten-loai.json ghi "Nghi quyet
' (ca biet)"/"Quyet dinh (ca biet)" nhung chinh van ban chi ghi hoa "NGHI QUYET"/"QUYET DINH" -
' xem docs/rules/03-chu-viet-tat-va-mau.md muc 1.
Private Function StripTrailingQualifier(ByVal s As String) As String
    StripTrailingQualifier = RegexReplace("\s*\([^)]*\)\s*$", s, "")
End Function

' Dau cau ket doan co the di kem ten loai khi nguoi soan go them dau hai cham/dau cham cuoi dong.
Private Function StripTrailingPunctuation(ByVal s As String) As String
    StripTrailingPunctuation = RegexReplace("[.:]+$", s, "")
End Function

Private Sub ResolveSearchWindow(ByVal paragraphs As Collection, ByRef startIndex As Long, _
        ByRef endIndex As Long)
    startIndex = 0
    endIndex = paragraphs.count

    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If RegexTest(PLACE_AND_DATE_PATTERN, p.text) Then
            startIndex = p.Index + 1
            Exit For
        End If
    Next p

    For Each p In paragraphs
        If p.Index >= startIndex Then
            If RegexTest(RECIPIENT_PATTERN, p.text) Or RegexTest(ARTICLE_PATTERN, p.text) Then
                endIndex = p.Index
                Exit For
            End If
        End If
    Next p
End Sub

' Loi tren MOT doan (vi du chuoi qua dai/ky tu la lam VBScript.RegExp gay loi) khong duoc lam hong
' ca luot nhan dien loai van ban - ghi log va coi nhu KHONG khop, di tiep doan khac. Cung nguyen
' tac voi DocumentSnapshot.CaptureImages (T-71).
Private Function RegexTest(ByVal pattern As String, ByVal s As String) As Boolean
    On Error GoTo ErrHandler
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern
    regex.IgnoreCase = False
    regex.Global = False
    RegexTest = regex.test(s)
    Exit Function
ErrHandler:
    DebugTrace.LogErr "DocumentTypeDetector.RegexTest", "Loi khi kiem tra mau '" & pattern & _
        "' - coi nhu khong khop", Err.number, Err.description
    RegexTest = False
End Function

Private Function RegexReplace(ByVal pattern As String, ByVal s As String, ByVal replacement As String) As String
    On Error GoTo ErrHandler
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern
    regex.IgnoreCase = False
    regex.Global = True
    RegexReplace = regex.Replace(s, replacement)
    Exit Function
ErrHandler:
    DebugTrace.LogErr "DocumentTypeDetector.RegexReplace", "Loi khi thay mau '" & pattern & _
        "' - giu nguyen chuoi goc", Err.number, Err.description
    RegexReplace = s
End Function

' ============================================================================
' MatchedTypeNameOrdinal â€” cho drop-down "Loai van ban" (nhom Khoi dong, thang 8/2026)
' ============================================================================

Public Function MatchedTypeNameOrdinal(ByVal paragraphs As Collection, ByVal evidenceIndex As Variant, _
        Optional ByVal regime As String = "ND30") As Long
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("DocumentTypeDetector.MatchedTypeNameOrdinal")
    On Error GoTo ErrHandler
    MatchedTypeNameOrdinal = 0
    If IsNull(evidenceIndex) Then Exit Function

    Dim p As ParagraphSnapshot
    Dim target As ParagraphSnapshot
    Set target = Nothing
    For Each p In paragraphs
        If p.Index = CLng(evidenceIndex) Then
            Set target = p
            Exit For
        End If
    Next p
    If target Is Nothing Then Exit Function

    Dim candidate As String
    candidate = StripTrailingPunctuation(Trim$(target.text))

    Dim matched As Object
    Set matched = FindBestFuzzyTypeNameEntry(candidate, regime)
    If Not matched Is Nothing Then MatchedTypeNameOrdinal = CLng(matched("ordinal"))
    Exit Function
ErrHandler:
    Err.Raise Err.number, "DocumentTypeDetector.MatchedTypeNameOrdinal", Err.description
End Function

