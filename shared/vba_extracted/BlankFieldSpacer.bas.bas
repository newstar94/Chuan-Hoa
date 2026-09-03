Attribute VB_Name = "BlankFieldSpacer"
'==============================================================
' Viettel Dang.md" muc 2, ke thua chot cua "chap nhan so/ngay de trong"): khi so/ky hieu hoac
' ngay/thang ban hanh duoc de TRONG (mau "Sá»‘: /QÄ�-TTg", "ngĂ y thĂ¡ng nÄƒm " - dien sau), NÄ�
' 30/Viettel doi hoi mot khoang cach TOI THIEU giua chu va dau phan cach de con cho vien tay: 10
' dau cach giua "Sá»‘" va "/" (Dieu 9 khoan 4), 6 dau cach giua "ngĂ y"/"thĂ¡ng" va giua "thĂ¡ng"/"nÄƒm"
' (Dieu 10 khoan 2) - hai gia tri lay tu shared/rules/thong-so-the-thuc.json/components
' (codeNumberNotation.blankNotationGapSpaces,
' placeAndIssuedDate.blankDayGapSpaces/blankMonthGapSpaces), KHONG hard-code (CLAUDE.md muc 3.1).
' Loai B (CLAUDE.md muc 2.2): chuyen doi co hoc, khong nhap nhang - regex CHI khop khi khoang giua
' hai moc la THUAN TUY khoang trang (khong co chu so nao), nghia la CHINH XAC truong hop "de
' trong" - mot ngay/so THAT (co chu so) khong bao gio khop, khong can danh sach loai tru rieng.
' Cung khuon voi EllipsisNormalizer.bas/MultiSpaceCollapser.bas: chup snapshot, tim vi tri TRONG
' TUNG DOAN qua VBScript.RegExp (dong khuon voi DecimalSeparatorConverter.bas), ap dung GIAM DAN
' theo vi tri de cac diem con lai (o TRUOC trong CUNG doan) khong bi lech khi mot diem o SAU da
' doi do dai.
' Diem vao goi TU DONG tu FindingReporter.RunCheckCore, cung vi tri voi MultiSpaceCollapser/
' EdgeWhitespaceTrimmer/EllipsisNormalizer/DashNormalizer/FontVariantNormalizer (truoc
' BuildCheckContext, "van ban phai on dinh truoc buoc nhan dien+kiem tra").
'==============================================================
Option Explicit

Private Function NewRegex(ByVal pattern As String) As Object
    Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern
    regex.IgnoreCase = False
    regex.Global = True
    Set NewRegex = regex
End Function

' Tim moi diem can dem trong mot doan text - tra Collection cua Dictionary {"Start" (0-based TRONG
' doan, vi tri BAT DAU khoang trang can thay), "CurLen" (do dai khoang trang hien co), "TargetLen"
' (do dai mong muon)} - CHI dua vao danh sach khi CurLen < TargetLen (da du toi thieu thi bo qua,
' dung y nghia "toi thieu N dau cach", khong ep dung bang N).
Private Function FindBlankGaps(ByVal text As String, ByVal notationGapTarget As Long, _
        ByVal dayGapTarget As Long, ByVal monthGapTarget As Long) As Collection
    Dim Result As New Collection

    ' "Sá»‘" [khoang trang tuy y] ":" [khoang trang - CAN DEM] "/" - dau hieu dung CHUNG voi
    ' codeNumberNotation.regexByRegime (dau-hieu-nhan-dien.json), chi khac o day CAN capture rieng
    ' phan khoang trang de biet do dai hien co.
    Dim reNumber As Object: Set reNumber = NewRegex("S" & ChrW(&H1ED1) & "\s*:(\s*)\/")
    Dim m As Object
    For Each m In reNumber.Execute(text)
        Dim gapText As String: gapText = m.SubMatches(0)
        If Len(gapText) < notationGapTarget Then
            ' Vi tri CUA nhom con (group 1) trong toan bo match: FirstIndex + do dai phan TRUOC
            ' group (o day la "Sá»‘...:", do dai = Len(m.Value) - Len(gapText) - 1 vi group nam ngay
            ' truoc ky tu "/" cuoi cung cua match).
            Dim gapStart As Long: gapStart = m.firstIndex + (Len(m.value) - Len(gapText) - 1)
            Dim item As Object: Set item = Utils.NewDictionary()
            item("Start") = gapStart
            item("CurLen") = Len(gapText)
            item("TargetLen") = notationGapTarget
            Result.Add item
        End If
    Next m

    ' "," [khoang trang] "ngĂ y" [khoang trang - CAN DEM] "thĂ¡ng" [khoang trang - CAN DEM] "nÄƒm"
    ' - dong bo dau hieu placeAndIssuedDate (dau-hieu-nhan-dien.json), neo bang dau phay truoc
    ' "ngĂ y" de tranh khop nham cum "ngĂ y...thĂ¡ng...nÄƒm" xuat hien ngau nhien trong than bai (dau
    ' phay + "ngĂ y" lien tiep CHI xuat hien o dong dia danh-ngay ban hanh).
    Dim reDate As Object: Set reDate = NewRegex(",\s*ng" & ChrW(&HE0) & "y(\s*)th" & _
        ChrW(&HE1) & "ng(\s*)n" & ChrW(&H103) & "m\b")
    For Each m In reDate.Execute(text)
        Dim dayGap As String: dayGap = m.SubMatches(0)
        Dim monthGap As String: monthGap = m.SubMatches(1)
        ' "thĂ¡ng" co do dai CO DINH (5 ky tu) - dung de tinh vi tri group 2 tu group 1.
        Const THANG_LEN As Long = 5
        Dim wholeStart As Long: wholeStart = m.firstIndex
        Dim afterCommaLen As Long: afterCommaLen = Len(m.value) - Len(dayGap) - THANG_LEN - _
            Len(monthGap) - 3 ' 3 = do dai "nÄƒm"
        If Len(dayGap) < dayGapTarget Then
            Dim dayItem As Object: Set dayItem = Utils.NewDictionary()
            dayItem("Start") = wholeStart + afterCommaLen
            dayItem("CurLen") = Len(dayGap)
            dayItem("TargetLen") = dayGapTarget
            Result.Add dayItem
        End If
        If Len(monthGap) < monthGapTarget Then
            Dim monthItem As Object: Set monthItem = Utils.NewDictionary()
            monthItem("Start") = wholeStart + afterCommaLen + Len(dayGap) + THANG_LEN
            monthItem("CurLen") = Len(monthGap)
            monthItem("TargetLen") = monthGapTarget
            Result.Add monthItem
        End If
    Next m

    Set FindBlankGaps = Result
End Function

' Sap xep MOT Collection cac Dictionary {"Start"...} GIAM DAN theo "Start" - sap xep chon don gian
' (so luong diem tren MOI doan rat nho, khong can thuat toan nhanh hon).
Private Function SortDescendingByStart(ByVal items As Collection) As Collection
    Dim arr() As Object
    ReDim arr(1 To items.count)
    Dim i As Long: i = 1
    Dim it As Variant
    For Each it In items
        Set arr(i) = it
        i = i + 1
    Next it

    Dim j As Long
    For i = 2 To UBound(arr)
        Dim tmp As Object: Set tmp = arr(i)
        Dim tmpStart As Long: tmpStart = CLng(tmp("Start"))
        j = i - 1
        Do While j >= 1
            If CLng(arr(j)("Start")) < tmpStart Then
                Set arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        Set arr(j + 1) = tmp
    Next i

    Dim Result As New Collection
    For i = 1 To UBound(arr)
        Result.Add arr(i)
    Next i
    Set SortDescendingByStart = Result
End Function

' Diem vao duy nhat -- FindingReporter.RunCheckCore goi TU DONG truoc moi lan "Kiem tra". Tra so
' cho da them khoang trang - dung cho nhat ky thao tac.
Public Function PadBlankFields() As Long
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("BlankFieldSpacer.PadBlankFields")
    Dim t0 As Double: t0 = Timer
    DebugTrace.Log "BlankFieldSpacer.PadBlankFields", "Bat dau"

    ' (chan doan hieu nang "Kiem tra chinh ta chay rat lau"): CHI can noi dung doan van de do
    ' khoang trang - dung ban chup NHE (CaptureParagraphsOnly), khong qua Sections/Tables/ Images.
    ' Tranh ganh chi phi cua CaptureImages khi tai lieu co InlineShape loi (xem ghi chu dau
    ' DocumentSnapshot.CaptureParagraphsOnly) - MOT lan capture nhu vay tren tai lieu do co the
    ' ton them 9-16 giay, va ham nay chay TRUOC BuildCheckContext (noi da co MOT lan capture DAY
    ' DU khac) nen truoc day ganh chi phi do HAI LAN moi luot "Kiem tra".
    On Error GoTo StepSnapshot
    Dim paragraphs As Collection: Set paragraphs = DocumentSnapshot.CaptureParagraphsOnly()
    DebugTrace.Log "BlankFieldSpacer.PadBlankFields", "CaptureParagraphsOnly xong, " & Format$(Timer - t0, "0.00") & "s"

    On Error GoTo ErrHandler
    Dim components As Object: Set components = RuleLoader.GetFormatSpec()("components")
    Dim notationGapTarget As Long: notationGapTarget = CLng(components("codeNumberNotation")("blankNotationGapSpaces"))
    Dim dayGapTarget As Long: dayGapTarget = CLng(components("placeAndIssuedDate")("blankDayGapSpaces"))
    Dim monthGapTarget As Long: monthGapTarget = CLng(components("placeAndIssuedDate")("blankMonthGapSpaces"))

    Dim targets As New Collection
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        Dim gaps As Collection: Set gaps = FindBlankGaps(p.text, notationGapTarget, dayGapTarget, monthGapTarget)
        If gaps.count > 0 Then
            Dim item As Object: Set item = Utils.NewDictionary()
            item("Index") = p.Index
            Set item("Gaps") = SortDescendingByStart(gaps)
            targets.Add item
        End If
    Next p

    If targets.count = 0 Then
        DebugTrace.Log "BlankFieldSpacer.PadBlankFields", "Khong co cho can them - thoat som"
        PadBlankFields = 0
        Exit Function
    End If

    Dim indexMap As Object: Set indexMap = DocumentSnapshot.BuildSnapshotIndexMap()
    DebugTrace.Log "BlankFieldSpacer.PadBlankFields", "BuildSnapshotIndexMap xong, " & Format$(Timer - t0, "0.00") & "s"

    Dim opName As String
    opName = "Th" & ChrW(&HEA) & "m d" & ChrW(&H1EA5) & "u c" & ChrW(&HE1) & "ch cho " & _
        ChrW(&HF4) & " tr" & ChrW(&H1ED1) & "ng"
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

            Dim g As Variant
            For Each g In t("Gaps")
                Dim gapStart As Long: gapStart = CLng(g("Start"))
                Dim curLen As Long: curLen = CLng(g("CurLen"))
                Dim targetLen As Long: targetLen = CLng(g("TargetLen"))
                ActiveDocument.Range(paraStart + gapStart, paraStart + gapStart + curLen).text = Space$(targetLen)
                fixedCount = fixedCount + 1
            Next g
        End If
    Next t

    Utils.EndOperation fixedCount, False
    DebugTrace.Log "BlankFieldSpacer.PadBlankFields", "Hoan tat, " & Format$(Timer - t0, "0.00") & "s, fixedCount=" & fixedCount
    PadBlankFields = fixedCount
    Exit Function

StepSnapshot:
    DebugTrace.LogErr "BlankFieldSpacer.PadBlankFields", "[CaptureParagraphsOnly]", Err.number, Err.description
    Err.Raise Err.number, "BlankFieldSpacer.PadBlankFields", "[CaptureParagraphsOnly] " & Err.description
ErrHandler:
    DebugTrace.LogErr "BlankFieldSpacer.PadBlankFields", "loi giua chung, " & Format$(Timer - t0, "0.00") & "s", Err.number, Err.description
    If opStarted Then
        Utils.AbortOperation Err.description
    Else
        Err.Raise Err.number, "BlankFieldSpacer.PadBlankFields", Err.description
    End If
    PadBlankFields = 0
End Function
