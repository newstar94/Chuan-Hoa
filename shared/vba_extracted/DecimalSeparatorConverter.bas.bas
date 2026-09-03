Attribute VB_Name = "DecimalSeparatorConverter"
'==============================================================
' ->,", nhom "Chinh ta"). Ap dung CA TRONG BANG (khac EdgeWhitespaceTrimmer/MultiSpaceCollapser --
' hai module do CHU Y loai tru bang).
' BAI TOAN NHAN DANG: mot chuoi "X.Y" (mot dau cham, hai ben la chu so) CO THE la (a) so thap phan
' that (vi du "12.5" = mot phay nam), HOAC (b) dau muc cua muc van ban / cot dau tien cua bang (vi
' du "1.1 Muc dich", "2.3" trong cot STT) -- HAI CACH DOC HOAN TOAN KHAC NHAU cho CUNG MOT chuoi
' ky tu, khong the phan biet chi bang regex don le.
' - Dau muc muc van ban: dung o DAU DOAN (ngoai bang), theo sau la khoang trang roi noi dung khac
'   (vi du "1.1 Muc dich yeu cau") -- BAO VE, khong doi.
' - Dau muc cot dau tien cua bang: chinh la TOAN BO noi dung (da trim) cua mot o thuoc CET 1 cua
'   bang (vi du o STT chi ghi "2.1") -- BAO VE, khong doi.
' - Con lai (nam GIUA cau, hoac trong o KHONG PHAI cot 1) -- coi la so thap phan THAT, chuyen doi.
' RIENG so co NHOM HANG NGHIN (it nhat mot dau phay ngan cach tung nhom 3 chu so, vi du
' "24,711,426,853") LUON duoc coi la so thuc su, KHONG BAO GIO la dau muc (dau muc khong bao gio
' dung dau phay) -- chuyen doi VO DIEU KIEN. RIENG ngay thang dang so co dau cham (vi du
' "05.3.2020", thay vi dang chuan "ngay.../.../..." dung dau gach cheo) -- BAO VE ca hai dau cham,
' tranh chuyen nham thanh so thap phan.
' thien ve BO SOT hon la CHUYEN NHAM mot dau muc thanh so thap phan.
' Ky thuat sua: MOI cho chuyen doi la MOT phep HOAN DOI 1 ky tu ('.' <-> ',') tai MOT vi tri --
' KHONG doi do dai doan, nen khong can quan tam thu tu ap dung (khac MultiSpaceCollapser, phep
' XOA, phai duyet giam dan). Duyet TRUC TIEP ActiveDocument.Paragraphs (khong qua
' DocumentSnapshot) vi module nay khong can vai tro the thuc, chi can biet doan co trong bang hay
' khong va o cot may.
' Moi thu tuc cong khai bat dau bang On Error GoTo ErrHandler.
'==============================================================
Option Explicit

Private Function RegexAllMatchesDS(ByVal pattern As String, ByVal text As String) As Object
    Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern
    regex.IgnoreCase = False
    regex.Global = True
    Set RegexAllMatchesDS = regex.Execute(text)
End Function

Private Function RegexFirstMatchDS(ByVal pattern As String, ByVal text As String) As Object
    Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern
    regex.IgnoreCase = False
    regex.Global = False
    Dim matches As Object: Set matches = regex.Execute(text)
    If matches.count = 0 Then
        Set RegexFirstMatchDS = Nothing
    Else
        Set RegexFirstMatchDS = matches(0)
    End If
End Function

Private Function RegexTestDS(ByVal pattern As String, ByVal text As String) As Boolean
    Dim regex As Object: Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern
    regex.IgnoreCase = False
    regex.Global = False
    RegexTestDS = regex.test(text)
End Function

Private Function MakeSpanDS(ByVal startPos As Long, ByVal endPos As Long) As Object
    Dim d As Object: Set d = Utils.NewDictionary()
    d("Start") = startPos
    d("End") = endPos
    Set MakeSpanDS = d
End Function

Private Function OverlapsAnySpanDS(ByVal spans As Collection, ByVal startPos As Long, ByVal endPos As Long) As Boolean
    Dim s As Variant
    For Each s In spans
        If startPos < CLng(s("End")) And endPos > CLng(s("Start")) Then
            OverlapsAnySpanDS = True
            Exit Function
        End If
    Next s
    OverlapsAnySpanDS = False
End Function

' "1.1", "2.3.1", tuy chon dau cham cuoi ("1.1.").
Private Function OutlinePattern() As String
    OutlinePattern = "\d{1,3}(\.\d{1,3})+\.?"
End Function

' "2.3.1", "1.2.3.4"... -- TU HAI DAU CHAM TRO LEN (3+ nhom chu so). Mot so thap phan THAT KHONG
' BAO GIO co qua MOT dau phan cach (giua phan nguyen va phan le) -- vi vay chuoi dang nay CHAC
' CHAN la dau muc/so hieu nhieu cap, KHONG PHAI so thap phan, BAT KE dung dau phay hay dau cham de
' phan cach ("2.3.1" hay "2,3,1" deu khong the la mot con so thap phan duy nhat). "2.3.1 khong the
' coi la mot so, du dung ',' hay '.' de phan tach thap phan") -- BAO VE VO DIEU KIEN, khong phu
' thuoc vi tri (khac OutlinePattern o tren, CHI bao ve khi o dau doan/toan bo o cot dau tien --
' rieng truong hop nay LUON bao ve, o BAT KY dau trong doan).
Private Function MultiDotPattern() As String
    MultiDotPattern = "\d{1,3}(\.\d{1,3}){2,}\.?"
End Function

' Ngay dang so dung dau cham thay vi gach cheo (vi du "05.3.2020") -- xem docs/rules/04-loi-go-
' may.md, hiem nhung co that trong van ban hanh chinh Viet Nam.
Private Function DatePattern() As String
    DatePattern = "\b\d{1,2}\.\d{1,2}\.\d{4}\b"
End Function

Private Function IsFirstTableColumn(ByVal rng As word.Range) As Boolean
    On Error Resume Next
    Dim ok As Boolean: ok = False
    Dim cel As word.Cell
    Set cel = rng.Cells(1)
    If Not cel Is Nothing Then ok = (cel.ColumnIndex = 1)
    On Error GoTo 0
    IsFirstTableColumn = ok
End Function

' Cac vung KHONG duoc dung toi (dau muc/ngay thang) -- xem ghi chu dau file.
Private Function DetectProtectedSpans(ByVal text As String, ByVal isInTable As Boolean, _
        ByVal isFirstColumn As Boolean) As Collection
    Dim spans As New Collection

    Dim dm As Object: Set dm = RegexAllMatchesDS(DatePattern(), text)
    Dim d As Object
    For Each d In dm
        spans.Add MakeSpanDS(CLng(d.firstIndex), CLng(d.firstIndex) + Len(CStr(d.value)))
    Next d

    ' Dau muc/so hieu nhieu cap (2+ dau cham) -- BAO VE VO DIEU KIEN, moi vi tri, khong phu thuoc
    ' trong/ngoai bang hay vi tri dau doan (xem ghi chu MultiDotPattern o tren).
    Dim mm As Object: Set mm = RegexAllMatchesDS(MultiDotPattern(), text)
    Dim md As Object
    For Each md In mm
        spans.Add MakeSpanDS(CLng(md.firstIndex), CLng(md.firstIndex) + Len(CStr(md.value)))
    Next md

    If Not isInTable Then
        Dim m As Object: Set m = RegexFirstMatchDS("^\s*" & OutlinePattern(), text)
        If Not (m Is Nothing) Then
            Dim endPos As Long: endPos = CLng(m.firstIndex) + Len(CStr(m.value))
            Dim nextChar As String: nextChar = ""
            If endPos < Len(text) Then nextChar = Mid$(text, endPos + 1, 1)
            If nextChar = "" Or nextChar = " " Or nextChar = vbTab Then
                spans.Add MakeSpanDS(CLng(m.firstIndex), endPos)
            End If
        End If
    End If

    If isInTable And isFirstColumn Then
        Dim trimmed As String: trimmed = Trim$(text)
        If Len(trimmed) > 0 Then
            If RegexTestDS("^" & OutlinePattern() & "$", trimmed) Then
                spans.Add MakeSpanDS(0, Len(text))
            End If
        End If
    End If

    Set DetectProtectedSpans = spans
End Function

' Tra ve Collection cac Long (0-based, vi tri ky tu ',' hoac '.' can hoan doi) cho MOT doan text.
Private Function ComputeEdits(ByVal text As String, ByVal isInTable As Boolean, _
        ByVal isFirstColumn As Boolean) As Collection
    Dim edits As New Collection
    Dim protectedSpans As Collection: Set protectedSpans = DetectProtectedSpans(text, isInTable, isFirstColumn)
    Dim handledSpans As New Collection

    ' Buoc 1 -- so co nhom hang nghin dung dau phay: vi du "24,711,426,853" hoac "1,234.56".
    Dim m1 As Object: Set m1 = RegexAllMatchesDS("\d{1,3}(,\d{3})+(\.\d+)?", text)
    Dim m As Object
    For Each m In m1
        Dim s1 As Long: s1 = CLng(m.firstIndex)
        Dim val1 As String: val1 = CStr(m.value)
        Dim e1 As Long: e1 = s1 + Len(val1)
        If Not OverlapsAnySpanDS(protectedSpans, s1, e1) Then
            Dim j As Long
            For j = 1 To Len(val1)
                Dim ch As String: ch = Mid$(val1, j, 1)
                If ch = "," Or ch = "." Then edits.Add s1 + j - 1
            Next j
            handledSpans.Add MakeSpanDS(s1, e1)
        End If
    Next m

    ' Buoc 2 -- so MOT dau cham don (khong nhom hang nghin), CHI khi KHONG bi bao ve/da xu ly o
    ' buoc 1: vi du "12.5" giua cau.
    Dim m2 As Object: Set m2 = RegexAllMatchesDS("\d+\.\d+", text)
    For Each m In m2
        Dim s2 As Long: s2 = CLng(m.firstIndex)
        Dim val2 As String: val2 = CStr(m.value)
        Dim e2 As Long: e2 = s2 + Len(val2)
        If Not OverlapsAnySpanDS(protectedSpans, s2, e2) And Not OverlapsAnySpanDS(handledSpans, s2, e2) Then
            Dim dotPos As Long: dotPos = InStr(val2, ".")
            If dotPos > 0 Then edits.Add s2 + dotPos - 1
        End If
    Next m

    Set ComputeEdits = edits
End Function

' Diem vao duy nhat -- RibbonCallbacks.OnDoiDauThapPhan goi vao day. Tra so doan da doi duoc it
' nhat mot vi tri.
Public Function ConvertDecimalSeparators() As Long
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("DecimalSeparatorConverter.ConvertDecimalSeparators")
    On Error GoTo ErrHandler

    Dim opName As String
    opName = ChrW(&H110) & ChrW(&H1ED5) & "i d" & ChrW(&H1EA5) & "u th" & ChrW(&H1EAD) & _
        "p ph" & ChrW(&HE2) & "n"
    Utils.BeginOperation opName

    Dim n As Long: n = ActiveDocument.paragraphs.count
    Dim i As Long
    Dim fixedCount As Long: fixedCount = 0
    For i = 1 To n
        Dim paraRng As word.Range: Set paraRng = ActiveDocument.paragraphs(i).Range
        Dim inTable As Boolean: inTable = paraRng.Information(wdWithInTable)
        Dim skipRow As Boolean: skipRow = False
        If inTable Then
            Dim rowMarkProbe As word.Range: Set rowMarkProbe = paraRng.Duplicate
            rowMarkProbe.Collapse wdCollapseStart
            If rowMarkProbe.Information(wdAtEndOfRowMarker) Then skipRow = True
        End If

        If Not skipRow Then
            Dim rng As word.Range: Set rng = paraRng.Duplicate
            rng.MoveEnd wdCharacter, -1
            Dim text As String: text = rng.text
            If Len(text) > 0 Then
                Dim isFirstCol As Boolean: isFirstCol = False
                If inTable Then isFirstCol = IsFirstTableColumn(rng)

                Dim edits As Collection: Set edits = ComputeEdits(text, inTable, isFirstCol)
                If edits.count > 0 Then
                    Dim paraStart As Long: paraStart = rng.Start
                    Dim e As Variant
                    For Each e In edits
                        Dim absPos As Long: absPos = paraStart + CLng(e)
                        Dim curRng As word.Range: Set curRng = ActiveDocument.Range(absPos, absPos + 1)
                        Dim curCh As String: curCh = curRng.text
                        If curCh = "," Then
                            curRng.text = "."
                        ElseIf curCh = "." Then
                            curRng.text = ","
                        End If
                    Next e
                    fixedCount = fixedCount + 1
                End If
            End If
        End If
    Next i

    Utils.EndOperation fixedCount, False
    ConvertDecimalSeparators = fixedCount
    Exit Function
ErrHandler:
    Utils.AbortOperation Err.description
    ConvertDecimalSeparators = -1
End Function
