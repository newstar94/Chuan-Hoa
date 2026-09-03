Attribute VB_Name = "ExcelPasteCleaner"
'==============================================================
' nhom "Bang bieu va hinh anh").
' - "O trong khong han la trong, ma la 1 nhom ky tu trang" -- xac nhan: cac o "trong" trong file
'   mau chua CHUOI KY TU U+00A0 (NON-BREAKING SPACE, ky hieu &nbsp;) -- KHONG PHAI dau cach thuong
'   (U+0020). Day la hanh vi PASTE cua Excel/Word: o Excel rong duoc dan vao Word thanh mot hoac
'   nhieu NBSP thay vi chuoi rong that.
' - "O chua so, phia truoc xuat hien rat nhieu ky tu trang" -- xac nhan: cac o so trong file mau
'   co toi 20-36 ky tu DAU CACH THUONG (U+0020, khong phai NBSP) truoc chu so -- Excel dung dau
'   cach de gia lap canh phai (right-align) khi dan vao Word dang van ban thuan, thay vi dung
'   thuoc tinh canh le that cua o.
' xoa CA HAI dau (dau/cuoi) cua MOI doan trong MOI o bang (moi story, ke ca bang long nhau) -- 3
' loai ky tu: dau cach (U+0020), tab (U+0009), NBSP (U+00A0). O "trong toan NBSP" tro thanh trong
' that (ca hai dau trim het). O so co dem dau cach truoc/sau deu duoc don sach, giu nguyen phan
' noi dung o giua.
' KHAC voi EdgeWhitespaceTrimmer.bas (tu dong, moi lan "Kiem tra", CHI ngoai bang): module nay LA
' MOT NUT RIENG tren ribbon, bam moi chay -- vi day la thao tac danh RIENG cho tinh huong dan bang
' tu Excel, khong phai loi go may thong thuong can quet MOI LAN kiem tra van ban.
' Ky thuat: duyet TRUC TIEP ActiveDocument.Paragraphs (KHONG qua DocumentSnapshot -- module nay
' CHI can biet doan co nam trong bang hay khong, khong can nhan dien vai tro the thuc), loc
' Information(wdWithInTable) = True, tinh do dai vung can xoa hai dau tu chinh Range.Text cua TUNG
' doan (doc truc tiep, khong qua snapshot trung gian), xoa qua Range tai vi tri tuyet doi -- cung
' ky thuat voi EdgeWhitespaceTrimmer.bas.
' Moi thu tuc cong khai bat dau bang On Error GoTo ErrHandler.
'==============================================================
Option Explicit

Private Function IsExcelPasteWhitespace(ByVal ch As String) As Boolean
    IsExcelPasteWhitespace = (ch = " " Or ch = vbTab Or ch = ChrW(&HA0))
End Function

Private Function TrailingLen(ByVal text As String) As Long
    Dim n As Long: n = 0
    Dim i As Long: i = Len(text)
    Do While i > 0
        If IsExcelPasteWhitespace(Mid$(text, i, 1)) Then
            n = n + 1
            i = i - 1
        Else
            Exit Do
        End If
    Loop
    TrailingLen = n
End Function

Private Function LeadingLen(ByVal text As String, ByVal capAt As Long) As Long
    Dim n As Long: n = 0
    Dim i As Long: i = 1
    Do While i <= capAt
        If IsExcelPasteWhitespace(Mid$(text, i, 1)) Then
            n = n + 1
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    LeadingLen = n
End Function

' Diem vao duy nhat -- RibbonCallbacks.OnXoaKyTuThuaBangExcel goi vao day (nut "Xoa ky tu thua
' bang Excel"). Tra so doan (trong bang) da xoa duoc it nhat mot mep.
Public Function CleanExcelPasteWhitespace() As Long
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ExcelPasteCleaner.CleanExcelPasteWhitespace")
    On Error GoTo ErrHandler

    Dim opName As String
    opName = "X" & ChrW(&HF3) & "a k" & ChrW(&HFD) & " t" & ChrW(&H1EF1) & " th" & ChrW(&H1EEB) & _
        "a b" & ChrW(&H1EA3) & "ng Excel"
    Utils.BeginOperation opName

    ' Duyet doan-theo-doan qua Range hien hanh (khong qua bien Word.Paragraph riĂªng, tranh giu
    ' tham chieu qua nhieu lan sua doi tai lieu giua chung).
    Dim n As Long: n = ActiveDocument.paragraphs.count
    Dim i As Long
    Dim fixedCount As Long: fixedCount = 0
    For i = 1 To n
        Dim paraRng As word.Range: Set paraRng = ActiveDocument.paragraphs(i).Range
        If paraRng.Information(wdWithInTable) Then
            ' Bo qua dau ket dong cua bang (end-of-row mark) -- khong mang noi dung that, cung quy
            ' uoc voi DocumentSnapshot.CaptureParagraphs.
            Dim rowMarkProbe As word.Range: Set rowMarkProbe = paraRng.Duplicate
            rowMarkProbe.Collapse wdCollapseStart
            If Not rowMarkProbe.Information(wdAtEndOfRowMarker) Then
                Dim rng As word.Range: Set rng = paraRng.Duplicate
                rng.MoveEnd wdCharacter, -1 ' bo dau ket doan/ket o o cuoi Range
                Dim text As String: text = rng.text
                If Len(text) > 0 Then
                    Dim trL As Long: trL = TrailingLen(text)
                    Dim leL As Long: leL = LeadingLen(text, Len(text) - trL)
                    If trL > 0 Or leL > 0 Then
                        Dim paraStart As Long: paraStart = rng.Start
                        Dim tLen As Long: tLen = Len(text)
                        If trL > 0 Then
                            ActiveDocument.Range(paraStart + tLen - trL, paraStart + tLen).text = ""
                        End If
                        If leL > 0 Then
                            ActiveDocument.Range(paraStart, paraStart + leL).text = ""
                        End If
                        fixedCount = fixedCount + 1
                    End If
                End If
            End If
        End If
    Next i

    Utils.EndOperation fixedCount, False
    CleanExcelPasteWhitespace = fixedCount
    Exit Function
ErrHandler:
    Utils.AbortOperation Err.description
    CleanExcelPasteWhitespace = -1
End Function
