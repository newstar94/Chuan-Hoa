Attribute VB_Name = "QrCodeGenerator"
'==============================================================
' QrCodeGenerator -- nut 7.1 "Chen QR". Sinh ma QR (chuan ISO/IEC 18004) tu noi dung
' nguoi dung tu nhap, roi chen anh vao vi tri con tro hien tai, dang inline text, mac dinh 5x5 cm.
' Ban Legacy KHONG co thu vien npm tuong duong "qrcode-generator" nen phai tu viet bo ma hoa QR
' bang VBA thuan -- day la RUI RO KY THUAT CHINH cua task nay. De giam rui ro toi da, TOAN BO
' logic duoi day duoc PORT TUNG HAM tu chinh ma nguon that cua thu vien "qrcode-generator"
' (Kazuhiko Arase, MIT, dist/qrcode.js) -- KHONG viet lai thuat toan tu dau theo tri nho/dac ta.
' Truoc khi port, toan bo logic da duoc dung lai trong mot ban "shadow" Node.js (cung cau truc thu
' tuc, khong dung closure, de de doi chieu tung ham voi VBA) va CHAY THU KHOP TUYET DOI (tung
' module, tung pixel) voi thu vien that qua 15 truong hop kiem thu -- do dai 1 den 800 ky tu, gom
' ca ASCII, tieng Viet co dau, emoji (surrogate pair UTF-16), chu Han/Nhat, ky tu dac biet, trai
' rong tu version 1 den version 23. Hai bang so lieu (PATTERN_POSITION_TABLE, RS_BLOCK_TABLE o
' EnsureTables ben duoi) duoc SINH BANG SCRIPT truc tiep tu dist/qrcode.js roi doi chieu vong lap
' nguoc (parse lai chuoi VBA, so khop JSON goc) -- KHONG go tay tung so, loai bo rui ro chep nham
' dang lo nhat cua mot bang 40 x 4 dong.
' MOT LOI DA PHAT HIEN va SUA khi doi chieu (ghi lai de nguoi doc sau khong lap lai): thu vien
' that dung MOT bien enum (QRErrorCorrectionLevel.M = 0) cho 15 bit "format info", nhung VI TRI
' HANG trong RS_BLOCK_TABLE (khoi 4 hang L/M/Q/H moi version) cua muc M la offset +1, HOAN TOAN
' DOC LAP voi gia tri enum 0 do -- nham lan hai con so nay se lay nham hang L (offset +0) thay vi
' hang M.
' VBA KHONG CO toan tu dich bit (<<, >>, >>>) nhu JavaScript -- xem BitShiftLeft/ BitShiftRight
' duoi day, cai dat bang nhan/chia luy thua 2. Da ra soat moi lan dich bit trong thuat toan nay:
' gia tri thuc te luon nho (duoi 20 bit) nen khong co rui ro tran so Long 32-bit co dau cua VBA
' (Long tran thi VBA nem loi runtime, khac JavaScript tu dong "wrap" theo ToInt32 -- neu co loi
' Overflow luc chay that, day la dieu can kiem tra dau tien).
' Muc sua loi QR CO DINH o muc M (khoi phuc ~15%), gioi han noi dung 800 ky tu -- PHAI khop tuyet
' doi. Dung AscW/ChrW, khong dung Asc/Chr (phu thuoc ma trang) -- CLAUDE.md muc 5.
'==============================================================
Option Explicit

' ============================================================================
' Hang so
' ============================================================================

Private Const QR_DEFAULT_SIZE_CM As Double = 5

Private Const QR_MAX_CONTENT_LENGTH As Long = 800

' Hai con so KHAC NHAU, de nham -- xem ghi chu dau file:
' - ECC_LEVEL_VALUE_M: gia tri enum dung trong 15 bit "format info" (SetupTypeInfo).
' - ECC_TABLE_ROW_OFFSET_M: vi tri HANG trong khoi 4 hang (L, M, Q, H) cua RS_BLOCK_TABLE cho moi
'   version -- DOC LAP voi gia tri enum o tren.
Private Const ECC_LEVEL_VALUE_M As Long = 0
Private Const ECC_TABLE_ROW_OFFSET_M As Long = 1

' Hang so BCH (dinh nghia chuan QR, khong phai bang du lieu can doi chieu -- gia tri co dinh, da
' tinh tay va kiem lai bang Node.js luc soan task): G15 =
' (1<<10)|(1<<8)|(1<<5)|(1<<4)|(1<<2)|(1<<1)|(1<<0) = 1335 G18 =
' (1<<12)|(1<<11)|(1<<10)|(1<<9)|(1<<8)|(1<<5)|(1<<2)|(1<<0) = 7973 G15_MASK=
' (1<<14)|(1<<12)|(1<<10)|(1<<4)|(1<<1) = 21522
Private Const G15 As Long = 1335
Private Const G18 As Long = 7973
Private Const G15_MASK As Long = 21522

Private Const PAD0 As Long = &HEC
Private Const PAD1 As Long = &H11

' ============================================================================
' Trang thai module-level cho MOT lan dung ma QR -- doi chieu bien closure "_modules" cua thu vien
' goc (nay cach duy nhat trong VBA khong dung Class module rieng cho tung the hien). Chi mot lan
' chen tai mot thoi diem (nguoi dung bam mot nut, cho popup dong) nen khong can tai vao lai (re-
' entrant).
' ============================================================================

Private mModules() As Long   ' -1 = chua dat, 0 = sang, 1 = toi -- doi chieu null/false/true cua JS
Private mN As Long           ' so module mot canh

Private mExpTable(0 To 255) As Long
Private mLogTable(0 To 255) As Long

Private mBufData() As Long
Private mBufByteCount As Long
Private mBufBitLen As Long

Private mPatternRows() As String
Private mRsRows() As String
Private mTablesReady As Boolean

' ============================================================================
' Diem vao ribbon -- nut 7.1 Chen QR. RibbonCallbacks.OnChenQrCode goi thang ham nay (mot dong,
' cung khuon voi ShapeFormatter.InsertUnderlines/LineCondenser.AutoCondenseLines).
' ============================================================================

Public Sub InsertQrCode()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("QrCodeGenerator.InsertQrCode")
    Dim tempPath As String
    tempPath = ""

    On Error GoTo ErrHandler

    frmQrCode.ShowQrCode

    Dim Result As String
    Result = frmQrCode.Result
    If Result <> "confirm" Then Exit Sub

    Dim Content As String
    Content = frmQrCode.Content

    Dim opName As String
    opName = "Ch" & ChrW(&HE8) & "n m" & ChrW(&HE3) & " QR"
    Utils.BeginOperation opName

    Dim moduleCount As Long
    Dim matrix As Variant
    matrix = GenerateQrModuleMatrix(Content, moduleCount)

    tempPath = NewTempBmpPath()
    WriteQrBmpFile tempPath, matrix, moduleCount

    Dim shp As InlineShape
    Set shp = Application.Selection.Range.InlineShapes.AddPicture( _
        fileName:=tempPath, LinkToFile:=False, SaveWithDocument:=True)
    shp.LockAspectRatio = msoTrue
    shp.width = Utils.CmToPoint(QR_DEFAULT_SIZE_CM)
    shp.Height = Utils.CmToPoint(QR_DEFAULT_SIZE_CM)

    DeleteTempFileIfExists tempPath
    tempPath = ""

    Utils.EndOperation 1, False
    Exit Sub
ErrHandler:
    DeleteTempFileIfExists tempPath
    Utils.AbortOperation Err.description
End Sub

Private Sub DeleteTempFileIfExists(ByVal filePath As String)
    If Len(filePath) = 0 Then Exit Sub
    On Error Resume Next
    If Len(Dir$(filePath)) > 0 Then Kill filePath
    On Error GoTo 0
End Sub

Private Function NewTempBmpPath() As String
    Static counter As Long
    counter = counter + 1
    NewTempBmpPath = Environ$("TEMP") & "\qr_the_thuc_" & Format$(Now, "yyyymmdd_hhnnss") & _
        "_" & CStr(counter) & ".bmp"
End Function

' ============================================================================
' ============================================================================

Public Function GenerateQrModuleMatrix(ByVal Content As String, ByRef outModuleCount As Long) As Variant
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("QrCodeGenerator.GenerateQrModuleMatrix")
    Dim trimmed As String
    trimmed = Trim$(Content)

    If Len(trimmed) = 0 Then
        Err.Raise vbObjectError + 1001, "QrCodeGenerator.GenerateQrModuleMatrix", _
            "Noi dung ma QR dang trong."
    End If
    If Len(trimmed) > QR_MAX_CONTENT_LENGTH Then
        Err.Raise vbObjectError + 1002, "QrCodeGenerator.GenerateQrModuleMatrix", _
            "Noi dung ma QR dai " & CStr(Len(trimmed)) & " ky tu, vuot gioi han " & _
            CStr(QR_MAX_CONTENT_LENGTH) & " ky tu."
    End If

    EnsureGfTables
    EnsureTables

    Dim dataBytes() As Long
    dataBytes = StringToUtf8Bytes(trimmed)

    Dim typeNumber As Long
    On Error GoTo CapacityErrHandler
    typeNumber = SelectVersion(dataBytes)
    On Error GoTo 0

    Dim maskPattern As Long
    maskPattern = GetBestMaskPattern(typeNumber, dataBytes)
    MakeImpl typeNumber, False, maskPattern, dataBytes

    outModuleCount = mN
    GenerateQrModuleMatrix = mModules
    Exit Function

CapacityErrHandler:
    If Err.number = vbObjectError + 1104 Then
        Err.Raise vbObjectError + 1003, "QrCodeGenerator.GenerateQrModuleMatrix", _
            "Noi dung ma QR chiem qua nhieu byte khi ma hoa (ky tu co dau/emoji chiem nhieu " & _
            "byte hon ky tu thuong) - vuot dung luong toi da cua ma QR. Hay rut ngan noi dung."
    Else
        Err.Raise Err.number, Err.source, Err.description
    End If
End Function

Private Function StringToUtf8Bytes(ByVal s As String) As Variant
    Dim n As Long
    n = Len(s)

    Dim bytes() As Long
    ReDim bytes(0 To n * 4)   ' can tren an toan (toi da 4 byte/ky tu UTF-16), cat gon o cuoi

    Dim count As Long
    count = 0

    Dim i As Long, charCode As Long, lowCode As Long
    i = 1
    Do While i <= n
        charCode = AscW(Mid$(s, i, 1))
        If charCode < 0 Then charCode = charCode + 65536

        If charCode < &H80 Then
            bytes(count) = charCode
            count = count + 1
        ElseIf charCode < &H800 Then
            bytes(count) = &HC0 Or BitShiftRight(charCode, 6)
            bytes(count + 1) = &H80 Or (charCode And &H3F)
            count = count + 2
        ElseIf charCode < &HD800 Or charCode >= &HE000 Then
            bytes(count) = &HE0 Or BitShiftRight(charCode, 12)
            bytes(count + 1) = &H80 Or (BitShiftRight(charCode, 6) And &H3F)
            bytes(count + 2) = &H80 Or (charCode And &H3F)
            count = count + 3
        Else
            ' Cap surrogate UTF-16 (ky tu ngoai Basic Multilingual Plane, vi du emoji) -- gop
            ' thanh mot codepoint 21-bit roi ma hoa UTF-8 bon byte.
            i = i + 1
            lowCode = AscW(Mid$(s, i, 1))
            If lowCode < 0 Then lowCode = lowCode + 65536
            charCode = &H10000 + (((charCode And &H3FF) * &H400) Or (lowCode And &H3FF))
            bytes(count) = &HF0 Or BitShiftRight(charCode, 18)
            bytes(count + 1) = &H80 Or (BitShiftRight(charCode, 12) And &H3F)
            bytes(count + 2) = &H80 Or (BitShiftRight(charCode, 6) And &H3F)
            bytes(count + 3) = &H80 Or (charCode And &H3F)
            count = count + 4
        End If
        i = i + 1
    Loop

    If count = 0 Then
        ' Khong le xay ra -- GenerateQrModuleMatrix da chan noi dung rong TRUOC khi goi toi day
        ' (Len(trimmed)=0 nem loi rieng). VBA khong ReDim duoc mang dong ve 0 phan tu (xem ghi chu
        ' ParseLongCsv) nen khong the "tra mang rong" o day nhu ban JS goc -- neu nhanh nay THAT
        ' SU chay toi, do la loi lap trinh can sua, khong phai truong hop hop le.
        Err.Raise vbObjectError + 1108, "QrCodeGenerator.StringToUtf8Bytes", _
            "chuoi rong khong the ma hoa UTF-8 - loi lap trinh, khong le xay ra"
    End If

    ReDim Preserve bytes(0 To count - 1)
    StringToUtf8Bytes = bytes
End Function

' ============================================================================
' GF(256) -- bang duoc SINH LUC CHAY bang dung mot cong thuc de quy (khong go tay 256 so -- giam
' rui ro chep nham xuong bang khong, chi con rui ro sai CONG THUC, de doi chieu hon nhieu). Port
' dung tu QRMath trong dist/qrcode.js.
' ============================================================================

Private Sub EnsureGfTables()
    Dim i As Long
    For i = 0 To 7
        mExpTable(i) = BitShiftLeft(1, i)
    Next i
    For i = 8 To 255
        mExpTable(i) = mExpTable(i - 4) Xor mExpTable(i - 5) Xor mExpTable(i - 6) Xor mExpTable(i - 8)
    Next i
    For i = 0 To 254
        mLogTable(mExpTable(i)) = i
    Next i
End Sub

Private Function Gexp(ByVal n As Long) As Long
    Do While n < 0
        n = n + 255
    Loop
    Do While n >= 256
        n = n - 255
    Loop
    Gexp = mExpTable(n)
End Function

Private Function Glog(ByVal n As Long) As Long
    If n < 1 Then
        Err.Raise vbObjectError + 1105, "QrCodeGenerator.Glog", "glog(" & CStr(n) & ")"
    End If
    Glog = mLogTable(n)
End Function

' ============================================================================
' Da thuc tren GF(256) -- port dung tu qrPolynomial trong dist/qrcode.js. Mang Long, chi so 0 la
' he so bac cao nhat (dung quy uoc cua thu vien goc).
' ============================================================================

Private Function PolyCreate(ByRef numArr() As Long, ByVal shift As Long) As Variant
    Dim n As Long
    n = UBound(numArr) - LBound(numArr) + 1

    Dim offset As Long
    offset = 0
    Do While offset < n
        If numArr(LBound(numArr) + offset) <> 0 Then Exit Do
        offset = offset + 1
    Loop

    Dim resultLen As Long
    resultLen = n - offset + shift

    Dim Result() As Long
    If resultLen <= 0 Then
        ReDim Result(0 To 0)
        Result(0) = 0
        PolyCreate = Result
        Exit Function
    End If

    ReDim Result(0 To resultLen - 1)
    Dim i As Long
    For i = 0 To n - offset - 1
        Result(i) = numArr(LBound(numArr) + i + offset)
    Next i
    PolyCreate = Result
End Function

Private Function PolyMultiply(ByRef a() As Long, ByRef b() As Long) As Variant
    Dim la As Long, lb As Long
    la = UBound(a) - LBound(a) + 1
    lb = UBound(b) - LBound(b) + 1

    Dim Result() As Long
    ReDim Result(0 To la + lb - 2)

    Dim i As Long, j As Long
    For i = 0 To la - 1
        For j = 0 To lb - 1
            Result(i + j) = Result(i + j) Xor Gexp(Glog(a(LBound(a) + i)) + Glog(b(LBound(b) + j)))
        Next j
    Next i

    PolyMultiply = PolyCreate(Result, 0)
End Function

Private Function PolyMod(ByRef a() As Long, ByRef b() As Long) As Variant
    Dim la As Long, lb As Long
    la = UBound(a) - LBound(a) + 1
    lb = UBound(b) - LBound(b) + 1

    If la - lb < 0 Then
        PolyMod = a
        Exit Function
    End If

    Dim ratio As Long
    ratio = Glog(a(LBound(a))) - Glog(b(LBound(b)))

    Dim num() As Long
    ReDim num(0 To la - 1)
    Dim i As Long
    For i = 0 To la - 1
        num(i) = a(LBound(a) + i)
    Next i
    For i = 0 To lb - 1
        num(i) = num(i) Xor Gexp(Glog(b(LBound(b) + i)) + ratio)
    Next i

    Dim reduced As Variant
    reduced = PolyCreate(num, 0)
    Dim reducedArr() As Long
    reducedArr = reduced

    PolyMod = PolyMod(reducedArr, b)
End Function

Private Function GetErrorCorrectPolynomial(ByVal ecLength As Long) As Variant
    Dim a() As Long
    ReDim a(0 To 0)
    a(0) = 1

    Dim factor(0 To 1) As Long
    Dim product As Variant
    Dim productArr() As Long
    Dim i As Long
    For i = 0 To ecLength - 1
        factor(0) = 1
        factor(1) = Gexp(i)
        product = PolyMultiply(a, factor)
        productArr = product
        a = productArr
    Next i

    GetErrorCorrectPolynomial = a
End Function

' ============================================================================
' Ma BCH (format info 15 bit, version info 18 bit) -- port dung tu QRUtil.getBCHTypeInfo/
' getBCHTypeNumber trong dist/qrcode.js.
' ============================================================================

Private Function BchDigit(ByVal data As Long) As Long
    Dim digit As Long
    digit = 0
    Do While data <> 0
        digit = digit + 1
        data = BitShiftRight(data, 1)
    Loop
    BchDigit = digit
End Function

Private Function GetBchTypeInfo(ByVal data As Long) As Long
    Dim d As Long
    d = BitShiftLeft(data, 10)
    Do While BchDigit(d) - BchDigit(G15) >= 0
        d = d Xor BitShiftLeft(G15, BchDigit(d) - BchDigit(G15))
    Loop
    GetBchTypeInfo = (BitShiftLeft(data, 10) Or d) Xor G15_MASK
End Function

Private Function GetBchTypeNumber(ByVal data As Long) As Long
    Dim d As Long
    d = BitShiftLeft(data, 12)
    Do While BchDigit(d) - BchDigit(G18) >= 0
        d = d Xor BitShiftLeft(G18, BchDigit(d) - BchDigit(G18))
    Loop
    GetBchTypeNumber = BitShiftLeft(data, 12) Or d
End Function

' ============================================================================
' Bo dem bit (bit buffer) -- port dung tu qrBitBuffer trong dist/qrcode.js. Trang thai module-
' level (mBufData/mBufByteCount/mBufBitLen) vi chi mot bo dem duoc dung tai mot thoi diem
' (BufReset luon goi dau CreateData).
' ============================================================================

Private Sub BufReset()
    ReDim mBufData(0 To 0)
    mBufByteCount = 0
    mBufBitLen = 0
End Sub

Private Sub BufPutBit(ByVal bit As Boolean)
    Dim bufIndex As Long
    bufIndex = mBufBitLen \ 8
    If mBufByteCount <= bufIndex Then
        ReDim Preserve mBufData(0 To bufIndex)
        mBufByteCount = bufIndex + 1
    End If
    If bit Then
        mBufData(bufIndex) = mBufData(bufIndex) Or BitShiftRight(&H80, mBufBitLen Mod 8)
    End If
    mBufBitLen = mBufBitLen + 1
End Sub

Private Sub BufPut(ByVal num As Long, ByVal bitLength As Long)
    Dim i As Long
    Dim bit As Boolean
    For i = 0 To bitLength - 1
        ' Tach rieng thanh bien truoc khi goi Sub -- goi "BufPutBit (bieu thuc) = 1" thang mot
        ' dong de VBA phan tich cu phap sai (nham (bieu thuc) la danh sach tham so boc ngoac cua
        ' mot Sub goi khong qua "Call", roi khong hieu duoc "= 1" con lai).
        bit = (BitShiftRight(num, bitLength - i - 1) And 1) = 1
        BufPutBit bit
    Next i
End Sub

' ============================================================================
' Do dai truong "length" theo che do Byte -- port dung tu QRUtil.getLengthInBits trong
' dist/qrcode.js, CHI nhanh MODE_8BIT_BYTE (module nay chi ho tro che do Byte -- xem dau file).
' ============================================================================

Private Function GetLengthInBitsByte(ByVal typeNumber As Long) As Long
    If typeNumber < 10 Then
        GetLengthInBitsByte = 8
    Else
        GetLengthInBitsByte = 16
    End If
End Function

' ============================================================================
' Bang so lieu -- SINH BANG SCRIPT truc tiep tu dist/qrcode.js (Kazuhiko Arase, MIT), roi doi
' chieu vong lap nguoc (parse lai, so khop JSON goc bang Node.js) -- KHONG go tay. Luu duoi dang
' chuoi phan cach "|" (giua cac version) va "," (giua cac so trong mot version), tach bang Split
' -- tranh cu phap khai bao mang 2 chieu lom xom (moi version dai khac nhau) dom hon nhieu so voi
' khai bao mang long dong.
' mPatternRows(v-1) = toa do alignment pattern cua version v (chuoi rong cho version 1).
' mRsRows(v-1) = "count,totalCount,dataCount[,count2,totalCount2,dataCount2]" cho MUC M cua
' version v (da loc san tu bang day du 4 muc L/M/Q/H -- xem ECC_TABLE_ROW_OFFSET_M).
' ============================================================================

Private Sub EnsureTables()
    If mTablesReady Then Exit Sub

    Dim patStr As String
    patStr = ""
    patStr = patStr & "|6,18|6,22|6,26|6,30|6,34|6,22,38|6,24,42|6,26,46|6,28,50|6,30,54|6,32"
    patStr = patStr & ",58|6,34,62|6,26,46,66|6,26,48,70|6,26,50,74|6,30,54,78|6,30,56,82|6,3"
    patStr = patStr & "0,58,86|6,34,62,90|6,28,50,72,94|6,26,50,74,98|6,30,54,78,102|6,28,54,"
    patStr = patStr & "80,106|6,32,58,84,110|6,30,58,86,114|6,34,62,90,118|6,26,50,74,98,122|"
    patStr = patStr & "6,30,54,78,102,126|6,26,52,78,104,130|6,30,56,82,108,134|6,34,60,86,11"
    patStr = patStr & "2,138|6,30,58,86,114,142|6,34,62,90,118,146|6,30,54,78,102,126,150|6,2"
    patStr = patStr & "4,50,76,102,128,154|6,28,54,80,106,132,158|6,32,58,84,110,136,162|6,26"
    patStr = patStr & ",54,82,110,138,166|6,30,58,86,114,142,170"
    mPatternRows = Split(patStr, "|")

    Dim rsStr As String
    rsStr = ""
    rsStr = rsStr & "1,26,16|1,44,28|1,70,44|2,50,32|2,67,43|4,43,27|4,49,31|2,60,38,2,61,3"
    rsStr = rsStr & "9|3,58,36,2,59,37|4,69,43,1,70,44|1,80,50,4,81,51|6,58,36,2,59,37|8,59"
    rsStr = rsStr & ",37,1,60,38|4,64,40,5,65,41|5,65,41,5,66,42|7,73,45,3,74,46|10,74,46,1"
    rsStr = rsStr & ",75,47|9,69,43,4,70,44|3,70,44,11,71,45|3,67,41,13,68,42|17,68,42|17,7"
    rsStr = rsStr & "4,46|4,75,47,14,76,48|6,73,45,14,74,46|8,75,47,13,76,48|19,74,46,4,75,"
    rsStr = rsStr & "47|22,73,45,3,74,46|3,73,45,23,74,46|21,73,45,7,74,46|19,75,47,10,76,4"
    rsStr = rsStr & "8|2,74,46,29,75,47|10,74,46,23,75,47|14,74,46,21,75,47|14,74,46,23,75,"
    rsStr = rsStr & "47|12,75,47,26,76,48|6,75,47,34,76,48|29,74,46,14,75,47|13,74,46,32,75"
    rsStr = rsStr & ",47|40,75,47,7,76,48|18,75,47,31,76,48"
    mRsRows = Split(rsStr, "|")

    mTablesReady = True
End Sub

' Tach mot chuoi "a,b,c" thanh mang Long -- tra qua ByRef outArr/outCount, KHONG dua vao
' UBound/LBound cua mang khong-phan-tu.
' VBA KHONG co cach hop le nao de ReDim mot mang dong ve dung 0 phan tu voi upper bound < lower
' bound. Vi vay outArr LUON co it nhat MOT phan tu (gia tri 0, khong dung toi) khi outCount=0 --
' noi goi PHAI kiem tra outCount TRUOC, KHONG duoc doc outArr khi outCount<=0 (xem
' SetupPositionAdjustPattern).
Private Sub ParseLongCsv(ByVal s As String, ByRef outArr() As Long, ByRef outCount As Long)
    If Len(s) = 0 Then
        outCount = 0
        ReDim outArr(0 To 0)
        Exit Sub
    End If

    Dim parts() As String
    parts = Split(s, ",")
    outCount = UBound(parts) + 1
    ReDim outArr(0 To outCount - 1)
    Dim i As Long
    For i = 0 To outCount - 1
        outArr(i) = CLng(Trim$(parts(i)))
    Next i
End Sub

Private Sub PatternPositionFor(ByVal typeNumber As Long, ByRef outArr() As Long, ByRef outCount As Long)
    ParseLongCsv mPatternRows(typeNumber - 1), outArr, outCount
End Sub

' Khai trien "count,totalCount,dataCount[,...]" thanh danh sach block (lap lai dung "count" lan
' moi nhom) -- port dung tu QRRSBlock.getRSBlocks trong dist/qrcode.js. Hang RS_BLOCK KHONG BAO
' GIO rong (moi version luon co it nhat mot nhom) nen khong can xu ly truong hop 0 nhu
' ParseLongCsv/PatternPositionFor.
Private Sub GetRsBlocks(ByVal typeNumber As Long, ByRef outTotalCounts() As Long, _
        ByRef outDataCounts() As Long, ByRef outBlockCount As Long)
    Dim rowNums() As Long
    Dim rowCount As Long
    ParseLongCsv mRsRows(typeNumber - 1), rowNums, rowCount

    Dim groups As Long
    groups = rowCount \ 3

    Dim g As Long, cnt As Long
    outBlockCount = 0
    For g = 0 To groups - 1
        outBlockCount = outBlockCount + rowNums(g * 3 + 0)
    Next g

    ReDim outTotalCounts(0 To outBlockCount - 1)
    ReDim outDataCounts(0 To outBlockCount - 1)

    Dim idx As Long, j As Long, tc As Long, dc As Long
    idx = 0
    For g = 0 To groups - 1
        cnt = rowNums(g * 3 + 0)
        tc = rowNums(g * 3 + 1)
        dc = rowNums(g * 3 + 2)
        For j = 1 To cnt
            outTotalCounts(idx) = tc
            outDataCounts(idx) = dc
            idx = idx + 1
        Next j
    Next g
End Sub

' ============================================================================
' Dung du lieu ma hoa cuoi (mode + do dai + du lieu + ket thuc + dem, roi sua loi Reed-Solomon
' theo tung khoi va xen khoi) -- port dung tu createData/createBytes trong dist/qrcode.js.
' ============================================================================

Private Function CreateData(ByVal typeNumber As Long, ByRef dataBytes() As Long) As Variant
    Dim totalCounts() As Long, dataCounts() As Long, blockCount As Long
    GetRsBlocks typeNumber, totalCounts, dataCounts, blockCount

    BufReset

    Dim lengthBits As Long
    lengthBits = GetLengthInBitsByte(typeNumber)
    Dim dataLen As Long
    dataLen = UBound(dataBytes) - LBound(dataBytes) + 1

    BufPut 4, 4   ' chi thi che do: Byte = 0100
    BufPut dataLen, lengthBits
    Dim i As Long
    For i = LBound(dataBytes) To UBound(dataBytes)
        BufPut dataBytes(i), 8
    Next i

    Dim totalDataCount As Long
    totalDataCount = 0
    Dim b As Long
    For b = 0 To blockCount - 1
        totalDataCount = totalDataCount + dataCounts(b)
    Next b

    If mBufBitLen > totalDataCount * 8 Then
        ' Khong le xay ra -- SelectVersion da chon dung version du cho noi dung nay TRUOC khi goi
        ' toi day, bang chinh cong thuc do dai bit nay. Ma so loi (1107) CO Y khac 1104 cua
        ' SelectVersion de GenerateQrModuleMatrix khong nham day voi truong hop "khong version nao
        ' vua" that su (xem CapacityErrHandler) -- day la loi lap trinh that su neu xay ra, khong
        ' phai loi noi dung nguoi dung nhap.
        Err.Raise vbObjectError + 1107, "QrCodeGenerator.CreateData", "code length overflow"
    End If

    If mBufBitLen + 4 <= totalDataCount * 8 Then BufPut 0, 4
    Do While mBufBitLen Mod 8 <> 0
        BufPutBit False
    Loop
    Do
        If mBufBitLen >= totalDataCount * 8 Then Exit Do
        BufPut PAD0, 8
        If mBufBitLen >= totalDataCount * 8 Then Exit Do
        BufPut PAD1, 8
    Loop

    CreateData = CreateBytes(totalCounts, dataCounts, blockCount)
End Function

Private Function CreateBytes(ByRef totalCounts() As Long, ByRef dataCounts() As Long, _
        ByVal blockCount As Long) As Variant
    Dim maxDcCount As Long, maxEcCount As Long
    maxDcCount = 0
    maxEcCount = 0

    Dim dcdata() As Variant, ecdata() As Variant
    ReDim dcdata(0 To blockCount - 1)
    ReDim ecdata(0 To blockCount - 1)

    Dim offset As Long
    offset = 0

    Dim r As Long, i As Long, dcCount As Long, ecCount As Long
    For r = 0 To blockCount - 1
        dcCount = dataCounts(r)
        ecCount = totalCounts(r) - dcCount
        If dcCount > maxDcCount Then maxDcCount = dcCount
        If ecCount > maxEcCount Then maxEcCount = ecCount

        Dim dcArr() As Long
        ReDim dcArr(0 To dcCount - 1)
        For i = 0 To dcCount - 1
            dcArr(i) = mBufData(i + offset) And &HFF
        Next i
        offset = offset + dcCount
        dcdata(r) = dcArr

        Dim rsPoly As Variant
        rsPoly = GetErrorCorrectPolynomial(ecCount)
        Dim rsPolyArr() As Long
        rsPolyArr = rsPoly
        Dim rsPolyLen As Long
        rsPolyLen = UBound(rsPolyArr) - LBound(rsPolyArr) + 1

        Dim rawPoly As Variant
        rawPoly = PolyCreate(dcArr, rsPolyLen - 1)
        Dim rawPolyArr() As Long
        rawPolyArr = rawPoly

        Dim modPoly As Variant
        modPoly = PolyMod(rawPolyArr, rsPolyArr)
        Dim modPolyArr() As Long
        modPolyArr = modPoly
        Dim modPolyLen As Long
        modPolyLen = UBound(modPolyArr) - LBound(modPolyArr) + 1

        Dim ecLen As Long
        ecLen = rsPolyLen - 1
        Dim ecArr() As Long
        ReDim ecArr(0 To ecLen - 1)
        Dim modIndex As Long
        For i = 0 To ecLen - 1
            modIndex = i + modPolyLen - ecLen
            If modIndex >= 0 Then
                ecArr(i) = modPolyArr(modIndex)
            Else
                ecArr(i) = 0
            End If
        Next i
        ecdata(r) = ecArr
    Next r

    Dim totalCodeCount As Long
    totalCodeCount = 0
    For r = 0 To blockCount - 1
        totalCodeCount = totalCodeCount + totalCounts(r)
    Next r

    Dim data() As Long
    ReDim data(0 To totalCodeCount - 1)
    Dim idx As Long
    idx = 0

    Dim k As Long
    For k = 0 To maxDcCount - 1
        For r = 0 To blockCount - 1
            Dim curDc() As Long
            curDc = dcdata(r)
            If k <= UBound(curDc) Then
                data(idx) = curDc(k)
                idx = idx + 1
            End If
        Next r
    Next k
    For k = 0 To maxEcCount - 1
        For r = 0 To blockCount - 1
            Dim curEc() As Long
            curEc = ecdata(r)
            If k <= UBound(curEc) Then
                data(idx) = curEc(k)
                idx = idx + 1
            End If
        Next r
    Next k

    CreateBytes = data
End Function

' ============================================================================
' Dat module (finder pattern, timing pattern, alignment pattern, version/format info, xen du lieu
' theo zigzag) -- port dung tu makeImpl va cac ham setupXxx trong dist/qrcode.js.
' ============================================================================

Private Sub SetupPositionProbePattern(ByVal row As Long, ByVal col As Long)
    Dim r As Long, c As Long
    For r = -1 To 7
        If Not (row + r <= -1 Or mN <= row + r) Then
            For c = -1 To 7
                If Not (col + c <= -1 Or mN <= col + c) Then
                    If (r >= 0 And r <= 6 And (c = 0 Or c = 6)) _
                        Or (c >= 0 And c <= 6 And (r = 0 Or r = 6)) _
                        Or (r >= 2 And r <= 4 And c >= 2 And c <= 4) Then
                        mModules(row + r, col + c) = 1
                    Else
                        mModules(row + r, col + c) = 0
                    End If
                End If
            Next c
        End If
    Next r
End Sub

Private Sub SetupTimingPattern()
    Dim r As Long, c As Long
    For r = 8 To mN - 8 - 1
        If mModules(r, 6) = -1 Then
            mModules(r, 6) = IIf(r Mod 2 = 0, 1, 0)
        End If
    Next r
    For c = 8 To mN - 8 - 1
        If mModules(6, c) = -1 Then
            mModules(6, c) = IIf(c Mod 2 = 0, 1, 0)
        End If
    Next c
End Sub

Private Sub SetupPositionAdjustPattern(ByVal typeNumber As Long)
    Dim pos() As Long
    Dim posLen As Long
    PatternPositionFor typeNumber, pos, posLen
    If posLen <= 0 Then Exit Sub

    Dim i As Long, j As Long, row As Long, col As Long, r As Long, c As Long
    For i = 0 To posLen - 1
        For j = 0 To posLen - 1
            row = pos(i)
            col = pos(j)
            If mModules(row, col) = -1 Then
                For r = -2 To 2
                    For c = -2 To 2
                        If r = -2 Or r = 2 Or c = -2 Or c = 2 Or (r = 0 And c = 0) Then
                            mModules(row + r, col + c) = 1
                        Else
                            mModules(row + r, col + c) = 0
                        End If
                    Next c
                Next r
            End If
        Next j
    Next i
End Sub

Private Sub SetupTypeNumber(ByVal typeNumber As Long, ByVal test As Boolean)
    Dim bits As Long
    bits = GetBchTypeNumber(typeNumber)

    Dim i As Long, mv As Long
    For i = 0 To 17
        mv = IIf((Not test) And ((BitShiftRight(bits, i) And 1) = 1), 1, 0)
        mModules(i \ 3, (i Mod 3) + mN - 8 - 3) = mv
    Next i
    For i = 0 To 17
        mv = IIf((Not test) And ((BitShiftRight(bits, i) And 1) = 1), 1, 0)
        mModules((i Mod 3) + mN - 8 - 3, i \ 3) = mv
    Next i
End Sub

Private Sub SetupTypeInfo(ByVal test As Boolean, ByVal maskPattern As Long, _
        ByVal errorCorrectionLevelValue As Long)
    Dim d As Long
    d = BitShiftLeft(errorCorrectionLevelValue, 3) Or maskPattern
    Dim bits As Long
    bits = GetBchTypeInfo(d)

    Dim i As Long, mv As Long
    For i = 0 To 14
        mv = IIf((Not test) And ((BitShiftRight(bits, i) And 1) = 1), 1, 0)
        If i < 6 Then
            mModules(i, 8) = mv
        ElseIf i < 8 Then
            mModules(i + 1, 8) = mv
        Else
            mModules(mN - 15 + i, 8) = mv
        End If
    Next i
    For i = 0 To 14
        mv = IIf((Not test) And ((BitShiftRight(bits, i) And 1) = 1), 1, 0)
        If i < 8 Then
            mModules(8, mN - i - 1) = mv
        ElseIf i < 9 Then
            mModules(8, 15 - i - 1 + 1) = mv
        Else
            mModules(8, 15 - i - 1) = mv
        End If
    Next i

    mModules(mN - 8, 8) = IIf(Not test, 1, 0)
End Sub

Private Function GetMaskBit(ByVal pattern As Long, ByVal i As Long, ByVal j As Long) As Boolean
    Select Case pattern
        Case 0: GetMaskBit = ((i + j) Mod 2 = 0)
        Case 1: GetMaskBit = (i Mod 2 = 0)
        Case 2: GetMaskBit = (j Mod 3 = 0)
        Case 3: GetMaskBit = ((i + j) Mod 3 = 0)
        Case 4: GetMaskBit = (((i \ 2) + (j \ 3)) Mod 2 = 0)
        Case 5: GetMaskBit = (((i * j) Mod 2) + ((i * j) Mod 3) = 0)
        Case 6: GetMaskBit = ((((i * j) Mod 2) + ((i * j) Mod 3)) Mod 2 = 0)
        Case 7: GetMaskBit = ((((i * j) Mod 3) + ((i + j) Mod 2)) Mod 2 = 0)
        Case Else
            Err.Raise vbObjectError + 1106, "QrCodeGenerator.GetMaskBit", "bad mask pattern"
    End Select
End Function

Private Sub MapData(ByRef data() As Long, ByVal maskPattern As Long)
    Dim dataLen As Long
    dataLen = UBound(data) - LBound(data) + 1

    Dim inc As Long, row As Long, bitIndex As Long, byteIndex As Long
    inc = -1
    row = mN - 1
    bitIndex = 7
    byteIndex = 0

    Dim col As Long, c As Long, dark As Boolean
    col = mN - 1
    Do While col > 0
        If col = 6 Then col = col - 1
        Do
            For c = 0 To 1
                If mModules(row, col - c) = -1 Then
                    dark = False
                    If byteIndex < dataLen Then
                        dark = (BitShiftRight(data(LBound(data) + byteIndex), bitIndex) And 1) = 1
                    End If
                    If GetMaskBit(maskPattern, row, col - c) Then dark = Not dark
                    mModules(row, col - c) = IIf(dark, 1, 0)
                    bitIndex = bitIndex - 1
                    If bitIndex = -1 Then
                        byteIndex = byteIndex + 1
                        bitIndex = 7
                    End If
                End If
            Next c
            row = row + inc
            If row < 0 Or mN <= row Then
                row = row - inc
                inc = -inc
                Exit Do
            End If
        Loop
        col = col - 2
    Loop
End Sub

Private Function ModIsDark(ByVal r As Long, ByVal c As Long) As Boolean
    ModIsDark = (mModules(r, c) = 1)
End Function

' Danh gia "diem mat" (lost point) cua ma tran hien tai theo 4 quy tac phat cua chuan QR -- dung
' de chon mask tot nhat trong 8 mask. Port dung tu QRUtil.getLostPoint trong dist/qrcode.js.
Private Function GetLostPoint() As Double
    Dim lostPoint As Double
    lostPoint = 0

    Dim row As Long, col As Long, r As Long, c As Long, sameCount As Long
    Dim dark As Boolean

    ' LEVEL1
    For row = 0 To mN - 1
        For col = 0 To mN - 1
            sameCount = 0
            dark = ModIsDark(row, col)
            For r = -1 To 1
                If row + r >= 0 And row + r < mN Then
                    For c = -1 To 1
                        If col + c >= 0 And col + c < mN Then
                            If Not (r = 0 And c = 0) Then
                                If dark = ModIsDark(row + r, col + c) Then
                                    sameCount = sameCount + 1
                                End If
                            End If
                        End If
                    Next c
                End If
            Next r
            If sameCount > 5 Then lostPoint = lostPoint + (3 + sameCount - 5)
        Next col
    Next row

    ' LEVEL2
    Dim cnt As Long
    For row = 0 To mN - 2
        For col = 0 To mN - 2
            cnt = 0
            If ModIsDark(row, col) Then cnt = cnt + 1
            If ModIsDark(row + 1, col) Then cnt = cnt + 1
            If ModIsDark(row, col + 1) Then cnt = cnt + 1
            If ModIsDark(row + 1, col + 1) Then cnt = cnt + 1
            If cnt = 0 Or cnt = 4 Then lostPoint = lostPoint + 3
        Next col
    Next row

    ' LEVEL3 -- ngang
    For row = 0 To mN - 1
        For col = 0 To mN - 7
            If ModIsDark(row, col) And Not ModIsDark(row, col + 1) And ModIsDark(row, col + 2) _
                And ModIsDark(row, col + 3) And ModIsDark(row, col + 4) _
                And Not ModIsDark(row, col + 5) And ModIsDark(row, col + 6) Then
                lostPoint = lostPoint + 40
            End If
        Next col
    Next row
    ' LEVEL3 -- doc
    For col = 0 To mN - 1
        For row = 0 To mN - 7
            If ModIsDark(row, col) And Not ModIsDark(row + 1, col) And ModIsDark(row + 2, col) _
                And ModIsDark(row + 3, col) And ModIsDark(row + 4, col) _
                And Not ModIsDark(row + 5, col) And ModIsDark(row + 6, col) Then
                lostPoint = lostPoint + 40
            End If
        Next row
    Next col

    ' LEVEL4
    Dim darkCount As Long
    darkCount = 0
    For col = 0 To mN - 1
        For row = 0 To mN - 1
            If ModIsDark(row, col) Then darkCount = darkCount + 1
        Next row
    Next col
    Dim ratio As Double
    ratio = Abs(100 * darkCount / mN / mN - 50) / 5
    lostPoint = lostPoint + ratio * 10

    GetLostPoint = lostPoint
End Function

Private Sub MakeImpl(ByVal typeNumber As Long, ByVal test As Boolean, ByVal maskPattern As Long, _
        ByRef dataBytes() As Long)
    mN = typeNumber * 4 + 17
    ReDim mModules(0 To mN - 1, 0 To mN - 1)
    Dim r As Long, c As Long
    For r = 0 To mN - 1
        For c = 0 To mN - 1
            mModules(r, c) = -1
        Next c
    Next r

    SetupPositionProbePattern 0, 0
    SetupPositionProbePattern mN - 7, 0
    SetupPositionProbePattern 0, mN - 7
    SetupPositionAdjustPattern typeNumber
    SetupTimingPattern
    SetupTypeInfo test, maskPattern, ECC_LEVEL_VALUE_M
    If typeNumber >= 7 Then SetupTypeNumber typeNumber, test

    Dim codewords As Variant
    codewords = CreateData(typeNumber, dataBytes)
    Dim codewordsArr() As Long
    codewordsArr = codewords
    MapData codewordsArr, maskPattern
End Sub

Private Function GetBestMaskPattern(ByVal typeNumber As Long, ByRef dataBytes() As Long) As Long
    Dim minLostPoint As Double
    Dim pattern As Long
    Dim i As Long, lostPoint As Double

    For i = 0 To 7
        MakeImpl typeNumber, True, i, dataBytes
        lostPoint = GetLostPoint()
        If i = 0 Or minLostPoint > lostPoint Then
            minLostPoint = lostPoint
            pattern = i
        End If
    Next i

    GetBestMaskPattern = pattern
End Function

' Chon version QR nho nhat (1..40) du chua noi dung o muc sua loi M -- port dung tu vong lap do
' trong _this.make cua dist/qrcode.js (uoc luong do dai bit truc tiep bang cong thuc, khong can
' dung that bo dem, vi ca hai cach cho cung mot ket qua so hoc).
Private Function SelectVersion(ByRef dataBytes() As Long) As Long
    Dim dataLen As Long
    dataLen = UBound(dataBytes) - LBound(dataBytes) + 1

    Dim typeNumber As Long
    For typeNumber = 1 To 40
        Dim totalCounts() As Long, dataCounts() As Long, blockCount As Long
        GetRsBlocks typeNumber, totalCounts, dataCounts, blockCount

        Dim bitLen As Long
        bitLen = 4 + GetLengthInBitsByte(typeNumber) + dataLen * 8

        Dim totalDataCount As Long
        totalDataCount = 0
        Dim b As Long
        For b = 0 To blockCount - 1
            totalDataCount = totalDataCount + dataCounts(b)
        Next b

        If bitLen <= totalDataCount * 8 Then
            SelectVersion = typeNumber
            Exit Function
        End If
    Next typeNumber

    Err.Raise vbObjectError + 1104, "QrCodeGenerator.SelectVersion", _
        "content too long for any QR version"
End Function

' ============================================================================
' Dich bit thu cong -- VBA KHONG CO toan tu <<, >>, >>> cua JavaScript. Gia tri thuc te trong
' thuat toan nay luon nho (duoi 20 bit) nen phep nhan/chia luy thua 2 khong tran Long 32-bit co
' dau -- xem phan tich chi tiet o dau file.
' ============================================================================

Private Function BitShiftLeft(ByVal x As Long, ByVal n As Long) As Long
    BitShiftLeft = CLng(x * (2 ^ n))
End Function

Private Function BitShiftRight(ByVal x As Long, ByVal n As Long) As Long
    BitShiftRight = x \ (2 ^ n)
End Function

Private Function WriteQrBmpFile(ByVal filePath As String, ByRef matrix As Variant, _
        ByVal moduleCount As Long) As Boolean
    Dim cellSize As Long
    cellSize = Int(900 / moduleCount + 0.5)
    If cellSize < 4 Then cellSize = 4
    Dim margin As Long
    margin = cellSize * 4

    Dim size As Long
    size = moduleCount * cellSize + margin * 2

    Dim rowBytes As Long
    rowBytes = ((size * 3 + 3) \ 4) * 4
    Dim pixelDataSize As Long
    pixelDataSize = rowBytes * size
    Dim fileSize As Long
    fileSize = 54 + pixelDataSize

    Dim buf() As Byte
    ReDim buf(0 To fileSize - 1)

    ' BITMAPFILEHEADER (14 byte)
    buf(0) = 66: buf(1) = 77   ' "BM"
    PutLE4 buf, 2, fileSize
    PutLE4 buf, 6, 0
    PutLE4 buf, 10, 54

    ' BITMAPINFOHEADER (40 byte)
    PutLE4 buf, 14, 40
    PutLE4 buf, 18, size
    PutLE4 buf, 22, size
    PutLE2 buf, 26, 1        ' planes
    PutLE2 buf, 28, 24       ' bitCount
    PutLE4 buf, 30, 0        ' compression = BI_RGB
    PutLE4 buf, 34, pixelDataSize
    PutLE4 buf, 38, 0
    PutLE4 buf, 42, 0
    PutLE4 buf, 46, 0
    PutLE4 buf, 50, 0

    ' Diem anh -- BMP luu tu DUOI LEN (fileRow 0 = hang duoi cung man hinh).
    Dim mat() As Long
    mat = matrix

    Dim fileRow As Long, px As Long, py As Long, base As Long
    Dim moduleRow As Long, moduleCol As Long, isDark As Boolean
    For fileRow = 0 To size - 1
        py = (size - 1) - fileRow
        base = 54 + fileRow * rowBytes
        For px = 0 To size - 1
            isDark = False
            If px >= margin And px < margin + moduleCount * cellSize _
                And py >= margin And py < margin + moduleCount * cellSize Then
                moduleCol = (px - margin) \ cellSize
                moduleRow = (py - margin) \ cellSize
                isDark = (mat(moduleRow, moduleCol) = 1)
            End If
            Dim pOff As Long
            pOff = base + px * 3
            If isDark Then
                buf(pOff) = 0: buf(pOff + 1) = 0: buf(pOff + 2) = 0
            Else
                buf(pOff) = 255: buf(pOff + 1) = 255: buf(pOff + 2) = 255
            End If
        Next px
    Next fileRow

    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Binary Access Write As #fileNum
    Put #fileNum, 1, buf
    Close #fileNum

    WriteQrBmpFile = True
End Function

Private Sub PutLE4(ByRef buf() As Byte, ByVal offset As Long, ByVal value As Long)
    buf(offset) = value And &HFF
    buf(offset + 1) = BitShiftRight(value, 8) And &HFF
    buf(offset + 2) = BitShiftRight(value, 16) And &HFF
    buf(offset + 3) = BitShiftRight(value, 24) And &HFF
End Sub

Private Sub PutLE2(ByRef buf() As Byte, ByVal offset As Long, ByVal value As Long)
    buf(offset) = value And &HFF
    buf(offset + 1) = BitShiftRight(value, 8) And &HFF
End Sub
