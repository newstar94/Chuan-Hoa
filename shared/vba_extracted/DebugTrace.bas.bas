Attribute VB_Name = "DebugTrace"
'==============================================================
' DebugTrace â€” Nhat ky GO LOI TAM THOI, ghi RA FILE (khac OperationLogger.bas: do la nhat ky
' NGUOI DUNG, chi ton tai trong phien, cam ghi ra file - NFR-SEC-03/CLAUDE.md muc 3.6). Phuc vu
' rieng viec chan doan crash/loi sai vi tri trong qua trinh phat trien - file.log ghi cuc bo trong
' thu muc TEMP cua nguoi dung, khong gui di dau, khong vi pham "hoan toan client-side, khong
' telemetry". Se tat (ENABLED = False) hoac go bo khi da chan doan xong.
' Dat trong Environ$("TEMP") - luon ghi duoc du tai lieu dang mo nam o dau (local, mang,
' OneDrive), khac loi "Bad file name or number" tung gap voi duong dan theo Document.Path.
' PHAM VI THEO DOI: moi thu tuc Public cua modules/, classes/, forms/ - callback ribbon, quy tac
' kiem tra, cac ham "Sua", dinh dang, nhan dien, chuyen doi bang ma, hop thoai. KHONG theo doi: du
' lieu sinh tu dong/tra cuu cache (RuleData.bas, RuleLoader.Get*), ham thuan chuoi/quy doi don vi
' cua Utils, plumbing thanh tien do, hang so chuoi, OperationLogger.bas va chinh
' DebugTrace/TraceScope (se de quy vo tan) - deu la ham thuan tinh toan chay day dac trong vong
' lap, ghi lai chi lam loang nhat ky ma khong giup dinh vi cho treo. Thu tuc Private la chi tiet
' ben trong mot thu tuc Public da theo doi: khi treo, dong "BAT DAU X" thieu "HOAN THANH X" di kem
' da khoanh dung X.
' CACH TAT TOAN BO: dat ENABLED = False ngay duoi - moi loi goi EnterScope tra ve Nothing va khong
' cham vao dia. Khong can go bo cac dong da chen.

Private Const ENABLED As Boolean = True
Private Const LOG_FILE_NAME As String = "ChuanHoaTheThuc-debug.log"

' So dong toi da giu trong bo dem truoc khi xa ra file. Cang nho thi cang it nguy co mat dong nhat
' ky cuoi khi Word chet dot ngot nhung cang nhieu lan cham dia (do duoc: mot lan mo/ghi/ dong file
' trong TEMP ~20ms, nen 25 dong/lan ~0,6 giay cho mot lan "Kiem tra" day du).
Private Const FLUSH_LINES As Long = 25

' Khoang cach toi thieu (giay) giua hai lan xa bo dem khi mot thao tac cap cao nhat vua ket thuc.
' Can nguong nay vi mot so thu tuc cap cao nhat chay rat day dac ngoai y nguoi dung (vi du
' AppEventsHost.OnSelectionChange chay moi lan con tro di chuyen) - khong co nguong thi moi phim
' go la mot lan cham dia. 0,2 giay du thua de chan dieu do, nhung van giu duoc vai dong cuoi truoc
' khi crash (gia tri chan doan quan trong nhat thuong nam o do).
Private Const FLUSH_MIN_SECONDS As Single = 0.2!

' Kich thuoc toi da cua file log truoc khi xoay vong (doi ten thanh...-prev.log, mo file moi). Ghi
' nhat ky VAO/RA cho MOI thu tuc sinh nhieu dong hon han truoc day nen phai co chan tren, neu
' khong file se phinh vo han qua nhieu ngay lam viec.
Private Const MAX_LOG_BYTES As Long = 8000000

Private mPathReady As Boolean
Private mPath As String

' Da xoa nhat ky cho phien VBA nay chua (xem ResetLogAtStartup). MOI khai bao cap module PHAI dung
' TRUOC toan bo Sub/Function - quy tac VBA, xem ghi chu dau Utils.bas.
Private mLogResetDone As Boolean

' --- Bo dem ghi ----------------------------------------------------------------------------- Ghi
' nhat ky BAT DAU/HOAN THANH cho moi thu tuc sinh hang nghin dong moi lan "Kiem tra" - mo/
' ghi/dong file TRON VEN cho MOI dong (nhu ban dau) se lap lai kieu "mo/dong file trong vong lap"
' tung la nguyen nhan loi "Not Responding" o mot dot sua truoc. Nen: gom dong vao bo nho, chi cham
' dia khi (a) day FLUSH_LINES dong, (b) mot thao tac cap cao nhat vua ket thuc (do sau ve 0), hoac
' (c) vua ghi mot loi VBA that su - dam bao du add-in treo/crash giua chung thi phan da ghi truoc
' do van nam tren dia.
Private mBuf As String
Private mBufLines As Long
Private mLastFlush As Single

' --- Do sau ngan xep goi --------------------------------------------------------------------
' Dung de thut le cac dong nhat ky theo cap goi (nhin ra ngay cay goi ham), va de biet khi nao mot
' thao tac cap cao nhat ket thuc (mDepth tro ve 0) ma xa bo dem ra file.
Private mDepth As Long

Private Sub EnsurePath()
    If mPathReady Then Exit Sub
    mPath = Environ$("TEMP") & "\" & LOG_FILE_NAME
    mPathReady = True
    RotateIfTooBig
End Sub

' Doi file log hien tai thanh "...-prev.log" khi no vuot MAX_LOG_BYTES - giu lai DUNG mot the he
' truoc, du de doi chieu ma khong de dia phinh vo han.
Private Sub RotateIfTooBig()
    On Error Resume Next
    If Len(Dir$(mPath)) = 0 Then Exit Sub
    If FileLen(mPath) < MAX_LOG_BYTES Then Exit Sub
    Dim prevPath As String
    prevPath = Replace$(mPath, ".log", "-prev.log")
    If Len(Dir$(prevPath)) > 0 Then Kill prevPath
    Name mPath As prevPath
    On Error GoTo 0
End Sub

' Doi ten file log hien tai (neu co) thanh "..._lastrun.log" (ghi de ban _lastrun.log cu neu co)
' roi bat dau file log MOI, trong, cho phien nay - moi khi add-in khoi dong. Goi dau tien trong
' RibbonCallbacks.RibbonOnLoad, truoc bat ky Log/EnterScope nao khac de khong doi ten nham luc
' dong vua ghi. An toan kha nang loi: chi Name/Kill khi file dang ton tai va khong bi khoa (vi du
' dang mo boi Notepad) - On Error Resume Next bo qua im lang.
' nhat ky cua PHIEN TRUOC (thuong la phien vua chet, dung thu can xem nhat) khong con bi mat trang
' khi Word mo lai/nap lai add-in - luon co CHINH XAC MOT ban "_lastrun.log" la nhat ky cua lan
' chay gan nhat truoc lan hien tai.
' CHI doi ten MOT LAN cho moi phien VBA (mLogResetDone): Word goi onLoad mot lan cho MOI cua so
' tai lieu, neu doi ten moi lan onLoad thi mo them mot cua so se lam trong nhat ky cua ca phien -
' tung lam mat bang chung chan doan crash o cac dong sau. Phien VBA cham dut khi Word dong han,
' nen dung mot lan la du "moi phien lam viec bat dau voi nhat ky sach".
Public Sub ResetLogAtStartup()
    On Error Resume Next
    If mLogResetDone Then Exit Sub
    mLogResetDone = True
    EnsurePath
    If Len(Dir$(mPath)) > 0 Then
        Dim lastRunPath As String
        lastRunPath = Replace$(mPath, ".log", "_lastrun.log")
        If Len(Dir$(lastRunPath)) > 0 Then Kill lastRunPath
        Name mPath As lastRunPath
    End If
    Dim prevPath As String
    prevPath = Replace$(mPath, ".log", "-prev.log")
    If Len(Dir$(prevPath)) > 0 Then Kill prevPath
    mBuf = vbNullString
    mBufLines = 0
    mDepth = 0
    On Error GoTo 0
End Sub

' Ghi MOT dong nhat ky - source dang "Module.Ham", message la noi dung tuy y. Chuoi Unicode co dau
' ghi THANG (khong can Utils.ToUnaccented) - Open...For Append ghi file van ban khong di qua API
' Win32 ANSI nhu MsgBox/UserForm.Caption nen khong bi loi "?" (xem Utils.ToUnaccented).
' Dong duoc NOI vao bo dem chu chua cham vao dia ngay - xem ghi chu "VI SAO CO BO DEM" o tren.
Public Sub Log(ByVal source As String, ByVal message As String)
    If Not ENABLED Then Exit Sub
    On Error Resume Next
    AppendLine Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & Indent() & "[" & source & "] " & message
    On Error GoTo 0
End Sub

' Chuoi thut le theo do sau ngan xep goi hien tai (2 khoang trang moi cap).
Private Function Indent() As String
    If mDepth <= 0 Then Exit Function
    Indent = Space$(mDepth * 2)
End Function

Private Sub AppendLine(ByVal lineText As String)
    If Len(mBuf) = 0 Then
        mBuf = lineText
    Else
        mBuf = mBuf & vbCrLf & lineText
    End If
    mBufLines = mBufLines + 1
    If mBufLines >= FLUSH_LINES Then Flush
End Sub

' Xa bo dem ra file. Goi tu dong khi day bo dem / khi mot thao tac cap cao nhat ket thuc / khi vua
' ghi mot loi VBA - nhung cung Public de goi tay tu Immediate Window luc chan doan.
Public Sub Flush()
    If Len(mBuf) = 0 Then Exit Sub
    On Error Resume Next
    EnsurePath
    Dim fnum As Integer: fnum = FreeFile
    Open mPath For Append As #fnum
    Print #fnum, mBuf
    Close #fnum
    mBuf = vbNullString
    mBufLines = 0
    mLastFlush = Timer
    On Error GoTo 0
End Sub

' Xa bo dem, nhung bo qua neu vua xa cach day chua toi FLUSH_MIN_SECONDS giay. Dung o cho ket thuc
' mot thao tac cap cao nhat - noi vua can du lieu nam tren dia som, vua khong duoc phep cham dia
' theo nhip go phim (xem ghi chu FLUSH_MIN_SECONDS).
Private Sub FlushThrottled()
    Dim elapsed As Single
    elapsed = Timer - mLastFlush
    ' Timer quay ve 0 luc nua dem - coi nhu qua han de khong ket bo dem lai qua mot ngay lam viec.
    If elapsed < 0 Then elapsed = FLUSH_MIN_SECONDS
    If elapsed < FLUSH_MIN_SECONDS Then Exit Sub
    Flush
End Sub

Public Function EnterScope(ByVal source As String) As TraceScope
    If Not ENABLED Then Exit Function
    On Error Resume Next
    Dim s As TraceScope
    Set s = New TraceScope
    s.Init source
    Set EnterScope = s
    On Error GoTo 0
End Function

' Goi tu TraceScope.Init - ghi dong vao, roi moi tang do sau (dong "vao" thut le ngang cap voi ham
' goi no, cac dong ben trong thut sau hon).
Public Sub TraceBegin(ByVal source As String)
    If Not ENABLED Then Exit Sub
    On Error Resume Next
    AppendLine Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & Indent() & "-> BAT DAU " & source
    mDepth = mDepth + 1
    On Error GoTo 0
End Sub

' Goi tu TraceScope.Class_Terminate - giam do sau roi moi ghi dong ra, va xa bo dem khi da ve toi
' cap cao nhat (mot thao tac vua xong tron ven thi phan nhat ky cua no phai nam tren dia).
Public Sub TraceEnd(ByVal source As String)
    If Not ENABLED Then Exit Sub
    On Error Resume Next
    mDepth = mDepth - 1
    If mDepth < 0 Then mDepth = 0
    AppendLine Format$(Now, "yyyy-mm-dd hh:nn:ss") & " " & Indent() & "<- HOAN THANH " & source
    If mDepth = 0 Then FlushThrottled
    On Error GoTo 0
End Sub

' Nhu Log, nhung KEM SO DONG/CHI TIET LOI (Err.Number/Err.Description) â€” dung trong cac nhanh
' ErrHandler de ghi ca ma loi VBA (vi du 438, 5825...), giup phan biet cac loai "crash" khac nhau
' ma khong can nguoi dung tu doc/go MsgBox lai cho dung tung chu.
Public Sub LogErr(ByVal source As String, ByVal context As String, ByVal errNumber As Long, ByVal errDescription As String)
    Log source, context & " - LOI #" & CStr(errNumber) & ": " & errDescription
    ' Xa bo dem NGAY: mot loi VBA that su la dau hieu manh cho thay phien lam viec co the ket thuc
    ' dot ngot ngay sau do - phan nhat ky dan toi no phai nam tren dia truoc luc ay.
    Flush
End Sub

' Dump toan bo LayoutMap (ParagraphIndex -> Role) doi chieu voi noi dung doan tai thoi diem
' snapshot duoc chup - dung trong ComponentDetector.DetectComponents de chan doan "nhan nham vai
' tro" (vi du legalBasis/recipientLabel/recipientList). Dump TAT CA doan, khong chi rieng doan da
' gan vai tro - doan chua gan hien "(chua gan)" de phan biet "khong nhan dien duoc" voi "khong ton
' tai trong tai lieu".
' GOM toan bo cac dong vao MOT chuoi buf, chi goi Log MOT LAN cho ca dump (thay vi mo/dong file
' trong vong lap For Each p) - DetectComponents co the bi goi nhieu lan trong mot luot "Kiem tra",
' mo/dong file trong vong lap tung la nguyen nhan cua mot loi hieu nang "Not Responding" o mot dot
' sua truoc. Mat timestamp rieng cho tung dong (dung chung mot timestamp cho ca khoi) - danh doi
' chap nhan duoc cho tinh nang go loi tam thoi.
Public Sub LogLayoutMap(ByVal source As String, ByVal paragraphs As Collection, ByVal layoutMap As Object)
    If Not ENABLED Then Exit Sub
    On Error Resume Next
    Dim buf As String
    buf = "--- LayoutMap: " & CStr(layoutMap.count) & "/" & CStr(paragraphs.count) & _
        " doan da gan vai tro (danh sach DUOI day la TOAN BO doan, ca chua gan) ---"
    Dim p As ParagraphSnapshot
    Dim role As String
    Dim preview As String
    For Each p In paragraphs
        If layoutMap.Exists(p.Index) Then
            role = CStr(layoutMap(p.Index))
        Else
            role = "(chua gan)"
        End If
        preview = p.text
        If Len(preview) > 70 Then preview = left$(preview, 70) & "..."
        Dim boldText As String
        If IsNull(p.bold) Then boldText = "MIXED" Else boldText = CStr(p.bold)
        buf = buf & vbCrLf & "  #" & CStr(p.Index) & " [" & role & "] size=" & CStr(p.FontSizePt) & _
            " style=" & p.styleName & " bold=" & boldText & " allCaps=" & CStr(p.AllCaps) & _
            " align=" & p.alignment & " """ & preview & """"
    Next p
    Log source, buf
    On Error GoTo 0
End Sub
