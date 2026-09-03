Attribute VB_Name = "SessionState"
'==============================================================
' "Bo tat ca ChuanHoaTheThuc_LoaiVanBanIndex
' ChuanHoaTheThuc_DaDocDuLieu, ChuanHoaTheThuc_CoBangMaCu da ghi an vao file... Toi muon moi lan
' mo add-in se coi nhu file moi hoan toan (phai doc lai tu dau, xac dinh loai van ban tu dau) Ly
' do: nguoi dung co the lay file cu roi viet thanh loai van ban khac, hoac viet cac du lieu moi/
' bang ma khac."
' Hai trang thai DocumentTypeState.bas (lua chon "Loai van ban") va DataReadState.bas (co "da Doc
' du lieu"/"co bang ma cu") TRUOC DAY luu qua Document.Variables â€” ghi THAT SU vao
' word/settings.xml cua file.docx, ton tai qua ca luc dong roi mo lai file. Nay chuyen sang luu
' THUAN TRONG BO NHO cua phien VBA hien tai (module nay) â€” KHONG con dong cham gi den OOXML cua
' tai lieu nua, dung dung yeu cau "khong ghi an du lieu cau hinh" (ngoai tru "con dau" duy nhat o
' DocumentSignature.bas, cung dot â€” xem module do).
' KHOA theo CStr(ObjPtr(doc)) â€” dia chi bo nho cua doi tuong Word.Document COM, ON DINH trong suot
' vong doi CUA MOT INSTANCE tai lieu dang mo (dung yeu cau "phan biet duoc nhieu tai lieu dang mo
' cung luc", moi ban tam quan trong xuyen suot /), nhung LUON KHAC nhau giua hai lan mo (dong roi
' mo lai CUNG mot file tao ra MOT INSTANCE COM MOI, ObjPtr moi) â€” dung y nghia "moi lan mo la file
' moi hoan toan" ma KHONG can code rieng de "reset" (ResetReadData/ DataReadState van giu de goi
' tu AppEventsHost.OnAnyDocumentOpen, nay chi con la mot lop bao hiem tuong minh, khong con la co
' che duy nhat dam bao "tai lieu moi thi trong").
' DANH DOI DA CHAP NHAN: ly thuyet Windows CO THE tai su dung dia chi bo nho sau khi mot Document
' COM bi giai phong, gay TRUNG ObjPtr giua hai tai lieu KHONG lien quan mo truoc/sau nhau trong
' cung phien Word â€” rui ro rat thap (can dung THOI DIEM giai phong bo nho) va hau qua NHE (hai gia
' tri luu o day chi la tien loi giao dien: drop-down "Loai van ban" hien san hay khong, nut "Kiem
' tra"/"Chuyen doi Unicode" mo hay mo â€” khong phai du lieu nghiep vu, nguoi dung tu sua duoc ngay
' o lan bam nut tiep theo). KHONG don du lieu khi mot tai lieu dong lai (them su kien
' DocumentBeforeClose chi cho viec nay se tang dien tich be mat cho crash COM ribbon â€” da tra gia
' dat qua nhieu lan... â€” trong khi vai chuc chuoi Key ton dong qua ca phien Word la khong dang
' ke).
'==============================================================
' nhat ky dung dung tai DataReadState.HasNonUnicodeEncoding -> SessionState.GetValue, ngay giua
' mot loat callback getEnabled cua ribbon phat ra luc mo tai lieu (dung thoi diem nhay cam da gay
' toan bo chuoi crash COM 0xc0000005 cua du an nay,...) â€” chi la lan nay lo qua MOT CUA MOI:
' DocKey doc Application.ActiveDocument tren GAN NHU MOI LAN goi getEnabled (rat nhieu lan lien
' tiep trong vai phan nghin giay ngay luc khoi dong), tang dien tich be mat cham vao trang thai
' COM cua Application dung luc no chua on dinh. Ap dung LAI cung mot ky thuat "vet dau chay" da
' chung minh hieu qua o RibbonHandle.bas/WinApiFormStyle.bas: neu buoc doc ActiveDocument giet
' Word MOT LAN, phien SAU se tu vo hieu hoa dung buoc do (tra ve khoa rong, coi nhu chua tung ghi
' gi â€” dung y nghia "khong the phan biet duoc tai lieu, an toan hon la chet"), khong con phai doan
' nua. File guard TACH RIENG (khong dung chung voi RibbonHandle).
'==============================================================
Option Explicit

Private mStore As Object ' Scripting.Dictionary, tao lazy â€” khong ton tai ngoai phien VBA nay.

Private Const GUARD_FILE_NAME As String = "ChuanHoaTheThuc-sessionstate-guard.tmp"
Private mGuardReady As Boolean
Private mBlockedLabel As String

Private Function StoreObj() As Object
    If mStore Is Nothing Then Set mStore = CreateObject("Scripting.Dictionary")
    Set StoreObj = mStore
End Function

' --- Vet dau chay (cung khuon RibbonHandle.bas/WinApiFormStyle.bas)

Private Function GuardFilePath() As String
    On Error Resume Next
    Dim tempDir As String
    tempDir = Environ$("TEMP")
    If Len(tempDir) = 0 Then tempDir = Environ$("TMP")
    If Len(tempDir) = 0 Then Exit Function
    GuardFilePath = tempDir & "\" & GUARD_FILE_NAME
End Function

Private Sub LoadGuardState()
    If mGuardReady Then Exit Sub
    mGuardReady = True
    On Error Resume Next
    Dim p As String: p = GuardFilePath()
    If Len(p) = 0 Then Exit Sub
    If Len(Dir$(p)) = 0 Then Exit Sub
    Dim fnum As Integer, lineText As String
    fnum = FreeFile
    Open p For Input As #fnum
    If Not EOF(fnum) Then Line Input #fnum, lineText
    Close #fnum
    Kill p
    mBlockedLabel = Trim$(lineText)
    If Len(mBlockedLabel) = 0 Then Exit Sub
    DebugTrace.Log "SessionState.LoadGuardState", "PHIEN TRUOC Word chet ngay trong buoc doc ActiveDocument """ & _
        mBlockedLabel & """ - VO HIEU HOA diem doc nay cho ca phien nay"
    DebugTrace.Flush
End Sub

Private Function CanCall(ByVal callLabel As String) As Boolean
    LoadGuardState
    If Len(mBlockedLabel) > 0 Then
        If left$(mBlockedLabel, Len(callLabel)) = callLabel Then
            DebugTrace.Log "SessionState.CanCall", callLabel & " - BO QUA (da giet Word o phien truoc: " & mBlockedLabel & ")"
            Exit Function
        End If
    End If
    CanCall = True
End Function

Private Sub ArmGuard(ByVal callLabel As String)
    On Error Resume Next
    Dim p As String: p = GuardFilePath()
    If Len(p) = 0 Then Exit Sub
    Dim fnum As Integer: fnum = FreeFile
    Open p For Output As #fnum
    Print #fnum, callLabel
    Close #fnum
End Sub

Private Sub DisarmGuard(ByVal callLabel As String)
    On Error Resume Next
    Dim p As String: p = GuardFilePath()
    If Len(p) > 0 Then
        If Len(Dir$(p)) > 0 Then Kill p
    End If
End Sub

' Tra "" (khoa rong, dung chung cho moi tai lieu khong phan biet duoc) neu buoc doc ActiveDocument
' da bi vo hieu hoa o phien truoc, HOAC neu chinh Application khong co tai lieu nao hop le luc nay
' (On Error bat duoc loi VBA thuong, khong bat duoc access violation â€” day la ly do can them lop
' vet dau chay ben ngoai On Error).
Private Function DocKey(ByVal varName As String) As String
    Const LABEL As String = "DocKey.ActiveDocument"
    If Not CanCall(LABEL) Then
        DocKey = "::" & varName
        Exit Function
    End If

    On Error GoTo ErrHandler
    ArmGuard LABEL
    Dim ptr As Variant: ptr = ObjPtr(ActiveDocument) ' Variant: bien tuc thi ca VBA6 (Long) lan VBA7 (LongPtr).
    DisarmGuard LABEL
    DocKey = CStr(ptr) & "::" & varName
    Exit Function
ErrHandler:
    DisarmGuard LABEL
    DocKey = "::" & varName
End Function

' Tra Empty neu chua tung ghi cho (ActiveDocument, varName) nay trong phien hien tai.
Public Function GetValue(ByVal varName As String) As Variant
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("SessionState.GetValue")
    On Error GoTo ErrHandler
    Dim key As String: key = DocKey(varName)
    If StoreObj().Exists(key) Then
        GetValue = StoreObj()(key)
    Else
        GetValue = Empty
    End If
    Exit Function
ErrHandler:
    ' Khong co ActiveDocument hop le (vi du dang dong tai lieu) - coi nhu chua tung ghi.
    GetValue = Empty
End Function

Public Sub SetValue(ByVal varName As String, ByVal value As Variant)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("SessionState.SetValue")
    On Error GoTo ErrHandler
    StoreObj()(DocKey(varName)) = value
    Exit Sub
ErrHandler:
    ' Nhu GetValue - khong lam gian doan luong goi vi mot buoc phu.
End Sub

Public Sub ClearValue(ByVal varName As String)
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("SessionState.ClearValue")
    On Error GoTo ErrHandler
    Dim key As String: key = DocKey(varName)
    If StoreObj().Exists(key) Then StoreObj().Remove key
    Exit Sub
ErrHandler:
    ' Nhu GetValue.
End Sub
