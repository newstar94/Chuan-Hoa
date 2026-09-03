Attribute VB_Name = "RibbonHandle"
'==============================================================
' RibbonHandle â€” CHU SO HUU DUY NHAT cua con tro IRibbonUI trong ca du an.
' VAN DE GOC (15 lan Word chet 0xc0000005 trong hai ngay 20-21/8/2026; bon dot sua deu truot:
' /23/24/30-31): Word co the HUY doi tuong IRibbonUI ma KHONG bao cho add-in va KHONG goi lai
' onLoad. Con tro cu tro vao vung nho da giai phong: "Is Nothing" van tra False, va "On Error
' Resume Next" KHONG bat duoc (access violation xay ra o tang COM cua Word, khong sinh loi VBA
' nao). Moi lan cham vao con tro do la mot lan Word co the chet ngay lap tuc.
' KHONG CO cach nao kiem tra mot con tro COM con song hay khong ma khong cham vao no. Do chinh la
' ly do bon dot sua truoc that bai: ca bon deu co gang DOAN thoi diem con tro chet (hoan lenh qua
' Application.OnTime, co chan de quy, so dang ky theo cua so, bo mot nhanh goi vo dieu kien) thay
' vi chap nhan rang khong the doan duoc. Module nay khong doan nua.
' BA LOP KIEM SOAT:
' 1. MOT CUA DUY NHAT. Chi file nay duoc khai bao IRibbonUI va goi Invalidate/InvalidateControl.
'   build/check-invariants.ps1 la cong kiem tra luc BUILD: bat ky file.bas/.cls/.frm nao khac cham
'   vao hai thu do se lam build DUNG LAI kem ten file va so dong. Khong con duong lot ra.
' 2. KHONG BAO GIO Release mot con tro nghi ngo. "Set mRibbon = Nothing" - VA ca phep gan de len
'   con tro cu trong "Set mRibbon = ribbon" o lan onLoad THU HAI - deu tu goi IUnknown::Release,
'   tuc cung cham vao vtable co the da chet. Day la mot duong crash chua tung duoc de y, du nhat
'   ky da ghi nhan onLoad chay 3 lan trong MOT phien. AbandonWithoutRelease xoa trang bien doi
'   tuong bang CopyMemory, khong Release: buong con tro voi rui ro bang 0.
' 3. VET DAU CHAY (breadcrumb) - lop cuoi, bat duoc CA truong hop chua ai luong truoc. Ngay TRUOC
'   moi lan cham vao COM: ghi mot file danh dau kem ten diem goi. Ngay SAU khi lenh tra ve: xoa
'   file. Neu Word chet giua chung, file con lai tren dia. Lan khoi dong SAU, Capture doc thay va
'   VO HIEU HOA DUNG diem goi da giet Word o phien truoc. Nghia la: mot kich ban stale-pointer
'   moi, chua tung biet, chi co the lam Word chet DUNG MOT LAN - lan thu hai no da bi chan san.
'   Day la co che kiem soat KHONG phu thuoc vao viec doan dung nguyen nhan.
' Khi mot diem goi bi vo hieu hoa (hoac chua co con tro), ribbon khong duoc ve lai ngay; cac nut
' cap nhat cham mot nhip theo dot quet getEnabled ma chinh Word tu phat rat thuong xuyen (nhat ky
' cho thay ~31 callback moi lan kich hoat cua so). Cham mot nhip van hon han mot lan Word chet.
'==============================================================
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByVal Destination As LongPtr, ByVal source As LongPtr, ByVal length As LongPtr)
#Else
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
#End If

Private Const GUARD_FILE_NAME As String = "ChuanHoaTheThuc-ribbon-guard.tmp"

' Con tro thuc su. KHONG duoc cham vao ngoai TouchRibbon; khong duoc gan Nothing (xem lop 2).
Private mRibbon As IRibbonUI
Private mHaveRibbon As Boolean

' Chan de quy do CHINH TA gay ra: True dung luc dang o giua mot loi goi Invalidate.
Private mInFlight As Boolean

' Do sau "dang o TRONG mot callback ribbon do CHINH WORD phat ra" - khac mInFlight (chi bat duoc
' de quy do ta gay ra). mCallbackDepthAt la moc thoi gian tu phuc hoi neu bo dem lo bi ket khac 0
' (bai hoc: mot co ket True da lam ribbon khong bao gio ve lai duoc trong ca phien).
Private mCallbackDepth As Long
Private mCallbackDepthAt As Single

' Trang thai doc tu vet dau chay cua phien truoc (lop 3).
Private mGuardReady As Boolean
Private mBlockedLabel As String

' --- Vong doi con tro

' customUI14.xml onLoad="RibbonOnLoad". DAT O DAY (khong phai RibbonCallbacks.bas) de kieu
' IRibbonUI chi xuat hien DUY NHAT trong file nay - nho vay cong kiem tra luc build (check-
' invariants.ps1) khong can bat ky ngoai le nao.
' onLoad phai TRA VE NHANH (khuyen cao cua Microsoft): moi viec cua add-in deu nam ben
' RibbonCallbacks.OnRibbonLoaded, va phan nang trong do lai duoc day tiep ra Application.OnTime.
Public Sub RibbonOnLoad(ByVal ribbon As IRibbonUI)
    ' Xoa nhat ky go loi cua phien truoc TRUOC bat ky Log/EnterScope nao.
    DebugTrace.ResetLogAtStartup
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RibbonHandle.RibbonOnLoad")

    Capture ribbon
    RibbonCallbacks.OnRibbonLoaded
End Sub

' Buong con tro cu TRUOC khi nhan con tro moi - khong duoc de phep gan tu Release con tro cu (xem
' lop 2 o dau file).
Private Sub Capture(ByVal ribbon As IRibbonUI)
    LoadGuardState

    AbandonWithoutRelease

    On Error Resume Next
    Set mRibbon = ribbon
    mHaveRibbon = (Err.number = 0)
    If Not mHaveRibbon Then Err.Clear
    On Error GoTo 0

    DebugTrace.Log "RibbonHandle.Capture", "Da nhan IRibbonUI moi, mHaveRibbon=" & CStr(mHaveRibbon)
End Sub

' Xoa trang bien doi tuong MA KHONG goi Release. Neu CopyMemory that bai vi bat ky ly do gi thi
' mHaveRibbon = False van la lop chan chinh: khong noi nao con cham vao con tro nua.
Private Sub AbandonWithoutRelease()
    On Error Resume Next
    If Not mHaveRibbon Then Exit Sub
    mHaveRibbon = False

    #If VBA7 Then
        Dim zeroBuf As LongPtr
    #Else
        Dim zeroBuf As Long
    #End If
    zeroBuf = 0
    CopyMemory VarPtr(mRibbon), VarPtr(zeroBuf), LenB(zeroBuf)
    DebugTrace.Log "RibbonHandle.AbandonWithoutRelease", "Da buong con tro ribbon (khong Release)"
End Sub

' --- Diem goi cong khai

Public Sub RequestInvalidateAll(ByVal callerLabel As String)
    TouchRibbon callerLabel, ""
End Sub

Public Sub RequestInvalidateControl(ByVal controlId As String)
    TouchRibbon "InvalidateControl:" & controlId, controlId
End Sub

' NOI DUY NHAT trong ca du an duoc phep cham vao mRibbon.Invalidate/InvalidateControl.
Private Sub TouchRibbon(ByVal callerLabel As String, ByVal controlId As String)
    If Not CanTouch(callerLabel) Then Exit Sub

    mInFlight = True
    On Error Resume Next

    ' Vet dau chay + xa nhat ky NGAY TRUOC loi goi: neu Word chet o tang COM (khong sinh loi VBA),
    ' ca hai deu con lai tren dia va khoanh dung thu pham.
    ArmGuard callerLabel
    DebugTrace.Log "RibbonHandle.TouchRibbon", callerLabel & " - truoc khi goi COM"
    DebugTrace.Flush

    If Len(controlId) = 0 Then
        mRibbon.Invalidate
    Else
        mRibbon.InvalidateControl controlId
    End If

    If Err.number <> 0 Then
        DebugTrace.LogErr "RibbonHandle.TouchRibbon", callerLabel & " - loi VBA, buong con tro", _
            Err.number, Err.description
        Err.Clear
        AbandonWithoutRelease
    Else
        DebugTrace.Log "RibbonHandle.TouchRibbon", callerLabel & " - xong"
    End If

    DisarmGuard
    ' Chan doan crash 6h12 27/8/2026 (xem CLAUDE.md): vet dau chay cho thay Invalidate lan do THUC
    ' SU tra ve thanh cong (file guard bi xoa dung) - Word chet sau do, trong luc chinh NO tu xu
    ' ly ket qua Invalidate (dot callback getEnabled/getVisible tu phat), khong con nam trong
    ' TouchRibbon nua. Truoc day khong Flush o day nen moi dong log sau "xong" van con ket trong
    ' bo dem, mat sach khi Word chet - xa NGAY de lan sau (neu con) thay duoc buoc ke tiep THAT SU
    ' chay la gi.
    DebugTrace.Flush
    On Error GoTo 0
    mInFlight = False
End Sub

Private Function CanTouch(ByVal callerLabel As String) As Boolean
    LoadGuardState

    If mInFlight Then
        DebugTrace.Log "RibbonHandle.CanTouch", callerLabel & " - BO QUA (mot lan Invalidate khac dang chay)"
        Exit Function
    End If
    If IsInsideCallback() Then
        ' Word dang tu phat mot vong callback getEnabled/getVisible va ta dang o trong vong do -
        ' bo qua, vong dang chay von da hoi lai toan bo trang thai roi.
        DebugTrace.Log "RibbonHandle.CanTouch", callerLabel & " - BO QUA (dang o trong callback cua Word)"
        Exit Function
    End If
    If Not mHaveRibbon Then
        DebugTrace.Log "RibbonHandle.CanTouch", callerLabel & " - BO QUA (chua co IRibbonUI, cho onLoad)"
        Exit Function
    End If
    If Len(mBlockedLabel) > 0 Then
        If StrComp(mBlockedLabel, callerLabel, vbTextCompare) = 0 Then
            DebugTrace.Log "RibbonHandle.CanTouch", callerLabel & " - BO QUA (diem goi nay da giet Word o phien truoc)"
            Exit Function
        End If
    End If

    CanTouch = True
End Function

' --- Cong "dang o trong callback ribbon cua Word" --------------------------------------------
' Dat o DAU va CUOI moi callback getEnabled/getVisible/getSelectedItemIndex. Than cac callback do
' deu duoc boc "On Error Resume Next" nen EndCallback CHAC CHAN chay.

Public Sub BeginCallback()
    On Error Resume Next
    mCallbackDepth = mCallbackDepth + 1
    mCallbackDepthAt = Timer
End Sub

Public Sub EndCallback()
    On Error Resume Next
    mCallbackDepth = mCallbackDepth - 1
    If mCallbackDepth < 0 Then mCallbackDepth = 0
End Sub

' Chot tu phuc hoi: neu mot duong thoat bat thuong lam bo dem khong ve duoc 0 thi sau 3 giay coi
' nhu da het vong callback (Word phat het ~31 callback trong vai phan nghin giay).
Private Function IsInsideCallback() As Boolean
    If mCallbackDepth <= 0 Then Exit Function
    Dim elapsed As Single
    elapsed = Timer - mCallbackDepthAt
    If elapsed < 0 Or elapsed > 3! Then
        mCallbackDepth = 0
        Exit Function
    End If
    IsInsideCallback = True
End Function

' --- Vet dau chay (lop 3)

Private Function GuardFilePath() As String
    On Error Resume Next
    Dim tempDir As String
    tempDir = Environ$("TEMP")
    If Len(tempDir) = 0 Then tempDir = Environ$("TMP")
    If Len(tempDir) = 0 Then Exit Function
    GuardFilePath = tempDir & "\" & GUARD_FILE_NAME
End Function

Private Sub ArmGuard(ByVal callerLabel As String)
    On Error Resume Next
    Dim p As String
    p = GuardFilePath()
    If Len(p) = 0 Then Exit Sub

    Dim fnum As Integer
    fnum = FreeFile
    Open p For Output As #fnum
    Print #fnum, callerLabel
    Close #fnum
End Sub

Private Sub DisarmGuard()
    On Error Resume Next
    Dim p As String
    p = GuardFilePath()
    If Len(p) = 0 Then Exit Sub
    If Len(Dir$(p)) > 0 Then Kill p
End Sub

' Doc MOT LAN cho ca phien, tu Capture (tuc sau DebugTrace.ResetLogAtStartup, nen dong log nay
' khong bi xoa mat).
Private Sub LoadGuardState()
    If mGuardReady Then Exit Sub
    mGuardReady = True

    On Error Resume Next
    Dim p As String
    p = GuardFilePath()
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

    DebugTrace.Log "RibbonHandle.LoadGuardState", "PHIEN TRUOC Word chet ngay trong loi goi ribbon """ & _
        mBlockedLabel & """ - VO HIEU HOA diem goi nay cho ca phien nay"
    DebugTrace.Flush
End Sub
