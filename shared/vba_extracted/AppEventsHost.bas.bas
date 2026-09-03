Attribute VB_Name = "AppEventsHost"
'==============================================================
' AppEventsHost â€” noi GIU THAM CHIEU SONG toi AppEvents.cls (xem chu thich dau file do) va la
' CAU NOI ma su kien cua Class Module goi duoc thang (KHONG can Application.Run - han che
' "TenForm.TenSub" chi anh huong UserForm, xem RibbonCallbacks.NotifyScanProgress, khong anh huong
' loi goi truc tiep tu Class Module nhu o day).
' nhan dien loai van ban NGAY khi mo/tao tai lieu, cap nhat drop-down "Loai van ban" (nhom Khoi
' dong) truoc khi nguoi dung bam "Kiem tra" lan dau - nut "Kiem tra" bi mo cho toi khi co mot loai
' van ban duoc chon (RibbonCallbacks.GetEnabledKiemTra doc DocumentTypeState.GetSelectedIndex).
'==============================================================
Option Explicit

Private mSink As AppEvents

' Cache cho AppEventsHost.OnSelectionChange/OnWindowActivate (item 2) - PHAI khai bao O DAY (dau
' file, TRUOC toan bo Sub/Function - quy tac VBA, xem ghi chu dau Utils.bas). Empty = chua biet;
' True/False = trang thai "rong hay khong" tai lan kiem gan nhat.
Private mLastKnownEmpty As Variant

' Goi tu RibbonCallbacks.RibbonOnLoad (customUI14.xml onLoad) - thoi diem add-in chac chan da nap
' xong. Goi lai nhieu lan (vi du Word nap lai ribbon) khong sao - chi tao instance moi neu chua
' co, tranh dang ky trung su kien nhieu lan.
Public Sub EnsureWired()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("AppEventsHost.EnsureWired")
    If mSink Is Nothing Then
        Set mSink = New AppEvents
        Set mSink.oApp = Application
    End If
End Sub

' Dang ky chay OnAnyDocumentOpen qua Application.OnTime NGAY KHI Word ranh (thay vi chay THANG
' trong su kien nhay cam DocumentOpen/NewDocument - xem canh bao dau AppEvents.cls).
' Application.OnTime CHI goi duoc thu tuc trong STANDARD MODULE qua ten chuoi (dung quy uoc
' "TenModule.TenSub", GIONG han che "Application.Run" da ghi o RibbonCallbacks.
' NotifyScanProgress) - AppEventsHost.OnAnyDocumentOpen thoa dieu kien nay.
Public Sub ScheduleOnAnyDocumentOpen()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("AppEventsHost.ScheduleOnAnyDocumentOpen")
    On Error GoTo ErrHandler
    DebugTrace.Log "AppEventsHost.ScheduleOnAnyDocumentOpen", "Dang ky OnTime cho OnAnyDocumentOpen"
    Application.OnTime Now, "AppEventsHost.OnAnyDocumentOpen"
    Exit Sub
ErrHandler:
    DebugTrace.LogErr "AppEventsHost.ScheduleOnAnyDocumentOpen", "Khong dang ky duoc OnTime", Err.number, Err.description
End Sub

' Chay khi mo BAT KY tai lieu nao (Document_Open) hoac tao moi (Application_NewDocument), VA mot
' lan thu cong tu RibbonCallbacks.RibbonOnLoad cho tai lieu DANG MO SAN luc add-in nap xong (thu
' tu nap khong dam bao Document_Open cua tai lieu dau tien xay ra SAU khi ribbon da nap). Ca hai
' noi goi di qua ScheduleOnAnyDocumentOpen (Application.OnTime), khong goi thang ham nay tu su
' kien nhay cam.
' Khong tu doc du lieu/nhan dien gi khi mo file ("dinh dang"/"trong hay khong" la phep tinh thuan
' tuy doc-tai-cho trong tung getEnabled/getVisible cua ribbon) - ham nay chi con hai viec: (1)
' buoc lai trang thai "da Doc du lieu chua" ve 0 (moi lan mo file phai Doc du lieu lai tu dau -
' khac DocumentTypeState van giu lua chon cu qua cac lan mo), (2) ve lai ribbon de moi
' getEnabled/getVisible duoc doc lai dung trang thai tai lieu vua mo/rong.
Public Sub OnAnyDocumentOpen()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("AppEventsHost.OnAnyDocumentOpen")
    On Error GoTo ErrHandler
    DebugTrace.Log "AppEventsHost.OnAnyDocumentOpen", "Bat dau"
    ' Chot an toan: day la thu tuc DUY NHAT con lai chay qua Application.OnTime. Neu vi ly do nao
    ' do no no ra dung luc Word dang dong het tai lieu thi thoat ngay, khong cham vao gi ca.
    If Application.Documents.count = 0 Then
        DebugTrace.Log "AppEventsHost.OnAnyDocumentOpen", "Khong con tai lieu nao mo - thoat som"
        Exit Sub
    End If
    If ActiveDocument Is Nothing Then
        DebugTrace.Log "AppEventsHost.OnAnyDocumentOpen", "ActiveDocument Is Nothing - thoat som"
        Exit Sub
    End If

    DebugTrace.Log "AppEventsHost.OnAnyDocumentOpen", "Buoc 1/2: DataReadState.ResetReadData"
    DataReadState.ResetReadData
    DebugTrace.Log "AppEventsHost.OnAnyDocumentOpen", "Buoc 1/2 xong"

    DebugTrace.Log "AppEventsHost.OnAnyDocumentOpen", "Buoc 2/2: RibbonCallbacks.InvalidateRibbon"
    RibbonCallbacks.InvalidateRibbon "AnyDocumentOpen"
    DebugTrace.Log "AppEventsHost.OnAnyDocumentOpen", "Hoan tat"
    DebugTrace.Flush
    Exit Sub
ErrHandler:
    ' Tinh nang phu, chay AM THAM luc mo tai lieu (nguoi dung chua chu dong yeu cau gi) - KHONG
    ' hien MsgBox (trai truc giac ngay luc vua mo tai lieu), nhung PHAI ghi log (SUA - truoc day
    ' im lang hoan toan, khong co dau vet chan doan).
    DebugTrace.LogErr "AppEventsHost.OnAnyDocumentOpen", "Loi giua chung", Err.number, Err.description
End Sub

' Goi tu AppEvents.cls's oApp_WindowSelectionChange - bon nut Luu thanh DOCX/Doc du lieu/ Chuyen
' doi Unicode/Kiem tra can bat ngay khi tai lieu bat dau co noi dung
' (RibbonCallbacks.GetVisibleRequiresNonEmptyDoc), nhung Ribbon khong tu ve lai getVisible khi noi
' dung doi - can mot diem kich hoat dong. WindowSelectionChange bat theo con tro (gan nhu moi phim
' go, khong lam viec COM nang) nhung chi thuc su goi InvalidateRibbon khi trang thai rong/co-noi-
' dung that su doi khac lan kiem gan nhat - tranh ve lai ribbon vo ich moi lan go.
Public Sub OnSelectionChange()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("AppEventsHost.OnSelectionChange")
    On Error Resume Next ' su kien tan suat cao - KHONG duoc lam gian doan thao tac go van ban
    If ActiveDocument Is Nothing Then Exit Sub

    Dim nowEmpty As Boolean
    nowEmpty = RibbonCallbacks.IsDocEmpty()

    If IsEmpty(mLastKnownEmpty) Then
        mLastKnownEmpty = nowEmpty
        Exit Sub ' lan dau kiem tra trong phien nay - chi ghi nhan, InvalidateRibbon da co san
                 ' o OnAnyDocumentOpen luc mo tai lieu, khong can ve lai them lan nua o day.
    End If

    If CBool(mLastKnownEmpty) <> nowEmpty Then
        mLastKnownEmpty = nowEmpty
        RibbonCallbacks.InvalidateRibbon "SelectionChange"
    End If
    ' Xa nhat ky ngay - cung ly do da them o OnWindowActivate va o day khi chan doan crash 6h12
    ' 27/8/2026: nhanh nay it khi thuc su goi InvalidateRibbon (chi khi trang thai rong/co-noi-
    ' dung THAT SU doi) nen xa o day khong ton hieu nang dang ke.
    DebugTrace.Flush
End Sub

Public Sub OnWindowActivate()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("AppEventsHost.OnWindowActivate")
    On Error Resume Next
    If ActiveDocument Is Nothing Then Exit Sub

    Dim nowEmpty As Boolean
    nowEmpty = RibbonCallbacks.IsDocEmpty()

    If IsEmpty(mLastKnownEmpty) Then
        mLastKnownEmpty = nowEmpty
        Exit Sub ' lan dau kiem tra trong phien nay - CHI ghi nhan, KHONG ep ve lai.
    End If

    If CBool(mLastKnownEmpty) <> nowEmpty Then
        mLastKnownEmpty = nowEmpty
        RibbonCallbacks.InvalidateRibbon "WindowActivate"
    End If
    DebugTrace.Flush
End Sub

' Goi tu AppEvents.cls's oApp_DocumentBeforeClose khi tai lieu SAP dong la tai lieu CUOI CUNG. day
' la ban thay the cho cap ScheduleOnWindowDeactivate/ OnDocumentCountBecameZero cu
' (Application.OnTime) - nguon crash 0xc0000005 ngay 25/8/2026.
' CO Y KHONG lam gi ngoai buoc lai cache: khong Invalidate, khong cham COM, khong timer. Thoi diem
' nay Word co the dang tu thao do chinh minh (Application.Quit cung di qua day, da do bang thuc
' nghiem) - moi thao tac COM o day deu la rui ro khong can thiet.
Public Sub OnLastDocumentClosing()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("AppEventsHost.OnLastDocumentClosing")
    On Error Resume Next
    mLastKnownEmpty = True
    ' Xa nhat ky NGAY: neu Word chet/thoat ngay sau day thi phan nhat ky dan toi no phai nam tren
    ' dia. Hai lan chan doan 25/8/2026 deu mat sach vai dong cuoi vi bo dem chua kip xa.
    DebugTrace.Flush
End Sub
