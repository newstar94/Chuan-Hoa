Attribute VB_Name = "CheckGate"
Option Explicit

Private mTextsReady As Boolean
Private TITLE_BEFORE_CHECK As String
Private MSG_DOCX_REQUIRED As String
Private ACTION_SAVE_AS_DOCX As String
Private MSG_UNICODE_REQUIRED_PREFIX As String
Private MSG_UNICODE_REQUIRED_MID As String
Private MSG_UNICODE_REQUIRED_SUFFIX As String
Private ACTION_CONVERT_UNICODE As String
Private REASON_DOCX_BYPASSED As String
Private REASON_UNICODE_BYPASSED As String
Private TITLE_ENCODING As String
Private LABEL_MIXED As String

Private Sub EnsureTexts()
    If mTextsReady Then Exit Sub

    ' "Truoc khi kiem tra" â€” vi.dialogs.gate.beforeCheckTitle, dung chung cho ca P2 va P3.
    TITLE_BEFORE_CHECK = "Tr" & ChrW(&H1B0) & ChrW(&H1EDB) & "c khi ki" & ChrW(&H1EC3) & "m tra"

    ' "Can luu thanh.docx truoc. Dung nut Luu thanh DOCX." â€” vi.dialogs.gate.docxRequiredMessage.
    MSG_DOCX_REQUIRED = "C" & ChrW(&H1EA7) & "n l" & ChrW(&H1B0) & "u th" & ChrW(&HE0) & "nh .docx tr"
    MSG_DOCX_REQUIRED = MSG_DOCX_REQUIRED & ChrW(&H1B0) & ChrW(&H1EDB) & "c. D" & ChrW(&HF9) & "ng n"
    MSG_DOCX_REQUIRED = MSG_DOCX_REQUIRED & ChrW(&HFA) & "t L" & ChrW(&H1B0) & "u th" & ChrW(&HE0) & "nh DOCX."

    ' "Luu thanh DOCX" â€” vi.dialogs.gate.docxRequiredActionLabel.
    ACTION_SAVE_AS_DOCX = "L" & ChrW(&H1B0) & "u th" & ChrW(&HE0) & "nh DOCX"

    ' Ba manh ghep "Phat hien {n} doan dung bang ma {label}. Dung nut Chuyen doi Unicode." â€”
    ' vi.dialogs.gate.unicodeRequiredMessage(count, encodingLabel).
    MSG_UNICODE_REQUIRED_PREFIX = "Ph" & ChrW(&HE1) & "t hi" & ChrW(&H1EC7) & "n "
    MSG_UNICODE_REQUIRED_MID = " " & ChrW(&H111) & "o" & ChrW(&H1EA1) & "n d" & ChrW(&HF9) & "ng b"
    MSG_UNICODE_REQUIRED_MID = MSG_UNICODE_REQUIRED_MID & ChrW(&H1EA3) & "ng m" & ChrW(&HE3) & " "
    MSG_UNICODE_REQUIRED_SUFFIX = ". D" & ChrW(&HF9) & "ng n" & ChrW(&HFA) & "t Chuy" & ChrW(&H1EC3) & "n "
    MSG_UNICODE_REQUIRED_SUFFIX = MSG_UNICODE_REQUIRED_SUFFIX & ChrW(&H111) & ChrW(&H1ED5) & "i Unicode."

    ' "Chuyen doi Unicode" â€” vi.dialogs.gate.unicodeRequiredActionLabel.
    ACTION_CONVERT_UNICODE = "Chuy" & ChrW(&H1EC3) & "n " & ChrW(&H111) & ChrW(&H1ED5) & "i Unicode"

    ' "Chua luu thanh.docx â€” bam "Van kiem tra" da bo qua buoc nay." â€”
    ' vi.checklist.gateBlockedReasons.docxBypassed.
    REASON_DOCX_BYPASSED = "Ch" & ChrW(&H1B0) & "a l" & ChrW(&H1B0) & "u th" & ChrW(&HE0) & "nh .docx "
    REASON_DOCX_BYPASSED = REASON_DOCX_BYPASSED & ChrW(&H2014) & " b" & ChrW(&H1EA5) & "m "
    REASON_DOCX_BYPASSED = REASON_DOCX_BYPASSED & ChrW(&H201C) & "V" & ChrW(&H1EAB) & "n ki"
    REASON_DOCX_BYPASSED = REASON_DOCX_BYPASSED & ChrW(&H1EC3) & "m tra" & ChrW(&H201D) & " " & ChrW(&H111)
    REASON_DOCX_BYPASSED = REASON_DOCX_BYPASSED & ChrW(&HE3) & " b" & ChrW(&H1ECF) & " qua b" & ChrW(&H1B0)
    REASON_DOCX_BYPASSED = REASON_DOCX_BYPASSED & ChrW(&H1EDB) & "c n" & ChrW(&HE0) & "y."

    ' "Van ban con doan dung bang ma cu â€” bam "Van kiem tra" da bo qua buoc chuyen Unicode." â€”
    ' vi.checklist.gateBlockedReasons.unicodeBypassed.
    REASON_UNICODE_BYPASSED = "V" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n c" & ChrW(&HF2) & "n "
    REASON_UNICODE_BYPASSED = REASON_UNICODE_BYPASSED & ChrW(&H111) & "o" & ChrW(&H1EA1) & "n d" & ChrW(&HF9)
    REASON_UNICODE_BYPASSED = REASON_UNICODE_BYPASSED & "ng b" & ChrW(&H1EA3) & "ng m" & ChrW(&HE3) & " c"
    REASON_UNICODE_BYPASSED = REASON_UNICODE_BYPASSED & ChrW(&H169) & " " & ChrW(&H2014) & " b" & ChrW(&H1EA5)
    REASON_UNICODE_BYPASSED = REASON_UNICODE_BYPASSED & "m " & ChrW(&H201C) & "V" & ChrW(&H1EAB) & "n ki"
    REASON_UNICODE_BYPASSED = REASON_UNICODE_BYPASSED & ChrW(&H1EC3) & "m tra" & ChrW(&H201D) & " " & ChrW(&H111)
    REASON_UNICODE_BYPASSED = REASON_UNICODE_BYPASSED & ChrW(&HE3) & " b" & ChrW(&H1ECF) & " qua b" & ChrW(&H1B0)
    REASON_UNICODE_BYPASSED = REASON_UNICODE_BYPASSED & ChrW(&H1EDB) & "c chuy" & ChrW(&H1EC3) & "n Unicode."

    ' "Chuyen bang ma sang Unicode" â€” vi.dialogs.warning.encodingConversionTitle, doi chieu
    ' RibbonCallbacks.TITLE_ENCODING (khong dung lai truc tiep vi bien do la Private cua module
    ' khac â€” VBA khong co "export" chon loc, phai dung lai cach dung TAI DAY).
    TITLE_ENCODING = "Chuy" & ChrW(&H1EC3) & "n b" & ChrW(&H1EA3) & "ng m" & ChrW(&HE3) & " sang Unicode"

    LABEL_MIXED = "TCVN3/VNI l" & ChrW(&H1EAB) & "n l" & ChrW(&H1ED9) & "n"

    mTextsReady = True
End Sub

' ============================================================================
' Diem vao duy nhat â€” FindingReporter.RunCheckAndReport goi khi bat dau vong "Kiem tra".
' ============================================================================

' Chay tuan tu cong (1) roi cong (2), DUNG THU TU. Tra False neu nguoi dung huy o bat ky dau
' (gateBlockedGroups khong duoc dung trong truong hop nay â€” noi goi phai giu nguyen trang thai
' cong chan). Tra True kem gateBlockedGroups (Dictionary groupId -> ly do hien thi, RONG neu
' khong cong nao bi bo qua) khi co the chay ComplianceChecker.RunAllChecks that su.
Public Function RunCheckGates(ByRef gateBlockedGroups As Object) As Boolean
    On Error GoTo ErrHandler
    EnsureTexts
    Set gateBlockedGroups = Utils.NewDictionary()

    Dim docxOutcome As String
    docxOutcome = RunDocxGate()
    If docxOutcome = "cancel" Then
        RunCheckGates = False
        Exit Function
    End If
    If docxOutcome = "bypassed" Then
        gateBlockedGroups("pageAndFont") = REASON_DOCX_BYPASSED
    End If

    Dim unicodeOutcome As String
    unicodeOutcome = RunUnicodeGate()
    If unicodeOutcome = "cancel" Then
        RunCheckGates = False
        Exit Function
    End If
    If unicodeOutcome = "bypassed" Then
        Dim groupIds As Variant
        groupIds = BlockedGroupIds()
        Dim i As Long
        For i = LBound(groupIds) To UBound(groupIds)
            gateBlockedGroups(groupIds(i)) = REASON_UNICODE_BYPASSED
        Next i
    End If

    RunCheckGates = True
    Exit Function
ErrHandler:
    Err.Raise Err.number, "CheckGate.RunCheckGates", Err.description
End Function

' ============================================================================
' Cong (1) â€”.docx
' ============================================================================

' "ok" = tiep tuc (da la.docx, hoac vua luu that thanh cong); "cancel" = huy toan bo Kiem tra
' (nguoi dung bam Huy, hoac luu that bai/bi huy o hop thoai trung ten cua DocxConverter);
' "bypassed" = "Van kiem tra".
Private Function RunDocxGate() As String
    If DocxConverter.IsCurrentDocumentDocx() Then
        RunDocxGate = "ok"
        Exit Function
    End If

    frmWarning.ShowGate TITLE_BEFORE_CHECK, MSG_DOCX_REQUIRED, ACTION_SAVE_AS_DOCX

    Select Case frmWarning.Result
        Case "action"
            ' Loi the rieng cua ban Legacy (xem dau file) Luu xong thi TIEP TUC luon, khong bat
            ' bam "Kiem tra" lan hai.
            If DocxConverter.SaveAsDocx() Then
                RunDocxGate = "ok"
            Else
                RunDocxGate = "cancel"
            End If
        Case "proceedAnyway"
            RunDocxGate = "bypassed"
        Case Else ' "cancel" hoac dong bang nut dieu khien he thong cua form
            RunDocxGate = "cancel"
    End Select
End Function

' ============================================================================
' Cong (2) â€” Unicode
' ============================================================================

Private Function RunUnicodeGate() As String
    Dim detection As Object
    Set detection = EncodingConverter.DetectEncoding()

    If CLng(detection("nonUnicodeCount")) = 0 Then
        RunUnicodeGate = "ok"
        Exit Function
    End If

    Dim affectedParagraphs As Long
    affectedParagraphs = CountAffectedParagraphs(detection("runs"))

    ' Ten bien PHAI khac ten ham EncodingLabel - VBA khong phan biet hoa/thuong nen trung ten se
    ' bi coi la goi lai chinh no ("Expected array" luc bien dich).
    Dim encodingText As String
    encodingText = EncodingLabel(CStr(detection("encoding")))

    Dim msg As String
    msg = MSG_UNICODE_REQUIRED_PREFIX & CStr(affectedParagraphs) & MSG_UNICODE_REQUIRED_MID & _
        encodingText & MSG_UNICODE_REQUIRED_SUFFIX

    frmWarning.ShowGate TITLE_BEFORE_CHECK, msg, ACTION_CONVERT_UNICODE

    Select Case frmWarning.Result
        Case "action"
            If ConfirmAndConvertUnicode() Then
                RunUnicodeGate = "ok"
            Else
                RunUnicodeGate = "cancel"
            End If
        Case "proceedAnyway"
            RunUnicodeGate = "bypassed"
        Case Else
            RunUnicodeGate = "cancel"
    End Select
End Function

Private Function ConfirmAndConvertUnicode() As Boolean
    Dim warning As Object
    Set warning = SafetyGuard.BuildWarning(SafetyGuard.HIGH_RISK_ENCODING_CONVERSION)

    frmWarning.ShowHighRisk TITLE_ENCODING, CStr(warning("whatWillHappen")), _
        CStr(warning("scope")), CStr(warning("undoability")), CStr(warning("saveReminder"))

    Select Case frmWarning.Result
        Case "saveAndRun"
            ActiveDocument.Save
        Case "runAnyway"
            ' Chay thang, khong luu.
        Case Else ' "cancel" hoac dong bang nut dieu khien he thong cua form
            ConfirmAndConvertUnicode = False
            Exit Function
    End Select

    Dim Result As Object
    Set Result = EncodingConverter.ConvertToUnicode()
    ConfirmAndConvertUnicode = Not CBool(Result("failed"))
End Function

' ============================================================================
' Tien ich
' ============================================================================

' So doan (paragraphIndex) co run KHONG PHAI Unicode â€” doan ngoai than bai (dau/chan trang, chu
' thich, paragraphIndex Null) gop chung thanh MOT don vi vi khong co chi so doan rieng de dem â€”
' xap xi hop ly cho mot con so hien thi, khong phai can cu phap ly.
Private Function CountAffectedParagraphs(ByVal runs As Object) As Long
    Dim seen As Object
    Set seen = Utils.NewDictionary()

    Dim r As Variant
    For Each r In runs
        If CStr(r("encoding")) <> "unicode" Then
            Dim key As String
            If IsNull(r("paragraphIndex")) Then
                key = "null"
            Else
                key = CStr(r("paragraphIndex"))
            End If
            If Not seen.Exists(key) Then seen(key) = True
        End If
    Next r

    CountAffectedParagraphs = seen.count
End Function

Private Function EncodingLabel(ByVal encoding As String) As String
    Select Case encoding
        Case "tcvn3": EncodingLabel = "TCVN3"
        Case "vni": EncodingLabel = "VNI"
        Case "mixed": EncodingLabel = LABEL_MIXED
        Case "unicode": EncodingLabel = "Unicode"
        Case Else: EncodingLabel = encoding
    End Select
End Function

' Moi nhom quy tac NGOAI "pageAndFont" phu thuoc doc dung noi dung chu (regex nhan dien thanh
' phan, viet hoa, chinh taâ€¦) â€” van ban con bang ma cu hien thi sai ky tu nen cac nhom nay khong
' dang tin neu bo qua cong Unicode. "pageAndFont" (kho giay/le/so do) khong phu thuoc noi dung chu
' nen van kiem duoc.
Private Function BlockedGroupIds() As Variant
    BlockedGroupIds = Array( _
        "nationalTitleAndMotto", "organName", "codeNumber", "placeAndDate", _
        "typeNameAndSubject", "bodyContent", "signer", "recipient", "markingsAndContact", _
        "appendix", "tableAndImage", "spellingConversion", "spellingLocalFix")
End Function
