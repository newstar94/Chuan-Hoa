Attribute VB_Name = "UnicodeNormalizer"
'==============================================================
' UnicodeNormalizer â€” Chuan hoa NFC (bang anh xa to hop dau roi -> dung san, vi VBA khong co
' ham normalize san), xoa ky tu an. Du lieu lay tu RuleLoader.GetUnicodeToNfc, nguon chan ly la
' shared/rules/unicode-to-nfc.json (CLAUDE.md muc 3.1) - KHONG hard-code bang o day.
' Loai B (CLAUDE.md muc 2.2): ap cho toan van ban, khong can xac nhan tung cho â€” phep bien doi xac
' dinh va khong co ngoai le.
' Moi thu tuc cong khai bat dau bang On Error GoTo ErrHandler. Dung AscW/ChrW, khong dung Asc/Chr
' (phu thuoc code page) â€” CLAUDE.md muc 5.
'==============================================================
Option Explicit

' ============================================================================
' NormalizeNfc â€” buoc 1 (docs/rules/04-loi-go-may.md muc 2): chuan hoa ve dang dung san NFC.
' Ban nay CHI co bang anh xa gioi han tieng Viet (shared/rules/unicode-to-
' nfc.json/combiningToPrecomposed, 132 to hop) â€” ky tu ngoai pham vi do (tieng nuoc ngoai, ky hieu
' toan hoc, emoji...) giu nguyen, dung quy uoc "Khong chac thi khong sua" (CLAUDE.md muc 5).
' THUAT TOAN BAT BUOC â€” "khop dai nhat tai moi vi tri": duyet chuoi tu trai qua phai, o moi vi tri
' thu khop khoa 3 ky tu (nguyen am + hai dau ket hop, vi du "e"+nang+mu) TRUOC, roi khoa 2 ky tu
' (nguyen am + mot dau, vi du "e"+nang). KHONG duoc thay tung khoa tuan tu (vi du lap Dictionary
' roi Replace$ tung khoa) â€” khoa 2 ky tu la TIEN TO cua khoa 3 ky tu tuong ung, thay tien to truoc
' se lam hong ket qua (vi du "Viet" se sai thanh "Viet" voi dau nang rieng + dau mu thua thay vi
' "Viet" dung mot ky tu "e" co du hai dau).
' Tra Dictionary {"text": <chuoi da chuan hoa>, "changedCount": <so to hop da gop>}
' ============================================================================
Public Function NormalizeNfc(ByVal s As String) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("UnicodeNormalizer.NormalizeNfc")
    On Error GoTo ErrHandler

    Dim mapDict As Object
    Set mapDict = RuleLoader.GetUnicodeToNfc()("combiningToPrecomposed")

    Dim n As Long
    n = Len(s)

    ' Moi vi tri trong s sinh toi da MOT phan tu trong pieces (mot ky tu goc, hoac mot ket qua
    ' gop) â€” kich thuoc n la CHAN TREN AN TOAN, khong can tinh chinh xac truoc. Cac o chua dung o
    ' cuoi mang giu gia tri mac dinh vbNullString, Join bo qua nen khong can cat.
    Dim pieces() As String
    If n > 0 Then ReDim pieces(0 To n - 1)
    Dim pieceCount As Long
    pieceCount = 0

    Dim changedCount As Long
    changedCount = 0

    Dim i As Long
    Dim key3 As String, key2 As String
    Dim matched As Boolean

    i = 1
    Do While i <= n
        matched = False

        If i <= n - 2 Then
            key3 = Mid$(s, i, 3)
            If mapDict.Exists(key3) Then
                pieces(pieceCount) = CStr(mapDict(key3))
                pieceCount = pieceCount + 1
                changedCount = changedCount + 1
                i = i + 3
                matched = True
            End If
        End If

        If Not matched And i <= n - 1 Then
            key2 = Mid$(s, i, 2)
            If mapDict.Exists(key2) Then
                pieces(pieceCount) = CStr(mapDict(key2))
                pieceCount = pieceCount + 1
                changedCount = changedCount + 1
                i = i + 2
                matched = True
            End If
        End If

        If Not matched Then
            pieces(pieceCount) = Mid$(s, i, 1)
            pieceCount = pieceCount + 1
            i = i + 1
        End If
    Loop

    Dim Result As Object
    Set Result = Utils.NewDictionary()
    If n > 0 Then
        Result("text") = Join(pieces, vbNullString)
    Else
        Result("text") = s
    End If
    Result("changedCount") = changedCount
    Set NormalizeNfc = Result
    Exit Function
ErrHandler:
    Err.Raise Err.number, "UnicodeNormalizer.NormalizeNfc", Err.description
End Function
