Attribute VB_Name = "RegimeDetector"
Option Explicit

' Tra Dictionary voi hai khoa: "RegimeCode" (String, "ND30"|"VIETTEL"|"DANG"), "Confident"
' (Boolean - False nghia la KHONG doc duoc du lieu de nhan dien chac chan, dang dung gia tri da
' nho lam du phong).
Public Function DetectRegime(ByVal snapshot As Object) As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RegimeDetector.DetectRegime")
    On Error GoTo ErrHandler

    Dim paragraphs As Collection
    Set paragraphs = snapshot("Paragraphs")

    Dim headerWindow As String
    headerWindow = DocumentSnapshot.BuildHeaderWindow(paragraphs)

    Dim detection As Object
    Set detection = RuleLoader.GetRegimeConfig()("detection")

    Dim nationalTitleTarget As String, partyHeaderTarget As String
    nationalTitleTarget = ComponentDetector.NormalizeForNationalTitle(CStr(detection("nationalTitleTarget")))
    partyHeaderTarget = ComponentDetector.NormalizeForNationalTitle(CStr(detection("partyHeaderTarget")))

    Dim hasNationalTitle As Boolean, hasPartyHeader As Boolean
    hasNationalTitle = ParagraphsOrWindowContain(paragraphs, headerWindow, nationalTitleTarget)
    hasPartyHeader = ParagraphsOrWindowContain(paragraphs, headerWindow, partyHeaderTarget)

    If Not hasNationalTitle And hasPartyHeader Then
        Set DetectRegime = BuildResult("DANG", True)
        Exit Function
    End If

    ' Chay mot luot nhan dien so bo (documentType="khongXacDinh" - khong bo qua bat ky tin hieu
    ' nao) chi de lay van ban cua organName/superiorOrganName - hai vai tro nay KHONG phu thuoc
    ' regime ve mat dau hieu (giong het nhau giua ND30/VIETTEL), nen dung tam "ND30" lam regime
    ' cho luot so bo nay la an toan.
    Dim preliminary As Object
    Set preliminary = ComponentDetector.DetectComponents(snapshot, "khongXacDinh", "ND30")
    Dim layoutMap As Object
    Set layoutMap = preliminary("LayoutMap")

    Dim organText As String
    organText = Trim$(TextOfRole(paragraphs, layoutMap, "organName") & " " & _
        TextOfRole(paragraphs, layoutMap, "superiorOrganName"))

    If Len(organText) = 0 Then
        Set DetectRegime = BuildResult(RegimeState.RememberedRegimeCode(), False)
        Exit Function
    End If

    Dim normalizedOrgan As String
    normalizedOrgan = ComponentDetector.NormalizeForNationalTitle(organText)

    Dim marker As Variant
    For Each marker In detection("viettelMarkers")
        If InStr(normalizedOrgan, ComponentDetector.NormalizeForNationalTitle(CStr(marker))) > 0 Then
            Set DetectRegime = BuildResult("VIETTEL", True)
            Exit Function
        End If
    Next marker

    Set DetectRegime = BuildResult("ND30", True)
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RegimeDetector.DetectRegime", Err.description
End Function

Private Function BuildResult(ByVal regimeCode As String, ByVal confident As Boolean) As Object
    Dim Result As Object
    Set Result = Utils.NewDictionary()
    Result.Add "RegimeCode", regimeCode
    Result.Add "Confident", confident
    Set BuildResult = Result
End Function

' True neu "target" (DA chuan hoa qua NormalizeForNationalTitle) xuat hien trong MOT doan RIENG LE
' (giong het phuong thuc "normalizedContains" cua ComponentDetector) HOAC trong ca headerWindow
' (B0.1 - bo tro, bat truong hop cum tu bi tach lech qua nhieu doan).
Private Function ParagraphsOrWindowContain(ByVal paragraphs As Collection, ByVal headerWindow As String, _
        ByVal target As String) As Boolean
    Dim p As ParagraphSnapshot
    For Each p In paragraphs
        If InStr(ComponentDetector.NormalizeForNationalTitle(p.text), target) > 0 Then
            ParagraphsOrWindowContain = True
            Exit Function
        End If
    Next p
    ParagraphsOrWindowContain = (InStr(ComponentDetector.NormalizeForNationalTitle(headerWindow), target) > 0)
End Function

' Noi dung (Trim) cua doan MANG vai tro "role" trong layoutMap - rong neu khong co doan nao. Chi
' tra doan DAU TIEN (FirstIndexWithRole) - organName/superiorOrganName chi co nhieu nhat mot doan
' moi vai tro theo thiet ke cua ComponentDetector.
Private Function TextOfRole(ByVal paragraphs As Collection, ByVal layoutMap As Object, ByVal role As String) As String
    Dim key As Variant
    For Each key In layoutMap.Keys
        If CStr(layoutMap(key)) = role Then
            Dim p As ParagraphSnapshot
            For Each p In paragraphs
                If p.Index = CLng(key) Then
                    TextOfRole = Trim$(p.text)
                    Exit Function
                End If
            Next p
        End If
    Next key
    TextOfRole = ""
End Function
