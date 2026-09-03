Attribute VB_Name = "RuleLoader"
Option Explicit

Private mFormatSpec As Object
Private mStyleSheet As Object
Private mCheckRules As Collection
Private mChecklistGroups As Collection
Private mToneMapping As Object
Private mIyMapping As Object
Private mTypoDictionary As Object
Private mPlaceNames As Object
Private mAdministrativeUnitNames As Object
Private mFontPatterns As Object
Private mDocTypeAbbreviations As Object
Private mTerrainPlaceNames As Object
Private mLunarYears As Object
Private mHolidays As Object
Private mSpecialCapitalizations As Object
Private mRegionNames As Object
Private mNonSentenceEndingAbbreviations As Object
Private mCommonOrganNames As Object
Private mTelexWhitelist As Object
Private mCitationRules As Object
Private mUnicodeToNfc As Object
Private mEncodingTables As Object
Private mComponentSignals As Object
Private mRegimeConfig As Object
Private mLoaded As Boolean

' ============================================================================
' Nap va cache â€” goi mot lan duy nhat, cac lan sau dung lai ket qua da nap
' ============================================================================

Private Sub EnsureLoaded()
    On Error GoTo ErrHandler
    If mLoaded Then Exit Sub

    Set mFormatSpec = RuleData.LoadRawFormatSpec()

    ' bo-styles.json khong co truong sourceLabel o goc (chi legalBasis/constraint/sizeToken/
    ' docDefaults/styles/...) - giong dia-danh-viet-nam.json, khong goi NormalizeSourceLabelField.
    Set mStyleSheet = RuleData.LoadRawStyleSheet()

    Dim rawCheckRules As Object
    Set rawCheckRules = RuleData.LoadRawCheckRules()
    Set mCheckRules = rawCheckRules("rules")
    Dim r As Variant
    For Each r In mCheckRules
        NormalizeSourceLabelField r
    Next r

    ' checklistGroups - 14 nhom quy tac, dung de loc khi chay "Kiem tra the thuc"/"Kiem tra chinh ta"
    ' muc 4.2. Khong co truong sourceLabel rieng.
    Set mChecklistGroups = rawCheckRules("checklistGroups")

    Set mToneMapping = RuleData.LoadRawToneMapping()
    NormalizeSourceLabelField mToneMapping

    Set mIyMapping = RuleData.LoadRawIyMapping()
    NormalizeSourceLabelField mIyMapping

    Set mTypoDictionary = RuleData.LoadRawTypoDictionary()
    NormalizeSourceLabelField mTypoDictionary

    ' dia-danh-viet-nam.json khong co truong sourceLabel o goc (chi schemaVersion/status/places).
    Set mPlaceNames = RuleData.LoadRawPlaceNames()

    Set mAdministrativeUnitNames = RuleData.LoadRawAdministrativeUnitNames()
    NormalizeSourceLabelField mAdministrativeUnitNames

    Set mFontPatterns = RuleData.LoadRawFontPatterns()
    NormalizeSourceLabelField mFontPatterns

    Set mDocTypeAbbreviations = RuleData.LoadRawDocTypeAbbreviations()
    NormalizeSourceLabelField mDocTypeAbbreviations

    Set mTerrainPlaceNames = RuleData.LoadRawTerrainPlaceNames()
    NormalizeSourceLabelField mTerrainPlaceNames

    Set mLunarYears = RuleData.LoadRawLunarYears()
    NormalizeSourceLabelField mLunarYears

    Set mHolidays = RuleData.LoadRawHolidays()
    NormalizeSourceLabelField mHolidays

    Set mSpecialCapitalizations = RuleData.LoadRawSpecialCapitalizations()
    NormalizeSourceLabelField mSpecialCapitalizations

    Set mRegionNames = RuleData.LoadRawRegionNames()
    NormalizeSourceLabelField mRegionNames

    Set mNonSentenceEndingAbbreviations = RuleData.LoadRawNonSentenceEndingAbbreviations()
    NormalizeSourceLabelField mNonSentenceEndingAbbreviations

    Set mCommonOrganNames = RuleData.LoadRawCommonOrganNames()
    NormalizeSourceLabelField mCommonOrganNames

    Set mTelexWhitelist = RuleData.LoadRawTelexWhitelist()
    NormalizeSourceLabelField mTelexWhitelist

    Set mCitationRules = RuleData.LoadRawCitationRules()
    NormalizeSourceLabelField mCitationRules

    Set mUnicodeToNfc = RuleData.LoadRawUnicodeToNfc()
    NormalizeSourceLabelField mUnicodeToNfc

    Set mComponentSignals = RuleData.LoadRawComponentSignals()
    NormalizeSourceLabelField mComponentSignals

    ' quy-dinh-che-do.json - ba che do quy dinh, dung boi RegimeDetector.bas.
    Set mRegimeConfig = RuleData.LoadRawRegimeConfig()
    NormalizeSourceLabelField mRegimeConfig

    LoadEncodingTables

    mLoaded = True
    ValidateRules
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleLoader.EnsureLoaded", Err.description
End Sub

Private Sub LoadEncodingTables()
    On Error GoTo ErrHandler

    Set mEncodingTables = Utils.NewDictionary()

    Dim tcvn3Lower As Object
    Set tcvn3Lower = RuleData.LoadRawTcvn3Lower()
    NormalizeSourceLabelField tcvn3Lower
    Set mEncodingTables("tcvn3Lower") = MakeEncodingTable(tcvn3Lower("map"), tcvn3Lower("toneMarks"))

    Dim tcvn3Upper As Object
    Set tcvn3Upper = RuleData.LoadRawTcvn3Upper()
    NormalizeSourceLabelField tcvn3Upper
    Set mEncodingTables("tcvn3Upper") = MakeEncodingTable(tcvn3Upper("map"), tcvn3Upper("toneMarks"))

    Dim vni As Object
    Set vni = RuleData.LoadRawVni()
    NormalizeSourceLabelField vni
    Set mEncodingTables("vni") = MakeEncodingTable(vni("standalone"), vni("combining"))

    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleLoader.LoadEncodingTables", Err.description
End Sub

Private Function MakeEncodingTable(ByVal directMap As Object, ByVal combiningMap As Object) As Object
    Dim t As Object
    Set t = Utils.NewDictionary()
    Set t("direct") = directMap
    Set t("combining") = combiningMap
    Set MakeEncodingTable = t
End Function

' Khong lam gi neu khong co truong.
Private Sub NormalizeSourceLabelField(ByVal d As Object)
    If Not d Is Nothing Then
        If d.Exists("sourceLabel") Then
            d("sourceLabel") = NormalizeSourceLabel(CStr(d("sourceLabel")))
        End If
    End If
End Sub

' ============================================================================
' Chuan hoa nhan nguon â€” dua ve bon gia tri chuan: Nem loi tieng Viet neu gap gia tri la, khong
' doan.
' ============================================================================

Public Function NormalizeSourceLabel(ByVal raw As String) As String
    On Error GoTo ErrHandler
    Dim s As String
    s = UCase$(Trim$(raw))

    ' Cac bien the co that trong shared/rules/*.json hien tai (kiem tra bang grep truoc khi viet
    ' ham nay): "ND30", "QD1989", "SUY RA", "THONG LE" (khong dau â€” quy-tac-kiem-tra.json) va
    ' "THONG LE" co dau "THĂ”NG Lá»†" (telex-whitelist.json, tu-dien-chinh-ta.json...).
    Dim thongLeCoDau As String
    thongLeCoDau = UCase$("TH" & ChrW(&HD4) & "NG L" & ChrW(&H1EC6))

    Select Case s
        Case "ND30"
            NormalizeSourceLabel = "ND30"
        Case "QD1989"
            NormalizeSourceLabel = "QD1989"
        Case "SUY RA"
            NormalizeSourceLabel = "SUY_RA"
        Case "THONG LE"
            NormalizeSourceLabel = "THONG_LE"
        Case thongLeCoDau
            NormalizeSourceLabel = "THONG_LE"
        Case Else
            Err.Raise vbObjectError + 513, "RuleLoader.NormalizeSourceLabel", _
                "Nhan nguon """ & raw & """ khong hop le, phai la mot trong: " & _
                "ND30, QD1989, SUY_RA, THONG_LE."
    End Select
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleLoader.NormalizeSourceLabel", Err.description
End Function

' ============================================================================
' ============================================================================

Public Function GetFormatSpec() As Object
    EnsureLoaded
    Set GetFormatSpec = mFormatSpec
End Function

Public Function GetCheckRules() As Collection
    EnsureLoaded
    Set GetCheckRules = mCheckRules
End Function

Public Function GetCheckRule(ByVal ruleCode As String) As Object
    On Error GoTo ErrHandler
    EnsureLoaded
    Dim r As Variant
    For Each r In mCheckRules
        If r("ruleCode") = ruleCode Then
            Set GetCheckRule = r
            Exit Function
        End If
    Next r
    Set GetCheckRule = Nothing
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleLoader.GetCheckRule", Err.description
End Function

' Moi noi AP hoac KIEM co chu (TextFormatter.ApplyFontSizeWholeDocument,
' ComplianceChecker.CheckFontSizeConsistency) PHAI goi ham nay THAY VI doc thang
' GetFontSizeSet(sizeSetKey)/GetFormatSpec("fontSizeSets")(sizeSetKey) â€” doc thang se BO SOT phan
' ghi de cho VIETTEL/DANG.
' Tra ve MOT DICTIONARY MOI (khong sua doi ban cache goc trong mFormatSpec) â€” vi mFormatSpec duoc
' EnsureLoaded NAP MOT LAN DUY NHAT va dung lai cho toan phien (dau file); sua thang len
' Dictionary goc se lam "ro" ghi de sang CA nhung lan goi voi che do KHAC sau do trong cung phien,
' mot loi kho phat hien vi chi loi khi goi HAI che do khac nhau trong cung mot phien Word (dung
' tinh huong nguoi dung doi Quy dinh giua chung phien lam viec).
Public Function GetEffectiveFontSizeSet(ByVal regimeCode As String, ByVal sizeSetKey As String) As Object
    EnsureLoaded
    Dim baseSet As Object: Set baseSet = mFormatSpec("fontSizeSets")(sizeSetKey)

    Dim overrides As Object: Set overrides = Nothing
    Dim regimes As Object: Set regimes = mRegimeConfig("regimes")
    If regimes.Exists(regimeCode) Then
        Dim regimeEntry As Object: Set regimeEntry = regimes(regimeCode)
        If regimeEntry.Exists("fontSizeOverrides") Then
            If regimeEntry("fontSizeOverrides").Exists(sizeSetKey) Then
                Set overrides = regimeEntry("fontSizeOverrides")(sizeSetKey)
            End If
        End If
    End If

    Dim merged As Object: Set merged = Utils.NewDictionary()
    Dim key As Variant
    For Each key In baseSet.Keys
        merged(key) = baseSet(key)
    Next key
    If Not overrides Is Nothing Then
        For Each key In overrides.Keys
            merged(key) = overrides(key)
        Next key
    End If

    Set GetEffectiveFontSizeSet = merged
End Function

' 2000 ky tu dau tai lieu).
Public Function GetHeaderWindowChars() As Long
    EnsureLoaded
    If mRegimeConfig.Exists("headerWindowChars") Then
        GetHeaderWindowChars = CLng(mRegimeConfig("headerWindowChars"))
    Else
        GetHeaderWindowChars = 2000
    End If
End Function

' Dau ket thuc phan noi dung van ban: ND30/VIETTEL dung "./.", DANG dung dau cham.
Public Function GetContentEndMark(ByVal regimeCode As String) As String
    EnsureLoaded
    GetContentEndMark = "./."
    Dim regimes As Object: Set regimes = mRegimeConfig("regimes")
    If regimes.Exists(regimeCode) Then
        If regimes(regimeCode).Exists("contentEndMark") Then
            GetContentEndMark = CStr(regimes(regimeCode)("contentEndMark"))
        End If
    End If
End Function

' Quy tac gian dong cho cac component mang lineSpacingZone = "body". Tra Dictionary hai khoa:
' "rule" ("single"|"exactly") va "exactlyPt" (0 khi rule la "single").
Public Function GetRegimeLineSpacing(ByVal regimeCode As String, ByVal sizeSetKey As String) As Object
    EnsureLoaded
    Dim Result As Object: Set Result = Utils.NewDictionary()
    Result.Add "rule", "single"
    Result.Add "exactlyPt", 0

    Dim regimes As Object: Set regimes = mRegimeConfig("regimes")
    If Not regimes.Exists(regimeCode) Then
        Set GetRegimeLineSpacing = Result
        Exit Function
    End If
    Dim entry As Object: Set entry = regimes(regimeCode)
    If Not entry.Exists("lineSpacing") Then
        Set GetRegimeLineSpacing = Result
        Exit Function
    End If

    Dim ls As Object: Set ls = entry("lineSpacing")
    If ls.Exists("bodyZoneRule") Then Result("rule") = CStr(ls("bodyZoneRule"))
    If ls.Exists("bodyZoneExactlyPtBySizeSet") Then
        Dim bySet As Object: Set bySet = ls("bodyZoneExactlyPtBySizeSet")
        If bySet.Exists(sizeSetKey) Then Result("exactlyPt") = CDbl(bySet(sizeSetKey))
    End If

    Set GetRegimeLineSpacing = Result
End Function

' Dac ta HIEU LUC cua mot component cho MOT che do â€” hop nhat thong-so-the-thuc.json/components
' voi quy-dinh-che-do.json/regimes/<che do>/componentOverrides. Moi noi AP hoac KIEM dinh dang
' component PHAI goi ham nay thay vi doc thang GetFormatSpec("components")(role).
' documentTypeName (tuy chon): ten loai van ban lay tu doan mang vai tro "typeName" â€” chi dung cho
' ghi de "styleByDocumentTypeGroup" (Viettel: can cu cua nghi quyet/quyet dinh/quy che/quy dinh in
' NGHIENG, cac loai con lai in DUNG).
' Tra ve Dictionary MOI, khong sua ban cache goc (mFormatSpec nap mot lan cho ca phien â€” sua thang
' se lam "ro" ghi de sang cac lan goi voi che do khac trong cung phien).
Public Function GetEffectiveComponentSpec(ByVal regimeCode As String, ByVal role As String, _
        Optional ByVal documentTypeName As String = "") As Object
    EnsureLoaded

    Dim merged As Object: Set merged = Utils.NewDictionary()
    Dim components As Object: Set components = mFormatSpec("components")
    If Not components.Exists(role) Then
        Set GetEffectiveComponentSpec = merged
        Exit Function
    End If

    Dim key As Variant
    Dim baseSpec As Object: Set baseSpec = components(role)
    For Each key In baseSpec.Keys
        If IsObject(baseSpec(key)) Then
            Set merged(key) = baseSpec(key)
        Else
            merged(key) = baseSpec(key)
        End If
    Next key

    Dim regimes As Object: Set regimes = mRegimeConfig("regimes")
    If Not regimes.Exists(regimeCode) Then
        Set GetEffectiveComponentSpec = merged
        Exit Function
    End If
    Dim entry As Object: Set entry = regimes(regimeCode)
    If Not entry.Exists("componentOverrides") Then
        Set GetEffectiveComponentSpec = merged
        Exit Function
    End If
    Dim overrides As Object: Set overrides = entry("componentOverrides")
    If Not overrides.Exists(role) Then
        Set GetEffectiveComponentSpec = merged
        Exit Function
    End If

    Dim roleOverride As Object: Set roleOverride = overrides(role)
    For Each key In roleOverride.Keys
        If left$(CStr(key), 1) = "$" Then GoTo ContinueKey
        If CStr(key) = "styleByDocumentTypeGroup" Then
            Dim groups As Object: Set groups = roleOverride(key)
            Dim groupKey As String
            groupKey = IIf(IsDecreeDocumentType(documentTypeName), "decree", "other")
            If groups.Exists(groupKey) Then
                ' Moi nhom la mot Dictionary ghi de con (vi du {"style":..., "lastLineEndChar":
                ' ","}) â€” hop nhat TUNG truong cua no vao merged, khong chi rieng "style".
                Dim groupOverride As Object: Set groupOverride = groups(groupKey)
                Dim gk As Variant
                For Each gk In groupOverride.Keys
                    If IsObject(groupOverride(gk)) Then
                        Set merged(gk) = groupOverride(gk)
                    Else
                        merged(gk) = groupOverride(gk)
                    End If
                Next gk
            End If
        ElseIf IsObject(roleOverride(key)) Then
            Set merged(key) = roleOverride(key)
        Else
            merged(key) = roleOverride(key)
        End If
ContinueKey:
    Next key

    Set GetEffectiveComponentSpec = merged
End Function

' Ten loai van ban co thuoc nhom "decree" (nghi quyet, quyet dinh, quy che, quy dinh) khong.
Private Function IsDecreeDocumentType(ByVal documentTypeName As String) As Boolean
    IsDecreeDocumentType = False
    If Len(Trim$(documentTypeName)) = 0 Then Exit Function
    If Not mRegimeConfig.Exists("decreeDocumentTypeNames") Then Exit Function

    Dim probe As String
    probe = Utils.ToUpperVn(Trim$(documentTypeName))

    Dim name As Variant
    For Each name In mRegimeConfig("decreeDocumentTypeNames")
        If probe = Utils.ToUpperVn(CStr(name)) Then
            IsDecreeDocumentType = True
            Exit Function
        End If
    Next name
End Function

' Bo 19 style theo huong dan OOXML - nguon chan ly cho nut 2.2 "Dung bo Styles" o CA HAI ban.
Public Function GetStyleSheet() As Object
    EnsureLoaded
    Set GetStyleSheet = mStyleSheet
End Function

Public Function GetToneMapping() As Object
    EnsureLoaded
    Set GetToneMapping = mToneMapping
End Function

Public Function GetIyMapping() As Object
    EnsureLoaded
    Set GetIyMapping = mIyMapping
End Function

Public Function GetTypoDictionary() As Object
    EnsureLoaded
    Set GetTypoDictionary = mTypoDictionary
End Function

Public Function GetPlaceNames() As Object
    EnsureLoaded
    Set GetPlaceNames = mPlaceNames
End Function

' don-vi-hanh-chinh.json â€” danh sach day du ten rieng tinh/thanh + xa/phuong, dung de doi chieu
' chinh ta ND30-PL2-M3-K1A (CheckAdministrativeUnitNameWarn) â€” KHAC mPlaceNames (dia-danh-viet-
' nam.json chi loc tap con co to hop oa/uy, dung cho ToneNormalizer).
Public Function GetAdministrativeUnitNames() As Object
    EnsureLoaded
    Set GetAdministrativeUnitNames = mAdministrativeUnitNames
End Function

Public Function GetFontPatterns() As Object
    EnsureLoaded
    Set GetFontPatterns = mFontPatterns
End Function

Public Function GetDocTypeAbbreviations() As Object
    EnsureLoaded
    Set GetDocTypeAbbreviations = mDocTypeAbbreviations
End Function

Public Function GetComponentSignals() As Object
    EnsureLoaded
    Set GetComponentSignals = mComponentSignals
End Function

' quy-dinh-che-do.json - ba che do quy dinh: nhan hien thi, dau hieu nhan dien che do
' (RegimeDetector.bas), danh sach khoa fontSizeSets kha dung cho tung che do.
Public Function GetRegimeConfig() As Object
    EnsureLoaded
    Set GetRegimeConfig = mRegimeConfig
End Function

Public Function GetTerrainPlaceNames() As Object
    EnsureLoaded
    Set GetTerrainPlaceNames = mTerrainPlaceNames
End Function

Public Function GetLunarYears() As Object
    EnsureLoaded
    Set GetLunarYears = mLunarYears
End Function

Public Function GetHolidays() As Object
    EnsureLoaded
    Set GetHolidays = mHolidays
End Function

Public Function GetSpecialCapitalizations() As Object
    EnsureLoaded
    Set GetSpecialCapitalizations = mSpecialCapitalizations
End Function

Public Function GetRegionNames() As Object
    EnsureLoaded
    Set GetRegionNames = mRegionNames
End Function

Public Function GetNonSentenceEndingAbbreviations() As Object
    EnsureLoaded
    Set GetNonSentenceEndingAbbreviations = mNonSentenceEndingAbbreviations
End Function

Public Function GetCommonOrganNames() As Object
    EnsureLoaded
    Set GetCommonOrganNames = mCommonOrganNames
End Function

Public Function GetTelexWhitelist() As Object
    EnsureLoaded
    Set GetTelexWhitelist = mTelexWhitelist
End Function

Public Function GetCitationRules() As Object
    EnsureLoaded
    Set GetCitationRules = mCitationRules
End Function

' shared/rules/unicode-to-nfc.json.
Public Function GetUnicodeToNfc() As Object
    EnsureLoaded
    Set GetUnicodeToNfc = mUnicodeToNfc
End Function

Public Function GetEncodingTable(ByVal encoding As String) As Object
    On Error GoTo ErrHandler
    EnsureLoaded
    If mEncodingTables.Exists(encoding) Then
        Set GetEncodingTable = mEncodingTables(encoding)
    Else
        Set GetEncodingTable = Nothing
    End If
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleLoader.GetEncodingTable", Err.description
End Function

' ============================================================================
' ValidateRules â€” kiem tra tinh toan ven luc khoi dong, doi chieu validateRules ben TS
' ============================================================================

Public Sub ValidateRules()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("RuleLoader.ValidateRules")
    On Error GoTo ErrHandler
    ' Goi doc lap duoc (vd tu test) â€” EnsureLoaded tu thoat ngay neu da nap roi, ke ca khi ham nay
    ' dang duoc goi TU BEN TRONG EnsureLoaded (mLoaded da = True truoc do).
    EnsureLoaded

    Dim errors As New Collection
    Dim seenRuleCodes As Object
    Set seenRuleCodes = CreateObject("Scripting.Dictionary")

    Dim validGroups As Object, validSeverities As Object, validActionTypes As Object
    Dim validCheckabilities As Object, validRiskLevels As Object
    Set validGroups = MakeStringSet(Array("pageSetup", "bodyText", "encoding", "component", _
        "capitalization", "citation", "structure", "spelling"))
    Set validSeverities = MakeStringSet(Array("error", "warning", "info"))
    Set validActionTypes = MakeStringSet(Array("A", "B", "C"))
    Set validCheckabilities = MakeStringSet(Array("full", "partial", "warnOnly"))
    Set validRiskLevels = MakeStringSet(Array("low", "high"))

    ' checklistGroups: dung DUNG 14 nhom, id 1..14 khong trung lap -
    ' giao-dien.md muc 4.2.
    Dim validChecklistGroupIds As Object
    Set validChecklistGroupIds = Utils.NewDictionary()
    Dim cg As Variant
    For Each cg In mChecklistGroups
        validChecklistGroupIds(CLng(cg("id"))) = True
    Next cg
    If mChecklistGroups.count <> 14 Or validChecklistGroupIds.count <> 14 Then
        errors.Add "checklistGroups phai co dung 14 nhom, id 1-14 khong trung lap (hien co " & _
            mChecklistGroups.count & ")."
    End If

    Dim r As Variant
    Dim ruleCode As String
    For Each r In mCheckRules
        ruleCode = r("ruleCode")

        If seenRuleCodes.Exists(ruleCode) Then
            errors.Add "Ma quy tac trung lap: """ & ruleCode & """."
        End If
        seenRuleCodes(ruleCode) = True

        If Not validGroups.Exists(CStr(r("group"))) Then
            errors.Add "Quy tac """ & ruleCode & """ co group khong hop le: """ & r("group") & """."
        End If
        If Not validSeverities.Exists(CStr(r("severity"))) Then
            errors.Add "Quy tac """ & ruleCode & """ co severity khong hop le: """ & r("severity") & """."
        End If
        If Not validActionTypes.Exists(CStr(r("actionType"))) Then
            errors.Add "Quy tac """ & ruleCode & """ co actionType khong hop le: """ & r("actionType") & """."
        End If
        If Not validCheckabilities.Exists(CStr(r("checkability"))) Then
            errors.Add "Quy tac """ & ruleCode & """ co checkability khong hop le: """ & r("checkability") & """."
        End If
        If Not validRiskLevels.Exists(CStr(r("riskLevel"))) Then
            errors.Add "Quy tac """ & ruleCode & """ co riskLevel khong hop le: """ & r("riskLevel") & """."
        End If
        If Not validChecklistGroupIds.Exists(CLng(r("checklistGroup"))) Then
            errors.Add "Quy tac """ & ruleCode & """ co checklistGroup khong hop le: " & r("checklistGroup") & "."
        End If
    Next r

    ' Moi sizeSetKey tham chieu (hien chi co pageNumber.sizeSetKey) phai ton tai trong
    ' fontSizeSets.set1, fontSizeSets.set2 va fontSizeRanges.
    Dim referencedSizeSetKey As String
    referencedSizeSetKey = mFormatSpec("pageNumber")("sizeSetKey")

    If Not mFormatSpec("fontSizeSets")("set1").Exists(referencedSizeSetKey) Then
        errors.Add "pageNumber.sizeSetKey """ & referencedSizeSetKey & _
            """ khong ton tai trong fontSizeSets.set1."
    End If
    If Not mFormatSpec("fontSizeSets")("set2").Exists(referencedSizeSetKey) Then
        errors.Add "pageNumber.sizeSetKey """ & referencedSizeSetKey & _
            """ khong ton tai trong fontSizeSets.set2."
    End If
    If Not mFormatSpec("fontSizeRanges").Exists(referencedSizeSetKey) Then
        errors.Add "pageNumber.sizeSetKey """ & referencedSizeSetKey & _
            """ khong ton tai trong fontSizeRanges."
    End If

    Dim validComponentSpecKeys As Object
    Set validComponentSpecKeys = MakeStringSet(mFormatSpec("components").Keys)

    Dim signalName As Variant
    Dim signal As Object
    Dim fieldName As Variant
    For Each signalName In mComponentSignals("signals").Keys
        Set signal = mComponentSignals("signals")(signalName)

        For Each fieldName In Array("role", "afterRole", "beforeRole")
            If signal.Exists(CStr(fieldName)) Then
                If Not IsNull(signal(fieldName)) Then
                    If Not validComponentSpecKeys.Exists(CStr(signal(fieldName))) Then
                        errors.Add "dau-hieu-nhan-dien.json/signals." & signalName & _
                            " tham chieu ComponentSpecKey khong ton tai: """ & signal(fieldName) & """."
                    End If
                End If
            End If
        Next fieldName

        ' "anchorRoles" la MANG cac ComponentSpecKey, khac ba truong don o tren.
        If signal.Exists("anchorRoles") Then
            Dim anchorRole As Variant
            For Each anchorRole In signal("anchorRoles")
                If Not validComponentSpecKeys.Exists(CStr(anchorRole)) Then
                    errors.Add "dau-hieu-nhan-dien.json/signals." & signalName & _
                        ".anchorRoles tham chieu ComponentSpecKey khong ton tai: """ & anchorRole & """."
                End If
            Next anchorRole
        End If

        If signal.Exists("regex") Then
            If Not CanCompileRegex(CStr(signal("regex"))) Then
                errors.Add "dau-hieu-nhan-dien.json/signals." & signalName & ".regex khong hop le: """ & _
                    signal("regex") & """."
            End If
        End If
        ' "regexByRegime" - the cho mot regex don khi cach nhan dien khac nhau giua
        ' ND30/VIETTEL/DANG (xem RegimeDetector.bas) - kiem BIEN DICH DUOC cho TUNG nhanh, khong
        ' chi mot nhanh dai dien.
        If signal.Exists("regexByRegime") Then
            Dim regimeKey As Variant
            For Each regimeKey In signal("regexByRegime").Keys
                If Not CanCompileRegex(CStr(signal("regexByRegime")(regimeKey))) Then
                    errors.Add "dau-hieu-nhan-dien.json/signals." & signalName & _
                        ".regexByRegime." & regimeKey & " khong hop le: """ & _
                        signal("regexByRegime")(regimeKey) & """."
                End If
            Next regimeKey
        End If
        If signal.Exists("excludeTrailingPunctuationRegex") Then
            If Not CanCompileRegex(CStr(signal("excludeTrailingPunctuationRegex"))) Then
                errors.Add "dau-hieu-nhan-dien.json/signals." & signalName & _
                    ".excludeTrailingPunctuationRegex khong hop le."
            End If
        End If
    Next signalName

    If errors.count > 0 Then
        Dim msg As String
        Dim e As Variant
        msg = "RuleLoader.ValidateRules(): du lieu shared/rules/ khong hop le:"
        For Each e In errors
            msg = msg & vbCrLf & "- " & e
        Next e
        Err.Raise vbObjectError + 514, "RuleLoader.ValidateRules", msg
    End If
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleLoader.ValidateRules", Err.description
End Sub

' Bien dich thu mot mau qua VBScript.RegExp - CHI kiem duoc loi cu phap PHIA VBA, khong bat duoc
' kieu loi khac giua hai engine.
Private Function CanCompileRegex(ByVal pattern As String) As Boolean
    On Error GoTo ErrHandler
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.pattern = pattern
    regex.test "" ' Ep dinh gia pattern - VBScript.RegExp chi thuc su bien dich luc dung den.
    CanCompileRegex = True
    Exit Function
ErrHandler:
    CanCompileRegex = False
End Function

Private Function MakeStringSet(ByVal items As Variant) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(items) To UBound(items)
        d(CStr(items(i))) = True
    Next i
    Set MakeStringSet = d
End Function
