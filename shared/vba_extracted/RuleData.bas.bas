Attribute VB_Name = "RuleData"
'==============================================================
' FILE NAY DUOC SINH TU DONG - KHONG SUA TAY
' Nguon: shared/rules/*.json (25 file)
' Hash SHA256 noi dung nguon: c587386ab8d6b5fef7b6a20a6c6646f8b6f453914f3565ffe942d8bb543f0c9d
' Sinh luc: 2026-08-27 22:06:30
' Muon doi gia tri: sua file JSON roi chay legacy/build/gen-rule-data.ps1
'
' Ky tu ngoai ASCII duoc sinh bang ChrW(&Hxxxx), KHONG dua vao ma trang cua
' VBE luc Import file .bas nay â€” xem phan .DESCRIPTION cua gen-rule-data.ps1
' de biet ly do (CLAUDE.md muc 5: dung AscW/ChrW, khong dung Asc/Chr).
'
' Moi ham LoadRaw<Ten>() phan anh dung cau truc shared/rules/<file>.json (doi
' RuleLoader.bas dung
' cac ham nay de dung API cong khai â€” KHONG goi truc tiep tu noi khac.
'==============================================================
Option Explicit

Public Const RULE_DATA_SOURCE_HASH As String = "c587386ab8d6b5fef7b6a20a6c6646f8b6f453914f3565ffe942d8bb543f0c9d"

Public Const PAGE_WIDTH_MM As Double = 210
Public Const PAGE_HEIGHT_MM As Double = 297
Public Const MARGIN_TOP_MM As Double = 25
Public Const MARGIN_BOTTOM_MM As Double = 20
Public Const MARGIN_LEFT_MM As Double = 35
Public Const MARGIN_RIGHT_MM As Double = 20
Public Const FONT_NAME As String = "Times New Roman"
Public Const FONT_COLOR As String = "#000000"
Public Const PAGE_NUMBER_START_AT As Long = 1

Private Function LoadRawTcvn3Lower_ToneMarks() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "B0", ChrW(&H300)
    d.Add "B1", ChrW(&H309)
    d.Add "B2", ChrW(&H303)
    d.Add "B3", ChrW(&H301)
    d.Add "B4", ChrW(&H323)
    Set LoadRawTcvn3Lower_ToneMarks = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTcvn3Lower_ToneMarks", Err.description
End Function

Private Function LoadRawTcvn3Lower_Map() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "A1", ChrW(&H102)
    d.Add "A2", ChrW(&HC2)
    d.Add "A3", ChrW(&HCA)
    d.Add "A4", ChrW(&HD4)
    d.Add "A5", ChrW(&H1A0)
    d.Add "A6", ChrW(&H1AF)
    d.Add "A7", ChrW(&H110)
    d.Add "A8", ChrW(&H103)
    d.Add "A9", ChrW(&HE2)
    d.Add "AA", ChrW(&HEA)
    d.Add "AB", ChrW(&HF4)
    d.Add "AC", ChrW(&H1A1)
    d.Add "AD", ChrW(&H1B0)
    d.Add "AE", ChrW(&H111)
    d.Add "AF", ChrW(&H1EB0)
    d.Add "B5", ChrW(&HE0)
    d.Add "B6", ChrW(&H1EA3)
    d.Add "B7", ChrW(&HE3)
    d.Add "B8", ChrW(&HE1)
    d.Add "B9", ChrW(&H1EA1)
    d.Add "BA", ChrW(&H1EB2)
    d.Add "BB", ChrW(&H1EB1)
    d.Add "BC", ChrW(&H1EB3)
    d.Add "BD", ChrW(&H1EB5)
    d.Add "BE", ChrW(&H1EAF)
    d.Add "BF", ChrW(&H1EB4)
    d.Add "C0", ChrW(&H1EAE)
    d.Add "C1", ChrW(&H1EA6)
    d.Add "C2", ChrW(&H1EA8)
    d.Add "C3", ChrW(&H1EAA)
    d.Add "C4", ChrW(&H1EA4)
    d.Add "C5", ChrW(&H1EC0)
    d.Add "C6", ChrW(&H1EB7)
    d.Add "C7", ChrW(&H1EA7)
    d.Add "C8", ChrW(&H1EA9)
    d.Add "C9", ChrW(&H1EAB)
    d.Add "CA", ChrW(&H1EA5)
    d.Add "CB", ChrW(&H1EAD)
    d.Add "CC", ChrW(&HE8)
    d.Add "CD", ChrW(&H1EC2)
    d.Add "CE", ChrW(&H1EBB)
    d.Add "CF", ChrW(&H1EBD)
    d.Add "D0", ChrW(&HE9)
    d.Add "D1", ChrW(&H1EB9)
    d.Add "D2", ChrW(&H1EC1)
    d.Add "D3", ChrW(&H1EC3)
    d.Add "D4", ChrW(&H1EC5)
    d.Add "D5", ChrW(&H1EBF)
    d.Add "D6", ChrW(&H1EC7)
    d.Add "D7", ChrW(&HEC)
    d.Add "D8", ChrW(&H1EC9)
    d.Add "D9", ChrW(&H1EC4)
    d.Add "DA", ChrW(&H1EBE)
    d.Add "DB", ChrW(&H1ED2)
    d.Add "DC", ChrW(&H129)
    d.Add "DD", ChrW(&HED)
    d.Add "DE", ChrW(&H1ECB)
    d.Add "DF", ChrW(&HF2)
    d.Add "E0", ChrW(&H1ED4)
    d.Add "E1", ChrW(&H1ECF)
    d.Add "E2", ChrW(&HF5)
    d.Add "E3", ChrW(&HF3)
    d.Add "E4", ChrW(&H1ECD)
    d.Add "E5", ChrW(&H1ED3)
    d.Add "E6", ChrW(&H1ED5)
    d.Add "E7", ChrW(&H1ED7)
    d.Add "E8", ChrW(&H1ED1)
    d.Add "E9", ChrW(&H1ED9)
    d.Add "EA", ChrW(&H1EDD)
    d.Add "EB", ChrW(&H1EDF)
    d.Add "EC", ChrW(&H1EE1)
    d.Add "ED", ChrW(&H1EDB)
    d.Add "EE", ChrW(&H1EE3)
    d.Add "EF", ChrW(&HF9)
    d.Add "F0", ChrW(&H1ED6)
    d.Add "F1", ChrW(&H1EE7)
    d.Add "F2", ChrW(&H169)
    d.Add "F3", ChrW(&HFA)
    d.Add "F4", ChrW(&H1EE5)
    d.Add "F5", ChrW(&H1EEB)
    d.Add "F6", ChrW(&H1EED)
    d.Add "F7", ChrW(&H1EEF)
    d.Add "F8", ChrW(&H1EE9)
    d.Add "F9", ChrW(&H1EF1)
    d.Add "FA", ChrW(&H1EF3)
    d.Add "FB", ChrW(&H1EF7)
    d.Add "FC", ChrW(&H1EF9)
    d.Add "FD", ChrW(&HFD)
    d.Add "FE", ChrW(&H1EF5)
    d.Add "FF", ChrW(&H1ED0)
    Set LoadRawTcvn3Lower_Map = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTcvn3Lower_Map", Err.description
End Function

Private Function LoadRawTcvn3Lower_MapVscii1Extension_Map() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "80", ChrW(&HC0)
    d.Add "81", ChrW(&H1EA2)
    d.Add "82", ChrW(&HC3)
    d.Add "83", ChrW(&HC1)
    d.Add "84", ChrW(&H1EA0)
    d.Add "85", ChrW(&H1EB6)
    d.Add "86", ChrW(&H1EAC)
    d.Add "87", ChrW(&HC8)
    d.Add "88", ChrW(&H1EBA)
    d.Add "89", ChrW(&H1EBC)
    d.Add "8A", ChrW(&HC9)
    d.Add "8B", ChrW(&H1EB8)
    d.Add "8C", ChrW(&H1EC6)
    d.Add "8D", ChrW(&HCC)
    d.Add "8E", ChrW(&H1EC8)
    d.Add "8F", ChrW(&H128)
    d.Add "90", ChrW(&HCD)
    d.Add "91", ChrW(&H1ECA)
    d.Add "92", ChrW(&HD2)
    d.Add "93", ChrW(&H1ECE)
    d.Add "94", ChrW(&HD5)
    d.Add "95", ChrW(&HD3)
    d.Add "96", ChrW(&H1ECC)
    d.Add "97", ChrW(&H1ED8)
    d.Add "98", ChrW(&H1EDC)
    d.Add "99", ChrW(&H1EDE)
    d.Add "9A", ChrW(&H1EE0)
    d.Add "9B", ChrW(&H1EDA)
    d.Add "9C", ChrW(&H1EE2)
    d.Add "9D", ChrW(&HD9)
    d.Add "9E", ChrW(&H1EE6)
    d.Add "9F", ChrW(&H168)
    Set LoadRawTcvn3Lower_MapVscii1Extension_Map = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTcvn3Lower_MapVscii1Extension_Map", Err.description
End Function

Private Function LoadRawTcvn3Lower_MapVscii1Extension() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "map", LoadRawTcvn3Lower_MapVscii1Extension_Map()
    Set LoadRawTcvn3Lower_MapVscii1Extension = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTcvn3Lower_MapVscii1Extension", Err.description
End Function

Public Function LoadRawTcvn3Lower() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "1.0"
    d.Add "sourceLabel", "ND30"
    d.Add "ruleCode", "ENC-TCVN3-LOWER"
    d.Add "actionType", "B"
    d.Add "toneMarks", LoadRawTcvn3Lower_ToneMarks()
    d.Add "map", LoadRawTcvn3Lower_Map()
    d.Add "mapVscii1Extension", LoadRawTcvn3Lower_MapVscii1Extension()
    Set LoadRawTcvn3Lower = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTcvn3Lower", Err.description
End Function

Private Function LoadRawTcvn3Upper_ToneMarks() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "B0", ChrW(&H300)
    d.Add "B1", ChrW(&H309)
    d.Add "B2", ChrW(&H303)
    d.Add "B3", ChrW(&H301)
    d.Add "B4", ChrW(&H323)
    Set LoadRawTcvn3Upper_ToneMarks = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTcvn3Upper_ToneMarks", Err.description
End Function

Private Function LoadRawTcvn3Upper_Map() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "A1", ChrW(&H102)
    d.Add "A2", ChrW(&HC2)
    d.Add "A3", ChrW(&HCA)
    d.Add "A4", ChrW(&HD4)
    d.Add "A5", ChrW(&H1A0)
    d.Add "A6", ChrW(&H1AF)
    d.Add "A7", ChrW(&H110)
    d.Add "A8", ChrW(&H102)
    d.Add "A9", ChrW(&HC2)
    d.Add "AA", ChrW(&HCA)
    d.Add "AB", ChrW(&HD4)
    d.Add "AC", ChrW(&H1A0)
    d.Add "AD", ChrW(&H1AF)
    d.Add "AE", ChrW(&H110)
    d.Add "AF", ChrW(&H1EB0)
    d.Add "B5", ChrW(&HC0)
    d.Add "B6", ChrW(&H1EA2)
    d.Add "B7", ChrW(&HC3)
    d.Add "B8", ChrW(&HC1)
    d.Add "B9", ChrW(&H1EA0)
    d.Add "BA", ChrW(&H1EB2)
    d.Add "BB", ChrW(&H1EB0)
    d.Add "BC", ChrW(&H1EB2)
    d.Add "BD", ChrW(&H1EB4)
    d.Add "BE", ChrW(&H1EAE)
    d.Add "BF", ChrW(&H1EB4)
    d.Add "C0", ChrW(&H1EAE)
    d.Add "C1", ChrW(&H1EA6)
    d.Add "C2", ChrW(&H1EA8)
    d.Add "C3", ChrW(&H1EAA)
    d.Add "C4", ChrW(&H1EA4)
    d.Add "C5", ChrW(&H1EC0)
    d.Add "C6", ChrW(&H1EB6)
    d.Add "C7", ChrW(&H1EA6)
    d.Add "C8", ChrW(&H1EA8)
    d.Add "C9", ChrW(&H1EAA)
    d.Add "CA", ChrW(&H1EA4)
    d.Add "CB", ChrW(&H1EAC)
    d.Add "CC", ChrW(&HC8)
    d.Add "CD", ChrW(&H1EC2)
    d.Add "CE", ChrW(&H1EBA)
    d.Add "CF", ChrW(&H1EBC)
    d.Add "D0", ChrW(&HC9)
    d.Add "D1", ChrW(&H1EB8)
    d.Add "D2", ChrW(&H1EC0)
    d.Add "D3", ChrW(&H1EC2)
    d.Add "D4", ChrW(&H1EC4)
    d.Add "D5", ChrW(&H1EBE)
    d.Add "D6", ChrW(&H1EC6)
    d.Add "D7", ChrW(&HCC)
    d.Add "D8", ChrW(&H1EC8)
    d.Add "D9", ChrW(&H1EC4)
    d.Add "DA", ChrW(&H1EBE)
    d.Add "DB", ChrW(&H1ED2)
    d.Add "DC", ChrW(&H128)
    d.Add "DD", ChrW(&HCD)
    d.Add "DE", ChrW(&H1ECA)
    d.Add "DF", ChrW(&HD2)
    d.Add "E0", ChrW(&H1ED4)
    d.Add "E1", ChrW(&H1ECE)
    d.Add "E2", ChrW(&HD5)
    d.Add "E3", ChrW(&HD3)
    d.Add "E4", ChrW(&H1ECC)
    d.Add "E5", ChrW(&H1ED2)
    d.Add "E6", ChrW(&H1ED4)
    d.Add "E7", ChrW(&H1ED6)
    d.Add "E8", ChrW(&H1ED0)
    d.Add "E9", ChrW(&H1ED8)
    d.Add "EA", ChrW(&H1EDC)
    d.Add "EB", ChrW(&H1EDE)
    d.Add "EC", ChrW(&H1EE0)
    d.Add "ED", ChrW(&H1EDA)
    d.Add "EE", ChrW(&H1EE2)
    d.Add "EF", ChrW(&HD9)
    d.Add "F0", ChrW(&H1ED6)
    d.Add "F1", ChrW(&H1EE6)
    d.Add "F2", ChrW(&H168)
    d.Add "F3", ChrW(&HDA)
    d.Add "F4", ChrW(&H1EE4)
    d.Add "F5", ChrW(&H1EEA)
    d.Add "F6", ChrW(&H1EEC)
    d.Add "F7", ChrW(&H1EEE)
    d.Add "F8", ChrW(&H1EE8)
    d.Add "F9", ChrW(&H1EF0)
    d.Add "FA", ChrW(&H1EF2)
    d.Add "FB", ChrW(&H1EF6)
    d.Add "FC", ChrW(&H1EF8)
    d.Add "FD", ChrW(&HDD)
    d.Add "FE", ChrW(&H1EF4)
    d.Add "FF", ChrW(&H1ED0)
    Set LoadRawTcvn3Upper_Map = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTcvn3Upper_Map", Err.description
End Function

Private Function LoadRawTcvn3Upper_MapVscii1Extension_Map() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "80", ChrW(&HC0)
    d.Add "81", ChrW(&H1EA2)
    d.Add "82", ChrW(&HC3)
    d.Add "83", ChrW(&HC1)
    d.Add "84", ChrW(&H1EA0)
    d.Add "85", ChrW(&H1EB6)
    d.Add "86", ChrW(&H1EAC)
    d.Add "87", ChrW(&HC8)
    d.Add "88", ChrW(&H1EBA)
    d.Add "89", ChrW(&H1EBC)
    d.Add "8A", ChrW(&HC9)
    d.Add "8B", ChrW(&H1EB8)
    d.Add "8C", ChrW(&H1EC6)
    d.Add "8D", ChrW(&HCC)
    d.Add "8E", ChrW(&H1EC8)
    d.Add "8F", ChrW(&H128)
    d.Add "90", ChrW(&HCD)
    d.Add "91", ChrW(&H1ECA)
    d.Add "92", ChrW(&HD2)
    d.Add "93", ChrW(&H1ECE)
    d.Add "94", ChrW(&HD5)
    d.Add "95", ChrW(&HD3)
    d.Add "96", ChrW(&H1ECC)
    d.Add "97", ChrW(&H1ED8)
    d.Add "98", ChrW(&H1EDC)
    d.Add "99", ChrW(&H1EDE)
    d.Add "9A", ChrW(&H1EE0)
    d.Add "9B", ChrW(&H1EDA)
    d.Add "9C", ChrW(&H1EE2)
    d.Add "9D", ChrW(&HD9)
    d.Add "9E", ChrW(&H1EE6)
    d.Add "9F", ChrW(&H168)
    Set LoadRawTcvn3Upper_MapVscii1Extension_Map = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTcvn3Upper_MapVscii1Extension_Map", Err.description
End Function

Private Function LoadRawTcvn3Upper_MapVscii1Extension() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "map", LoadRawTcvn3Upper_MapVscii1Extension_Map()
    Set LoadRawTcvn3Upper_MapVscii1Extension = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTcvn3Upper_MapVscii1Extension", Err.description
End Function

Public Function LoadRawTcvn3Upper() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "1.0"
    d.Add "sourceLabel", "ND30"
    d.Add "ruleCode", "ENC-TCVN3-UPPER"
    d.Add "actionType", "B"
    d.Add "toneMarks", LoadRawTcvn3Upper_ToneMarks()
    d.Add "map", LoadRawTcvn3Upper_Map()
    d.Add "mapVscii1Extension", LoadRawTcvn3Upper_MapVscii1Extension()
    Set LoadRawTcvn3Upper = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTcvn3Upper", Err.description
End Function

Private Function LoadRawVni_Standalone() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "C6", ChrW(&H1EC8)
    d.Add "CC", ChrW(&HCC)
    d.Add "CD", ChrW(&HCD)
    d.Add "CE", ChrW(&H1EF4)
    d.Add "D1", ChrW(&H110)
    d.Add "D2", ChrW(&H1ECA)
    d.Add "D3", ChrW(&H128)
    d.Add "D4", ChrW(&H1A0)
    d.Add "D6", ChrW(&H1AF)
    d.Add "E6", ChrW(&H1EC9)
    d.Add "EC", ChrW(&HEC)
    d.Add "ED", ChrW(&HED)
    d.Add "EE", ChrW(&H1EF5)
    d.Add "F1", ChrW(&H111)
    d.Add "F2", ChrW(&H1ECB)
    d.Add "F3", ChrW(&H129)
    d.Add "F4", ChrW(&H1A1)
    d.Add "F6", ChrW(&H1B0)
    Set LoadRawVni_Standalone = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawVni_Standalone", Err.description
End Function

Private Function LoadRawVni_Combining() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "C0", ChrW(&H302) & ChrW(&H300)
    d.Add "C1", ChrW(&H302) & ChrW(&H301)
    d.Add "C2", ChrW(&H302)
    d.Add "C3", ChrW(&H302) & ChrW(&H303)
    d.Add "C4", ChrW(&H323) & ChrW(&H302)
    d.Add "C5", ChrW(&H302) & ChrW(&H309)
    d.Add "C8", ChrW(&H306) & ChrW(&H300)
    d.Add "C9", ChrW(&H306) & ChrW(&H301)
    d.Add "CA", ChrW(&H306)
    d.Add "CB", ChrW(&H323) & ChrW(&H306)
    d.Add "CF", ChrW(&H323)
    d.Add "D5", ChrW(&H303)
    d.Add "D8", ChrW(&H300)
    d.Add "D9", ChrW(&H301)
    d.Add "DA", ChrW(&H306) & ChrW(&H309)
    d.Add "DB", ChrW(&H309)
    d.Add "DC", ChrW(&H306) & ChrW(&H303)
    d.Add "E0", ChrW(&H302) & ChrW(&H300)
    d.Add "E1", ChrW(&H302) & ChrW(&H301)
    d.Add "E2", ChrW(&H302)
    d.Add "E3", ChrW(&H302) & ChrW(&H303)
    d.Add "E4", ChrW(&H323) & ChrW(&H302)
    d.Add "E5", ChrW(&H302) & ChrW(&H309)
    d.Add "E8", ChrW(&H306) & ChrW(&H300)
    d.Add "E9", ChrW(&H306) & ChrW(&H301)
    d.Add "EA", ChrW(&H306)
    d.Add "EB", ChrW(&H323) & ChrW(&H306)
    d.Add "EF", ChrW(&H323)
    d.Add "F5", ChrW(&H303)
    d.Add "F8", ChrW(&H300)
    d.Add "F9", ChrW(&H301)
    d.Add "FA", ChrW(&H306) & ChrW(&H309)
    d.Add "FB", ChrW(&H309)
    d.Add "FC", ChrW(&H306) & ChrW(&H303)
    Set LoadRawVni_Combining = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawVni_Combining", Err.description
End Function

Public Function LoadRawVni() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "1.0"
    d.Add "sourceLabel", "THONG LE"
    d.Add "ruleCode", "ENC-VNI"
    d.Add "actionType", "B"
    d.Add "standalone", LoadRawVni_Standalone()
    d.Add "combining", LoadRawVni_Combining()
    Set LoadRawVni = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawVni", Err.description
End Function

Private Function LoadRawStyleSheet_LegalBasis() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "source", "references/huong-dan-template-nd30-ooxml.md"
    d.Add "sections", ChrW(&HA7) & "2.1 - " & ChrW(&HA7) & "2.6, Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c A, Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c B"
    d.Add "providedDate", "2026-08-11"
    Set LoadRawStyleSheet_LegalBasis = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_LegalBasis", Err.description
End Function

Private Function LoadRawStyleSheet_Constraint() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "noNewStyles", True
    d.Add "styleCount", 19
    d.Add "builtInCount", 9
    d.Add "latentCount", 10
    Set LoadRawStyleSheet_Constraint = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Constraint", Err.description
End Function

Private Function LoadRawStyleSheet_SizeToken_BySet_Set1() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "label", "C" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & " 14"
    d.Add "halfPoint", 28
    d.Add "pt", 14
    Set LoadRawStyleSheet_SizeToken_BySet_Set1 = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_SizeToken_BySet_Set1", Err.description
End Function

Private Function LoadRawStyleSheet_SizeToken_BySet_Set2() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "label", "C" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & " 13"
    d.Add "halfPoint", 26
    d.Add "pt", 13
    Set LoadRawStyleSheet_SizeToken_BySet_Set2 = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_SizeToken_BySet_Set2", Err.description
End Function

Private Function LoadRawStyleSheet_SizeToken_BySet_Set3() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "label", "C" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & " 15"
    d.Add "halfPoint", 30
    d.Add "pt", 15
    Set LoadRawStyleSheet_SizeToken_BySet_Set3 = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_SizeToken_BySet_Set3", Err.description
End Function

Private Function LoadRawStyleSheet_SizeToken_BySet() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "set1", LoadRawStyleSheet_SizeToken_BySet_Set1()
    d.Add "set2", LoadRawStyleSheet_SizeToken_BySet_Set2()
    d.Add "set3", LoadRawStyleSheet_SizeToken_BySet_Set3()
    Set LoadRawStyleSheet_SizeToken_BySet = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_SizeToken_BySet", Err.description
End Function

Private Function LoadRawStyleSheet_SizeToken() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "placeholder", "${SZ_MAIN}"
    d.Add "bySet", LoadRawStyleSheet_SizeToken_BySet()
    Set LoadRawStyleSheet_SizeToken = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_SizeToken", Err.description
End Function

Private Function LoadRawStyleSheet_DocDefaults_RunFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "font", "Times New Roman"
    d.Add "color", "auto"
    d.Add "kern", 0
    d.Add "sz", "${SZ_MAIN}"
    d.Add "szCs", "${SZ_MAIN}"
    d.Add "lang", "vi-VN"
    d.Add "ligatures", "none"
    Set LoadRawStyleSheet_DocDefaults_RunFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_DocDefaults_RunFormat", Err.description
End Function

Private Function LoadRawStyleSheet_DocDefaults_ParagraphFormat_Spacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "before", 0
    d.Add "after", 0
    d.Add "line", 240
    d.Add "lineRule", "auto"
    Set LoadRawStyleSheet_DocDefaults_ParagraphFormat_Spacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_DocDefaults_ParagraphFormat_Spacing", Err.description
End Function

Private Function LoadRawStyleSheet_DocDefaults_ParagraphFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "widowControl", True
    d.Add "spacing", LoadRawStyleSheet_DocDefaults_ParagraphFormat_Spacing()
    Set LoadRawStyleSheet_DocDefaults_ParagraphFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_DocDefaults_ParagraphFormat", Err.description
End Function

Private Function LoadRawStyleSheet_DocDefaults() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "runFormat", LoadRawStyleSheet_DocDefaults_RunFormat()
    d.Add "paragraphFormat", LoadRawStyleSheet_DocDefaults_ParagraphFormat()
    Set LoadRawStyleSheet_DocDefaults = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_DocDefaults", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_0_ParagraphFormat_Spacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "before", 240
    d.Add "after", 0
    d.Add "line", 240
    d.Add "lineRule", "auto"
    Set LoadRawStyleSheet_Styles_0_ParagraphFormat_Spacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_0_ParagraphFormat_Spacing", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_0_ParagraphFormat_Indent() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "firstLine", 567
    Set LoadRawStyleSheet_Styles_0_ParagraphFormat_Indent = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_0_ParagraphFormat_Indent", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_0_ParagraphFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "widowControl", True
    d.Add "spacing", LoadRawStyleSheet_Styles_0_ParagraphFormat_Spacing()
    d.Add "indent", LoadRawStyleSheet_Styles_0_ParagraphFormat_Indent()
    d.Add "jc", "both"
    Set LoadRawStyleSheet_Styles_0_ParagraphFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_0_ParagraphFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_0_RunFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "sz", "${SZ_MAIN}"
    d.Add "szCs", "${SZ_MAIN}"
    Set LoadRawStyleSheet_Styles_0_RunFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_0_RunFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_1_ParagraphFormat_Spacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "before", 240
    d.Add "after", 0
    d.Add "line", 240
    d.Add "lineRule", "auto"
    Set LoadRawStyleSheet_Styles_1_ParagraphFormat_Spacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_1_ParagraphFormat_Spacing", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_1_ParagraphFormat_Indent() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "firstLine", 0
    Set LoadRawStyleSheet_Styles_1_ParagraphFormat_Indent = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_1_ParagraphFormat_Indent", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_1_ParagraphFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "keepNext", True
    d.Add "keepLines", True
    d.Add "spacing", LoadRawStyleSheet_Styles_1_ParagraphFormat_Spacing()
    d.Add "indent", LoadRawStyleSheet_Styles_1_ParagraphFormat_Indent()
    d.Add "jc", "center"
    d.Add "outlineLvl", 0
    Set LoadRawStyleSheet_Styles_1_ParagraphFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_1_ParagraphFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_1_RunFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "b", True
    d.Add "bCs", True
    d.Add "sz", "${SZ_MAIN}"
    d.Add "szCs", "${SZ_MAIN}"
    Set LoadRawStyleSheet_Styles_1_RunFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_1_RunFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_2_ParagraphFormat_Spacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "before", 240
    d.Add "after", 0
    d.Add "line", 240
    d.Add "lineRule", "auto"
    Set LoadRawStyleSheet_Styles_2_ParagraphFormat_Spacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_2_ParagraphFormat_Spacing", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_2_ParagraphFormat_Indent() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "firstLine", 0
    Set LoadRawStyleSheet_Styles_2_ParagraphFormat_Indent = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_2_ParagraphFormat_Indent", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_2_ParagraphFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "keepNext", True
    d.Add "keepLines", True
    d.Add "spacing", LoadRawStyleSheet_Styles_2_ParagraphFormat_Spacing()
    d.Add "indent", LoadRawStyleSheet_Styles_2_ParagraphFormat_Indent()
    d.Add "jc", "center"
    d.Add "outlineLvl", 1
    Set LoadRawStyleSheet_Styles_2_ParagraphFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_2_ParagraphFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_2_RunFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "b", True
    d.Add "bCs", True
    d.Add "sz", "${SZ_MAIN}"
    d.Add "szCs", "${SZ_MAIN}"
    Set LoadRawStyleSheet_Styles_2_RunFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_2_RunFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_3_ParagraphFormat_Spacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "before", 240
    d.Add "after", 0
    d.Add "line", 240
    d.Add "lineRule", "auto"
    Set LoadRawStyleSheet_Styles_3_ParagraphFormat_Spacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_3_ParagraphFormat_Spacing", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_3_ParagraphFormat_Indent() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "firstLine", 0
    Set LoadRawStyleSheet_Styles_3_ParagraphFormat_Indent = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_3_ParagraphFormat_Indent", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_3_ParagraphFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "keepNext", True
    d.Add "keepLines", True
    d.Add "spacing", LoadRawStyleSheet_Styles_3_ParagraphFormat_Spacing()
    d.Add "indent", LoadRawStyleSheet_Styles_3_ParagraphFormat_Indent()
    d.Add "jc", "center"
    d.Add "outlineLvl", 2
    Set LoadRawStyleSheet_Styles_3_ParagraphFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_3_ParagraphFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_3_RunFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "b", True
    d.Add "bCs", True
    d.Add "sz", "${SZ_MAIN}"
    d.Add "szCs", "${SZ_MAIN}"
    Set LoadRawStyleSheet_Styles_3_RunFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_3_RunFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_4_ParagraphFormat_Spacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "before", 240
    d.Add "after", 0
    d.Add "line", 240
    d.Add "lineRule", "auto"
    Set LoadRawStyleSheet_Styles_4_ParagraphFormat_Spacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_4_ParagraphFormat_Spacing", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_4_ParagraphFormat_Indent() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "firstLine", 0
    Set LoadRawStyleSheet_Styles_4_ParagraphFormat_Indent = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_4_ParagraphFormat_Indent", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_4_ParagraphFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "keepNext", True
    d.Add "keepLines", True
    d.Add "spacing", LoadRawStyleSheet_Styles_4_ParagraphFormat_Spacing()
    d.Add "indent", LoadRawStyleSheet_Styles_4_ParagraphFormat_Indent()
    d.Add "jc", "center"
    d.Add "outlineLvl", 3
    Set LoadRawStyleSheet_Styles_4_ParagraphFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_4_ParagraphFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_4_RunFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "b", True
    d.Add "bCs", True
    d.Add "sz", "${SZ_MAIN}"
    d.Add "szCs", "${SZ_MAIN}"
    Set LoadRawStyleSheet_Styles_4_RunFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_4_RunFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_5_ParagraphFormat_Spacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "before", 240
    d.Add "after", 0
    d.Add "line", 240
    d.Add "lineRule", "auto"
    Set LoadRawStyleSheet_Styles_5_ParagraphFormat_Spacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_5_ParagraphFormat_Spacing", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_5_ParagraphFormat_Indent() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "firstLine", 567
    Set LoadRawStyleSheet_Styles_5_ParagraphFormat_Indent = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_5_ParagraphFormat_Indent", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_5_ParagraphFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "keepNext", True
    d.Add "keepLines", True
    d.Add "spacing", LoadRawStyleSheet_Styles_5_ParagraphFormat_Spacing()
    d.Add "indent", LoadRawStyleSheet_Styles_5_ParagraphFormat_Indent()
    d.Add "jc", "both"
    d.Add "outlineLvl", 4
    Set LoadRawStyleSheet_Styles_5_ParagraphFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_5_ParagraphFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_5_RunFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "b", True
    d.Add "bCs", True
    d.Add "sz", "${SZ_MAIN}"
    d.Add "szCs", "${SZ_MAIN}"
    Set LoadRawStyleSheet_Styles_5_RunFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_5_RunFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_6_ParagraphFormat_Spacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "before", 240
    d.Add "after", 0
    d.Add "line", 240
    d.Add "lineRule", "auto"
    Set LoadRawStyleSheet_Styles_6_ParagraphFormat_Spacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_6_ParagraphFormat_Spacing", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_6_ParagraphFormat_Indent() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "firstLine", 567
    Set LoadRawStyleSheet_Styles_6_ParagraphFormat_Indent = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_6_ParagraphFormat_Indent", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_6_ParagraphFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "keepNext", True
    d.Add "keepLines", True
    d.Add "spacing", LoadRawStyleSheet_Styles_6_ParagraphFormat_Spacing()
    d.Add "indent", LoadRawStyleSheet_Styles_6_ParagraphFormat_Indent()
    d.Add "jc", "both"
    d.Add "outlineLvl", 5
    Set LoadRawStyleSheet_Styles_6_ParagraphFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_6_ParagraphFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_6_RunFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "b", True
    d.Add "bCs", True
    d.Add "sz", "${SZ_MAIN}"
    d.Add "szCs", "${SZ_MAIN}"
    Set LoadRawStyleSheet_Styles_6_RunFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_6_RunFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_7_ParagraphFormat_Spacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "before", 480
    d.Add "after", 0
    d.Add "line", 240
    d.Add "lineRule", "auto"
    Set LoadRawStyleSheet_Styles_7_ParagraphFormat_Spacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_7_ParagraphFormat_Spacing", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_7_ParagraphFormat_Indent() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "firstLine", 0
    Set LoadRawStyleSheet_Styles_7_ParagraphFormat_Indent = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_7_ParagraphFormat_Indent", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_7_ParagraphFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "keepNext", True
    d.Add "keepLines", True
    d.Add "spacing", LoadRawStyleSheet_Styles_7_ParagraphFormat_Spacing()
    d.Add "indent", LoadRawStyleSheet_Styles_7_ParagraphFormat_Indent()
    d.Add "jc", "center"
    d.Add "outlineLvl", 9
    Set LoadRawStyleSheet_Styles_7_ParagraphFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_7_ParagraphFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_7_RunFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "b", True
    d.Add "bCs", True
    d.Add "color", "auto"
    d.Add "sz", "${SZ_MAIN}"
    d.Add "szCs", "${SZ_MAIN}"
    Set LoadRawStyleSheet_Styles_7_RunFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_7_RunFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_8_ParagraphFormat_Spacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "before", 0
    d.Add "after", 0
    d.Add "line", 240
    d.Add "lineRule", "auto"
    Set LoadRawStyleSheet_Styles_8_ParagraphFormat_Spacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_8_ParagraphFormat_Spacing", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_8_ParagraphFormat_Indent() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "firstLine", 0
    Set LoadRawStyleSheet_Styles_8_ParagraphFormat_Indent = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_8_ParagraphFormat_Indent", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_8_ParagraphFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "keepNext", True
    d.Add "keepLines", True
    d.Add "spacing", LoadRawStyleSheet_Styles_8_ParagraphFormat_Spacing()
    d.Add "indent", LoadRawStyleSheet_Styles_8_ParagraphFormat_Indent()
    d.Add "jc", "center"
    d.Add "outlineLvl", 9
    Set LoadRawStyleSheet_Styles_8_ParagraphFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_8_ParagraphFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_8_RunFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "b", True
    d.Add "bCs", True
    d.Add "color", "auto"
    d.Add "sz", "${SZ_MAIN}"
    d.Add "szCs", "${SZ_MAIN}"
    Set LoadRawStyleSheet_Styles_8_RunFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_8_RunFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_9_ParagraphFormat_Spacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "before", 0
    d.Add "after", 0
    d.Add "line", 240
    d.Add "lineRule", "auto"
    Set LoadRawStyleSheet_Styles_9_ParagraphFormat_Spacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_9_ParagraphFormat_Spacing", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_9_ParagraphFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "spacing", LoadRawStyleSheet_Styles_9_ParagraphFormat_Spacing()
    Set LoadRawStyleSheet_Styles_9_ParagraphFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_9_ParagraphFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_9_TableFormat_Borders_Top() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "val", "single"
    d.Add "sz", 4
    d.Add "space", 0
    d.Add "color", "auto"
    Set LoadRawStyleSheet_Styles_9_TableFormat_Borders_Top = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_9_TableFormat_Borders_Top", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_9_TableFormat_Borders_Left() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "val", "single"
    d.Add "sz", 4
    d.Add "space", 0
    d.Add "color", "auto"
    Set LoadRawStyleSheet_Styles_9_TableFormat_Borders_Left = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_9_TableFormat_Borders_Left", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_9_TableFormat_Borders_Bottom() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "val", "single"
    d.Add "sz", 4
    d.Add "space", 0
    d.Add "color", "auto"
    Set LoadRawStyleSheet_Styles_9_TableFormat_Borders_Bottom = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_9_TableFormat_Borders_Bottom", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_9_TableFormat_Borders_Right() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "val", "single"
    d.Add "sz", 4
    d.Add "space", 0
    d.Add "color", "auto"
    Set LoadRawStyleSheet_Styles_9_TableFormat_Borders_Right = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_9_TableFormat_Borders_Right", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_9_TableFormat_Borders_InsideH() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "val", "single"
    d.Add "sz", 4
    d.Add "space", 0
    d.Add "color", "auto"
    Set LoadRawStyleSheet_Styles_9_TableFormat_Borders_InsideH = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_9_TableFormat_Borders_InsideH", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_9_TableFormat_Borders_InsideV() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "val", "single"
    d.Add "sz", 4
    d.Add "space", 0
    d.Add "color", "auto"
    Set LoadRawStyleSheet_Styles_9_TableFormat_Borders_InsideV = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_9_TableFormat_Borders_InsideV", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_9_TableFormat_Borders() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "top", LoadRawStyleSheet_Styles_9_TableFormat_Borders_Top()
    d.Add "left", LoadRawStyleSheet_Styles_9_TableFormat_Borders_Left()
    d.Add "bottom", LoadRawStyleSheet_Styles_9_TableFormat_Borders_Bottom()
    d.Add "right", LoadRawStyleSheet_Styles_9_TableFormat_Borders_Right()
    d.Add "insideH", LoadRawStyleSheet_Styles_9_TableFormat_Borders_InsideH()
    d.Add "insideV", LoadRawStyleSheet_Styles_9_TableFormat_Borders_InsideV()
    Set LoadRawStyleSheet_Styles_9_TableFormat_Borders = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_9_TableFormat_Borders", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_9_TableFormat_FirstRow_ParagraphFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "keepNext", True
    d.Add "jc", "center"
    Set LoadRawStyleSheet_Styles_9_TableFormat_FirstRow_ParagraphFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_9_TableFormat_FirstRow_ParagraphFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_9_TableFormat_FirstRow_RunFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "b", True
    d.Add "bCs", True
    Set LoadRawStyleSheet_Styles_9_TableFormat_FirstRow_RunFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_9_TableFormat_FirstRow_RunFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_9_TableFormat_FirstRow() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "paragraphFormat", LoadRawStyleSheet_Styles_9_TableFormat_FirstRow_ParagraphFormat()
    d.Add "runFormat", LoadRawStyleSheet_Styles_9_TableFormat_FirstRow_RunFormat()
    d.Add "tblHeader", True
    d.Add "vAlign", "center"
    Set LoadRawStyleSheet_Styles_9_TableFormat_FirstRow = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_9_TableFormat_FirstRow", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_9_TableFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "borders", LoadRawStyleSheet_Styles_9_TableFormat_Borders()
    d.Add "firstRow", LoadRawStyleSheet_Styles_9_TableFormat_FirstRow()
    Set LoadRawStyleSheet_Styles_9_TableFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_9_TableFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_10_ParagraphFormat_Spacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "before", 0
    d.Add "after", 0
    d.Add "line", 240
    d.Add "lineRule", "auto"
    Set LoadRawStyleSheet_Styles_10_ParagraphFormat_Spacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_10_ParagraphFormat_Spacing", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_10_ParagraphFormat_Indent() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "left", 0
    d.Add "right", 0
    d.Add "firstLine", 0
    Set LoadRawStyleSheet_Styles_10_ParagraphFormat_Indent = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_10_ParagraphFormat_Indent", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_10_ParagraphFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "spacing", LoadRawStyleSheet_Styles_10_ParagraphFormat_Spacing()
    d.Add "indent", LoadRawStyleSheet_Styles_10_ParagraphFormat_Indent()
    d.Add "jc", "left"
    Set LoadRawStyleSheet_Styles_10_ParagraphFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_10_ParagraphFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_10_RunFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "sz", "${SZ_MAIN}"
    d.Add "szCs", "${SZ_MAIN}"
    Set LoadRawStyleSheet_Styles_10_RunFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_10_RunFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_11_ParagraphFormat_Spacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "before", 120
    d.Add "after", 120
    d.Add "line", 240
    d.Add "lineRule", "auto"
    Set LoadRawStyleSheet_Styles_11_ParagraphFormat_Spacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_11_ParagraphFormat_Spacing", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_11_ParagraphFormat_Indent() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "firstLine", 0
    Set LoadRawStyleSheet_Styles_11_ParagraphFormat_Indent = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_11_ParagraphFormat_Indent", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_11_ParagraphFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "spacing", LoadRawStyleSheet_Styles_11_ParagraphFormat_Spacing()
    d.Add "indent", LoadRawStyleSheet_Styles_11_ParagraphFormat_Indent()
    d.Add "jc", "center"
    Set LoadRawStyleSheet_Styles_11_ParagraphFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_11_ParagraphFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_11_RunFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "b", True
    d.Add "bCs", True
    d.Add "color", "auto"
    d.Add "sz", "${SZ_MAIN}"
    d.Add "szCs", "${SZ_MAIN}"
    Set LoadRawStyleSheet_Styles_11_RunFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_11_RunFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_12_ParagraphFormat_Spacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "before", 0
    d.Add "after", 0
    d.Add "line", 240
    d.Add "lineRule", "auto"
    Set LoadRawStyleSheet_Styles_12_ParagraphFormat_Spacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_12_ParagraphFormat_Spacing", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_12_ParagraphFormat_Indent() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "firstLine", 0
    Set LoadRawStyleSheet_Styles_12_ParagraphFormat_Indent = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_12_ParagraphFormat_Indent", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_12_ParagraphFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "spacing", LoadRawStyleSheet_Styles_12_ParagraphFormat_Spacing()
    d.Add "indent", LoadRawStyleSheet_Styles_12_ParagraphFormat_Indent()
    d.Add "jc", "center"
    Set LoadRawStyleSheet_Styles_12_ParagraphFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_12_ParagraphFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_12_RunFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "sz", "${SZ_MAIN}"
    d.Add "szCs", "${SZ_MAIN}"
    Set LoadRawStyleSheet_Styles_12_RunFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_12_RunFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_13_RunFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "font", "Times New Roman"
    d.Add "sz", "${SZ_MAIN}"
    d.Add "szCs", "${SZ_MAIN}"
    Set LoadRawStyleSheet_Styles_13_RunFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_13_RunFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_14_ParagraphFormat_Spacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "before", 0
    d.Add "after", 0
    d.Add "line", 240
    d.Add "lineRule", "auto"
    Set LoadRawStyleSheet_Styles_14_ParagraphFormat_Spacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_14_ParagraphFormat_Spacing", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_14_ParagraphFormat_Indent() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "firstLine", 567
    Set LoadRawStyleSheet_Styles_14_ParagraphFormat_Indent = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_14_ParagraphFormat_Indent", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_14_ParagraphFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "spacing", LoadRawStyleSheet_Styles_14_ParagraphFormat_Spacing()
    d.Add "indent", LoadRawStyleSheet_Styles_14_ParagraphFormat_Indent()
    d.Add "jc", "both"
    Set LoadRawStyleSheet_Styles_14_ParagraphFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_14_ParagraphFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_14_RunFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "sz", 22
    d.Add "szCs", 22
    Set LoadRawStyleSheet_Styles_14_RunFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_14_RunFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_15_ParagraphFormat_Spacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "before", 480
    d.Add "after", 240
    d.Add "line", 240
    d.Add "lineRule", "auto"
    Set LoadRawStyleSheet_Styles_15_ParagraphFormat_Spacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_15_ParagraphFormat_Spacing", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_15_ParagraphFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "spacing", LoadRawStyleSheet_Styles_15_ParagraphFormat_Spacing()
    d.Add "outlineLvl", 9
    Set LoadRawStyleSheet_Styles_15_ParagraphFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_15_ParagraphFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_16_ParagraphFormat_Tabs() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim d18 As Object
    Set d18 = CreateObject("Scripting.Dictionary")
    d18.Add "val", "right"
    d18.Add "leader", "dot"
    d18.Add "pos", 8788
    c.Add d18
    Set LoadRawStyleSheet_Styles_16_ParagraphFormat_Tabs = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_16_ParagraphFormat_Tabs", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_16_ParagraphFormat_Spacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "before", 120
    d.Add "after", 0
    d.Add "line", 240
    d.Add "lineRule", "auto"
    Set LoadRawStyleSheet_Styles_16_ParagraphFormat_Spacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_16_ParagraphFormat_Spacing", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_16_ParagraphFormat_Indent() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "left", 0
    d.Add "right", 0
    d.Add "firstLine", 0
    Set LoadRawStyleSheet_Styles_16_ParagraphFormat_Indent = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_16_ParagraphFormat_Indent", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_16_ParagraphFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "tabs", LoadRawStyleSheet_Styles_16_ParagraphFormat_Tabs()
    d.Add "spacing", LoadRawStyleSheet_Styles_16_ParagraphFormat_Spacing()
    d.Add "indent", LoadRawStyleSheet_Styles_16_ParagraphFormat_Indent()
    d.Add "jc", "left"
    Set LoadRawStyleSheet_Styles_16_ParagraphFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_16_ParagraphFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_16_RunFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "b", True
    d.Add "bCs", True
    d.Add "sz", "${SZ_MAIN}"
    d.Add "szCs", "${SZ_MAIN}"
    Set LoadRawStyleSheet_Styles_16_RunFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_16_RunFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_17_ParagraphFormat_Tabs() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim d20 As Object
    Set d20 = CreateObject("Scripting.Dictionary")
    d20.Add "val", "right"
    d20.Add "leader", "dot"
    d20.Add "pos", 8788
    c.Add d20
    Set LoadRawStyleSheet_Styles_17_ParagraphFormat_Tabs = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_17_ParagraphFormat_Tabs", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_17_ParagraphFormat_Spacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "before", 0
    d.Add "after", 0
    d.Add "line", 240
    d.Add "lineRule", "auto"
    Set LoadRawStyleSheet_Styles_17_ParagraphFormat_Spacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_17_ParagraphFormat_Spacing", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_17_ParagraphFormat_Indent() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "left", 284
    d.Add "right", 0
    d.Add "firstLine", 0
    Set LoadRawStyleSheet_Styles_17_ParagraphFormat_Indent = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_17_ParagraphFormat_Indent", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_17_ParagraphFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "tabs", LoadRawStyleSheet_Styles_17_ParagraphFormat_Tabs()
    d.Add "spacing", LoadRawStyleSheet_Styles_17_ParagraphFormat_Spacing()
    d.Add "indent", LoadRawStyleSheet_Styles_17_ParagraphFormat_Indent()
    d.Add "jc", "left"
    Set LoadRawStyleSheet_Styles_17_ParagraphFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_17_ParagraphFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_17_RunFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "sz", "${SZ_MAIN}"
    d.Add "szCs", "${SZ_MAIN}"
    Set LoadRawStyleSheet_Styles_17_RunFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_17_RunFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_18_ParagraphFormat_Tabs() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim d22 As Object
    Set d22 = CreateObject("Scripting.Dictionary")
    d22.Add "val", "right"
    d22.Add "leader", "dot"
    d22.Add "pos", 8788
    c.Add d22
    Set LoadRawStyleSheet_Styles_18_ParagraphFormat_Tabs = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_18_ParagraphFormat_Tabs", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_18_ParagraphFormat_Spacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "before", 0
    d.Add "after", 0
    d.Add "line", 240
    d.Add "lineRule", "auto"
    Set LoadRawStyleSheet_Styles_18_ParagraphFormat_Spacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_18_ParagraphFormat_Spacing", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_18_ParagraphFormat_Indent() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "left", 567
    d.Add "right", 0
    d.Add "firstLine", 0
    Set LoadRawStyleSheet_Styles_18_ParagraphFormat_Indent = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_18_ParagraphFormat_Indent", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_18_ParagraphFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "tabs", LoadRawStyleSheet_Styles_18_ParagraphFormat_Tabs()
    d.Add "spacing", LoadRawStyleSheet_Styles_18_ParagraphFormat_Spacing()
    d.Add "indent", LoadRawStyleSheet_Styles_18_ParagraphFormat_Indent()
    d.Add "jc", "left"
    Set LoadRawStyleSheet_Styles_18_ParagraphFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_18_ParagraphFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles_18_RunFormat() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "sz", "${SZ_MAIN}"
    d.Add "szCs", "${SZ_MAIN}"
    Set LoadRawStyleSheet_Styles_18_RunFormat = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles_18_RunFormat", Err.description
End Function

Private Function LoadRawStyleSheet_Styles() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim d1 As Object
    Set d1 = CreateObject("Scripting.Dictionary")
    d1.Add "styleId", "Normal"
    d1.Add "name", "Normal"
    d1.Add "type", "paragraph"
    d1.Add "group", "A"
    d1.Add "role", "N" & ChrW(&H1ED9) & "i dung v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n (" & ChrW(&HF4) & " 6)"
    d1.Add "default", True
    d1.Add "uiPriority", 0
    d1.Add "qFormat", True
    d1.Add "paragraphFormat", LoadRawStyleSheet_Styles_0_ParagraphFormat()
    d1.Add "runFormat", LoadRawStyleSheet_Styles_0_RunFormat()
    d1.Add "note", "Style nen cua toan bo template. C" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & " phai bang c" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & " noi dung; jc=both theo yeu cau canh deu hai le; after=0 vi khoang cach giua doan chi dat o mot trong hai (before hoac after), khong duoc cong don."
    c.Add d1
    Dim d2 As Object
    Set d2 = CreateObject("Scripting.Dictionary")
    d2.Add "styleId", "Heading1"
    d2.Add "name", "heading 1"
    d2.Add "type", "paragraph"
    d2.Add "group", "A"
    d2.Add "role", "Ti" & ChrW(&HEA) & "u " & ChrW(&H111) & ChrW(&H1EC1) & " c" & ChrW(&H1EE7) & "a Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c v" & ChrW(&HE0) & " Ti" & ChrW(&HEA) & "u " & ChrW(&H111) & ChrW(&H1EC1) & " c" & ChrW(&H1EE7) & "a Ph" & ChrW(&H1EA7) & "n"
    d2.Add "basedOn", "Normal"
    d2.Add "next", "Normal"
    d2.Add "link", "Heading1Char"
    d2.Add "uiPriority", 9
    d2.Add "qFormat", True
    d2.Add "paragraphFormat", LoadRawStyleSheet_Styles_1_ParagraphFormat()
    d2.Add "runFormat", LoadRawStyleSheet_Styles_1_RunFormat()
    d2.Add "note", "Phuc vu hai vai tro vi N" & ChrW(&H110) & " 30 quy dinh dinh dang giong het nhau: tieu de Phan trong than van ban, va tieu de Phu luc."
    c.Add d2
    Dim d3 As Object
    Set d3 = CreateObject("Scripting.Dictionary")
    d3.Add "styleId", "Heading2"
    d3.Add "name", "heading 2"
    d3.Add "type", "paragraph"
    d3.Add "group", "A"
    d3.Add "role", "Ti" & ChrW(&HEA) & "u " & ChrW(&H111) & ChrW(&H1EC1) & " c" & ChrW(&H1EE7) & "a Ch" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    d3.Add "basedOn", "Normal"
    d3.Add "next", "Normal"
    d3.Add "link", "Heading2Char"
    d3.Add "uiPriority", 9
    d3.Add "qFormat", True
    d3.Add "paragraphFormat", LoadRawStyleSheet_Styles_2_ParagraphFormat()
    d3.Add "runFormat", LoadRawStyleSheet_Styles_2_RunFormat()
    c.Add d3
    Dim d4 As Object
    Set d4 = CreateObject("Scripting.Dictionary")
    d4.Add "styleId", "Heading3"
    d4.Add "name", "heading 3"
    d4.Add "type", "paragraph"
    d4.Add "group", "A"
    d4.Add "role", "Ti" & ChrW(&HEA) & "u " & ChrW(&H111) & ChrW(&H1EC1) & " c" & ChrW(&H1EE7) & "a M" & ChrW(&H1EE5) & "c"
    d4.Add "basedOn", "Normal"
    d4.Add "next", "Normal"
    d4.Add "link", "Heading3Char"
    d4.Add "uiPriority", 9
    d4.Add "qFormat", True
    d4.Add "paragraphFormat", LoadRawStyleSheet_Styles_3_ParagraphFormat()
    d4.Add "runFormat", LoadRawStyleSheet_Styles_3_RunFormat()
    c.Add d4
    Dim d5 As Object
    Set d5 = CreateObject("Scripting.Dictionary")
    d5.Add "styleId", "Heading4"
    d5.Add "name", "heading 4"
    d5.Add "type", "paragraph"
    d5.Add "group", "A"
    d5.Add "role", "Ti" & ChrW(&HEA) & "u " & ChrW(&H111) & ChrW(&H1EC1) & " c" & ChrW(&H1EE7) & "a Ti" & ChrW(&H1EC3) & "u m" & ChrW(&H1EE5) & "c"
    d5.Add "basedOn", "Normal"
    d5.Add "next", "Normal"
    d5.Add "link", "Heading4Char"
    d5.Add "uiPriority", 9
    d5.Add "qFormat", True
    d5.Add "paragraphFormat", LoadRawStyleSheet_Styles_4_ParagraphFormat()
    d5.Add "runFormat", LoadRawStyleSheet_Styles_4_RunFormat()
    c.Add d5
    Dim d6 As Object
    Set d6 = CreateObject("Scripting.Dictionary")
    d6.Add "styleId", "Heading5"
    d6.Add "name", "heading 5"
    d6.Add "type", "paragraph"
    d6.Add "group", "A"
    d6.Add "role", ChrW(&H110) & "i" & ChrW(&H1EC1) & "u"
    d6.Add "basedOn", "Normal"
    d6.Add "next", "Normal"
    d6.Add "link", "Heading5Char"
    d6.Add "uiPriority", 9
    d6.Add "qFormat", True
    d6.Add "paragraphFormat", LoadRawStyleSheet_Styles_5_ParagraphFormat()
    d6.Add "runFormat", LoadRawStyleSheet_Styles_5_RunFormat()
    d6.Add "note", "Khac Heading1-Heading4 o ba diem theo dung N" & ChrW(&H110) & " 30: chu in thuong, lui dau dong 1 cm thay vi canh giua, giu jc=both."
    c.Add d6
    Dim d7 As Object
    Set d7 = CreateObject("Scripting.Dictionary")
    d7.Add "styleId", "Heading6"
    d7.Add "name", "heading 6"
    d7.Add "type", "paragraph"
    d7.Add "group", "A"
    d7.Add "role", "Kho" & ChrW(&H1EA3) & "n c" & ChrW(&HF3) & " ti" & ChrW(&HEA) & "u " & ChrW(&H111) & ChrW(&H1EC1)
    d7.Add "basedOn", "Normal"
    d7.Add "next", "Normal"
    d7.Add "link", "Heading6Char"
    d7.Add "uiPriority", 9
    d7.Add "qFormat", True
    d7.Add "paragraphFormat", LoadRawStyleSheet_Styles_6_ParagraphFormat()
    d7.Add "runFormat", LoadRawStyleSheet_Styles_6_RunFormat()
    d7.Add "note", "Chi dung cho khoan CO tieu de. Khoan khong co tieu de dung Normal."
    c.Add d7
    Dim d8 As Object
    Set d8 = CreateObject("Scripting.Dictionary")
    d8.Add "styleId", "Title"
    d8.Add "name", "Title"
    d8.Add "type", "paragraph"
    d8.Add "group", "A"
    d8.Add "role", "T" & ChrW(&HEA) & "n lo" & ChrW(&H1EA1) & "i v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n (" & ChrW(&HF4) & " 5a)"
    d8.Add "basedOn", "Normal"
    d8.Add "next", "Subtitle"
    d8.Add "link", "TitleChar"
    d8.Add "uiPriority", 10
    d8.Add "qFormat", True
    d8.Add "paragraphFormat", LoadRawStyleSheet_Styles_7_ParagraphFormat()
    d8.Add "runFormat", LoadRawStyleSheet_Styles_7_RunFormat()
    d8.Add "note", "next=Subtitle tao chuoi tu dong Ten loai -> Trich yeu."
    c.Add d8
    Dim d9 As Object
    Set d9 = CreateObject("Scripting.Dictionary")
    d9.Add "styleId", "Subtitle"
    d9.Add "name", "Subtitle"
    d9.Add "type", "paragraph"
    d9.Add "group", "A"
    d9.Add "role", "Tr" & ChrW(&HED) & "ch y" & ChrW(&H1EBF) & "u n" & ChrW(&H1ED9) & "i dung v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n (" & ChrW(&HF4) & " 5a)"
    d9.Add "basedOn", "Normal"
    d9.Add "next", "Normal"
    d9.Add "link", "SubtitleChar"
    d9.Add "uiPriority", 11
    d9.Add "qFormat", True
    d9.Add "paragraphFormat", LoadRawStyleSheet_Styles_8_ParagraphFormat()
    d9.Add "runFormat", LoadRawStyleSheet_Styles_8_RunFormat()
    d9.Add "note", "before=0 vi trich yeu luon nam ngay duoi ten loai van ban."
    c.Add d9
    Dim d10 As Object
    Set d10 = CreateObject("Scripting.Dictionary")
    d10.Add "styleId", "TableGrid"
    d10.Add "name", "Table Grid"
    d10.Add "type", "table"
    d10.Add "group", "B"
    d10.Add "role", "B" & ChrW(&H1EA3) & "ng d" & ChrW(&H1EEF) & " li" & ChrW(&H1EC7) & "u trong ph" & ChrW(&H1EA7) & "n th" & ChrW(&HE2) & "n"
    d10.Add "basedOn", "TableNormal"
    d10.Add "uiPriority", 39
    d10.Add "paragraphFormat", LoadRawStyleSheet_Styles_9_ParagraphFormat()
    d10.Add "tableFormat", LoadRawStyleSheet_Styles_9_TableFormat()
    d10.Add "note", "Vien sz=4 (0,5pt) mau auto la duong ke mac dinh cua Word. VAN DE CHUA XAC MINH: w:tblHeader khai trong tblStylePr[firstRow]/trPr co duoc Word lap lai dong dau khi sang trang khong - can kiem chung bang bang du lieu dai hon mot trang (xem " & ChrW(&HA7) & "2.4 B2 cua tai lieu nguon). Neu khong, bo sung w:tblHeader truc tiep vao trPr cua dong dau tung bang."
    c.Add d10
    Dim d11 As Object
    Set d11 = CreateObject("Scripting.Dictionary")
    d11.Add "styleId", "NoSpacing"
    d11.Add "name", "No Spacing"
    d11.Add "type", "paragraph"
    d11.Add "group", "B"
    d11.Add "role", "M" & ChrW(&H1ECD) & "i " & ChrW(&H111) & "o" & ChrW(&H1EA1) & "n v" & ChrW(&H103) & "n n" & ChrW(&H1EB1) & "m trong " & ChrW(&HF4) & " b" & ChrW(&H1EA3) & "ng"
    d11.Add "latent", True
    d11.Add "basedOn", "Normal"
    d11.Add "next", "NoSpacing"
    d11.Add "uiPriority", 1
    d11.Add "qFormat", True
    d11.Add "paragraphFormat", LoadRawStyleSheet_Styles_10_ParagraphFormat()
    d11.Add "runFormat", LoadRawStyleSheet_Styles_10_RunFormat()
    d11.Add "note", "BAT BUOC ap cho moi doan van nam trong o bang. Ly do: paragraph style nam SAU table style trong thu tu ap dung OOXML, nen Normal se ghi de TableGrid neu khong ap NoSpacing rieng cho tung doan."
    c.Add d11
    Dim d12 As Object
    Set d12 = CreateObject("Scripting.Dictionary")
    d12.Add "styleId", "Caption"
    d12.Add "name", "caption"
    d12.Add "type", "paragraph"
    d12.Add "group", "B"
    d12.Add "role", "Ch" & ChrW(&HFA) & " th" & ChrW(&HED) & "ch B" & ChrW(&H1EA3) & "ng v" & ChrW(&HE0) & " ch" & ChrW(&HFA) & " th" & ChrW(&HED) & "ch H" & ChrW(&HEC) & "nh"
    d12.Add "latent", True
    d12.Add "basedOn", "Normal"
    d12.Add "next", "Normal"
    d12.Add "uiPriority", 35
    d12.Add "qFormat", True
    d12.Add "paragraphFormat", LoadRawStyleSheet_Styles_11_ParagraphFormat()
    d12.Add "runFormat", LoadRawStyleSheet_Styles_11_RunFormat()
    d12.Add "note", "Khong mang keepNext. Chu thich Bang (phia tren bang) can them keepNext tren chinh doan chu thich; chu thich Hinh (phia duoi hinh) can dat keepNext tren doan CHUA hinh, khong dat tren doan chu thich."
    c.Add d12
    Dim d13 As Object
    Set d13 = CreateObject("Scripting.Dictionary")
    d13.Add "styleId", "Header"
    d13.Add "name", "header"
    d13.Add "type", "paragraph"
    d13.Add "group", "C"
    d13.Add "role", ChrW(&H110) & "o" & ChrW(&H1EA1) & "n v" & ChrW(&H103) & "n ch" & ChrW(&H1EE9) & "a s" & ChrW(&H1ED1) & " trang, t" & ChrW(&H1EA1) & "i l" & ChrW(&H1EC1) & " tr" & ChrW(&HEA) & "n"
    d13.Add "latent", True
    d13.Add "basedOn", "Normal"
    d13.Add "next", "Normal"
    d13.Add "uiPriority", 99
    d13.Add "unhideWhenUsed", True
    d13.Add "paragraphFormat", LoadRawStyleSheet_Styles_12_ParagraphFormat()
    d13.Add "runFormat", LoadRawStyleSheet_Styles_12_RunFormat()
    d13.Add "note", "QUYET DINH 2026-08-11 (docs/design/02-dac-ta-giao-dien.md muc 6.4): BO HAN khoi <w:tabs> ma huong dan OOXML goc dat tai 4394/8788 twip. Ly do: hai gia tri do tinh theo vung trinh bay DOC 155mm, sai o section trang NGANG (242mm); N" & ChrW(&H110) & " 30 chi yeu cau canh giua theo chieu ngang, da do jc=center dam nhiem tron ven; bo tab thi style dung o moi huong giay ma van giu dung 19 style. Danh doi: ai muon them chu vao header phai tu dat tab, ngoai pham vi v1."
    c.Add d13
    Dim d14 As Object
    Set d14 = CreateObject("Scripting.Dictionary")
    d14.Add "styleId", "PageNumber"
    d14.Add "name", "page number"
    d14.Add "type", "character"
    d14.Add "group", "C"
    d14.Add "role", "Run ch" & ChrW(&H1EE9) & "a field PAGE"
    d14.Add "latent", True
    d14.Add "basedOn", "DefaultParagraphFont"
    d14.Add "uiPriority", 99
    d14.Add "unhideWhenUsed", True
    d14.Add "runFormat", LoadRawStyleSheet_Styles_13_RunFormat()
    d14.Add "note", "Character style, khong mang duoc jc/ind. Bo sung cho Header (lo vi tri doan) chu khong thay the: Header lo vi tri, PageNumber lo phong chu va co chu cua con so."
    c.Add d14
    Dim d15 As Object
    Set d15 = CreateObject("Scripting.Dictionary")
    d15.Add "styleId", "FootnoteText"
    d15.Add "name", "footnote text"
    d15.Add "type", "paragraph"
    d15.Add "group", "C"
    d15.Add "role", "Ghi ch" & ChrW(&HFA) & " ch" & ChrW(&HE2) & "n trang"
    d15.Add "latent", True
    d15.Add "basedOn", "Normal"
    d15.Add "next", "FootnoteText"
    d15.Add "uiPriority", 99
    d15.Add "semiHidden", True
    d15.Add "unhideWhenUsed", True
    d15.Add "paragraphFormat", LoadRawStyleSheet_Styles_14_ParagraphFormat()
    d15.Add "runFormat", LoadRawStyleSheet_Styles_14_RunFormat()
    d15.Add "note", "N" & ChrW(&H110) & " 30 khong quy dinh ghi chu chan trang. Co 11pt (sz=22, GHI CUNG khong dung token SZ_MAIN) chon theo nhom chu nho nhat ma N" & ChrW(&H110) & " 30 su dung - nhom C."
    c.Add d15
    Dim d16 As Object
    Set d16 = CreateObject("Scripting.Dictionary")
    d16.Add "styleId", "TOCHeading"
    d16.Add "name", "TOC Heading"
    d16.Add "type", "paragraph"
    d16.Add "group", "D"
    d16.Add "role", "Ti" & ChrW(&HEA) & "u " & ChrW(&H111) & ChrW(&H1EC1) & " m" & ChrW(&H1EE5) & "c l" & ChrW(&H1EE5) & "c"
    d16.Add "latent", True
    d16.Add "basedOn", "Heading1"
    d16.Add "next", "Normal"
    d16.Add "uiPriority", 39
    d16.Add "qFormat", True
    d16.Add "paragraphFormat", LoadRawStyleSheet_Styles_15_ParagraphFormat()
    d16.Add "note", "Ke thua keepNext, keepLines, canh giua, in dam va co chu tu Heading1 (basedOn); chi ghi de spacing va outlineLvl. GIU basedOn=Heading1 va outlineLvl=9 dung mac dinh cua Word - dong 'MUC LUC' KHONG tu chen minh vao muc luc vi outlineLvl=9 khong phai gia tri ke thua tu Heading1 (outlineLvl=0)."
    c.Add d16
    Dim d17 As Object
    Set d17 = CreateObject("Scripting.Dictionary")
    d17.Add "styleId", "TOC1"
    d17.Add "name", "toc 1"
    d17.Add "type", "paragraph"
    d17.Add "group", "D"
    d17.Add "role", "Th" & ChrW(&HE2) & "n m" & ChrW(&H1EE5) & "c l" & ChrW(&H1EE5) & "c c" & ChrW(&H1EA5) & "p 1"
    d17.Add "latent", True
    d17.Add "basedOn", "Normal"
    d17.Add "next", "Normal"
    d17.Add "autoRedefine", True
    d17.Add "uiPriority", 39
    d17.Add "unhideWhenUsed", True
    d17.Add "paragraphFormat", LoadRawStyleSheet_Styles_16_ParagraphFormat()
    d17.Add "runFormat", LoadRawStyleSheet_Styles_16_RunFormat()
    c.Add d17
    Dim d19 As Object
    Set d19 = CreateObject("Scripting.Dictionary")
    d19.Add "styleId", "TOC2"
    d19.Add "name", "toc 2"
    d19.Add "type", "paragraph"
    d19.Add "group", "D"
    d19.Add "role", "Th" & ChrW(&HE2) & "n m" & ChrW(&H1EE5) & "c l" & ChrW(&H1EE5) & "c c" & ChrW(&H1EA5) & "p 2"
    d19.Add "latent", True
    d19.Add "basedOn", "Normal"
    d19.Add "next", "Normal"
    d19.Add "autoRedefine", True
    d19.Add "uiPriority", 39
    d19.Add "unhideWhenUsed", True
    d19.Add "paragraphFormat", LoadRawStyleSheet_Styles_17_ParagraphFormat()
    d19.Add "runFormat", LoadRawStyleSheet_Styles_17_RunFormat()
    c.Add d19
    Dim d21 As Object
    Set d21 = CreateObject("Scripting.Dictionary")
    d21.Add "styleId", "TOC3"
    d21.Add "name", "toc 3"
    d21.Add "type", "paragraph"
    d21.Add "group", "D"
    d21.Add "role", "Th" & ChrW(&HE2) & "n m" & ChrW(&H1EE5) & "c l" & ChrW(&H1EE5) & "c c" & ChrW(&H1EA5) & "p 3"
    d21.Add "latent", True
    d21.Add "basedOn", "Normal"
    d21.Add "next", "Normal"
    d21.Add "autoRedefine", True
    d21.Add "uiPriority", 39
    d21.Add "unhideWhenUsed", True
    d21.Add "paragraphFormat", LoadRawStyleSheet_Styles_18_ParagraphFormat()
    d21.Add "runFormat", LoadRawStyleSheet_Styles_18_RunFormat()
    c.Add d21
    Set LoadRawStyleSheet_Styles = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_Styles", Err.description
End Function

Private Function LoadRawStyleSheet_CharacterStyleSync_Affects() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "Heading1"
    c.Add "Heading2"
    c.Add "Heading3"
    c.Add "Heading4"
    c.Add "Heading5"
    c.Add "Heading6"
    c.Add "Title"
    c.Add "Subtitle"
    Set LoadRawStyleSheet_CharacterStyleSync_Affects = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_CharacterStyleSync_Affects", Err.description
End Function

Private Function LoadRawStyleSheet_CharacterStyleSync_Template() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "type", "character"
    d.Add "customStyle", True
    d.Add "basedOn", "DefaultParagraphFont"
    d.Add "uiPriority", 9
    d.Add "semiHidden", True
    Set LoadRawStyleSheet_CharacterStyleSync_Template = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_CharacterStyleSync_Template", Err.description
End Function

Private Function LoadRawStyleSheet_CharacterStyleSync() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "namingPattern", "<styleId>Char"
    d.Add "affects", LoadRawStyleSheet_CharacterStyleSync_Affects()
    d.Add "template", LoadRawStyleSheet_CharacterStyleSync_Template()
    Set LoadRawStyleSheet_CharacterStyleSync = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_CharacterStyleSync", Err.description
End Function

Private Function LoadRawStyleSheet_LatentStyleActivation() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim d23 As Object
    Set d23 = CreateObject("Scripting.Dictionary")
    d23.Add "name", "header"
    d23.Add "uiPriority", 99
    d23.Add "unhideWhenUsed", True
    c.Add d23
    Dim d24 As Object
    Set d24 = CreateObject("Scripting.Dictionary")
    d24.Add "name", "page number"
    d24.Add "uiPriority", 99
    d24.Add "unhideWhenUsed", True
    c.Add d24
    Dim d25 As Object
    Set d25 = CreateObject("Scripting.Dictionary")
    d25.Add "name", "No Spacing"
    d25.Add "uiPriority", 1
    d25.Add "qFormat", True
    c.Add d25
    Dim d26 As Object
    Set d26 = CreateObject("Scripting.Dictionary")
    d26.Add "name", "Table Grid"
    d26.Add "uiPriority", 39
    c.Add d26
    Dim d27 As Object
    Set d27 = CreateObject("Scripting.Dictionary")
    d27.Add "name", "caption"
    d27.Add "uiPriority", 35
    d27.Add "qFormat", True
    c.Add d27
    Dim d28 As Object
    Set d28 = CreateObject("Scripting.Dictionary")
    d28.Add "name", "footnote text"
    d28.Add "uiPriority", 99
    d28.Add "semiHidden", True
    d28.Add "unhideWhenUsed", True
    c.Add d28
    Dim d29 As Object
    Set d29 = CreateObject("Scripting.Dictionary")
    d29.Add "name", "TOC Heading"
    d29.Add "uiPriority", 39
    d29.Add "qFormat", True
    c.Add d29
    Dim d30 As Object
    Set d30 = CreateObject("Scripting.Dictionary")
    d30.Add "name", "toc 1"
    d30.Add "uiPriority", 39
    c.Add d30
    Dim d31 As Object
    Set d31 = CreateObject("Scripting.Dictionary")
    d31.Add "name", "toc 2"
    d31.Add "uiPriority", 39
    c.Add d31
    Dim d32 As Object
    Set d32 = CreateObject("Scripting.Dictionary")
    d32.Add "name", "toc 3"
    d32.Add "uiPriority", 39
    c.Add d32
    Set LoadRawStyleSheet_LatentStyleActivation = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_LatentStyleActivation", Err.description
End Function

Private Function LoadRawStyleSheet_OutOfScope() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "reference", "references/huong-dan-template-nd30-ooxml.md " & ChrW(&HA7) & "2.6"
    Set LoadRawStyleSheet_OutOfScope = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet_OutOfScope", Err.description
End Function

Public Function LoadRawStyleSheet() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "1.0.0"
    d.Add "legalBasis", LoadRawStyleSheet_LegalBasis()
    d.Add "constraint", LoadRawStyleSheet_Constraint()
    d.Add "sizeToken", LoadRawStyleSheet_SizeToken()
    d.Add "docDefaults", LoadRawStyleSheet_DocDefaults()
    d.Add "styles", LoadRawStyleSheet_Styles()
    d.Add "characterStyleSync", LoadRawStyleSheet_CharacterStyleSync()
    d.Add "latentStyleActivation", LoadRawStyleSheet_LatentStyleActivation()
    d.Add "outOfScope", LoadRawStyleSheet_OutOfScope()
    Set LoadRawStyleSheet = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawStyleSheet", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart1_0_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart1_0_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart1_0_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart1_1_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart1_1_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart1_1_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart1_2_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart1_2_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart1_2_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart1_3_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart1_3_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart1_3_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart1_4_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart1_4_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart1_4_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart1_5_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart1_5_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart1_5_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart1_6_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart1_6_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart1_6_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart1_7_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart1_7_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart1_7_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart1_8_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart1_8_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart1_8_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart1_9_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart1_9_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart1_9_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart1_10_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart1_10_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart1_10_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart1_11_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart1_11_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart1_11_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart1_12_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart1_12_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart1_12_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart1_13_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart1_13_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart1_13_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart1_14_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart1_14_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart1_14_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart1_15_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart1_15_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart1_15_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart1_16_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart1_16_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart1_16_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart1_17_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart1_17_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart1_17_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart1_18_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart1_18_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart1_18_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart1_19_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart1_19_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart1_19_Regimes", Err.description
End Function

Private Sub LoadRawDocTypeAbbreviations_DocumentTypesPart1(ByRef c As Collection)
    On Error GoTo ErrHandler
    Dim d33 As Object
    Set d33 = CreateObject("Scripting.Dictionary")
    d33.Add "ordinal", 1
    d33.Add "typeName", "Ngh" & ChrW(&H1ECB) & " quy" & ChrW(&H1EBF) & "t (c" & ChrW(&HE1) & " bi" & ChrW(&H1EC7) & "t)"
    d33.Add "abbreviation", "NQ"
    d33.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart1_0_Regimes()
    c.Add d33
    Dim d34 As Object
    Set d34 = CreateObject("Scripting.Dictionary")
    d34.Add "ordinal", 2
    d34.Add "typeName", "Quy" & ChrW(&H1EBF) & "t " & ChrW(&H111) & ChrW(&H1ECB) & "nh (c" & ChrW(&HE1) & " bi" & ChrW(&H1EC7) & "t)"
    d34.Add "abbreviation", "Q" & ChrW(&H110)
    d34.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart1_1_Regimes()
    c.Add d34
    Dim d35 As Object
    Set d35 = CreateObject("Scripting.Dictionary")
    d35.Add "ordinal", 3
    d35.Add "typeName", "Ch" & ChrW(&H1EC9) & " th" & ChrW(&H1ECB)
    d35.Add "abbreviation", "CT"
    d35.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart1_2_Regimes()
    c.Add d35
    Dim d36 As Object
    Set d36 = CreateObject("Scripting.Dictionary")
    d36.Add "ordinal", 4
    d36.Add "typeName", "Quy ch" & ChrW(&H1EBF)
    d36.Add "abbreviation", "QC"
    d36.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart1_3_Regimes()
    c.Add d36
    Dim d37 As Object
    Set d37 = CreateObject("Scripting.Dictionary")
    d37.Add "ordinal", 5
    d37.Add "typeName", "Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh"
    d37.Add "abbreviation", "Qy" & ChrW(&H110)
    d37.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart1_4_Regimes()
    c.Add d37
    Dim d38 As Object
    Set d38 = CreateObject("Scripting.Dictionary")
    d38.Add "ordinal", 6
    d38.Add "typeName", "Th" & ChrW(&HF4) & "ng c" & ChrW(&HE1) & "o"
    d38.Add "abbreviation", "TC"
    d38.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart1_5_Regimes()
    c.Add d38
    Dim d39 As Object
    Set d39 = CreateObject("Scripting.Dictionary")
    d39.Add "ordinal", 7
    d39.Add "typeName", "Th" & ChrW(&HF4) & "ng b" & ChrW(&HE1) & "o"
    d39.Add "abbreviation", "TB"
    d39.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart1_6_Regimes()
    c.Add d39
    Dim d40 As Object
    Set d40 = CreateObject("Scripting.Dictionary")
    d40.Add "ordinal", 8
    d40.Add "typeName", "H" & ChrW(&H1B0) & ChrW(&H1EDB) & "ng d" & ChrW(&H1EAB) & "n"
    d40.Add "abbreviation", "HD"
    d40.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart1_7_Regimes()
    c.Add d40
    Dim d41 As Object
    Set d41 = CreateObject("Scripting.Dictionary")
    d41.Add "ordinal", 9
    d41.Add "typeName", "Ch" & ChrW(&H1B0) & ChrW(&H1A1) & "ng tr" & ChrW(&HEC) & "nh"
    d41.Add "abbreviation", "CTr"
    d41.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart1_8_Regimes()
    c.Add d41
    Dim d42 As Object
    Set d42 = CreateObject("Scripting.Dictionary")
    d42.Add "ordinal", 10
    d42.Add "typeName", "K" & ChrW(&H1EBF) & " ho" & ChrW(&H1EA1) & "ch"
    d42.Add "abbreviation", "KH"
    d42.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart1_9_Regimes()
    c.Add d42
    Dim d43 As Object
    Set d43 = CreateObject("Scripting.Dictionary")
    d43.Add "ordinal", 11
    d43.Add "typeName", "Ph" & ChrW(&H1B0) & ChrW(&H1A1) & "ng " & ChrW(&HE1) & "n"
    d43.Add "abbreviation", "PA"
    d43.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart1_10_Regimes()
    c.Add d43
    Dim d44 As Object
    Set d44 = CreateObject("Scripting.Dictionary")
    d44.Add "ordinal", 12
    d44.Add "typeName", ChrW(&H110) & ChrW(&H1EC1) & " " & ChrW(&HE1) & "n"
    d44.Add "abbreviation", ChrW(&H110) & "A"
    d44.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart1_11_Regimes()
    c.Add d44
    Dim d45 As Object
    Set d45 = CreateObject("Scripting.Dictionary")
    d45.Add "ordinal", 13
    d45.Add "typeName", "D" & ChrW(&H1EF1) & " " & ChrW(&HE1) & "n"
    d45.Add "abbreviation", "DA"
    d45.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart1_12_Regimes()
    c.Add d45
    Dim d46 As Object
    Set d46 = CreateObject("Scripting.Dictionary")
    d46.Add "ordinal", 14
    d46.Add "typeName", "B" & ChrW(&HE1) & "o c" & ChrW(&HE1) & "o"
    d46.Add "abbreviation", "BC"
    d46.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart1_13_Regimes()
    c.Add d46
    Dim d47 As Object
    Set d47 = CreateObject("Scripting.Dictionary")
    d47.Add "ordinal", 15
    d47.Add "typeName", "Bi" & ChrW(&HEA) & "n b" & ChrW(&H1EA3) & "n"
    d47.Add "abbreviation", "BB"
    d47.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart1_14_Regimes()
    c.Add d47
    Dim d48 As Object
    Set d48 = CreateObject("Scripting.Dictionary")
    d48.Add "ordinal", 16
    d48.Add "typeName", "T" & ChrW(&H1EDD) & " tr" & ChrW(&HEC) & "nh"
    d48.Add "abbreviation", "TTr"
    d48.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart1_15_Regimes()
    c.Add d48
    Dim d49 As Object
    Set d49 = CreateObject("Scripting.Dictionary")
    d49.Add "ordinal", 17
    d49.Add "typeName", "H" & ChrW(&H1EE3) & "p " & ChrW(&H111) & ChrW(&H1ED3) & "ng"
    d49.Add "abbreviation", "H" & ChrW(&H110)
    d49.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart1_16_Regimes()
    c.Add d49
    Dim d50 As Object
    Set d50 = CreateObject("Scripting.Dictionary")
    d50.Add "ordinal", 18
    d50.Add "typeName", "C" & ChrW(&HF4) & "ng " & ChrW(&H111) & "i" & ChrW(&H1EC7) & "n"
    d50.Add "abbreviation", "C" & ChrW(&H110)
    d50.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart1_17_Regimes()
    c.Add d50
    Dim d51 As Object
    Set d51 = CreateObject("Scripting.Dictionary")
    d51.Add "ordinal", 19
    d51.Add "typeName", "B" & ChrW(&H1EA3) & "n ghi nh" & ChrW(&H1EDB)
    d51.Add "abbreviation", "BGN"
    d51.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart1_18_Regimes()
    c.Add d51
    Dim d52 As Object
    Set d52 = CreateObject("Scripting.Dictionary")
    d52.Add "ordinal", 20
    d52.Add "typeName", "B" & ChrW(&H1EA3) & "n th" & ChrW(&H1ECF) & "a thu" & ChrW(&H1EAD) & "n"
    d52.Add "abbreviation", "BTT"
    d52.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart1_19_Regimes()
    c.Add d52
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart1", Err.description
End Sub

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart2_20_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart2_20_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart2_20_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart2_21_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart2_21_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart2_21_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart2_22_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart2_22_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart2_22_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart2_23_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart2_23_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart2_23_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart2_24_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart2_24_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart2_24_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart2_25_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart2_25_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart2_25_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart2_26_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ND30"
    c.Add "VIETTEL"
    c.Add "DANG"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart2_26_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart2_26_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart2_27_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "VIETTEL"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart2_27_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart2_27_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart2_28_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "VIETTEL"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart2_28_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart2_28_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart2_29_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "VIETTEL"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart2_29_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart2_29_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart2_30_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "VIETTEL"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart2_30_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart2_30_Regimes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_DocumentTypesPart2_31_Regimes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "VIETTEL"
    Set LoadRawDocTypeAbbreviations_DocumentTypesPart2_31_Regimes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart2_31_Regimes", Err.description
End Function

Private Sub LoadRawDocTypeAbbreviations_DocumentTypesPart2(ByRef c As Collection)
    On Error GoTo ErrHandler
    Dim d53 As Object
    Set d53 = CreateObject("Scripting.Dictionary")
    d53.Add "ordinal", 21
    d53.Add "typeName", "Gi" & ChrW(&H1EA5) & "y " & ChrW(&H1EE7) & "y quy" & ChrW(&H1EC1) & "n"
    d53.Add "abbreviation", "GUQ"
    d53.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart2_20_Regimes()
    c.Add d53
    Dim d54 As Object
    Set d54 = CreateObject("Scripting.Dictionary")
    d54.Add "ordinal", 22
    d54.Add "typeName", "Gi" & ChrW(&H1EA5) & "y m" & ChrW(&H1EDD) & "i"
    d54.Add "abbreviation", "GM"
    d54.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart2_21_Regimes()
    c.Add d54
    Dim d55 As Object
    Set d55 = CreateObject("Scripting.Dictionary")
    d55.Add "ordinal", 23
    d55.Add "typeName", "Gi" & ChrW(&H1EA5) & "y gi" & ChrW(&H1EDB) & "i thi" & ChrW(&H1EC7) & "u"
    d55.Add "abbreviation", "GGT"
    d55.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart2_22_Regimes()
    c.Add d55
    Dim d56 As Object
    Set d56 = CreateObject("Scripting.Dictionary")
    d56.Add "ordinal", 24
    d56.Add "typeName", "Gi" & ChrW(&H1EA5) & "y ngh" & ChrW(&H1EC9) & " ph" & ChrW(&HE9) & "p"
    d56.Add "abbreviation", "GNP"
    d56.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart2_23_Regimes()
    c.Add d56
    Dim d57 As Object
    Set d57 = CreateObject("Scripting.Dictionary")
    d57.Add "ordinal", 25
    d57.Add "typeName", "Phi" & ChrW(&H1EBF) & "u g" & ChrW(&H1EED) & "i"
    d57.Add "abbreviation", "PG"
    d57.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart2_24_Regimes()
    c.Add d57
    Dim d58 As Object
    Set d58 = CreateObject("Scripting.Dictionary")
    d58.Add "ordinal", 26
    d58.Add "typeName", "Phi" & ChrW(&H1EBF) & "u chuy" & ChrW(&H1EC3) & "n"
    d58.Add "abbreviation", "PC"
    d58.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart2_25_Regimes()
    c.Add d58
    Dim d59 As Object
    Set d59 = CreateObject("Scripting.Dictionary")
    d59.Add "ordinal", 27
    d59.Add "typeName", "Phi" & ChrW(&H1EBF) & "u b" & ChrW(&HE1) & "o"
    d59.Add "abbreviation", "PB"
    d59.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart2_26_Regimes()
    c.Add d59
    Dim d60 As Object
    Set d60 = CreateObject("Scripting.Dictionary")
    d60.Add "ordinal", 28
    d60.Add "typeName", "Gi" & ChrW(&H1EA5) & "y c" & ChrW(&HF4) & "ng t" & ChrW(&HE1) & "c"
    d60.Add "abbreviation", "GCT"
    d60.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart2_27_Regimes()
    d60.Add "note", "Rieng Phu luc I cua Viettel, khong co trong Phu luc III N" & ChrW(&H110) & " 30."
    c.Add d60
    Dim d61 As Object
    Set d61 = CreateObject("Scripting.Dictionary")
    d61.Add "ordinal", 29
    d61.Add "typeName", "Phi" & ChrW(&H1EBF) & "u Nh" & ChrW(&H1EAD) & "n x" & ChrW(&HE9) & "t"
    d61.Add "abbreviation", "PNX"
    d61.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart2_28_Regimes()
    d61.Add "note", "Rieng Phu luc I cua Viettel."
    c.Add d61
    Dim d62 As Object
    Set d62 = CreateObject("Scripting.Dictionary")
    d62.Add "ordinal", 30
    d62.Add "typeName", "Quy tr" & ChrW(&HEC) & "nh"
    d62.Add "abbreviation", "QT"
    d62.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart2_29_Regimes()
    d62.Add "note", "Rieng Phu luc I cua Viettel."
    c.Add d62
    Dim d63 As Object
    Set d63 = CreateObject("Scripting.Dictionary")
    d63.Add "ordinal", 31
    d63.Add "typeName", "Guideline"
    d63.Add "abbreviation", "GL"
    d63.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart2_30_Regimes()
    d63.Add "note", "Rieng Phu luc I cua Viettel."
    c.Add d63
    Dim d64 As Object
    Set d64 = CreateObject("Scripting.Dictionary")
    d64.Add "ordinal", 32
    d64.Add "typeName", "Ti" & ChrW(&HEA) & "u chu" & ChrW(&H1EA9) & "n"
    d64.Add "abbreviation", "TC"
    d64.Add "regimes", LoadRawDocTypeAbbreviations_DocumentTypesPart2_31_Regimes()
    d64.Add "note", "Rieng Phu luc I cua Viettel. LUU Y: trung ky hieu 'TC' voi 'Thong cao' (ordinal 6) - chu du an da xac nhan van ap dung, khong anh huong (danh muc tra theo TEN LOAI, khong tra nguoc tu ky hieu)."
    c.Add d64
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypesPart2", Err.description
End Sub

Private Function LoadRawDocTypeAbbreviations_DocumentTypes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    LoadRawDocTypeAbbreviations_DocumentTypesPart1 c
    LoadRawDocTypeAbbreviations_DocumentTypesPart2 c
    Set LoadRawDocTypeAbbreviations_DocumentTypes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_DocumentTypes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_CopyTypes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim d65 As Object
    Set d65 = CreateObject("Scripting.Dictionary")
    d65.Add "typeName", "B" & ChrW(&H1EA3) & "n sao y"
    d65.Add "abbreviation", "SY"
    c.Add d65
    Dim d66 As Object
    Set d66 = CreateObject("Scripting.Dictionary")
    d66.Add "typeName", "B" & ChrW(&H1EA3) & "n tr" & ChrW(&HED) & "ch sao"
    d66.Add "abbreviation", "TrS"
    c.Add d66
    Dim d67 As Object
    Set d67 = CreateObject("Scripting.Dictionary")
    d67.Add "typeName", "B" & ChrW(&H1EA3) & "n sao l" & ChrW(&H1EE5) & "c"
    d67.Add "abbreviation", "SL"
    c.Add d67
    Set LoadRawDocTypeAbbreviations_CopyTypes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_CopyTypes", Err.description
End Function

Private Function LoadRawDocTypeAbbreviations_NoTypeAbbreviation() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "C" & ChrW(&HF4) & "ng v" & ChrW(&H103) & "n"
    c.Add "Th" & ChrW(&H1B0) & " c" & ChrW(&HF4) & "ng"
    Set LoadRawDocTypeAbbreviations_NoTypeAbbreviation = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations_NoTypeAbbreviation", Err.description
End Function

Public Function LoadRawDocTypeAbbreviations() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "1.0.0"
    d.Add "status", "day du - 27 dong goc + 5 dong rieng cua Viettel + 3 loai ban sao"
    d.Add "sourceLabel", "ND30"
    d.Add "documentTypes", LoadRawDocTypeAbbreviations_DocumentTypes()
    d.Add "copyTypes", LoadRawDocTypeAbbreviations_CopyTypes()
    d.Add "noTypeAbbreviation", LoadRawDocTypeAbbreviations_NoTypeAbbreviation()
    Set LoadRawDocTypeAbbreviations = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawDocTypeAbbreviations", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_NationalTitle() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "nationalTitle"
    d.Add "confidence", "high"
    d.Add "method", "normalizedFuzzyContains"
    d.Add "normalizedTarget", "CONGHOAXAHOICHUNGHIAVIETNAM"
    d.Add "maxEditDistance", 3
    d.Add "headerWindowOnly", True
    d.Add "firstMatchOnly", True
    d.Add "note", "Khop GAN DUNG (cho phep lech toi da maxEditDistance ky tu chen/xoa/thay o BAT KY vi tri nao trong cum, khong chi cuoi cum) thay vi khop CHUA nguyen van " & ChrW(&H2014) & " tai lieu that hay go thieu/sai mot vai ky tu cua Quoc hieu (vi du thieu chu 'M' cuoi 'VIET NAM'). 3 ky tu tren tong 27 ky tu la nguong an toan: cum nay rat dac trung, khong the trung voi doan van nao khac du cho phep lech chung do. CHI dung cho ND30/VIETTEL " & ChrW(&H2014) & " DANG dung partyHeader."
    Set LoadRawComponentSignals_Signals_NationalTitle = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_NationalTitle", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_NationalMotto() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "nationalMotto"
    d.Add "confidence", "high"
    d.Add "method", "normalizedFuzzyContainsAfterRole"
    d.Add "afterRole", "nationalTitle"
    d.Add "normalizedTarget", "DOCLAP-TUDO-HANHPHUC"
    d.Add "maxEditDistance", 2
    d.Add "headerWindowOnly", True
    d.Add "firstMatchOnly", True
    d.Add "skipBlankParagraphs", True
    d.Add "note", "skipBlankParagraphs: bo qua cac doan Word RONG xen giua Quoc hieu va Tieu ngu (tai lieu that hay chen dong trong de tao khoang cach trong o bang) roi moi kiem tra doan khong rong dau tien sau Quoc hieu. Khop GAN DUNG (nhu nationalTitle) thay vi regex dung nguyen van tung tu " & ChrW(&H2014) & " tai lieu that co the go sai/thieu mot vai ky tu cua Tieu ngu."
    Set LoadRawComponentSignals_Signals_NationalMotto = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_NationalMotto", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_PartyHeader() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "partyHeader"
    d.Add "confidence", "high"
    d.Add "method", "normalizedContains"
    d.Add "normalizedTarget", ChrW(&H110) & ChrW(&H1EA2) & "NGC" & ChrW(&H1ED8) & "NGS" & ChrW(&H1EA2) & "NVI" & ChrW(&H1EC6) & "TNAM"
    d.Add "headerWindowOnly", True
    d.Add "firstMatchOnly", True
    d.Add "note", "CHI dung cho che do DANG " & ChrW(&H2014) & " thay cho nationalTitle + nationalMotto."
    Set LoadRawComponentSignals_Signals_PartyHeader = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_PartyHeader", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_CodeNumberNotation_RegexByRegime() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "ND30", "^\s*S" & ChrW(&H1ED1) & "\s*:\s*\d*\s*\/"
    d.Add "VIETTEL", "^\s*S" & ChrW(&H1ED1) & "\s*:\s*\d*\s*\/"
    d.Add "DANG", "^\s*S" & ChrW(&H1ED1) & "\s*:?\s*(\d+|[\-\/]|$)"
    Set LoadRawComponentSignals_Signals_CodeNumberNotation_RegexByRegime = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_CodeNumberNotation_RegexByRegime", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_CodeNumberNotation() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "codeNumberNotation"
    d.Add "confidence", "high"
    d.Add "method", "regex"
    d.Add "headerWindowOnly", True
    d.Add "firstMatchOnly", True
    d.Add "regexByRegime", LoadRawComponentSignals_Signals_CodeNumberNotation_RegexByRegime()
    d.Add "note", "Chap nhan ca truong hop so de trong ('S" & ChrW(&H1ED1) & ":          /Q" & ChrW(&H110) & "-TTg'). DANG khac: khong bat buoc dau hai cham, ky hieu dung '-' thay '/'."
    Set LoadRawComponentSignals_Signals_CodeNumberNotation = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_CodeNumberNotation", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_StarSeparator() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "starSeparator"
    d.Add "confidence", "medium"
    d.Add "method", "regexBeforeRole"
    d.Add "beforeRole", "codeNumberNotation"
    d.Add "regex", "^\s*\*\s*$"
    d.Add "headerWindowOnly", True
    d.Add "note", "CHI dung cho che do DANG " & ChrW(&H2014) & " dong chi co dau sao, giua ten co quan ban hanh va so/ky hieu."
    Set LoadRawComponentSignals_Signals_StarSeparator = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_StarSeparator", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_PlaceAndIssuedDate() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "placeAndIssuedDate"
    d.Add "confidence", "high"
    d.Add "method", "regex"
    d.Add "regex", ",\s*ng" & ChrW(&HE0) & "y.*?th" & ChrW(&HE1) & "ng.*?n" & ChrW(&H103) & "m(\s+\d*)?"
    d.Add "headerWindowOnly", True
    d.Add "firstMatchOnly", True
    d.Add "note", "Neo dau phay truoc 'ng" & ChrW(&HE0) & "y' de khong khop nham dong 'C" & ChrW(&H103) & "n c" & ChrW(&H1EE9) & " ... ng" & ChrW(&HE0) & "y ... th" & ChrW(&HE1) & "ng ... n" & ChrW(&H103) & "m ...'. Dung .*? (khong phai \s+\d*\s+) giua 'ng" & ChrW(&HE0) & "y'-'th" & ChrW(&HE1) & "ng' va 'th" & ChrW(&HE1) & "ng'-'n" & ChrW(&H103) & "m' de khop moi kieu khoang trang/so ngay-thang, ke ca khi de trong chi go 1 khoang trang don (khong dem 2 khoang trang nhu mau chuan)."
    Set LoadRawComponentSignals_Signals_PlaceAndIssuedDate = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_PlaceAndIssuedDate", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_SubjectOfficialLetter() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "subjectOfficialLetter"
    d.Add "confidence", "high"
    d.Add "method", "regex"
    d.Add "regex", "^\s*V\/v"
    d.Add "headerWindowOnly", True
    d.Add "firstMatchOnly", True
    Set LoadRawComponentSignals_Signals_SubjectOfficialLetter = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_SubjectOfficialLetter", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_AppendixLabel() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "appendixLabel"
    d.Add "confidence", "high"
    d.Add "method", "regex"
    d.Add "regex", "^\s*Ph" & ChrW(&H1EE5) & "\s+l" & ChrW(&H1EE5) & "c(\s+[IVXLCDM0-9]+)?\s*$"
    d.Add "note", "Doan CHI co chu 'Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c' kem so thu tu (hoac khong) " & ChrW(&H2014) & " khong khop cau than bai co chua chu 'Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c'."
    Set LoadRawComponentSignals_Signals_AppendixLabel = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_AppendixLabel", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_AppendixTitle() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "appendixTitle"
    d.Add "confidence", "medium"
    d.Add "method", "positionalAfterRoleStyled"
    d.Add "afterRole", "appendixLabel"
    d.Add "requireAllCaps", True
    d.Add "requireNonEmpty", True
    Set LoadRawComponentSignals_Signals_AppendixTitle = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_AppendixTitle", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_AppendixReference() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "appendixReference"
    d.Add "confidence", "medium"
    d.Add "method", "regexAfterRole"
    d.Add "afterRole", "appendixTitle"
    d.Add "regex", "^\s*\(?\s*K" & ChrW(&HE8) & "m\s+theo"
    Set LoadRawComponentSignals_Signals_AppendixReference = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_AppendixReference", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_RecipientSalutation_RegexByRegime() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "ND30", "^\s*K" & ChrW(&HED) & "nh\s+g" & ChrW(&H1EED) & "i\s*:\s*$"
    d.Add "VIETTEL", "^\s*K" & ChrW(&HED) & "nh\s+g" & ChrW(&H1EED) & "i\s*:\s*$"
    d.Add "DANG", "^\s*K" & ChrW(&HED) & "nh\s+(g" & ChrW(&H1EED) & "i|tr" & ChrW(&HEC) & "nh)\s*:\s*$"
    Set LoadRawComponentSignals_Signals_RecipientSalutation_RegexByRegime = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_RecipientSalutation_RegexByRegime", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_RecipientSalutation() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "recipientSalutation"
    d.Add "confidence", "high"
    d.Add "method", "regex"
    d.Add "firstMatchOnly", True
    d.Add "regexByRegime", LoadRawComponentSignals_Signals_RecipientSalutation_RegexByRegime()
    d.Add "note", "Dong CHI co nhan, khong co noi dung phia sau dau hai cham. DANG chap nhan them 'K" & ChrW(&HED) & "nh tr" & ChrW(&HEC) & "nh' (dung cho To trinh)."
    Set LoadRawComponentSignals_Signals_RecipientSalutation = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_RecipientSalutation", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_RecipientSalutationInline_RegexByRegime() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "ND30", "^\s*K" & ChrW(&HED) & "nh\s+g" & ChrW(&H1EED) & "i\s*:\s*\S"
    d.Add "VIETTEL", "^\s*K" & ChrW(&HED) & "nh\s+g" & ChrW(&H1EED) & "i\s*:\s*\S"
    d.Add "DANG", "^\s*K" & ChrW(&HED) & "nh\s+(g" & ChrW(&H1EED) & "i|tr" & ChrW(&HEC) & "nh)\s*:\s*\S"
    Set LoadRawComponentSignals_Signals_RecipientSalutationInline_RegexByRegime = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_RecipientSalutationInline_RegexByRegime", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_RecipientSalutationInline() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "recipientSalutationInline"
    d.Add "confidence", "high"
    d.Add "method", "regex"
    d.Add "firstMatchOnly", True
    d.Add "regexByRegime", LoadRawComponentSignals_Signals_RecipientSalutationInline_RegexByRegime()
    d.Add "note", "Dong co noi dung ngay sau dau hai cham. Phan noi dung do mang vai tro recipientSalutationInlineContent, tach theo ky tu trong CUNG doan (khong phai doan rieng) " & ChrW(&H2014) & " xem TextFormatter.InlineContentRole."
    Set LoadRawComponentSignals_Signals_RecipientSalutationInline = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_RecipientSalutationInline", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_RecipientSalutationList() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "recipientSalutationList"
    d.Add "confidence", "high"
    d.Add "method", "contiguousRunAfterRole"
    d.Add "afterRole", "recipientSalutation"
    d.Add "regex", "^\s*-"
    d.Add "note", "Khoi LIEN TUC cac dong bat dau bang gach noi, ngay sau dong 'K" & ChrW(&HED) & "nh g" & ChrW(&H1EED) & "i:' dung mot minh."
    Set LoadRawComponentSignals_Signals_RecipientSalutationList = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_RecipientSalutationList", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_RecipientLabel() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "recipientLabel"
    d.Add "confidence", "high"
    d.Add "method", "normalizedFuzzyContains"
    d.Add "normalizedTarget", "NOINHAN"
    d.Add "maxEditDistance", 1
    d.Add "maxSearchChars", 12
    d.Add "lastMatchOnly", True
    d.Add "stopAtRole", "appendixLabel"
    d.Add "note", "Chot cua chu du an: chi lay ket qua CUOI CUNG cua van ban, KHONG tinh phan phu luc. Khop GAN DUNG, gioi han 12 ky tu DAU doan (maxSearchChars) " & ChrW(&H2014) & " nhan 'N" & ChrW(&H1A1) & "i nh" & ChrW(&H1EAD) & "n' bi go sai/thieu nhung PHAI nam o dau doan, tranh khop nham mot cau than bai vo tinh nhac toi cum tu nay o giua doan."
    Set LoadRawComponentSignals_Signals_RecipientLabel = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_RecipientLabel", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_RecipientListClosing_RegexByRegime() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "ND30", "^\s*-?\s*L" & ChrW(&H1B0) & "u\s*:\s*VT"
    d.Add "VIETTEL", "^\s*-?\s*L" & ChrW(&H1B0) & "u\s*:\s*VT"
    d.Add "DANG", "^\s*-?\s*L" & ChrW(&H1B0) & "u\s+\S+\."
    Set LoadRawComponentSignals_Signals_RecipientListClosing_RegexByRegime = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_RecipientListClosing_RegexByRegime", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_RecipientListClosing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "recipientList"
    d.Add "confidence", "high"
    d.Add "method", "regex"
    d.Add "regexByRegime", LoadRawComponentSignals_Signals_RecipientListClosing_RegexByRegime()
    d.Add "note", "Dong 'L" & ChrW(&H1B0) & "u...' " & ChrW(&H2014) & " dong cuoi cung cua danh sach noi nhan. DANG khong dung dau hai cham va khong co 'VT'; regex doi hoi mot dau cham trong cum sau chu 'L" & ChrW(&H1B0) & "u' de khong khop nham cau than bai kieu 'L" & ChrW(&H1B0) & "u " & ChrW(&HFD) & ":'."
    Set LoadRawComponentSignals_Signals_RecipientListClosing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_RecipientListClosing", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_RecipientListAfterLabel() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "recipientList"
    d.Add "confidence", "medium"
    d.Add "method", "positionalAfterRole"
    d.Add "afterRole", "recipientLabel"
    d.Add "uptoRole", "recipientList"
    d.Add "skipIfAnchorMatchesRegex", "L" & ChrW(&H1B0) & "u\s*:\s*VT"
    d.Add "note", "Moi doan sau recipientLabel chua duoc gan vai tro nao khac. uptoRole: chan tren la dong 'L" & ChrW(&H1B0) & "u...'. skipIfAnchorMatchesRegex: neu CHINH doan recipientLabel da tu chua dong 'L" & ChrW(&H1B0) & "u' (nguoi soan gop ca khoi vao mot doan Word qua xuong dong thu cong) thi khong lay them doan nao."
    Set LoadRawComponentSignals_Signals_RecipientListAfterLabel = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_RecipientListAfterLabel", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_TypeName() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "typeName"
    d.Add "confidence", "medium"
    d.Add "method", "typeNameDictionary"
    d.Add "note", "Toan hoa, canh giua, khop danh muc shared/rules/chu-viet-tat-ten-loai.json (loc theo che do qua truong 'regimes')."
    Set LoadRawComponentSignals_Signals_TypeName = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_TypeName", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_Subject_StopIfMatchesRegexByRegime() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "ND30", "^\s*C" & ChrW(&H103) & "n\s+c" & ChrW(&H1EE9)
    d.Add "VIETTEL", "^\s*C" & ChrW(&H103) & "n\s+c" & ChrW(&H1EE9)
    d.Add "DANG", "^\s*-?\s*C" & ChrW(&H103) & "n\s+c" & ChrW(&H1EE9)
    Set LoadRawComponentSignals_Signals_Subject_StopIfMatchesRegexByRegime = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_Subject_StopIfMatchesRegexByRegime", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_Subject() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "subject"
    d.Add "confidence", "medium"
    d.Add "method", "contiguousRunAfterRole"
    d.Add "afterRole", "typeName"
    d.Add "requireNonEmpty", True
    d.Add "stopIfMatchesRegexByRegime", LoadRawComponentSignals_Signals_Subject_StopIfMatchesRegexByRegime()
    Dim tmp1 As String
    tmp1 = "Khoi LIEN TUC cac doan khong rong, chua duoc gan vai tro khac, bat dau ngay sau typeName. Trich yeu ngoai doi co the tach thanh NHIEU doan Word that (nguoi soan an Enter thay vi xuong dong thu cong Shift+Enter) " & ChrW(&H2014) & " dung het khi gap doan RONG, doan da duoc gan vai tro khac (vi du 'K" & ChrW(&HED) & "nh g" & ChrW(&H1EED) & "i:'), hoac doan doc nhu mo dau legalBasis ('C" & ChrW(&H103) & "n c" & ChrW(&H1EE9) & "...') " & ChrW(&H2014) & " legalBasis dung SAU subject trong SignalOrder nen chua duoc gan luc nay, phai chan rieng bang stopIfMatchesRegex. Chi dua vao VI TRI, khong doi hoi regex noi dung rieng cho ban than trich yeu " & ChrW(&H2014)
    tmp1 = tmp1 & " tai lieu that co trich yeu canh giua nhung KHONG in dam."
    d.Add "note", tmp1
    Set LoadRawComponentSignals_Signals_Subject = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_Subject", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_OrganName() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "organName"
    d.Add "confidence", "medium"
    d.Add "method", "allCapsBeforeRole"
    d.Add "beforeRole", "codeNumberNotation"
    d.Add "requireNonEmpty", True
    d.Add "excludeIfMatchesRegex", ChrW(&H110) & ChrW(&H1ED9) & "c\s*l" & ChrW(&H1EAD) & "p.*T" & ChrW(&H1EF1) & "\s*do.*H" & ChrW(&H1EA1) & "nh\s*ph" & ChrW(&HFA) & "c"
    d.Add "headerWindowOnly", True
    d.Add "note", "Doan KHONG RONG, GAN NHAT truoc codeNumberNotation, chua duoc gan vai tro khac. Khong dua vao kieu chu: tai lieu that dat khoi nay trong o bang va CANH GIUA. Ten method giu nguyen (chi la nhan dispatch). excludeIfMatchesRegex: khong duoc nhan doan doc nhu Tieu ngu " & ChrW(&H2014) & " phong khi Quoc hieu/Tieu ngu bi go sai/thieu qua muc nen chua duoc gan vai tro rieng, doan Tieu ngu con 'trong' luc dau hieu nay chay va se bi vo nham lam ten co quan neu khong loai truoc."
    Set LoadRawComponentSignals_Signals_OrganName = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_OrganName", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_SuperiorOrganName() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "superiorOrganName"
    d.Add "confidence", "medium"
    d.Add "method", "immediatelyBeforeRole"
    d.Add "beforeRole", "organName"
    d.Add "requireNonEmpty", True
    d.Add "headerWindowOnly", True
    d.Add "note", "Doan NGAY TRUOC organName. Chot cua chu du an: khoi ten co quan co 2 dong thi dong tren la chu quan; chi 1 dong thi KHONG co chu quan " & ChrW(&H2014) & " dieu kien 'khong rong + chua duoc gan + trong header window' chinh la cach loai truong hop 1 dong."
    Set LoadRawComponentSignals_Signals_SuperiorOrganName = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_SuperiorOrganName", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_LegalBasis_AnchorRoles() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "recipientSalutationList"
    c.Add "recipientSalutationInline"
    c.Add "recipientSalutation"
    c.Add "subject"
    Set LoadRawComponentSignals_Signals_LegalBasis_AnchorRoles = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_LegalBasis_AnchorRoles", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_LegalBasis_RegexByRegime() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "ND30", "^\s*C" & ChrW(&H103) & "n\s+c" & ChrW(&H1EE9)
    d.Add "VIETTEL", "^\s*C" & ChrW(&H103) & "n\s+c" & ChrW(&H1EE9)
    d.Add "DANG", "^\s*-?\s*C" & ChrW(&H103) & "n\s+c" & ChrW(&H1EE9)
    Set LoadRawComponentSignals_Signals_LegalBasis_RegexByRegime = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_LegalBasis_RegexByRegime", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_LegalBasis() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "legalBasis"
    d.Add "confidence", "high"
    d.Add "method", "anchoredContiguousRun"
    d.Add "anchorRoles", LoadRawComponentSignals_Signals_LegalBasis_AnchorRoles()
    d.Add "regexByRegime", LoadRawComponentSignals_Signals_LegalBasis_RegexByRegime()
    d.Add "note", "Khoi LIEN TUC, neo vao doan NGAY SAU mot trong cac anchorRoles (thu theo dung thu tu do). Mot doan 'C" & ChrW(&H103) & "n c" & ChrW(&H1EE9) & "' dung le loi giua than bai KHONG duoc gan vai tro nay. DANG: moi dong can cu co them gach noi dau dong."
    Set LoadRawComponentSignals_Signals_LegalBasis = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_LegalBasis", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_SignerAuthority_RegexByRegime() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "ND30", "^\s*(TM|KT|TL|TUQ|Q)\."
    d.Add "VIETTEL", "^\s*(TM|KT|TL|TUQ|Q)\."
    d.Add "DANG", "^\s*(T\/M|K\/T|T\/L|Q\.)"
    Set LoadRawComponentSignals_Signals_SignerAuthority_RegexByRegime = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_SignerAuthority_RegexByRegime", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_SignerAuthority() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "signerAuthority"
    d.Add "confidence", "medium"
    d.Add "method", "regex"
    d.Add "regexByRegime", LoadRawComponentSignals_Signals_SignerAuthority_RegexByRegime()
    d.Add "note", "DANG dung dang co dau gach cheo va KHONG co 'TUQ.'."
    Set LoadRawComponentSignals_Signals_SignerAuthority = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_SignerAuthority", Err.description
End Function

Private Function LoadRawComponentSignals_Signals_SignerAuthorityTitle() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "role", "signerAuthorityTitle"
    d.Add "confidence", "medium"
    d.Add "method", "positionalAfterRoleStyled"
    d.Add "afterRole", "signerAuthority"
    d.Add "requireAllCaps", True
    d.Add "requireNonEmpty", True
    d.Add "note", "Dong THU HAI cua khoi ky ten " & ChrW(&H2014) & " chuc vu cu the (vi du 'PH" & ChrW(&HD3) & " GI" & ChrW(&HC1) & "M " & ChrW(&H110) & ChrW(&H1ED0) & "C' ngay sau 'TM. GI" & ChrW(&HC1) & "M " & ChrW(&H110) & ChrW(&H1ED0) & "C'). requireAllCaps phan biet voi ho ten nguoi ky (chi hoa chu cai dau moi tu, khong duoc nhan dien/xu ly theo chot cua chu du an)."
    Set LoadRawComponentSignals_Signals_SignerAuthorityTitle = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals_SignerAuthorityTitle", Err.description
End Function

Private Function LoadRawComponentSignals_Signals() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "nationalTitle", LoadRawComponentSignals_Signals_NationalTitle()
    d.Add "nationalMotto", LoadRawComponentSignals_Signals_NationalMotto()
    d.Add "partyHeader", LoadRawComponentSignals_Signals_PartyHeader()
    d.Add "codeNumberNotation", LoadRawComponentSignals_Signals_CodeNumberNotation()
    d.Add "starSeparator", LoadRawComponentSignals_Signals_StarSeparator()
    d.Add "placeAndIssuedDate", LoadRawComponentSignals_Signals_PlaceAndIssuedDate()
    d.Add "subjectOfficialLetter", LoadRawComponentSignals_Signals_SubjectOfficialLetter()
    d.Add "appendixLabel", LoadRawComponentSignals_Signals_AppendixLabel()
    d.Add "appendixTitle", LoadRawComponentSignals_Signals_AppendixTitle()
    d.Add "appendixReference", LoadRawComponentSignals_Signals_AppendixReference()
    d.Add "recipientSalutation", LoadRawComponentSignals_Signals_RecipientSalutation()
    d.Add "recipientSalutationInline", LoadRawComponentSignals_Signals_RecipientSalutationInline()
    d.Add "recipientSalutationList", LoadRawComponentSignals_Signals_RecipientSalutationList()
    d.Add "recipientLabel", LoadRawComponentSignals_Signals_RecipientLabel()
    d.Add "recipientListClosing", LoadRawComponentSignals_Signals_RecipientListClosing()
    d.Add "recipientListAfterLabel", LoadRawComponentSignals_Signals_RecipientListAfterLabel()
    d.Add "typeName", LoadRawComponentSignals_Signals_TypeName()
    d.Add "subject", LoadRawComponentSignals_Signals_Subject()
    d.Add "organName", LoadRawComponentSignals_Signals_OrganName()
    d.Add "superiorOrganName", LoadRawComponentSignals_Signals_SuperiorOrganName()
    d.Add "legalBasis", LoadRawComponentSignals_Signals_LegalBasis()
    d.Add "signerAuthority", LoadRawComponentSignals_Signals_SignerAuthority()
    d.Add "signerAuthorityTitle", LoadRawComponentSignals_Signals_SignerAuthorityTitle()
    Set LoadRawComponentSignals_Signals = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals_Signals", Err.description
End Function

Public Function LoadRawComponentSignals() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "2.0.0"
    d.Add "sourceLabel", "ND30"
    d.Add "signals", LoadRawComponentSignals_Signals()
    Set LoadRawComponentSignals = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawComponentSignals", Err.description
End Function

Private Function LoadRawToneMapping_Styles_ToneOnMainVowel_Examples() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ho" & ChrW(&HE0)
    c.Add "thu" & ChrW(&H1EF7)
    c.Add "kho" & ChrW(&H1EBB)
    c.Add "ho" & ChrW(&HE1)
    c.Add "to" & ChrW(&HE0)
    c.Add "ngu" & ChrW(&H1EF5)
    Set LoadRawToneMapping_Styles_ToneOnMainVowel_Examples = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_Styles_ToneOnMainVowel_Examples", Err.description
End Function

Private Function LoadRawToneMapping_Styles_ToneOnMainVowel() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "label", "Ki" & ChrW(&H1EC3) & "u o" & ChrW(&HE0) & ", u" & ChrW(&HFD)
    d.Add "examples", LoadRawToneMapping_Styles_ToneOnMainVowel_Examples()
    d.Add "isStandard", True
    d.Add "standardCitation", "Quy" & ChrW(&H1EBF) & "t " & ChrW(&H111) & ChrW(&H1ECB) & "nh s" & ChrW(&H1ED1) & " 1989/Q" & ChrW(&H110) & "-BGD" & ChrW(&H110) & "T ng" & ChrW(&HE0) & "y 25 th" & ChrW(&HE1) & "ng 5 n" & ChrW(&H103) & "m 2018, " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "u 8"
    Set LoadRawToneMapping_Styles_ToneOnMainVowel = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_Styles_ToneOnMainVowel", Err.description
End Function

Private Function LoadRawToneMapping_Styles_ToneOnFirstVowel_Examples() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "h" & ChrW(&HF2) & "a"
    c.Add "th" & ChrW(&H1EE7) & "y"
    c.Add "kh" & ChrW(&H1ECF) & "e"
    c.Add "h" & ChrW(&HF3) & "a"
    c.Add "t" & ChrW(&HF2) & "a"
    c.Add "ng" & ChrW(&H1EE5) & "y"
    Set LoadRawToneMapping_Styles_ToneOnFirstVowel_Examples = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_Styles_ToneOnFirstVowel_Examples", Err.description
End Function

Private Function LoadRawToneMapping_Styles_ToneOnFirstVowel() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "label", "Ki" & ChrW(&H1EC3) & "u " & ChrW(&HF2) & "a, " & ChrW(&HFA) & "y"
    d.Add "examples", LoadRawToneMapping_Styles_ToneOnFirstVowel_Examples()
    d.Add "isStandard", False
    Set LoadRawToneMapping_Styles_ToneOnFirstVowel = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_Styles_ToneOnFirstVowel", Err.description
End Function

Private Function LoadRawToneMapping_Styles() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "toneOnMainVowel", LoadRawToneMapping_Styles_ToneOnMainVowel()
    d.Add "toneOnFirstVowel", LoadRawToneMapping_Styles_ToneOnFirstVowel()
    Set LoadRawToneMapping_Styles = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_Styles", Err.description
End Function

Private Function LoadRawToneMapping_MapToMainVowel() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add ChrW(&HF2) & "a", "o" & ChrW(&HE0)
    d.Add ChrW(&HF3) & "a", "o" & ChrW(&HE1)
    d.Add ChrW(&H1ECF) & "a", "o" & ChrW(&H1EA3)
    d.Add ChrW(&HF5) & "a", "o" & ChrW(&HE3)
    d.Add ChrW(&H1ECD) & "a", "o" & ChrW(&H1EA1)
    d.Add ChrW(&HF2) & "e", "o" & ChrW(&HE8)
    d.Add ChrW(&HF3) & "e", "o" & ChrW(&HE9)
    d.Add ChrW(&H1ECF) & "e", "o" & ChrW(&H1EBB)
    d.Add ChrW(&HF5) & "e", "o" & ChrW(&H1EBD)
    d.Add ChrW(&H1ECD) & "e", "o" & ChrW(&H1EB9)
    d.Add ChrW(&HF9) & "y", "u" & ChrW(&H1EF3)
    d.Add ChrW(&HFA) & "y", "u" & ChrW(&HFD)
    d.Add ChrW(&H1EE7) & "y", "u" & ChrW(&H1EF7)
    d.Add ChrW(&H169) & "y", "u" & ChrW(&H1EF9)
    d.Add ChrW(&H1EE5) & "y", "u" & ChrW(&H1EF5)
    Set LoadRawToneMapping_MapToMainVowel = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_MapToMainVowel", Err.description
End Function

Private Function LoadRawToneMapping_MapToFirstVowel() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "o" & ChrW(&HE0), ChrW(&HF2) & "a"
    d.Add "o" & ChrW(&HE1), ChrW(&HF3) & "a"
    d.Add "o" & ChrW(&H1EA3), ChrW(&H1ECF) & "a"
    d.Add "o" & ChrW(&HE3), ChrW(&HF5) & "a"
    d.Add "o" & ChrW(&H1EA1), ChrW(&H1ECD) & "a"
    d.Add "o" & ChrW(&HE8), ChrW(&HF2) & "e"
    d.Add "o" & ChrW(&HE9), ChrW(&HF3) & "e"
    d.Add "o" & ChrW(&H1EBB), ChrW(&H1ECF) & "e"
    d.Add "o" & ChrW(&H1EBD), ChrW(&HF5) & "e"
    d.Add "o" & ChrW(&H1EB9), ChrW(&H1ECD) & "e"
    d.Add "u" & ChrW(&H1EF3), ChrW(&HF9) & "y"
    d.Add "u" & ChrW(&HFD), ChrW(&HFA) & "y"
    d.Add "u" & ChrW(&H1EF7), ChrW(&H1EE7) & "y"
    d.Add "u" & ChrW(&H1EF9), ChrW(&H169) & "y"
    d.Add "u" & ChrW(&H1EF5), ChrW(&H1EE5) & "y"
    Set LoadRawToneMapping_MapToFirstVowel = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_MapToFirstVowel", Err.description
End Function

Private Function LoadRawToneMapping_Conditions() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "mustBeSyllableFinal", True
    d.Add "precedingConsonantsForOGroup", "bcd" & ChrW(&H111) & "ghklmnpqrstvx"
    d.Add "precedingConsonantsForUGroup", "bcd" & ChrW(&H111) & "ghklmnprstvx"
    Set LoadRawToneMapping_Conditions = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_Conditions", Err.description
End Function

Private Function LoadRawToneMapping_Regex_ToMainVowel() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "oGroup", "(?<=[bcd" & ChrW(&H111) & "ghklmnpqrstvx])(" & ChrW(&HF2) & "a|" & ChrW(&HF3) & "a|" & ChrW(&H1ECF) & "a|" & ChrW(&HF5) & "a|" & ChrW(&H1ECD) & "a|" & ChrW(&HF2) & "e|" & ChrW(&HF3) & "e|" & ChrW(&H1ECF) & "e|" & ChrW(&HF5) & "e|" & ChrW(&H1ECD) & "e)(?![\p{L}])"
    d.Add "uGroup", "(?<=[bcd" & ChrW(&H111) & "ghklmnprstvx])(" & ChrW(&HF9) & "y|" & ChrW(&HFA) & "y|" & ChrW(&H1EE7) & "y|" & ChrW(&H169) & "y|" & ChrW(&H1EE5) & "y)(?![\p{L}])"
    Set LoadRawToneMapping_Regex_ToMainVowel = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_Regex_ToMainVowel", Err.description
End Function

Private Function LoadRawToneMapping_Regex_ToFirstVowel() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "oGroup", "(?<=[bcd" & ChrW(&H111) & "ghklmnpqrstvx])(o" & ChrW(&HE0) & "|o" & ChrW(&HE1) & "|o" & ChrW(&H1EA3) & "|o" & ChrW(&HE3) & "|o" & ChrW(&H1EA1) & "|o" & ChrW(&HE8) & "|o" & ChrW(&HE9) & "|o" & ChrW(&H1EBB) & "|o" & ChrW(&H1EBD) & "|o" & ChrW(&H1EB9) & ")(?![\p{L}])"
    d.Add "uGroup", "(?<=[bcd" & ChrW(&H111) & "ghklmnprstvx])(u" & ChrW(&H1EF3) & "|u" & ChrW(&HFD) & "|u" & ChrW(&H1EF7) & "|u" & ChrW(&H1EF9) & "|u" & ChrW(&H1EF5) & ")(?![\p{L}])"
    Set LoadRawToneMapping_Regex_ToFirstVowel = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_Regex_ToFirstVowel", Err.description
End Function

Private Function LoadRawToneMapping_Regex() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "toMainVowel", LoadRawToneMapping_Regex_ToMainVowel()
    d.Add "toFirstVowel", LoadRawToneMapping_Regex_ToFirstVowel()
    Set LoadRawToneMapping_Regex = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_Regex", Err.description
End Function

Private Function LoadRawToneMapping_ProperNounHandling_SurnameGuard_Surnames() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "Nguy" & ChrW(&H1EC5) & "n"
    c.Add "Tr" & ChrW(&H1EA7) & "n"
    c.Add "L" & ChrW(&HEA)
    c.Add "Ph" & ChrW(&H1EA1) & "m"
    c.Add "Ho" & ChrW(&HE0) & "ng"
    c.Add "Hu" & ChrW(&H1EF3) & "nh"
    c.Add "Phan"
    c.Add "V" & ChrW(&H169)
    c.Add "V" & ChrW(&HF5)
    c.Add ChrW(&H110) & ChrW(&H1EB7) & "ng"
    c.Add "B" & ChrW(&HF9) & "i"
    c.Add ChrW(&H110) & ChrW(&H1ED7)
    c.Add "H" & ChrW(&H1ED3)
    c.Add "Ng" & ChrW(&HF4)
    c.Add "D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "L" & ChrW(&HFD)
    Set LoadRawToneMapping_ProperNounHandling_SurnameGuard_Surnames = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_ProperNounHandling_SurnameGuard_Surnames", Err.description
End Function

Private Function LoadRawToneMapping_ProperNounHandling_SurnameGuard() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "enabled", True
    d.Add "surnames", LoadRawToneMapping_ProperNounHandling_SurnameGuard_Surnames()
    Set LoadRawToneMapping_ProperNounHandling_SurnameGuard = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_ProperNounHandling_SurnameGuard", Err.description
End Function

Private Function LoadRawToneMapping_ProperNounHandling() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "lowercase", "convert"
    d.Add "capitalizedAtSentenceStart", "convert"
    d.Add "capitalizedMidSentence", "checkPlaceNameList"
    d.Add "placeNameListFile", "dia-danh-viet-nam.json"
    d.Add "matchWholePhrase", True
    d.Add "surnameGuard", LoadRawToneMapping_ProperNounHandling_SurnameGuard()
    Set LoadRawToneMapping_ProperNounHandling = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_ProperNounHandling", Err.description
End Function

Private Function LoadRawToneMapping_WordInitialForms_ToFirstVowel() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "o" & ChrW(&HE0), ChrW(&HF2) & "a"
    d.Add "o" & ChrW(&HE1), ChrW(&HF3) & "a"
    d.Add "o" & ChrW(&H1EA1), ChrW(&H1ECD) & "a"
    d.Add "u" & ChrW(&H1EF7), ChrW(&H1EE7) & "y"
    Set LoadRawToneMapping_WordInitialForms_ToFirstVowel = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_WordInitialForms_ToFirstVowel", Err.description
End Function

Private Function LoadRawToneMapping_WordInitialForms_ToMainVowel() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add ChrW(&HF2) & "a", "o" & ChrW(&HE0)
    d.Add ChrW(&HF3) & "a", "o" & ChrW(&HE1)
    d.Add ChrW(&H1ECD) & "a", "o" & ChrW(&H1EA1)
    d.Add ChrW(&H1EE7) & "y", "u" & ChrW(&H1EF7)
    Set LoadRawToneMapping_WordInitialForms_ToMainVowel = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_WordInitialForms_ToMainVowel", Err.description
End Function

Private Function LoadRawToneMapping_WordInitialForms() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "enabled", True
    d.Add "toFirstVowel", LoadRawToneMapping_WordInitialForms_ToFirstVowel()
    d.Add "toMainVowel", LoadRawToneMapping_WordInitialForms_ToMainVowel()
    Set LoadRawToneMapping_WordInitialForms = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_WordInitialForms", Err.description
End Function

Private Function LoadRawToneMapping_ExcludeComponentRoles_Roles() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "nationalTitle"
    c.Add "nationalMotto"
    c.Add "organName"
    c.Add "superiorOrganName"
    c.Add "signerName"
    Set LoadRawToneMapping_ExcludeComponentRoles_Roles = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_ExcludeComponentRoles_Roles", Err.description
End Function

Private Function LoadRawToneMapping_ExcludeComponentRoles() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "roles", LoadRawToneMapping_ExcludeComponentRoles_Roles()
    Set LoadRawToneMapping_ExcludeComponentRoles = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_ExcludeComponentRoles", Err.description
End Function

Private Function LoadRawToneMapping_TestCases_ToMainVowel() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim c68 As New Collection
    c68.Add "h" & ChrW(&HF2) & "a b" & ChrW(&HEC) & "nh"
    c68.Add "ho" & ChrW(&HE0) & " b" & ChrW(&HEC) & "nh"
    c.Add c68
    Dim c69 As New Collection
    c69.Add "th" & ChrW(&H1EE7) & "y l" & ChrW(&H1EE3) & "i"
    c69.Add "thu" & ChrW(&H1EF7) & " l" & ChrW(&H1EE3) & "i"
    c.Add c69
    Dim c70 As New Collection
    c70.Add "kh" & ChrW(&H1ECF) & "e m" & ChrW(&H1EA1) & "nh"
    c70.Add "kho" & ChrW(&H1EBB) & " m" & ChrW(&H1EA1) & "nh"
    c.Add c70
    Dim c71 As New Collection
    c71.Add "t" & ChrW(&HF2) & "a " & ChrW(&HE1) & "n"
    c71.Add "to" & ChrW(&HE0) & " " & ChrW(&HE1) & "n"
    c.Add c71
    Dim c72 As New Collection
    c72.Add "ng" & ChrW(&H1EE5) & "y trang"
    c72.Add "ngu" & ChrW(&H1EF5) & " trang"
    c.Add c72
    Dim c73 As New Collection
    c73.Add "h" & ChrW(&H1ECD) & "a s" & ChrW(&H129)
    c73.Add "ho" & ChrW(&H1EA1) & " s" & ChrW(&H129)
    c.Add c73
    Dim c74 As New Collection
    c74.Add ChrW(&H1EE7) & "y ban"
    c74.Add "u" & ChrW(&H1EF7) & " ban"
    c.Add c74
    Dim c75 As New Collection
    c75.Add "t" & ChrW(&HF9) & "y theo"
    c75.Add "tu" & ChrW(&H1EF3) & " theo"
    c.Add c75
    Dim c76 As New Collection
    c76.Add "x" & ChrW(&HF3) & "a b" & ChrW(&H1ECF)
    c76.Add "xo" & ChrW(&HE1) & " b" & ChrW(&H1ECF)
    c.Add c76
    Dim c77 As New Collection
    c77.Add "l" & ChrW(&HF2) & "e lo" & ChrW(&H1EB9) & "t"
    c77.Add "lo" & ChrW(&HE8) & " lo" & ChrW(&H1EB9) & "t"
    c.Add c77
    Dim c78 As New Collection
    c78.Add "qu" & ChrW(&HFD) & " h" & ChrW(&HF3) & "a"
    c78.Add "qu" & ChrW(&HFD) & " ho" & ChrW(&HE1)
    c.Add c78
    Set LoadRawToneMapping_TestCases_ToMainVowel = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_TestCases_ToMainVowel", Err.description
End Function

Private Function LoadRawToneMapping_TestCases_ToFirstVowel() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim c79 As New Collection
    c79.Add "ho" & ChrW(&HE0) & " b" & ChrW(&HEC) & "nh"
    c79.Add "h" & ChrW(&HF2) & "a b" & ChrW(&HEC) & "nh"
    c.Add c79
    Dim c80 As New Collection
    c80.Add "thu" & ChrW(&H1EF7) & " l" & ChrW(&H1EE3) & "i"
    c80.Add "th" & ChrW(&H1EE7) & "y l" & ChrW(&H1EE3) & "i"
    c.Add c80
    Dim c81 As New Collection
    c81.Add "kho" & ChrW(&H1EBB) & " m" & ChrW(&H1EA1) & "nh"
    c81.Add "kh" & ChrW(&H1ECF) & "e m" & ChrW(&H1EA1) & "nh"
    c.Add c81
    Dim c82 As New Collection
    c82.Add "to" & ChrW(&HE0) & " " & ChrW(&HE1) & "n"
    c82.Add "t" & ChrW(&HF2) & "a " & ChrW(&HE1) & "n"
    c.Add c82
    Dim c83 As New Collection
    c83.Add "ngu" & ChrW(&H1EF5) & " trang"
    c83.Add "ng" & ChrW(&H1EE5) & "y trang"
    c.Add c83
    Dim c84 As New Collection
    c84.Add "ho" & ChrW(&H1EA1) & " s" & ChrW(&H129)
    c84.Add "h" & ChrW(&H1ECD) & "a s" & ChrW(&H129)
    c.Add c84
    Dim c85 As New Collection
    c85.Add "u" & ChrW(&H1EF7) & " ban"
    c85.Add ChrW(&H1EE7) & "y ban"
    c.Add c85
    Dim c86 As New Collection
    c86.Add "tu" & ChrW(&H1EF3) & " theo"
    c86.Add "t" & ChrW(&HF9) & "y theo"
    c.Add c86
    Dim c87 As New Collection
    c87.Add "xo" & ChrW(&HE1) & " b" & ChrW(&H1ECF)
    c87.Add "x" & ChrW(&HF3) & "a b" & ChrW(&H1ECF)
    c.Add c87
    Dim c88 As New Collection
    c88.Add "lo" & ChrW(&HE8) & " lo" & ChrW(&H1EB9) & "t"
    c88.Add "l" & ChrW(&HF2) & "e lo" & ChrW(&H1EB9) & "t"
    c.Add c88
    Dim c89 As New Collection
    c89.Add "qu" & ChrW(&HFD) & " ho" & ChrW(&HE1)
    c89.Add "qu" & ChrW(&HFD) & " h" & ChrW(&HF3) & "a"
    c.Add c89
    Set LoadRawToneMapping_TestCases_ToFirstVowel = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_TestCases_ToFirstVowel", Err.description
End Function

Private Function LoadRawToneMapping_TestCases_MustKeepBothDirections() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ho" & ChrW(&HE0) & "n th" & ChrW(&HE0) & "nh"
    c.Add "kho" & ChrW(&H1EA3) & "n 3"
    c.Add "ngo" & ChrW(&HE0) & "i ra"
    c.Add "ngo" & ChrW(&HE1) & "y"
    c.Add "khu" & ChrW(&H1EF7) & "u tay"
    c.Add "qu" & ChrW(&HFD) & " I"
    c.Add "qu" & ChrW(&H1EF9) & " l" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "qu" & ChrW(&H1EF7)
    c.Add "quy" & ChrW(&H1EBF) & "t " & ChrW(&H111) & ChrW(&H1ECB) & "nh"
    c.Add "chuy" & ChrW(&H1EC3) & "n giao"
    c.Add "thuy" & ChrW(&H1EC1) & "n"
    c.Add "nguy" & ChrW(&HEA) & "n t" & ChrW(&H1EAF) & "c"
    c.Add "Nguy" & ChrW(&H1EC5) & "n Th" & ChrW(&H1ECB) & " Ho" & ChrW(&HE0)
    c.Add "Nguy" & ChrW(&H1EC5) & "n Th" & ChrW(&H1ECB) & " H" & ChrW(&HF2) & "a"
    c.Add "Nguy" & ChrW(&H1EC5) & "n H" & ChrW(&HF2) & "a B" & ChrW(&HEC) & "nh"
    Set LoadRawToneMapping_TestCases_MustKeepBothDirections = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_TestCases_MustKeepBothDirections", Err.description
End Function

Private Function LoadRawToneMapping_TestCases_PlaceNamesMustConvert() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim c90 As New Collection
    c90.Add "t" & ChrW(&H1EC9) & "nh Ho" & ChrW(&HE0) & " B" & ChrW(&HEC) & "nh"
    c90.Add "t" & ChrW(&H1EC9) & "nh H" & ChrW(&HF2) & "a B" & ChrW(&HEC) & "nh"
    c.Add c90
    Dim c91 As New Collection
    c91.Add "t" & ChrW(&H1EC9) & "nh Kh" & ChrW(&HE1) & "nh Ho" & ChrW(&HE0)
    c91.Add "t" & ChrW(&H1EC9) & "nh Kh" & ChrW(&HE1) & "nh H" & ChrW(&HF2) & "a"
    c.Add c91
    Dim c92 As New Collection
    c92.Add "t" & ChrW(&H1EC9) & "nh Thanh Ho" & ChrW(&HE1)
    c92.Add "t" & ChrW(&H1EC9) & "nh Thanh H" & ChrW(&HF3) & "a"
    c.Add c92
    Set LoadRawToneMapping_TestCases_PlaceNamesMustConvert = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_TestCases_PlaceNamesMustConvert", Err.description
End Function

Private Function LoadRawToneMapping_TestCases() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "toMainVowel", LoadRawToneMapping_TestCases_ToMainVowel()
    d.Add "toFirstVowel", LoadRawToneMapping_TestCases_ToFirstVowel()
    d.Add "mustKeepBothDirections", LoadRawToneMapping_TestCases_MustKeepBothDirections()
    d.Add "placeNamesMustConvert", LoadRawToneMapping_TestCases_PlaceNamesMustConvert()
    Set LoadRawToneMapping_TestCases = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping_TestCases", Err.description
End Function

Public Function LoadRawToneMapping() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "2.0.0"
    d.Add "sourceLabel", "QD1989"
    d.Add "ruleCode", "QD1989-D8-TONE"
    d.Add "actionType", "B"
    d.Add "styles", LoadRawToneMapping_Styles()
    d.Add "defaultStyle", "toneOnMainVowel"
    d.Add "mapToMainVowel", LoadRawToneMapping_MapToMainVowel()
    d.Add "mapToFirstVowel", LoadRawToneMapping_MapToFirstVowel()
    d.Add "conditions", LoadRawToneMapping_Conditions()
    d.Add "regex", LoadRawToneMapping_Regex()
    d.Add "properNounHandling", LoadRawToneMapping_ProperNounHandling()
    d.Add "wordInitialForms", LoadRawToneMapping_WordInitialForms()
    d.Add "excludeComponentRoles", LoadRawToneMapping_ExcludeComponentRoles()
    d.Add "testCases", LoadRawToneMapping_TestCases()
    Set LoadRawToneMapping = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawToneMapping", Err.description
End Function

Private Function LoadRawTerrainPlaceNames_Places() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "C" & ChrW(&H1EED) & "a L" & ChrW(&HF2)
    c.Add "C" & ChrW(&H1EED) & "a " & ChrW(&HD4) & "ng"
    c.Add "C" & ChrW(&H1EED) & "a Vi" & ChrW(&H1EC7) & "t"
    c.Add "V" & ChrW(&H169) & "ng T" & ChrW(&HE0) & "u"
    c.Add "V" & ChrW(&H169) & "ng " & ChrW(&HC1) & "ng"
    c.Add "L" & ChrW(&H1EA1) & "ch Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "V" & ChrW(&HE0) & "m C" & ChrW(&H1ECF)
    c.Add "V" & ChrW(&HE0) & "m L" & ChrW(&HE1) & "ng"
    c.Add "C" & ChrW(&H1EA7) & "u Gi" & ChrW(&H1EA5) & "y"
    c.Add "C" & ChrW(&H1EA7) & "u Di" & ChrW(&H1EC5) & "n"
    c.Add "B" & ChrW(&H1EBF) & "n Th" & ChrW(&HE0) & "nh"
    c.Add "B" & ChrW(&H1EBF) & "n Ngh" & ChrW(&HE9)
    c.Add "Ch" & ChrW(&H1EE3) & " L" & ChrW(&H1EDB) & "n"
    c.Add "H" & ChrW(&HF2) & "n Gai"
    c.Add "H" & ChrW(&HF2) & "n Ch" & ChrW(&H1ED3) & "ng"
    Set LoadRawTerrainPlaceNames_Places = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTerrainPlaceNames_Places", Err.description
End Function

Public Function LoadRawTerrainPlaceNames() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "1.0.0"
    d.Add "status", "SEED - chua day du, can bo sung"
    d.Add "sourceLabel", "ND30"
    d.Add "actionType", "TU SUA"
    d.Add "places", LoadRawTerrainPlaceNames_Places()
    Set LoadRawTerrainPlaceNames = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTerrainPlaceNames", Err.description
End Function

Private Sub LoadRawPlaceNames_PlacesPart1(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "An H" & ChrW(&HF2) & "a"
    c.Add "An Th" & ChrW(&H1EA1) & "nh Th" & ChrW(&H1EE7) & "y"
    c.Add "Bi" & ChrW(&HEA) & "n H" & ChrW(&HF2) & "a"
    c.Add "B" & ChrW(&HEC) & "nh H" & ChrW(&HF2) & "a"
    c.Add "B" & ChrW(&HEC) & "nh H" & ChrW(&H1B0) & "ng H" & ChrW(&HF2) & "a"
    c.Add "B" & ChrW(&HEC) & "nh Th" & ChrW(&H1EE7) & "y"
    c.Add "B" & ChrW(&H1EAF) & "c Ninh H" & ChrW(&HF2) & "a"
    c.Add "B" & ChrW(&H1EAF) & "c Th" & ChrW(&H1EE5) & "y Anh"
    c.Add "Chi" & ChrW(&HEA) & "m H" & ChrW(&HF3) & "a"
    c.Add "Ch" & ChrW(&HE1) & "nh Ph" & ChrW(&HFA) & " H" & ChrW(&HF2) & "a"
    c.Add "Ch" & ChrW(&HE2) & "u H" & ChrW(&HF2) & "a"
    c.Add "C" & ChrW(&H1EA3) & "nh Th" & ChrW(&H1EE5) & "y"
    c.Add "C" & ChrW(&H1EA9) & "m Th" & ChrW(&H1EE7) & "y"
    c.Add "D" & ChrW(&HE2) & "n H" & ChrW(&HF2) & "a"
    c.Add "D" & ChrW(&HE2) & "n H" & ChrW(&HF3) & "a"
    c.Add "D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng H" & ChrW(&HF2) & "a"
    c.Add "Gia H" & ChrW(&HF2) & "a"
    c.Add "Giao H" & ChrW(&HF2) & "a"
    c.Add "Giao Th" & ChrW(&H1EE7) & "y"
    c.Add "Hi" & ChrW(&H1EC7) & "p H" & ChrW(&HF2) & "a"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawPlaceNames_PlacesPart1", Err.description
End Sub

Private Sub LoadRawPlaceNames_PlacesPart2(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ho" & ChrW(&H1EB1) & "ng H" & ChrW(&HF3) & "a"
    c.Add "H" & ChrW(&HF2) & "a An"
    c.Add "H" & ChrW(&HF2) & "a B" & ChrW(&HEC) & "nh"
    c.Add "H" & ChrW(&HF2) & "a B" & ChrW(&H1EAF) & "c"
    c.Add "H" & ChrW(&HF2) & "a C" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "H" & ChrW(&HF2) & "a Hi" & ChrW(&H1EC7) & "p"
    c.Add "H" & ChrW(&HF2) & "a H" & ChrW(&H1B0) & "ng"
    c.Add "H" & ChrW(&HF2) & "a H" & ChrW(&H1EA3) & "i"
    c.Add "H" & ChrW(&HF2) & "a H" & ChrW(&H1ED9) & "i"
    c.Add "H" & ChrW(&HF2) & "a Kh" & ChrW(&HE1) & "nh"
    c.Add "H" & ChrW(&HF2) & "a Li" & ChrW(&HEA) & "n"
    c.Add "H" & ChrW(&HF2) & "a Long"
    c.Add "H" & ChrW(&HF2) & "a L" & ChrW(&H1EA1) & "c"
    c.Add "H" & ChrW(&HF2) & "a L" & ChrW(&H1EE3) & "i"
    c.Add "H" & ChrW(&HF2) & "a Minh"
    c.Add "H" & ChrW(&HF2) & "a M" & ChrW(&H1EF9)
    c.Add "H" & ChrW(&HF2) & "a Nh" & ChrW(&H1A1) & "n"
    c.Add "H" & ChrW(&HF2) & "a Ninh"
    c.Add "H" & ChrW(&HF2) & "a Phong"
    c.Add "H" & ChrW(&HF2) & "a Ph" & ChrW(&HFA)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawPlaceNames_PlacesPart2", Err.description
End Sub

Private Sub LoadRawPlaceNames_PlacesPart3(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "H" & ChrW(&HF2) & "a S" & ChrW(&H1A1) & "n"
    c.Add "H" & ChrW(&HF2) & "a Thu" & ChrW(&H1EAD) & "n"
    c.Add "H" & ChrW(&HF2) & "a Th" & ChrW(&HE0) & "nh"
    c.Add "H" & ChrW(&HF2) & "a Th" & ChrW(&H1EAF) & "ng"
    c.Add "H" & ChrW(&HF2) & "a Th" & ChrW(&H1ECB) & "nh"
    c.Add "H" & ChrW(&HF2) & "a Ti" & ChrW(&H1EBF) & "n"
    c.Add "H" & ChrW(&HF2) & "a Tr" & ChrW(&HED)
    c.Add "H" & ChrW(&HF2) & "a Tr" & ChrW(&H1EA1) & "ch"
    c.Add "H" & ChrW(&HF2) & "a T" & ChrW(&HFA)
    c.Add "H" & ChrW(&HF2) & "a Vang"
    c.Add "H" & ChrW(&HF2) & "a Xu" & ChrW(&HE2) & "n"
    c.Add "H" & ChrW(&HF2) & "a X" & ChrW(&HE1)
    c.Add "H" & ChrW(&HF2) & "a " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "n"
    c.Add "H" & ChrW(&HF3) & "a Ch" & ChrW(&HE2) & "u"
    c.Add "H" & ChrW(&HF3) & "a Qu" & ChrW(&H1EF3)
    c.Add "H" & ChrW(&HF3) & "a Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng"
    c.Add "H" & ChrW(&HF9) & "ng H" & ChrW(&HF2) & "a"
    c.Add "H" & ChrW(&H1B0) & ChrW(&H1A1) & "ng Th" & ChrW(&H1EE7) & "y"
    c.Add "H" & ChrW(&H1EA1) & " H" & ChrW(&HF2) & "a"
    c.Add "H" & ChrW(&H1EA3) & "i H" & ChrW(&HF2) & "a"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawPlaceNames_PlacesPart3", Err.description
End Sub

Private Sub LoadRawPlaceNames_PlacesPart4(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "H" & ChrW(&H1ECF) & "a L" & ChrW(&H1EF1) & "u"
    c.Add "Kh" & ChrW(&HE1) & "nh H" & ChrW(&HF2) & "a"
    c.Add "Ki" & ChrW(&H1EBF) & "n Th" & ChrW(&H1EE5) & "y"
    c.Add "Lai H" & ChrW(&HF2) & "a"
    c.Add "Li" & ChrW(&HEA) & "n H" & ChrW(&HF2) & "a"
    c.Add "Long H" & ChrW(&HF2) & "a"
    c.Add "L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng H" & ChrW(&HF2) & "a"
    c.Add "L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng H" & ChrW(&HF2) & "a L" & ChrW(&H1EA1) & "c"
    c.Add "L" & ChrW(&H1EA1) & "c Th" & ChrW(&H1EE7) & "y"
    c.Add "L" & ChrW(&H1EC7) & " Th" & ChrW(&H1EE7) & "y"
    c.Add "Minh H" & ChrW(&HF2) & "a"
    c.Add "Minh H" & ChrW(&HF3) & "a"
    c.Add "M" & ChrW(&H1ED9) & "c H" & ChrW(&HF3) & "a"
    c.Add "M" & ChrW(&H1EF9) & " Ch" & ChrW(&HE1) & "nh H" & ChrW(&HF2) & "a"
    c.Add "M" & ChrW(&H1EF9) & " H" & ChrW(&HF2) & "a H" & ChrW(&H1B0) & "ng"
    c.Add "M" & ChrW(&H1EF9) & " Th" & ChrW(&H1EE7) & "y"
    c.Add "Nam H" & ChrW(&HF2) & "a"
    c.Add "Nam Ninh H" & ChrW(&HF2) & "a"
    c.Add "Nam Th" & ChrW(&H1EE5) & "y Anh"
    c.Add "Ngh" & ChrW(&H129) & "a H" & ChrW(&HF2) & "a"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawPlaceNames_PlacesPart4", Err.description
End Sub

Private Sub LoadRawPlaceNames_PlacesPart5(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Nguy" & ChrW(&H1EC7) & "t H" & ChrW(&HF3) & "a"
    c.Add "Ng" & ChrW(&H1EE5) & "y Nh" & ChrW(&H1B0)
    c.Add "Nh" & ChrW(&HE2) & "n H" & ChrW(&HF2) & "a"
    c.Add "Nh" & ChrW(&H1A1) & "n H" & ChrW(&HF2) & "a L" & ChrW(&H1EAD) & "p"
    c.Add "Nh" & ChrW(&H1EA5) & "t H" & ChrW(&HF2) & "a"
    c.Add "Ninh H" & ChrW(&HF2) & "a"
    c.Add "Phong H" & ChrW(&HF2) & "a"
    c.Add "Ph" & ChrW(&HFA) & " H" & ChrW(&HF2) & "a"
    c.Add "Ph" & ChrW(&HFA) & " H" & ChrW(&HF2) & "a 1"
    c.Add "Ph" & ChrW(&HFA) & " H" & ChrW(&HF2) & "a 2"
    c.Add "Ph" & ChrW(&HFA) & " H" & ChrW(&HF2) & "a " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "Ph" & ChrW(&HFA) & " Th" & ChrW(&H1ECD) & " H" & ChrW(&HF2) & "a"
    c.Add "Ph" & ChrW(&HFA) & " Th" & ChrW(&H1EE7) & "y"
    c.Add "Ph" & ChrW(&HFA) & "c H" & ChrW(&HF2) & "a"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c H" & ChrW(&HF2) & "a"
    c.Add "Ph" & ChrW(&H1EE5) & "c H" & ChrW(&HF2) & "a"
    c.Add "Qu" & ChrW(&HFD) & " H" & ChrW(&HF2) & "a"
    c.Add "Qu" & ChrW(&H1EA3) & "ng H" & ChrW(&HF2) & "a"
    c.Add "Qu" & ChrW(&H1EA3) & "ng H" & ChrW(&HF3) & "a"
    c.Add "S" & ChrW(&HF4) & "ng L" & ChrW(&H169) & "y"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawPlaceNames_PlacesPart5", Err.description
End Sub

Private Sub LoadRawPlaceNames_PlacesPart6(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "S" & ChrW(&H1A1) & "n H" & ChrW(&HF2) & "a"
    c.Add "S" & ChrW(&H1A1) & "n Th" & ChrW(&H1EE7) & "y"
    c.Add "Thanh H" & ChrW(&HF2) & "a"
    c.Add "Thanh H" & ChrW(&HF3) & "a"
    c.Add "Thanh Th" & ChrW(&H1EE7) & "y"
    c.Add "Thi" & ChrW(&H1EC7) & "n H" & ChrW(&HF2) & "a"
    c.Add "Thi" & ChrW(&H1EC7) & "u H" & ChrW(&HF3) & "a"
    c.Add "Thu" & ChrW(&H1EAD) & "n H" & ChrW(&HF2) & "a"
    c.Add "Thu" & ChrW(&H1EAD) & "n H" & ChrW(&HF3) & "a"
    c.Add "Th" & ChrW(&HE1) & "i H" & ChrW(&HF2) & "a"
    c.Add "Th" & ChrW(&HE1) & "i Th" & ChrW(&H1EE5) & "y"
    c.Add "Th" & ChrW(&HFA) & "y S" & ChrW(&H1A1) & "n"
    c.Add "Th" & ChrW(&H1EA1) & "nh H" & ChrW(&HF2) & "a"
    c.Add "Th" & ChrW(&H1EA1) & "nh H" & ChrW(&HF3) & "a"
    c.Add "Th" & ChrW(&H1EDB) & "i H" & ChrW(&HF2) & "a"
    c.Add "Th" & ChrW(&H1EE5) & "y Anh"
    c.Add "Th" & ChrW(&H1EE5) & "y H" & ChrW(&HF9) & "ng"
    c.Add "Th" & ChrW(&H1EE7) & "y Nguy" & ChrW(&HEA) & "n"
    c.Add "Th" & ChrW(&H1EE7) & "y Xu" & ChrW(&HE2) & "n"
    c.Add "Th" & ChrW(&H1EE7) & "y " & ChrW(&H110) & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawPlaceNames_PlacesPart6", Err.description
End Sub

Private Sub LoadRawPlaceNames_PlacesPart7(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ti" & ChrW(&HEA) & "n Th" & ChrW(&H1EE7) & "y"
    c.Add "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Long H" & ChrW(&HF2) & "a"
    c.Add "Tuy H" & ChrW(&HF2) & "a"
    c.Add "Tuy" & ChrW(&HEA) & "n H" & ChrW(&HF3) & "a"
    c.Add "T" & ChrW(&HE2) & "n H" & ChrW(&HF2) & "a"
    c.Add "T" & ChrW(&HE2) & "n S" & ChrW(&H1A1) & "n H" & ChrW(&HF2) & "a"
    c.Add "T" & ChrW(&HE2) & "n Th" & ChrW(&H1EE7) & "y"
    c.Add "T" & ChrW(&HE2) & "y H" & ChrW(&HF2) & "a"
    c.Add "T" & ChrW(&HE2) & "y Ninh H" & ChrW(&HF2) & "a"
    c.Add "T" & ChrW(&HE2) & "y Th" & ChrW(&H1EE5) & "y Anh"
    c.Add "Vi" & ChrW(&H1EC7) & "t H" & ChrW(&HF2) & "a"
    c.Add "V" & ChrW(&HE2) & "n H" & ChrW(&HF2) & "a"
    c.Add "V" & ChrW(&H129) & "nh H" & ChrW(&HF2) & "a"
    c.Add "V" & ChrW(&H129) & "nh H" & ChrW(&HF2) & "a H" & ChrW(&H1B0) & "ng"
    c.Add "V" & ChrW(&H129) & "nh Th" & ChrW(&H1EE7) & "y"
    c.Add "V" & ChrW(&H1ECB) & " Th" & ChrW(&H1EE7) & "y"
    c.Add "Xu" & ChrW(&HE2) & "n H" & ChrW(&HF2) & "a"
    c.Add "Xu" & ChrW(&HE2) & "n Th" & ChrW(&H1EE7) & "y"
    c.Add "Y" & ChrW(&HEA) & "n H" & ChrW(&HF2) & "a"
    c.Add "Y" & ChrW(&HEA) & "n Th" & ChrW(&H1EE7) & "y"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawPlaceNames_PlacesPart7", Err.description
End Sub

Private Sub LoadRawPlaceNames_PlacesPart8(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add ChrW(&H110) & "i" & ChrW(&H1EC1) & "m Th" & ChrW(&H1EE5) & "y"
    c.Add ChrW(&H110) & ChrW(&HE0) & "m Th" & ChrW(&H1EE7) & "y"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng H" & ChrW(&HF2) & "a"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Ninh H" & ChrW(&HF2) & "a"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Th" & ChrW(&H1EE5) & "y Anh"
    c.Add ChrW(&H110) & ChrW(&H1EB7) & "ng Th" & ChrW(&HF9) & "y Tr" & ChrW(&HE2) & "m"
    c.Add ChrW(&H110) & ChrW(&H1ECB) & "nh H" & ChrW(&HF2) & "a"
    c.Add ChrW(&H110) & ChrW(&H1ECB) & "nh H" & ChrW(&HF3) & "a"
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c H" & ChrW(&HF2) & "a"
    c.Add ChrW(&H1EE8) & "ng H" & ChrW(&HF2) & "a"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawPlaceNames_PlacesPart8", Err.description
End Sub

Private Function LoadRawPlaceNames_Places() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    LoadRawPlaceNames_PlacesPart1 c
    LoadRawPlaceNames_PlacesPart2 c
    LoadRawPlaceNames_PlacesPart3 c
    LoadRawPlaceNames_PlacesPart4 c
    LoadRawPlaceNames_PlacesPart5 c
    LoadRawPlaceNames_PlacesPart6 c
    LoadRawPlaceNames_PlacesPart7 c
    LoadRawPlaceNames_PlacesPart8 c
    Set LoadRawPlaceNames_Places = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawPlaceNames_Places", Err.description
End Function

Public Function LoadRawPlaceNames() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "1.1.0"
    d.Add "status", "Day du theo cau truc hanh chinh 2 cap (tinh/thanh + xa/phuong) sau sap nhap 2025 - loc tu danh sach 3.321 xa/phuong + 34 tinh/thanh do chu du an cung cap (11/8/2026). Da loai bo 4 muc sai trong hat giong cu (Chuong My, Quynh Luu, Quynh Nhai, Quynh Phu - khong thuc su thuoc pham vi to hop oa/oe/uy vi la am tiet dong, khong bao gio bi doi kieu)."
    d.Add "places", LoadRawPlaceNames_Places()
    Set LoadRawPlaceNames = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawPlaceNames", Err.description
End Function

Private Sub LoadRawAdministrativeUnitNames_ProvincesPart1(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "An Giang"
    c.Add "B" & ChrW(&H1EAF) & "c Ninh"
    c.Add "Cao B" & ChrW(&H1EB1) & "ng"
    c.Add "C" & ChrW(&HE0) & " Mau"
    c.Add "C" & ChrW(&H1EA7) & "n Th" & ChrW(&H1A1)
    c.Add "Gia Lai"
    c.Add "Hu" & ChrW(&H1EBF)
    c.Add "H" & ChrW(&HE0) & " N" & ChrW(&H1ED9) & "i"
    c.Add "H" & ChrW(&HE0) & " T" & ChrW(&H129) & "nh"
    c.Add "H" & ChrW(&H1B0) & "ng Y" & ChrW(&HEA) & "n"
    c.Add "H" & ChrW(&H1EA3) & "i Ph" & ChrW(&HF2) & "ng"
    c.Add "Kh" & ChrW(&HE1) & "nh H" & ChrW(&HF2) & "a"
    c.Add "Lai Ch" & ChrW(&HE2) & "u"
    c.Add "L" & ChrW(&HE0) & "o Cai"
    c.Add "L" & ChrW(&HE2) & "m " & ChrW(&H110) & ChrW(&H1ED3) & "ng"
    c.Add "L" & ChrW(&H1EA1) & "ng S" & ChrW(&H1A1) & "n"
    c.Add "Ngh" & ChrW(&H1EC7) & " An"
    c.Add "Ninh B" & ChrW(&HEC) & "nh"
    c.Add "Ph" & ChrW(&HFA) & " Th" & ChrW(&H1ECD)
    c.Add "Qu" & ChrW(&H1EA3) & "ng Ng" & ChrW(&HE3) & "i"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_ProvincesPart1", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_ProvincesPart2(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Qu" & ChrW(&H1EA3) & "ng Ninh"
    c.Add "Qu" & ChrW(&H1EA3) & "ng Tr" & ChrW(&H1ECB)
    c.Add "S" & ChrW(&H1A1) & "n La"
    c.Add "Thanh H" & ChrW(&HF3) & "a"
    c.Add "Th" & ChrW(&HE0) & "nh ph" & ChrW(&H1ED1) & " H" & ChrW(&H1ED3) & " Ch" & ChrW(&HED) & " Minh"
    c.Add "Th" & ChrW(&HE1) & "i Nguy" & ChrW(&HEA) & "n"
    c.Add "Tuy" & ChrW(&HEA) & "n Quang"
    c.Add "T" & ChrW(&HE2) & "y Ninh"
    c.Add "V" & ChrW(&H129) & "nh Long"
    c.Add ChrW(&H110) & "i" & ChrW(&H1EC7) & "n Bi" & ChrW(&HEA) & "n"
    c.Add ChrW(&H110) & ChrW(&HE0) & " N" & ChrW(&H1EB5) & "ng"
    c.Add ChrW(&H110) & ChrW(&H1EAF) & "k L" & ChrW(&H1EAF) & "k"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng Nai"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng Th" & ChrW(&HE1) & "p"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_ProvincesPart2", Err.description
End Sub

Private Function LoadRawAdministrativeUnitNames_Provinces() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    LoadRawAdministrativeUnitNames_ProvincesPart1 c
    LoadRawAdministrativeUnitNames_ProvincesPart2 c
    Set LoadRawAdministrativeUnitNames_Provinces = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_Provinces", Err.description
End Function

Private Sub LoadRawAdministrativeUnitNames_CommunesPart1(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "1 B" & ChrW(&H1EA3) & "o L" & ChrW(&H1ED9) & "c"
    c.Add "2 B" & ChrW(&H1EA3) & "o L" & ChrW(&H1ED9) & "c"
    c.Add "3 B" & ChrW(&H1EA3) & "o L" & ChrW(&H1ED9) & "c"
    c.Add "A D" & ChrW(&H1A1) & "i"
    c.Add "A L" & ChrW(&H1B0) & ChrW(&H1EDB) & "i 1"
    c.Add "A L" & ChrW(&H1B0) & ChrW(&H1EDB) & "i 2"
    c.Add "A L" & ChrW(&H1B0) & ChrW(&H1EDB) & "i 3"
    c.Add "A L" & ChrW(&H1B0) & ChrW(&H1EDB) & "i 4"
    c.Add "A L" & ChrW(&H1B0) & ChrW(&H1EDB) & "i 5"
    c.Add "A M" & ChrW(&HFA) & " Sung"
    c.Add "A S" & ChrW(&HE0) & "o"
    c.Add "Al B" & ChrW(&HE1)
    c.Add "An Bi" & ChrW(&HEA) & "n"
    c.Add "An B" & ChrW(&HEC) & "nh"
    c.Add "An Ch" & ChrW(&HE2) & "u"
    c.Add "An C" & ChrW(&H1B0)
    c.Add "An C" & ChrW(&H1EF1) & "u"
    c.Add "An D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "An Hi" & ChrW(&H1EC7) & "p"
    c.Add "An H" & ChrW(&HF2) & "a"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart1", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart2(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "An H" & ChrW(&H1B0) & "ng"
    c.Add "An H" & ChrW(&H1EA3) & "i"
    c.Add "An H" & ChrW(&H1ED9) & "i"
    c.Add "An H" & ChrW(&H1ED9) & "i T" & ChrW(&HE2) & "y"
    c.Add "An H" & ChrW(&H1ED9) & "i " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "An H" & ChrW(&H1EEF) & "u"
    c.Add "An Kh" & ChrW(&HE1) & "nh"
    c.Add "An Kh" & ChrW(&HEA)
    c.Add "An Long"
    c.Add "An L" & ChrW(&HE3) & "o"
    c.Add "An L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "An L" & ChrW(&H1EA1) & "c"
    c.Add "An L" & ChrW(&H1EA1) & "c Th" & ChrW(&HF4) & "n"
    c.Add "An L" & ChrW(&H1ED9) & "c"
    c.Add "An L" & ChrW(&H1EE5) & "c Long"
    c.Add "An Minh"
    c.Add "An Ngh" & ChrW(&H129) & "a"
    c.Add "An Ng" & ChrW(&HE3) & "i Trung"
    c.Add "An Nh" & ChrW(&H1A1) & "n"
    c.Add "An Nh" & ChrW(&H1A1) & "n B" & ChrW(&H1EAF) & "c"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart2", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart3(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "An Nh" & ChrW(&H1A1) & "n Nam"
    c.Add "An Nh" & ChrW(&H1A1) & "n T" & ChrW(&HE2) & "y"
    c.Add "An Nh" & ChrW(&H1A1) & "n " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "An Ninh"
    c.Add "An N" & ChrW(&HF4) & "ng"
    c.Add "An Phong"
    c.Add "An Ph" & ChrW(&HFA)
    c.Add "An Ph" & ChrW(&HFA) & " T" & ChrW(&HE2) & "n"
    c.Add "An Ph" & ChrW(&HFA) & " " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "An Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add "An Quang"
    c.Add "An Qui"
    c.Add "An Sinh"
    c.Add "An Th" & ChrW(&HE0) & "nh"
    c.Add "An Th" & ChrW(&H1EA1) & "nh"
    c.Add "An Th" & ChrW(&H1EA1) & "nh Th" & ChrW(&H1EE7) & "y"
    c.Add "An Th" & ChrW(&H1EAF) & "ng"
    c.Add "An Th" & ChrW(&H1EDB) & "i " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "An To" & ChrW(&HE0) & "n"
    c.Add "An Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart3", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart4(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "An Tr" & ChrW(&H1EA1) & "ch"
    c.Add "An T" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "An T" & ChrW(&H1ECB) & "nh"
    c.Add "An Vinh"
    c.Add "An Vi" & ChrW(&H1EC5) & "n"
    c.Add "An Xuy" & ChrW(&HEA) & "n"
    c.Add "An " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "An " & ChrW(&H110) & ChrW(&H1ECB) & "nh"
    c.Add "Anh D" & ChrW(&H169) & "ng"
    c.Add "Anh S" & ChrW(&H1A1) & "n"
    c.Add "Anh S" & ChrW(&H1A1) & "n " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "Av" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Ayun"
    c.Add "Ayun Pa"
    c.Add "Ba B" & ChrW(&H1EC3)
    c.Add "Ba Ch" & ChrW(&HFA) & "c"
    c.Add "Ba Ch" & ChrW(&H1EBD)
    c.Add "Ba Dinh"
    c.Add "Ba Gia"
    c.Add "Ba L" & ChrW(&HF2) & "ng"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart4", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart5(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ba Ng" & ChrW(&HF2) & "i"
    c.Add "Ba Sao"
    c.Add "Ba S" & ChrW(&H1A1) & "n"
    c.Add "Ba Tri"
    c.Add "Ba T" & ChrW(&HF4)
    c.Add "Ba T" & ChrW(&H1A1)
    c.Add "Ba Vinh"
    c.Add "Ba V" & ChrW(&HEC)
    c.Add "Ba Xa"
    c.Add "Ba " & ChrW(&H110) & ChrW(&HEC) & "nh"
    c.Add "Ba " & ChrW(&H110) & ChrW(&H1ED3) & "n"
    c.Add "Ba " & ChrW(&H110) & ChrW(&H1ED9) & "ng"
    c.Add "Bao La"
    c.Add "Bi" & ChrW(&HEA) & "n H" & ChrW(&HF2) & "a"
    c.Add "Bi" & ChrW(&HEA) & "n S" & ChrW(&H1A1) & "n"
    c.Add "Bi" & ChrW(&H1EC3) & "n B" & ChrW(&H1EA1) & "ch"
    c.Add "Bi" & ChrW(&H1EC3) & "n H" & ChrW(&H1ED3)
    c.Add "Bi" & ChrW(&H1EC3) & "n " & ChrW(&H110) & ChrW(&H1ED9) & "ng"
    c.Add "Bi" & ChrW(&H1EC7) & "n Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng"
    c.Add "Bom Bo"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart5", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart6(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Bum N" & ChrW(&H1B0) & "a"
    c.Add "Bum T" & ChrW(&H1EDF)
    c.Add "Bu" & ChrW(&HF4) & "n H" & ChrW(&H1ED3)
    c.Add "Bu" & ChrW(&HF4) & "n Ma Thu" & ChrW(&H1ED9) & "t"
    c.Add "Bu" & ChrW(&HF4) & "n " & ChrW(&H110) & ChrW(&HF4) & "n"
    c.Add "B" & ChrW(&HE0) & " N" & ChrW(&HE0)
    c.Add "B" & ChrW(&HE0) & " R" & ChrW(&H1ECB) & "a"
    c.Add "B" & ChrW(&HE0) & " " & ChrW(&H110) & "i" & ChrW(&H1EC3) & "m"
    c.Add "B" & ChrW(&HE0) & "n C" & ChrW(&H1EDD)
    c.Add "B" & ChrW(&HE0) & "n Th" & ChrW(&H1EA1) & "ch"
    c.Add "B" & ChrW(&HE0) & "u B" & ChrW(&HE0) & "ng"
    c.Add "B" & ChrW(&HE0) & "u C" & ChrW(&H1EA1) & "n"
    c.Add "B" & ChrW(&HE0) & "u H" & ChrW(&HE0) & "m"
    c.Add "B" & ChrW(&HE0) & "u L" & ChrW(&HE2) & "m"
    c.Add "B" & ChrW(&HE1) & " Th" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add "B" & ChrW(&HE1) & " Xuy" & ChrW(&HEA) & "n"
    c.Add "B" & ChrW(&HE1) & "c " & ChrW(&HC1) & "i"
    c.Add "B" & ChrW(&HE1) & "c " & ChrW(&HC1) & "i T" & ChrW(&HE2) & "y"
    c.Add "B" & ChrW(&HE1) & "c " & ChrW(&HC1) & "i " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "B" & ChrW(&HE1) & "ch Quang"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart6", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart7(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "B" & ChrW(&HE1) & "t M" & ChrW(&H1ECD) & "t"
    c.Add "B" & ChrW(&HE1) & "t Tr" & ChrW(&HE0) & "ng"
    c.Add "B" & ChrW(&HE1) & "t X" & ChrW(&HE1) & "t"
    c.Add "B" & ChrW(&HE3) & "i Ch" & ChrW(&HE1) & "y"
    c.Add "B" & ChrW(&HEC) & "nh An"
    c.Add "B" & ChrW(&HEC) & "nh Ca"
    c.Add "B" & ChrW(&HEC) & "nh Chu" & ChrW(&H1EA9) & "n"
    c.Add "B" & ChrW(&HEC) & "nh Ch" & ChrW(&HE1) & "nh"
    c.Add "B" & ChrW(&HEC) & "nh Ch" & ChrW(&HE2) & "u"
    c.Add "B" & ChrW(&HEC) & "nh Ch" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "B" & ChrW(&HEC) & "nh C" & ChrW(&H1A1)
    c.Add "B" & ChrW(&HEC) & "nh D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "B" & ChrW(&HEC) & "nh Gia"
    c.Add "B" & ChrW(&HEC) & "nh Giang"
    c.Add "B" & ChrW(&HEC) & "nh Gi" & ChrW(&HE3)
    c.Add "B" & ChrW(&HEC) & "nh Hi" & ChrW(&H1EC7) & "p"
    c.Add "B" & ChrW(&HEC) & "nh H" & ChrW(&HE0) & "ng Trung"
    c.Add "B" & ChrW(&HEC) & "nh H" & ChrW(&HF2) & "a"
    c.Add "B" & ChrW(&HEC) & "nh H" & ChrW(&H1B0) & "ng"
    c.Add "B" & ChrW(&HEC) & "nh H" & ChrW(&H1B0) & "ng H" & ChrW(&HF2) & "a"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart7", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart8(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "B" & ChrW(&HEC) & "nh Kh" & ChrW(&HE1) & "nh"
    c.Add "B" & ChrW(&HEC) & "nh Kh" & ChrW(&HEA)
    c.Add "B" & ChrW(&HEC) & "nh Ki" & ChrW(&H1EBF) & "n"
    c.Add "B" & ChrW(&HEC) & "nh Li" & ChrW(&HEA) & "u"
    c.Add "B" & ChrW(&HEC) & "nh Long"
    c.Add "B" & ChrW(&HEC) & "nh L" & ChrW(&H1B0)
    c.Add "B" & ChrW(&HEC) & "nh L" & ChrW(&H1ED9) & "c"
    c.Add "B" & ChrW(&HEC) & "nh L" & ChrW(&H1EE3) & "i"
    c.Add "B" & ChrW(&HEC) & "nh L" & ChrW(&H1EE3) & "i Trung"
    c.Add "B" & ChrW(&HEC) & "nh L" & ChrW(&H1EE5) & "c"
    c.Add "B" & ChrW(&HEC) & "nh Minh"
    c.Add "B" & ChrW(&HEC) & "nh M" & ChrW(&H1EF9)
    c.Add "B" & ChrW(&HEC) & "nh Nguy" & ChrW(&HEA) & "n"
    c.Add "B" & ChrW(&HEC) & "nh Ninh"
    c.Add "B" & ChrW(&HEC) & "nh Ph" & ChrW(&HFA)
    c.Add "B" & ChrW(&HEC) & "nh Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add "B" & ChrW(&HEC) & "nh Qu" & ChrW(&H1EDB) & "i"
    c.Add "B" & ChrW(&HEC) & "nh S" & ChrW(&H1A1) & "n"
    c.Add "B" & ChrW(&HEC) & "nh Thanh"
    c.Add "B" & ChrW(&HEC) & "nh Thu" & ChrW(&H1EAD) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart8", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart9(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "B" & ChrW(&HEC) & "nh Th" & ChrW(&HE0) & "nh"
    c.Add "B" & ChrW(&HEC) & "nh Th" & ChrW(&H1EA1) & "nh"
    c.Add "B" & ChrW(&HEC) & "nh Th" & ChrW(&H1EA1) & "nh " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "B" & ChrW(&HEC) & "nh Th" & ChrW(&H1EDB) & "i"
    c.Add "B" & ChrW(&HEC) & "nh Th" & ChrW(&H1EE7) & "y"
    c.Add "B" & ChrW(&HEC) & "nh Ti" & ChrW(&HEA) & "n"
    c.Add "B" & ChrW(&HEC) & "nh Tr" & ChrW(&H1B0) & "ng"
    c.Add "B" & ChrW(&HEC) & "nh Tr" & ChrW(&H1ECB) & " " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "B" & ChrW(&HEC) & "nh Tuy" & ChrW(&H1EC1) & "n"
    c.Add "B" & ChrW(&HEC) & "nh T" & ChrW(&HE2) & "n"
    c.Add "B" & ChrW(&HEC) & "nh T" & ChrW(&HE2) & "y"
    c.Add "B" & ChrW(&HEC) & "nh Xa"
    c.Add "B" & ChrW(&HEC) & "nh Xuy" & ChrW(&HEA) & "n"
    c.Add "B" & ChrW(&HEC) & "nh Xu" & ChrW(&HE2) & "n"
    c.Add "B" & ChrW(&HEC) & "nh Y" & ChrW(&HEA) & "n"
    c.Add "B" & ChrW(&HEC) & "nh " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "n"
    c.Add "B" & ChrW(&HEC) & "nh " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "B" & ChrW(&HEC) & "nh " & ChrW(&H110) & ChrW(&H1EA1) & "i"
    c.Add "B" & ChrW(&HEC) & "nh " & ChrW(&H110) & ChrW(&H1ECB) & "nh"
    c.Add "B" & ChrW(&HEC) & "nh " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart9", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart10(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "B" & ChrW(&HED) & "ch H" & ChrW(&HE0) & "o"
    c.Add "B" & ChrW(&HF3) & " Sinh"
    c.Add "B" & ChrW(&HF9) & " Gia M" & ChrW(&H1EAD) & "p"
    c.Add "B" & ChrW(&HF9) & " " & ChrW(&H110) & ChrW(&H103) & "ng"
    c.Add "B" & ChrW(&HFA) & "ng Lao"
    c.Add "B" & ChrW(&H1EA1) & "c Li" & ChrW(&HEA) & "u"
    c.Add "B" & ChrW(&H1EA1) & "ch H" & ChrW(&HE0)
    c.Add "B" & ChrW(&H1EA1) & "ch Long V" & ChrW(&H129)
    c.Add "B" & ChrW(&H1EA1) & "ch Mai"
    c.Add "B" & ChrW(&H1EA1) & "ch Ng" & ChrW(&H1ECD) & "c"
    c.Add "B" & ChrW(&H1EA1) & "ch Th" & ChrW(&HF4) & "ng"
    c.Add "B" & ChrW(&H1EA1) & "ch Xa"
    c.Add "B" & ChrW(&H1EA1) & "ch " & ChrW(&H110) & ChrW(&HED) & "ch"
    c.Add "B" & ChrW(&H1EA1) & "ch " & ChrW(&H110) & ChrW(&H1EB1) & "ng"
    c.Add "B" & ChrW(&H1EA3) & "n Bo"
    c.Add "B" & ChrW(&H1EA3) & "n H" & ChrW(&H1ED3)
    c.Add "B" & ChrW(&H1EA3) & "n Li" & ChrW(&H1EC1) & "n"
    c.Add "B" & ChrW(&H1EA3) & "n L" & ChrW(&H1EA7) & "u"
    c.Add "B" & ChrW(&H1EA3) & "n M" & ChrW(&HE1) & "y"
    c.Add "B" & ChrW(&H1EA3) & "n Nguy" & ChrW(&HEA) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart10", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart11(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "B" & ChrW(&H1EA3) & "n X" & ChrW(&HE8) & "o"
    c.Add "B" & ChrW(&H1EA3) & "o An"
    c.Add "B" & ChrW(&H1EA3) & "o H" & ChrW(&HE0)
    c.Add "B" & ChrW(&H1EA3) & "o L" & ChrW(&HE2) & "m"
    c.Add "B" & ChrW(&H1EA3) & "o L" & ChrW(&HE2) & "m 1"
    c.Add "B" & ChrW(&H1EA3) & "o L" & ChrW(&HE2) & "m 2"
    c.Add "B" & ChrW(&H1EA3) & "o L" & ChrW(&HE2) & "m 3"
    c.Add "B" & ChrW(&H1EA3) & "o L" & ChrW(&HE2) & "m 4"
    c.Add "B" & ChrW(&H1EA3) & "o L" & ChrW(&HE2) & "m 5"
    c.Add "B" & ChrW(&H1EA3) & "o L" & ChrW(&H1EA1) & "c"
    c.Add "B" & ChrW(&H1EA3) & "o Nhai"
    c.Add "B" & ChrW(&H1EA3) & "o Thu" & ChrW(&H1EAD) & "n"
    c.Add "B" & ChrW(&H1EA3) & "o Th" & ChrW(&H1EA1) & "nh"
    c.Add "B" & ChrW(&H1EA3) & "o Th" & ChrW(&H1EAF) & "ng"
    c.Add "B" & ChrW(&H1EA3) & "o Vinh"
    c.Add "B" & ChrW(&H1EA3) & "o Y" & ChrW(&HEA) & "n"
    c.Add "B" & ChrW(&H1EA3) & "o " & ChrW(&HC1) & "i"
    c.Add "B" & ChrW(&H1EA3) & "o " & ChrW(&H110) & ChrW(&HE0) & "i"
    c.Add "B" & ChrW(&H1EA3) & "y Hi" & ChrW(&H1EC1) & "n"
    c.Add "B" & ChrW(&H1EA5) & "t B" & ChrW(&H1EA1) & "t"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart11", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart12(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "B" & ChrW(&H1EAF) & "c An Ph" & ChrW(&H1EE5)
    c.Add "B" & ChrW(&H1EAF) & "c B" & ChrW(&HEC) & "nh"
    c.Add "B" & ChrW(&H1EAF) & "c Cam Ranh"
    c.Add "B" & ChrW(&H1EAF) & "c Gia Ngh" & ChrW(&H129) & "a"
    c.Add "B" & ChrW(&H1EAF) & "c Giang"
    c.Add "B" & ChrW(&H1EAF) & "c Gianh"
    c.Add "B" & ChrW(&H1EAF) & "c H" & ChrW(&HE0)
    c.Add "B" & ChrW(&H1EAF) & "c H" & ChrW(&H1ED3) & "ng L" & ChrW(&H129) & "nh"
    c.Add "B" & ChrW(&H1EAF) & "c Kh" & ChrW(&HE1) & "nh V" & ChrW(&H129) & "nh"
    c.Add "B" & ChrW(&H1EAF) & "c K" & ChrW(&H1EA1) & "n"
    c.Add "B" & ChrW(&H1EAF) & "c L" & ChrW(&HFD)
    c.Add "B" & ChrW(&H1EAF) & "c L" & ChrW(&H169) & "ng"
    c.Add "B" & ChrW(&H1EAF) & "c M" & ChrW(&HEA)
    c.Add "B" & ChrW(&H1EAF) & "c Nha Trang"
    c.Add "B" & ChrW(&H1EAF) & "c Ninh H" & ChrW(&HF2) & "a"
    c.Add "B" & ChrW(&H1EAF) & "c Quang"
    c.Add "B" & ChrW(&H1EAF) & "c Ru" & ChrW(&H1ED9) & "ng"
    c.Add "B" & ChrW(&H1EAF) & "c Thanh Mi" & ChrW(&H1EC7) & "n"
    c.Add "B" & ChrW(&H1EAF) & "c Th" & ChrW(&HE1) & "i Ninh"
    c.Add "B" & ChrW(&H1EAF) & "c Th" & ChrW(&H1EE5) & "y Anh"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart12", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart13(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "B" & ChrW(&H1EAF) & "c Ti" & ChrW(&HEA) & "n H" & ChrW(&H1B0) & "ng"
    c.Add "B" & ChrW(&H1EAF) & "c Tr" & ChrW(&H1EA1) & "ch"
    c.Add "B" & ChrW(&H1EAF) & "c T" & ChrW(&HE2) & "n Uy" & ChrW(&HEA) & "n"
    c.Add "B" & ChrW(&H1EAF) & "c Y" & ChrW(&HEA) & "n"
    c.Add "B" & ChrW(&H1EAF) & "c " & ChrW(&H110) & ChrW(&HF4) & "ng H" & ChrW(&H1B0) & "ng"
    c.Add "B" & ChrW(&H1EAF) & "c " & ChrW(&H110) & ChrW(&HF4) & "ng Quan"
    c.Add "B" & ChrW(&H1EB1) & "ng H" & ChrW(&HE0) & "nh"
    c.Add "B" & ChrW(&H1EB1) & "ng Lang"
    c.Add "B" & ChrW(&H1EB1) & "ng Lu" & ChrW(&HE2) & "n"
    c.Add "B" & ChrW(&H1EB1) & "ng M" & ChrW(&H1EA1) & "c"
    c.Add "B" & ChrW(&H1EB1) & "ng Th" & ChrW(&HE0) & "nh"
    c.Add "B" & ChrW(&H1EB1) & "ng V" & ChrW(&HE2) & "n"
    c.Add "B" & ChrW(&H1EBF) & " V" & ChrW(&H103) & "n " & ChrW(&H110) & ChrW(&HE0) & "n"
    c.Add "B" & ChrW(&H1EBF) & "n C" & ChrW(&HE1) & "t"
    c.Add "B" & ChrW(&H1EBF) & "n C" & ChrW(&H1EA7) & "u"
    c.Add "B" & ChrW(&H1EBF) & "n Gi" & ChrW(&H1EB1) & "ng"
    c.Add "B" & ChrW(&H1EBF) & "n Hi" & ChrW(&HEA) & "n"
    c.Add "B" & ChrW(&H1EBF) & "n H" & ChrW(&H1EA3) & "i"
    c.Add "B" & ChrW(&H1EBF) & "n L" & ChrW(&H1EE9) & "c"
    c.Add "B" & ChrW(&H1EBF) & "n Quan"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart13", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart14(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "B" & ChrW(&H1EBF) & "n Th" & ChrW(&HE0) & "nh"
    c.Add "B" & ChrW(&H1EBF) & "n Tre"
    c.Add "B" & ChrW(&H1EC9) & "m S" & ChrW(&H1A1) & "n"
    c.Add "B" & ChrW(&H1ED1) & " H" & ChrW(&H1EA1)
    c.Add "B" & ChrW(&H1ED1) & " Tr" & ChrW(&H1EA1) & "ch"
    c.Add "B" & ChrW(&H1ED3) & " " & ChrW(&H110) & ChrW(&H1EC1)
    c.Add "B" & ChrW(&H1ED3) & "ng Lai"
    c.Add "B" & ChrW(&H1ED3) & "ng S" & ChrW(&H1A1) & "n"
    c.Add "B" & ChrW(&H1EDD) & " Ngoong"
    c.Add "B" & ChrW(&H1EDD) & " Y"
    c.Add "B" & ChrW(&H2019) & "Lao"
    c.Add "Ca Th" & ChrW(&HE0) & "nh"
    c.Add "Cai Kinh"
    c.Add "Cai L" & ChrW(&H1EAD) & "y"
    c.Add "Cam An"
    c.Add "Cam Hi" & ChrW(&H1EC7) & "p"
    c.Add "Cam H" & ChrW(&H1ED3) & "ng"
    c.Add "Cam Linh"
    c.Add "Cam Ly - " & ChrW(&H110) & ChrW(&HE0) & " L" & ChrW(&H1EA1) & "t"
    c.Add "Cam L" & ChrW(&HE2) & "m"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart14", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart15(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Cam L" & ChrW(&H1ED9)
    c.Add "Cam Ph" & ChrW(&H1EE5) & "c"
    c.Add "Cam Ranh"
    c.Add "Cam " & ChrW(&H110) & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "Can L" & ChrW(&H1ED9) & "c"
    c.Add "Canh Li" & ChrW(&HEA) & "n"
    c.Add "Canh T" & ChrW(&HE2) & "n"
    c.Add "Canh Vinh"
    c.Add "Cao B" & ChrW(&H1ED3)
    c.Add "Cao D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Cao L" & ChrW(&HE3) & "nh"
    c.Add "Cao L" & ChrW(&H1ED9) & "c"
    c.Add "Cao Minh"
    c.Add "Cao Phong"
    c.Add "Cao S" & ChrW(&H1A1) & "n"
    c.Add "Cao Xanh"
    c.Add "Cao " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    c.Add "Chi L" & ChrW(&H103) & "ng"
    c.Add "Chi" & ChrW(&HEA) & "m Ho" & ChrW(&HE1)
    c.Add "Chi" & ChrW(&HEA) & "n " & ChrW(&H110) & ChrW(&HE0) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart15", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart16(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Chi" & ChrW(&HEA) & "u L" & ChrW(&H1B0) & "u"
    c.Add "Chi" & ChrW(&H1EBF) & "n Th" & ChrW(&H1EAF) & "ng"
    c.Add "Chi" & ChrW(&H1EC1) & "ng An"
    c.Add "Chi" & ChrW(&H1EC1) & "ng C" & ChrW(&H1A1) & "i"
    c.Add "Chi" & ChrW(&H1EC1) & "ng Hoa"
    c.Add "Chi" & ChrW(&H1EC1) & "ng H" & ChrW(&H1EB7) & "c"
    c.Add "Chi" & ChrW(&H1EC1) & "ng Ken"
    c.Add "Chi" & ChrW(&H1EC1) & "ng Khoong"
    c.Add "Chi" & ChrW(&H1EC1) & "ng Kh" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Chi" & ChrW(&H1EC1) & "ng La"
    c.Add "Chi" & ChrW(&H1EC1) & "ng Lao"
    c.Add "Chi" & ChrW(&H1EC1) & "ng Mai"
    c.Add "Chi" & ChrW(&H1EC1) & "ng Mung"
    c.Add "Chi" & ChrW(&H1EC1) & "ng Sinh"
    c.Add "Chi" & ChrW(&H1EC1) & "ng Sung"
    c.Add "Chi" & ChrW(&H1EC1) & "ng S" & ChrW(&H1A1)
    c.Add "Chi" & ChrW(&H1EC1) & "ng S" & ChrW(&H1A1) & "n"
    c.Add "Chi" & ChrW(&H1EC1) & "ng S" & ChrW(&H1EA1) & "i"
    c.Add "Chu V" & ChrW(&H103) & "n An"
    c.Add "Chuy" & ChrW(&HEA) & "n M" & ChrW(&H1EF9)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart16", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart17(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ch" & ChrW(&HE0) & " T" & ChrW(&H1EDF)
    c.Add "Ch" & ChrW(&HE1) & "nh Hi" & ChrW(&H1EC7) & "p"
    c.Add "Ch" & ChrW(&HE1) & "nh H" & ChrW(&H1B0) & "ng"
    c.Add "Ch" & ChrW(&HE1) & "nh Ph" & ChrW(&HFA) & " H" & ChrW(&HF2) & "a"
    c.Add "Ch" & ChrW(&HE2) & "n M" & ChrW(&HE2) & "y - L" & ChrW(&H103) & "ng C" & ChrW(&HF4)
    c.Add "Ch" & ChrW(&HE2) & "n M" & ChrW(&H1ED9) & "ng"
    c.Add "Ch" & ChrW(&HE2) & "u B" & ChrW(&HEC) & "nh"
    c.Add "Ch" & ChrW(&HE2) & "u H" & ChrW(&HF2) & "a"
    c.Add "Ch" & ChrW(&HE2) & "u H" & ChrW(&H1B0) & "ng"
    c.Add "Ch" & ChrW(&HE2) & "u H" & ChrW(&H1ED3) & "ng"
    c.Add "Ch" & ChrW(&HE2) & "u Kh" & ChrW(&HEA)
    c.Add "Ch" & ChrW(&HE2) & "u L" & ChrW(&H1ED9) & "c"
    c.Add "Ch" & ChrW(&HE2) & "u Ninh"
    c.Add "Ch" & ChrW(&HE2) & "u Pha"
    c.Add "Ch" & ChrW(&HE2) & "u Phong"
    c.Add "Ch" & ChrW(&HE2) & "u Ph" & ChrW(&HFA)
    c.Add "Ch" & ChrW(&HE2) & "u Qu" & ChrW(&H1EBF)
    c.Add "Ch" & ChrW(&HE2) & "u S" & ChrW(&H1A1) & "n"
    c.Add "Ch" & ChrW(&HE2) & "u Th" & ChrW(&HE0) & "nh"
    c.Add "Ch" & ChrW(&HE2) & "u Th" & ChrW(&H1EDB) & "i"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart17", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart18(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ch" & ChrW(&HE2) & "u Ti" & ChrW(&H1EBF) & "n"
    c.Add "Ch" & ChrW(&HE2) & "u " & ChrW(&H110) & ChrW(&H1ED1) & "c"
    c.Add "Ch" & ChrW(&HE2) & "u " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    c.Add "Ch" & ChrW(&HED) & " Linh"
    c.Add "Ch" & ChrW(&HED) & " Minh"
    c.Add "Ch" & ChrW(&HED) & " Ti" & ChrW(&HEA) & "n"
    c.Add "Ch" & ChrW(&HED) & " " & ChrW(&H110) & ChrW(&HE1) & "m"
    c.Add "Ch" & ChrW(&H169)
    c.Add "Ch" & ChrW(&H1A1) & " Long"
    c.Add "Ch" & ChrW(&H1A1) & "n Th" & ChrW(&HE0) & "nh"
    c.Add "Ch" & ChrW(&H1B0) & " A Thai"
    c.Add "Ch" & ChrW(&H1B0) & " Krey"
    c.Add "Ch" & ChrW(&H1B0) & " Pr" & ChrW(&HF4) & "ng"
    c.Add "Ch" & ChrW(&H1B0) & " P" & ChrW(&H103) & "h"
    c.Add "Ch" & ChrW(&H1B0) & " P" & ChrW(&H1B0) & "h"
    c.Add "Ch" & ChrW(&H1B0) & " S" & ChrW(&HEA)
    c.Add "Ch" & ChrW(&H1B0) & ChrW(&H1A1) & "ng D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Ch" & ChrW(&H1B0) & ChrW(&H1A1) & "ng M" & ChrW(&H1EF9)
    c.Add "Ch" & ChrW(&H1EA5) & "n H" & ChrW(&H1B0) & "ng"
    c.Add "Ch" & ChrW(&H1EA5) & "n Th" & ChrW(&H1ECB) & "nh"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart18", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart19(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ch" & ChrW(&H1EA5) & "t B" & ChrW(&HEC) & "nh"
    c.Add "Ch" & ChrW(&H1EBF) & " T" & ChrW(&H1EA1) & "o"
    c.Add "Ch" & ChrW(&H1EE3) & " G" & ChrW(&H1EA1) & "o"
    c.Add "Ch" & ChrW(&H1EE3) & " L" & ChrW(&HE1) & "ch"
    c.Add "Ch" & ChrW(&H1EE3) & " L" & ChrW(&H1EDB) & "n"
    c.Add "Ch" & ChrW(&H1EE3) & " M" & ChrW(&H1EDB) & "i"
    c.Add "Ch" & ChrW(&H1EE3) & " Qu" & ChrW(&HE1) & "n"
    c.Add "Ch" & ChrW(&H1EE3) & " R" & ChrW(&HE3)
    c.Add "Ch" & ChrW(&H1EE3) & " V" & ChrW(&HE0) & "m"
    c.Add "Ch" & ChrW(&H1EE3) & " " & ChrW(&H110) & ChrW(&H1ED3) & "n"
    c.Add "Co M" & ChrW(&H1EA1)
    c.Add "Con Cu" & ChrW(&HF4) & "ng"
    c.Add "Cu" & ChrW(&HF4) & "r " & ChrW(&H110) & ChrW(&H103) & "ng"
    c.Add "C" & ChrW(&HE0) & " N" & ChrW(&HE1)
    c.Add "C" & ChrW(&HE0) & " " & ChrW(&H110) & "am"
    c.Add "C" & ChrW(&HE0) & "ng Long"
    c.Add "C" & ChrW(&HE1) & "c S" & ChrW(&H1A1) & "n"
    c.Add "C" & ChrW(&HE1) & "i B" & ChrW(&HE8)
    c.Add "C" & ChrW(&HE1) & "i Chi" & ChrW(&HEA) & "n"
    c.Add "C" & ChrW(&HE1) & "i Kh" & ChrW(&H1EBF)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart19", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart20(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "C" & ChrW(&HE1) & "i Ngang"
    c.Add "C" & ChrW(&HE1) & "i Nhum"
    c.Add "C" & ChrW(&HE1) & "i N" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add "C" & ChrW(&HE1) & "i R" & ChrW(&H103) & "ng"
    c.Add "C" & ChrW(&HE1) & "i V" & ChrW(&H1ED3) & "n"
    c.Add "C" & ChrW(&HE1) & "i " & ChrW(&H110) & ChrW(&HF4) & "i V" & ChrW(&HE0) & "m"
    c.Add "C" & ChrW(&HE1) & "n T" & ChrW(&H1EF7)
    c.Add "C" & ChrW(&HE1) & "t H" & ChrW(&H1EA3) & "i"
    c.Add "C" & ChrW(&HE1) & "t L" & ChrW(&HE1) & "i"
    c.Add "C" & ChrW(&HE1) & "t Ng" & ChrW(&H1EA1) & "n"
    c.Add "C" & ChrW(&HE1) & "t Th" & ChrW(&HE0) & "nh"
    c.Add "C" & ChrW(&HE1) & "t Th" & ChrW(&H1ECB) & "nh"
    c.Add "C" & ChrW(&HE1) & "t Ti" & ChrW(&HEA) & "n"
    c.Add "C" & ChrW(&HE1) & "t Ti" & ChrW(&HEA) & "n 2"
    c.Add "C" & ChrW(&HE1) & "t Ti" & ChrW(&HEA) & "n 3"
    c.Add "C" & ChrW(&HE1) & "t Ti" & ChrW(&H1EBF) & "n"
    c.Add "C" & ChrW(&HF4) & " Ba"
    c.Add "C" & ChrW(&HF4) & " T" & ChrW(&HF4)
    c.Add "C" & ChrW(&HF4) & "n L" & ChrW(&HF4) & "n"
    c.Add "C" & ChrW(&HF4) & "n Minh"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart20", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart21(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "C" & ChrW(&HF4) & "n " & ChrW(&H110) & ChrW(&H1EA3) & "o"
    c.Add "C" & ChrW(&HF4) & "ng Ch" & ChrW(&HED) & "nh"
    c.Add "C" & ChrW(&HF4) & "ng H" & ChrW(&H1EA3) & "i"
    c.Add "C" & ChrW(&HF4) & "ng S" & ChrW(&H1A1) & "n"
    c.Add "C" & ChrW(&HF9) & " Lao Dung"
    c.Add "C" & ChrW(&HF9) & " Lao Gi" & ChrW(&HEA) & "ng"
    c.Add "C" & ChrW(&HFA) & "c Ph" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "C" & ChrW(&H1B0) & " Bao"
    c.Add "C" & ChrW(&H1B0) & " J" & ChrW(&HFA) & "t"
    c.Add "C" & ChrW(&H1B0) & " M" & ChrW(&H2019) & "gar"
    c.Add "C" & ChrW(&H1B0) & " M" & ChrW(&H2019) & "ta"
    c.Add "C" & ChrW(&H1B0) & " Prao"
    c.Add "C" & ChrW(&H1B0) & " Pui"
    c.Add "C" & ChrW(&H1B0) & " P" & ChrW(&H1A1) & "ng"
    c.Add "C" & ChrW(&H1B0) & " Yang"
    c.Add "C" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng L" & ChrW(&H1EE3) & "i"
    c.Add "C" & ChrW(&H1EA3) & "m Nh" & ChrW(&HE2) & "n"
    c.Add "C" & ChrW(&H1EA3) & "nh Th" & ChrW(&H1EE5) & "y"
    c.Add "C" & ChrW(&H1EA7) & "n Giu" & ChrW(&H1ED9) & "c"
    c.Add "C" & ChrW(&H1EA7) & "n Gi" & ChrW(&H1EDD)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart21", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart22(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "C" & ChrW(&H1EA7) & "n Y" & ChrW(&HEA) & "n"
    c.Add "C" & ChrW(&H1EA7) & "n " & ChrW(&H110) & ChrW(&H103) & "ng"
    c.Add "C" & ChrW(&H1EA7) & "n " & ChrW(&H110) & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add "C" & ChrW(&H1EA7) & "u Gi" & ChrW(&H1EA5) & "y"
    c.Add "C" & ChrW(&H1EA7) & "u Kh" & ChrW(&H1EDF) & "i"
    c.Add "C" & ChrW(&H1EA7) & "u Ki" & ChrW(&H1EC7) & "u"
    c.Add "C" & ChrW(&H1EA7) & "u K" & ChrW(&HE8)
    c.Add "C" & ChrW(&H1EA7) & "u Ngang"
    c.Add "C" & ChrW(&H1EA7) & "u Thia"
    c.Add "C" & ChrW(&H1EA7) & "u " & ChrW(&HD4) & "ng L" & ChrW(&HE3) & "nh"
    c.Add "C" & ChrW(&H1EA9) & "m B" & ChrW(&HEC) & "nh"
    c.Add "C" & ChrW(&H1EA9) & "m Du" & ChrW(&H1EC7)
    c.Add "C" & ChrW(&H1EA9) & "m Giang"
    c.Add "C" & ChrW(&H1EA9) & "m Gi" & ChrW(&HE0) & "ng"
    c.Add "C" & ChrW(&H1EA9) & "m H" & ChrW(&H1B0) & "ng"
    c.Add "C" & ChrW(&H1EA9) & "m Kh" & ChrW(&HEA)
    c.Add "C" & ChrW(&H1EA9) & "m L" & ChrW(&HFD)
    c.Add "C" & ChrW(&H1EA9) & "m L" & ChrW(&H1EA1) & "c"
    c.Add "C" & ChrW(&H1EA9) & "m L" & ChrW(&H1EC7)
    c.Add "C" & ChrW(&H1EA9) & "m M" & ChrW(&H1EF9)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart22", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart23(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "C" & ChrW(&H1EA9) & "m Ph" & ChrW(&H1EA3)
    c.Add "C" & ChrW(&H1EA9) & "m Th" & ChrW(&HE0) & "nh"
    c.Add "C" & ChrW(&H1EA9) & "m Th" & ChrW(&H1EA1) & "ch"
    c.Add "C" & ChrW(&H1EA9) & "m Th" & ChrW(&H1EE7) & "y"
    c.Add "C" & ChrW(&H1EA9) & "m Trung"
    c.Add "C" & ChrW(&H1EA9) & "m T" & ChrW(&HE2) & "n"
    c.Add "C" & ChrW(&H1EA9) & "m T" & ChrW(&HFA)
    c.Add "C" & ChrW(&H1EA9) & "m V" & ChrW(&HE2) & "n"
    c.Add "C" & ChrW(&H1EA9) & "m Xuy" & ChrW(&HEA) & "n"
    c.Add "C" & ChrW(&H1ED1) & "c L" & ChrW(&H1EA7) & "u"
    c.Add "C" & ChrW(&H1ED1) & "c P" & ChrW(&HE0) & "ng"
    c.Add "C" & ChrW(&H1ED1) & "c San"
    c.Add "C" & ChrW(&H1ED3) & "n C" & ChrW(&H1ECF)
    c.Add "C" & ChrW(&H1ED3) & "n Ti" & ChrW(&HEA) & "n"
    c.Add "C" & ChrW(&H1ED5) & " L" & ChrW(&H169) & "ng"
    c.Add "C" & ChrW(&H1ED5) & " L" & ChrW(&H1EC5)
    c.Add "C" & ChrW(&H1ED5) & " " & ChrW(&H110) & ChrW(&HF4)
    c.Add "C" & ChrW(&H1ED5) & " " & ChrW(&H110) & ChrW(&H1EA1) & "m"
    c.Add "C" & ChrW(&H1EDD) & " " & ChrW(&H110) & ChrW(&H1ECF)
    c.Add "C" & ChrW(&H1EE7) & " Chi"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart23", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart24(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "C" & ChrW(&H1EED) & "a L" & ChrW(&HF2)
    c.Add "C" & ChrW(&H1EED) & "a Nam"
    c.Add "C" & ChrW(&H1EED) & "a T" & ChrW(&HF9) & "ng"
    c.Add "C" & ChrW(&H1EED) & "a Vi" & ChrW(&H1EC7) & "t"
    c.Add "C" & ChrW(&H1EED) & "a " & ChrW(&HD4) & "ng"
    c.Add "C" & ChrW(&H1EED) & "u An"
    c.Add "C" & ChrW(&H1EF1) & " " & ChrW(&H110) & ChrW(&H1ED3) & "ng"
    c.Add "Dang Kang"
    c.Add "Di Linh"
    c.Add "Di" & ChrW(&HEA) & "n H" & ChrW(&HE0)
    c.Add "Di" & ChrW(&HEA) & "n H" & ChrW(&H1ED3) & "ng"
    c.Add "Di" & ChrW(&HEA) & "n Kh" & ChrW(&HE1) & "nh"
    c.Add "Di" & ChrW(&HEA) & "n L" & ChrW(&HE2) & "m"
    c.Add "Di" & ChrW(&HEA) & "n L" & ChrW(&H1EA1) & "c"
    c.Add "Di" & ChrW(&HEA) & "n Sanh"
    c.Add "Di" & ChrW(&HEA) & "n Th" & ChrW(&H1ECD)
    c.Add "Di" & ChrW(&HEA) & "n " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "n"
    c.Add "Di" & ChrW(&H1EC5) & "n Ch" & ChrW(&HE2) & "u"
    c.Add "Dli" & ChrW(&HEA) & " Ya"
    c.Add "Dray Bh" & ChrW(&H103) & "ng"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart24", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart25(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Du Gi" & ChrW(&HE0)
    c.Add "Dur Km" & ChrW(&H103) & "l"
    c.Add "Duy H" & ChrW(&HE0)
    c.Add "Duy Ngh" & ChrW(&H129) & "a"
    c.Add "Duy Ti" & ChrW(&HEA) & "n"
    c.Add "Duy T" & ChrW(&HE2) & "n"
    c.Add "Duy Xuy" & ChrW(&HEA) & "n"
    c.Add "Duy" & ChrW(&HEA) & "n H" & ChrW(&H1EA3) & "i"
    c.Add "D" & ChrW(&HE0) & "o San"
    c.Add "D" & ChrW(&HE2) & "n Ch" & ChrW(&H1EE7)
    c.Add "D" & ChrW(&HE2) & "n H" & ChrW(&HF2) & "a"
    c.Add "D" & ChrW(&HE2) & "n H" & ChrW(&HF3) & "a"
    c.Add "D" & ChrW(&HE2) & "n Ti" & ChrW(&H1EBF) & "n"
    c.Add "D" & ChrW(&H129) & " An"
    c.Add "D" & ChrW(&H169) & "ng Ti" & ChrW(&H1EBF) & "n"
    c.Add "D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng H" & ChrW(&HF2) & "a"
    c.Add "D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng H" & ChrW(&H1B0) & "u"
    c.Add "D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng Kinh"
    c.Add "D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng Minh Ch" & ChrW(&HE2) & "u"
    c.Add "D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng N" & ChrW(&H1ED7)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart25", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart26(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng N" & ChrW(&H1ED9) & "i"
    c.Add "D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng Qu" & ChrW(&H1EF3)
    c.Add "D" & ChrW(&H1EA7) & "u Gi" & ChrW(&HE2) & "y"
    c.Add "D" & ChrW(&H1EA7) & "u Ti" & ChrW(&H1EBF) & "ng"
    c.Add "D" & ChrW(&H1EC1) & "n S" & ChrW(&HE1) & "ng"
    c.Add "D" & ChrW(&H1EE5) & "c N" & ChrW(&HF4) & "ng"
    c.Add "D" & ChrW(&H2019) & "Ran"
    c.Add "Ea Bung"
    c.Add "Ea B" & ChrW(&HE1)
    c.Add "Ea Dr" & ChrW(&HF4) & "ng"
    c.Add "Ea Dr" & ChrW(&H103) & "ng"
    c.Add "Ea Hiao"
    c.Add "Ea H" & ChrW(&H2019) & "Leo"
    c.Add "Ea Kao"
    c.Add "Ea Kar"
    c.Add "Ea Kh" & ChrW(&H103) & "l"
    c.Add "Ea Ki" & ChrW(&H1EBF) & "t"
    c.Add "Ea Kly"
    c.Add "Ea Knu" & ChrW(&H1EBF) & "c"
    c.Add "Ea Kn" & ChrW(&H1ED1) & "p"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart26", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart27(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ea Ktur"
    c.Add "Ea Ly"
    c.Add "Ea M" & ChrW(&H2019) & "Droh"
    c.Add "Ea Na"
    c.Add "Ea Ning"
    c.Add "Ea Nu" & ChrW(&HF4) & "l"
    c.Add "Ea Ph" & ChrW(&HEA)
    c.Add "Ea P" & ChrW(&H103) & "l"
    c.Add "Ea Ri" & ChrW(&HEA) & "ng"
    c.Add "Ea R" & ChrW(&H1ED1) & "k"
    c.Add "Ea S" & ChrW(&HFA) & "p"
    c.Add "Ea Trang"
    c.Add "Ea Tul"
    c.Add "Ea Wer"
    c.Add "Ea Wy"
    c.Add "Ea " & ChrW(&HD4)
    c.Add "Gia B" & ChrW(&HEC) & "nh"
    c.Add "Gia Hanh"
    c.Add "Gia Hi" & ChrW(&H1EC7) & "p"
    c.Add "Gia H" & ChrW(&HF2) & "a"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart27", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart28(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Gia H" & ChrW(&H1B0) & "ng"
    c.Add "Gia H" & ChrW(&H1ED9) & "i"
    c.Add "Gia Ki" & ChrW(&H1EC7) & "m"
    c.Add "Gia L" & ChrW(&HE2) & "m"
    c.Add "Gia L" & ChrW(&H1ED9) & "c"
    c.Add "Gia Phong"
    c.Add "Gia Ph" & ChrW(&HF9)
    c.Add "Gia Ph" & ChrW(&HFA)
    c.Add "Gia Ph" & ChrW(&HFA) & "c"
    c.Add "Gia S" & ChrW(&HE0) & "ng"
    c.Add "Gia Thu" & ChrW(&H1EAD) & "n"
    c.Add "Gia Tr" & ChrW(&H1EA5) & "n"
    c.Add "Gia T" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "Gia Vi" & ChrW(&HEA) & "n"
    c.Add "Gia Vi" & ChrW(&H1EC5) & "n"
    c.Add "Gia V" & ChrW(&HE2) & "n"
    c.Add "Gia " & ChrW(&H110) & ChrW(&H1ECB) & "nh"
    c.Add "Giai L" & ChrW(&H1EA1) & "c"
    c.Add "Giai Xu" & ChrW(&HE2) & "n"
    c.Add "Giang Th" & ChrW(&HE0) & "nh"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart28", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart29(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Giao An"
    c.Add "Giao B" & ChrW(&HEC) & "nh"
    c.Add "Giao Ho" & ChrW(&HE0)
    c.Add "Giao H" & ChrW(&H1B0) & "ng"
    c.Add "Giao Long"
    c.Add "Giao Minh"
    c.Add "Giao Ninh"
    c.Add "Giao Ph" & ChrW(&HFA) & "c"
    c.Add "Giao Thu" & ChrW(&H1EF7)
    c.Add "Gio Linh"
    c.Add "Gi" & ChrW(&HE1) & " Rai"
    c.Add "Gi" & ChrW(&HE1) & "p Trung"
    c.Add "Gi" & ChrW(&H1EA3) & "ng V" & ChrW(&HF5)
    c.Add "Gi" & ChrW(&H1ED3) & "ng Ri" & ChrW(&H1EC1) & "ng"
    c.Add "Gi" & ChrW(&H1ED3) & "ng Tr" & ChrW(&HF4) & "m"
    c.Add "G" & ChrW(&HE0) & "nh H" & ChrW(&HE0) & "o"
    c.Add "G" & ChrW(&HE0) & "o"
    c.Add "G" & ChrW(&HF2) & " C" & ChrW(&HF4) & "ng"
    c.Add "G" & ChrW(&HF2) & " C" & ChrW(&HF4) & "ng " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "G" & ChrW(&HF2) & " D" & ChrW(&H1EA7) & "u"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart29", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart30(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "G" & ChrW(&HF2) & " N" & ChrW(&H1ED5) & "i"
    c.Add "G" & ChrW(&HF2) & " Quao"
    c.Add "G" & ChrW(&HF2) & " V" & ChrW(&H1EA5) & "p"
    c.Add "Hai B" & ChrW(&HE0) & " Tr" & ChrW(&H1B0) & "ng"
    c.Add "Hi" & ChrW(&H1EBF) & "u Giang"
    c.Add "Hi" & ChrW(&H1EBF) & "u Ph" & ChrW(&H1EE5) & "ng"
    c.Add "Hi" & ChrW(&H1EBF) & "u Th" & ChrW(&HE0) & "nh"
    c.Add "Hi" & ChrW(&H1EC1) & "n Ki" & ChrW(&H1EC7) & "t"
    c.Add "Hi" & ChrW(&H1EC1) & "n L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Hi" & ChrW(&H1EC1) & "n Quan"
    c.Add "Hi" & ChrW(&H1EC3) & "n Kh" & ChrW(&HE1) & "nh"
    c.Add "Hi" & ChrW(&H1EC7) & "p B" & ChrW(&HEC) & "nh"
    c.Add "Hi" & ChrW(&H1EC7) & "p C" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "Hi" & ChrW(&H1EC7) & "p Ho" & ChrW(&HE0)
    c.Add "Hi" & ChrW(&H1EC7) & "p H" & ChrW(&HF2) & "a"
    c.Add "Hi" & ChrW(&H1EC7) & "p H" & ChrW(&H1B0) & "ng"
    c.Add "Hi" & ChrW(&H1EC7) & "p L" & ChrW(&H1EF1) & "c"
    c.Add "Hi" & ChrW(&H1EC7) & "p M" & ChrW(&H1EF9)
    c.Add "Hi" & ChrW(&H1EC7) & "p Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add "Hi" & ChrW(&H1EC7) & "p Th" & ChrW(&HE0) & "nh"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart30", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart31(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Hi" & ChrW(&H1EC7) & "p Th" & ChrW(&H1EA1) & "nh"
    c.Add "Hi" & ChrW(&H1EC7) & "p " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    c.Add "Hoa L" & ChrW(&H1B0)
    c.Add "Hoa L" & ChrW(&H1ED9) & "c"
    c.Add "Hoa Qu" & ChrW(&HE2) & "n"
    c.Add "Hoa Th" & ChrW(&HE1) & "m"
    c.Add "Ho" & ChrW(&HE0) & " An"
    c.Add "Ho" & ChrW(&HE0) & "i Nh" & ChrW(&H1A1) & "n"
    c.Add "Ho" & ChrW(&HE0) & "i Nh" & ChrW(&H1A1) & "n B" & ChrW(&H1EAF) & "c"
    c.Add "Ho" & ChrW(&HE0) & "i Nh" & ChrW(&H1A1) & "n Nam"
    c.Add "Ho" & ChrW(&HE0) & "i Nh" & ChrW(&H1A1) & "n T" & ChrW(&HE2) & "y"
    c.Add "Ho" & ChrW(&HE0) & "i Nh" & ChrW(&H1A1) & "n " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "Ho" & ChrW(&HE0) & "i " & ChrW(&HC2) & "n"
    c.Add "Ho" & ChrW(&HE0) & "i " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    c.Add "Ho" & ChrW(&HE0) & "n Ki" & ChrW(&H1EBF) & "m"
    c.Add "Ho" & ChrW(&HE0) & "n Long"
    c.Add "Ho" & ChrW(&HE0) & "n L" & ChrW(&HE3) & "o"
    c.Add "Ho" & ChrW(&HE0) & "ng An"
    c.Add "Ho" & ChrW(&HE0) & "ng C" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Ho" & ChrW(&HE0) & "ng Hoa Th" & ChrW(&HE1) & "m"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart31", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart32(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ho" & ChrW(&HE0) & "ng Li" & ChrW(&H1EC7) & "t"
    c.Add "Ho" & ChrW(&HE0) & "ng Mai"
    c.Add "Ho" & ChrW(&HE0) & "ng Qu" & ChrW(&H1EBF)
    c.Add "Ho" & ChrW(&HE0) & "ng Su Ph" & ChrW(&HEC)
    c.Add "Ho" & ChrW(&HE0) & "ng V" & ChrW(&HE2) & "n"
    c.Add "Ho" & ChrW(&HE0) & "ng V" & ChrW(&H103) & "n Th" & ChrW(&H1EE5)
    c.Add "Ho" & ChrW(&HE0) & "nh B" & ChrW(&H1ED3)
    c.Add "Ho" & ChrW(&HE0) & "nh M" & ChrW(&HF4)
    c.Add "Ho" & ChrW(&HE0) & "nh S" & ChrW(&H1A1) & "n"
    c.Add "Ho" & ChrW(&H1EA1) & "t Giang"
    c.Add "Ho" & ChrW(&H1EB1) & "ng Ch" & ChrW(&HE2) & "u"
    c.Add "Ho" & ChrW(&H1EB1) & "ng Giang"
    c.Add "Ho" & ChrW(&H1EB1) & "ng H" & ChrW(&HF3) & "a"
    c.Add "Ho" & ChrW(&H1EB1) & "ng L" & ChrW(&H1ED9) & "c"
    c.Add "Ho" & ChrW(&H1EB1) & "ng Ph" & ChrW(&HFA)
    c.Add "Ho" & ChrW(&H1EB1) & "ng S" & ChrW(&H1A1) & "n"
    c.Add "Ho" & ChrW(&H1EB1) & "ng Thanh"
    c.Add "Ho" & ChrW(&H1EB1) & "ng Ti" & ChrW(&H1EBF) & "n"
    c.Add "Hra"
    c.Add "Hua Bum"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart32", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart33(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Huy Gi" & ChrW(&HE1) & "p"
    c.Add "Hu" & ChrW(&H1ED3) & "i T" & ChrW(&H1EE5)
    c.Add "Hu" & ChrW(&H1ED5) & "i M" & ChrW(&H1ED9) & "t"
    c.Add "Hy C" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "H" & ChrW(&HE0) & " An"
    c.Add "H" & ChrW(&HE0) & " B" & ChrW(&H1EAF) & "c"
    c.Add "H" & ChrW(&HE0) & " Giang 1"
    c.Add "H" & ChrW(&HE0) & " Giang 2"
    c.Add "H" & ChrW(&HE0) & " Huy T" & ChrW(&H1EAD) & "p"
    c.Add "H" & ChrW(&HE0) & " Linh"
    c.Add "H" & ChrW(&HE0) & " Long"
    c.Add "H" & ChrW(&HE0) & " L" & ChrW(&H1EA7) & "m"
    c.Add "H" & ChrW(&HE0) & " Nam"
    c.Add "H" & ChrW(&HE0) & " Nha"
    c.Add "H" & ChrW(&HE0) & " Qu" & ChrW(&H1EA3) & "ng"
    c.Add "H" & ChrW(&HE0) & " Ti" & ChrW(&HEA) & "n"
    c.Add "H" & ChrW(&HE0) & " Trung"
    c.Add "H" & ChrW(&HE0) & " Tu"
    c.Add "H" & ChrW(&HE0) & " T" & ChrW(&HE2) & "y"
    c.Add "H" & ChrW(&HE0) & " " & ChrW(&H110) & ChrW(&HF4) & "ng"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart33", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart34(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "H" & ChrW(&HE0) & "m Giang"
    c.Add "H" & ChrW(&HE0) & "m Ki" & ChrW(&H1EC7) & "m"
    c.Add "H" & ChrW(&HE0) & "m Li" & ChrW(&HEA) & "m"
    c.Add "H" & ChrW(&HE0) & "m R" & ChrW(&H1ED3) & "ng"
    c.Add "H" & ChrW(&HE0) & "m Thu" & ChrW(&H1EAD) & "n"
    c.Add "H" & ChrW(&HE0) & "m Thu" & ChrW(&H1EAD) & "n B" & ChrW(&H1EAF) & "c"
    c.Add "H" & ChrW(&HE0) & "m Thu" & ChrW(&H1EAD) & "n Nam"
    c.Add "H" & ChrW(&HE0) & "m Th" & ChrW(&H1EA1) & "nh"
    c.Add "H" & ChrW(&HE0) & "m Th" & ChrW(&H1EAF) & "ng"
    c.Add "H" & ChrW(&HE0) & "m T" & ChrW(&HE2) & "n"
    c.Add "H" & ChrW(&HE0) & "m Y" & ChrW(&HEA) & "n"
    c.Add "H" & ChrW(&HE0) & "ng G" & ChrW(&HF2) & "n"
    c.Add "H" & ChrW(&HE1) & "t M" & ChrW(&HF4) & "n"
    c.Add "H" & ChrW(&HF2) & "a An"
    c.Add "H" & ChrW(&HF2) & "a B" & ChrW(&HEC) & "nh"
    c.Add "H" & ChrW(&HF2) & "a B" & ChrW(&H1EAF) & "c"
    c.Add "H" & ChrW(&HF2) & "a C" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "H" & ChrW(&HF2) & "a Hi" & ChrW(&H1EC7) & "p"
    c.Add "H" & ChrW(&HF2) & "a H" & ChrW(&H1B0) & "ng"
    c.Add "H" & ChrW(&HF2) & "a H" & ChrW(&H1ED9) & "i"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart34", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart35(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "H" & ChrW(&HF2) & "a Kh" & ChrW(&HE1) & "nh"
    c.Add "H" & ChrW(&HF2) & "a Long"
    c.Add "H" & ChrW(&HF2) & "a L" & ChrW(&H1EA1) & "c"
    c.Add "H" & ChrW(&HF2) & "a L" & ChrW(&H1EE3) & "i"
    c.Add "H" & ChrW(&HF2) & "a Minh"
    c.Add "H" & ChrW(&HF2) & "a M" & ChrW(&H1EF9)
    c.Add "H" & ChrW(&HF2) & "a Ninh"
    c.Add "H" & ChrW(&HF2) & "a Ph" & ChrW(&HFA)
    c.Add "H" & ChrW(&HF2) & "a S" & ChrW(&H1A1) & "n"
    c.Add "H" & ChrW(&HF2) & "a Thu" & ChrW(&H1EAD) & "n"
    c.Add "H" & ChrW(&HF2) & "a Th" & ChrW(&HE0) & "nh"
    c.Add "H" & ChrW(&HF2) & "a Th" & ChrW(&H1EAF) & "ng"
    c.Add "H" & ChrW(&HF2) & "a Th" & ChrW(&H1ECB) & "nh"
    c.Add "H" & ChrW(&HF2) & "a Ti" & ChrW(&H1EBF) & "n"
    c.Add "H" & ChrW(&HF2) & "a Tr" & ChrW(&HED)
    c.Add "H" & ChrW(&HF2) & "a Tr" & ChrW(&H1EA1) & "ch"
    c.Add "H" & ChrW(&HF2) & "a T" & ChrW(&HFA)
    c.Add "H" & ChrW(&HF2) & "a Vang"
    c.Add "H" & ChrW(&HF2) & "a Xu" & ChrW(&HE2) & "n"
    c.Add "H" & ChrW(&HF2) & "a X" & ChrW(&HE1)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart35", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart36(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "H" & ChrW(&HF2) & "a " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "n"
    c.Add "H" & ChrW(&HF2) & "n Ngh" & ChrW(&H1EC7)
    c.Add "H" & ChrW(&HF2) & "n " & ChrW(&H110) & ChrW(&H1EA5) & "t"
    c.Add "H" & ChrW(&HF3) & "a Ch" & ChrW(&HE2) & "u"
    c.Add "H" & ChrW(&HF3) & "a Qu" & ChrW(&H1EF3)
    c.Add "H" & ChrW(&HF3) & "c M" & ChrW(&HF4) & "n"
    c.Add "H" & ChrW(&HF9) & "ng An"
    c.Add "H" & ChrW(&HF9) & "ng Ch" & ChrW(&HE2) & "n"
    c.Add "H" & ChrW(&HF9) & "ng Ch" & ChrW(&HE2) & "u"
    c.Add "H" & ChrW(&HF9) & "ng H" & ChrW(&HF2) & "a"
    c.Add "H" & ChrW(&HF9) & "ng L" & ChrW(&H1EE3) & "i"
    c.Add "H" & ChrW(&HF9) & "ng S" & ChrW(&H1A1) & "n"
    c.Add "H" & ChrW(&HF9) & "ng Th" & ChrW(&H1EAF) & "ng"
    c.Add "H" & ChrW(&HF9) & "ng Vi" & ChrW(&H1EC7) & "t"
    c.Add "H" & ChrW(&HF9) & "ng " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    c.Add "H" & ChrW(&H1B0) & "ng H" & ChrW(&HE0)
    c.Add "H" & ChrW(&H1B0) & "ng H" & ChrW(&H1ED9) & "i"
    c.Add "H" & ChrW(&H1B0) & "ng Kh" & ChrW(&HE1) & "nh"
    c.Add "H" & ChrW(&H1B0) & "ng Kh" & ChrW(&HE1) & "nh Trung"
    c.Add "H" & ChrW(&H1B0) & "ng Long"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart36", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart37(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "H" & ChrW(&H1B0) & "ng L" & ChrW(&H1ED9) & "c"
    c.Add "H" & ChrW(&H1B0) & "ng M" & ChrW(&H1EF9)
    c.Add "H" & ChrW(&H1B0) & "ng Nguy" & ChrW(&HEA) & "n"
    c.Add "H" & ChrW(&H1B0) & "ng Nguy" & ChrW(&HEA) & "n Nam"
    c.Add "H" & ChrW(&H1B0) & "ng Nh" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng"
    c.Add "H" & ChrW(&H1B0) & "ng Ph" & ChrW(&HFA)
    c.Add "H" & ChrW(&H1B0) & "ng Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add "H" & ChrW(&H1B0) & "ng Thu" & ChrW(&H1EAD) & "n"
    c.Add "H" & ChrW(&H1B0) & "ng Th" & ChrW(&H1EA1) & "nh"
    c.Add "H" & ChrW(&H1B0) & "ng Th" & ChrW(&H1ECB) & "nh"
    c.Add "H" & ChrW(&H1B0) & "ng V" & ChrW(&H169)
    c.Add "H" & ChrW(&H1B0) & "ng " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "n"
    c.Add "H" & ChrW(&H1B0) & "ng " & ChrW(&H110) & ChrW(&H1EA1) & "o"
    c.Add "H" & ChrW(&H1B0) & ChrW(&H1A1) & "ng An"
    c.Add "H" & ChrW(&H1B0) & ChrW(&H1A1) & "ng B" & ChrW(&HEC) & "nh"
    c.Add "H" & ChrW(&H1B0) & ChrW(&H1A1) & "ng C" & ChrW(&H1EA7) & "n"
    c.Add "H" & ChrW(&H1B0) & ChrW(&H1A1) & "ng Kh" & ChrW(&HEA)
    c.Add "H" & ChrW(&H1B0) & ChrW(&H1A1) & "ng M" & ChrW(&H1EF9)
    c.Add "H" & ChrW(&H1B0) & ChrW(&H1A1) & "ng Ph" & ChrW(&H1ED1)
    c.Add "H" & ChrW(&H1B0) & ChrW(&H1A1) & "ng S" & ChrW(&H1A1) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart37", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart38(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "H" & ChrW(&H1B0) & ChrW(&H1A1) & "ng Th" & ChrW(&H1EE7) & "y"
    c.Add "H" & ChrW(&H1B0) & ChrW(&H1A1) & "ng Tr" & ChrW(&HE0)
    c.Add "H" & ChrW(&H1B0) & ChrW(&H1A1) & "ng Xu" & ChrW(&HE2) & "n"
    c.Add "H" & ChrW(&H1B0) & ChrW(&H1A1) & "ng " & ChrW(&H110) & ChrW(&HF4)
    c.Add "H" & ChrW(&H1B0) & ChrW(&H1EDB) & "ng Hi" & ChrW(&H1EC7) & "p"
    c.Add "H" & ChrW(&H1B0) & ChrW(&H1EDB) & "ng L" & ChrW(&H1EAD) & "p"
    c.Add "H" & ChrW(&H1B0) & ChrW(&H1EDB) & "ng Ph" & ChrW(&HF9) & "ng"
    c.Add "H" & ChrW(&H1EA1) & " B" & ChrW(&H1EB1) & "ng"
    c.Add "H" & ChrW(&H1EA1) & " H" & ChrW(&HF2) & "a"
    c.Add "H" & ChrW(&H1EA1) & " Lang"
    c.Add "H" & ChrW(&H1EA1) & " Long"
    c.Add "H" & ChrW(&H1EA1) & "c Th" & ChrW(&HE0) & "nh"
    c.Add "H" & ChrW(&H1EA1) & "nh L" & ChrW(&HE2) & "m"
    c.Add "H" & ChrW(&H1EA1) & "nh Ph" & ChrW(&HFA) & "c"
    c.Add "H" & ChrW(&H1EA1) & "nh Th" & ChrW(&HF4) & "ng"
    c.Add "H" & ChrW(&H1EA1) & "p L" & ChrW(&H129) & "nh"
    c.Add "H" & ChrW(&H1EA3) & "i An"
    c.Add "H" & ChrW(&H1EA3) & "i Anh"
    c.Add "H" & ChrW(&H1EA3) & "i B" & ChrW(&HEC) & "nh"
    c.Add "H" & ChrW(&H1EA3) & "i Ch" & ChrW(&HE2) & "u"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart38", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart39(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "H" & ChrW(&H1EA3) & "i D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "H" & ChrW(&H1EA3) & "i H" & ChrW(&HF2) & "a"
    c.Add "H" & ChrW(&H1EA3) & "i H" & ChrW(&H1B0) & "ng"
    c.Add "H" & ChrW(&H1EA3) & "i H" & ChrW(&H1EAD) & "u"
    c.Add "H" & ChrW(&H1EA3) & "i L" & ChrW(&H103) & "ng"
    c.Add "H" & ChrW(&H1EA3) & "i L" & ChrW(&H129) & "nh"
    c.Add "H" & ChrW(&H1EA3) & "i L" & ChrW(&H1EA1) & "ng"
    c.Add "H" & ChrW(&H1EA3) & "i L" & ChrW(&H1ED9) & "c"
    c.Add "H" & ChrW(&H1EA3) & "i L" & ChrW(&H1EF1) & "u"
    c.Add "H" & ChrW(&H1EA3) & "i Ninh"
    c.Add "H" & ChrW(&H1EA3) & "i Quang"
    c.Add "H" & ChrW(&H1EA3) & "i S" & ChrW(&H1A1) & "n"
    c.Add "H" & ChrW(&H1EA3) & "i Th" & ChrW(&H1ECB) & "nh"
    c.Add "H" & ChrW(&H1EA3) & "i Ti" & ChrW(&H1EBF) & "n"
    c.Add "H" & ChrW(&H1EA3) & "i V" & ChrW(&HE2) & "n"
    c.Add "H" & ChrW(&H1EA3) & "i Xu" & ChrW(&HE2) & "n"
    c.Add "H" & ChrW(&H1EA3) & "o " & ChrW(&H110) & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add "H" & ChrW(&H1EAD) & "u L" & ChrW(&H1ED9) & "c"
    c.Add "H" & ChrW(&H1EAD) & "u M" & ChrW(&H1EF9)
    c.Add "H" & ChrW(&H1EAD) & "u Ngh" & ChrW(&H129) & "a"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart39", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart40(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "H" & ChrW(&H1EAD) & "u Th" & ChrW(&H1EA1) & "nh"
    c.Add "H" & ChrW(&H1ECF) & "a L" & ChrW(&H1EF1) & "u"
    c.Add "H" & ChrW(&H1ED1) & " Nai"
    c.Add "H" & ChrW(&H1ED3) & " Th" & ChrW(&H1EA7) & "u"
    c.Add "H" & ChrW(&H1ED3) & " Th" & ChrW(&H1ECB) & " K" & ChrW(&H1EF7)
    c.Add "H" & ChrW(&H1ED3) & " Tr" & ChrW(&HE0) & "m"
    c.Add "H" & ChrW(&H1ED3) & " V" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "H" & ChrW(&H1ED3) & " " & ChrW(&H110) & ChrW(&H1EAF) & "c Ki" & ChrW(&H1EC7) & "n"
    c.Add "H" & ChrW(&H1ED3) & "i Xu" & ChrW(&HE2) & "n"
    c.Add "H" & ChrW(&H1ED3) & "ng An"
    c.Add "H" & ChrW(&H1ED3) & "ng B" & ChrW(&HE0) & "ng"
    c.Add "H" & ChrW(&H1ED3) & "ng Ch" & ChrW(&HE2) & "u"
    c.Add "H" & ChrW(&H1ED3) & "ng D" & ChrW(&HE2) & "n"
    c.Add "H" & ChrW(&H1ED3) & "ng Gai"
    c.Add "H" & ChrW(&H1ED3) & "ng H" & ChrW(&HE0)
    c.Add "H" & ChrW(&H1ED3) & "ng L" & ChrW(&H1ED9) & "c"
    c.Add "H" & ChrW(&H1ED3) & "ng Minh"
    c.Add "H" & ChrW(&H1ED3) & "ng Ng" & ChrW(&H1EF1)
    c.Add "H" & ChrW(&H1ED3) & "ng Phong"
    c.Add "H" & ChrW(&H1ED3) & "ng Quang"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart40", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart41(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "H" & ChrW(&H1ED3) & "ng S" & ChrW(&H1A1) & "n"
    c.Add "H" & ChrW(&H1ED3) & "ng Thu"
    c.Add "H" & ChrW(&H1ED3) & "ng Th" & ChrW(&HE1) & "i"
    c.Add "H" & ChrW(&H1ED3) & "ng V" & ChrW(&HE2) & "n"
    c.Add "H" & ChrW(&H1ED3) & "ng V" & ChrW(&H169)
    c.Add "H" & ChrW(&H1ED9) & "i An"
    c.Add "H" & ChrW(&H1ED9) & "i An T" & ChrW(&HE2) & "y"
    c.Add "H" & ChrW(&H1ED9) & "i An " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "H" & ChrW(&H1ED9) & "i C" & ChrW(&H1B0)
    c.Add "H" & ChrW(&H1ED9) & "i Hoan"
    c.Add "H" & ChrW(&H1ED9) & "i Ph" & ChrW(&HFA)
    c.Add "H" & ChrW(&H1ED9) & "i S" & ChrW(&H1A1) & "n"
    c.Add "H" & ChrW(&H1ED9) & "i Th" & ChrW(&H1ECB) & "nh"
    c.Add "H" & ChrW(&H1EE3) & "p Kim"
    c.Add "H" & ChrW(&H1EE3) & "p L" & ChrW(&HFD)
    c.Add "H" & ChrW(&H1EE3) & "p Minh"
    c.Add "H" & ChrW(&H1EE3) & "p Th" & ChrW(&HE0) & "nh"
    c.Add "H" & ChrW(&H1EE3) & "p Th" & ChrW(&H1ECB) & "nh"
    c.Add "H" & ChrW(&H1EE3) & "p Ti" & ChrW(&H1EBF) & "n"
    c.Add "H" & ChrW(&H1EEF) & "u Khu" & ChrW(&HF4) & "ng"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart41", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart42(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "H" & ChrW(&H1EEF) & "u Ki" & ChrW(&H1EC7) & "m"
    c.Add "H" & ChrW(&H1EEF) & "u Li" & ChrW(&HEA) & "n"
    c.Add "H" & ChrW(&H1EEF) & "u L" & ChrW(&H169) & "ng"
    c.Add "Ia Bo" & ChrW(&HF2) & "ng"
    c.Add "Ia B" & ChrW(&H103) & "ng"
    c.Add "Ia Chia"
    c.Add "Ia Chim"
    c.Add "Ia Dom"
    c.Add "Ia Dreh"
    c.Add "Ia D" & ChrW(&H1A1) & "k"
    c.Add "Ia Grai"
    c.Add "Ia Hiao"
    c.Add "Ia Hrung"
    c.Add "Ia Hr" & ChrW(&HFA)
    c.Add "Ia Kh" & ChrW(&H1B0) & ChrW(&H1A1) & "l"
    c.Add "Ia Ko"
    c.Add "Ia Kr" & ChrW(&HE1) & "i"
    c.Add "Ia Kr" & ChrW(&HEA) & "l"
    c.Add "Ia Le"
    c.Add "Ia Ly"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart42", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart43(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ia L" & ChrW(&HE2) & "u"
    c.Add "Ia L" & ChrW(&H1ED1) & "p"
    c.Add "Ia M" & ChrW(&H1A1)
    c.Add "Ia Nan"
    c.Add "Ia O"
    c.Add "Ia Pa"
    c.Add "Ia Ph" & ChrW(&HED)
    c.Add "Ia Pia"
    c.Add "Ia Pn" & ChrW(&HF4) & "n"
    c.Add "Ia P" & ChrW(&HFA) & "ch"
    c.Add "Ia Rbol"
    c.Add "Ia Rsai"
    c.Add "Ia Rv" & ChrW(&HEA)
    c.Add "Ia Sao"
    c.Add "Ia Tul"
    c.Add "Ia T" & ChrW(&HF4) & "r"
    c.Add "Ia T" & ChrW(&H1A1) & "i"
    c.Add "Ia " & ChrW(&H110) & "al"
    c.Add "KDang"
    c.Add "Ka " & ChrW(&H110) & ChrW(&HF4)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart43", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart44(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Kbang"
    c.Add "Keng " & ChrW(&H110) & "u"
    c.Add "Kha S" & ChrW(&H1A1) & "n"
    c.Add "Khao Mang"
    c.Add "Khe Sanh"
    c.Add "Khe Tre"
    c.Add "Khoen On"
    c.Add "Kho" & ChrW(&HE1) & "i Ch" & ChrW(&HE2) & "u"
    c.Add "Khun H" & ChrW(&HE1)
    c.Add "Khu" & ChrW(&HF4) & "n L" & ChrW(&HF9) & "ng"
    c.Add "Khu" & ChrW(&H1EA5) & "t X" & ChrW(&HE1)
    c.Add "Kh" & ChrW(&HE1) & "ng Chi" & ChrW(&H1EBF) & "n"
    c.Add "Kh" & ChrW(&HE1) & "nh An"
    c.Add "Kh" & ChrW(&HE1) & "nh B" & ChrW(&HEC) & "nh"
    c.Add "Kh" & ChrW(&HE1) & "nh C" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "Kh" & ChrW(&HE1) & "nh H" & ChrW(&HF2) & "a"
    c.Add "Kh" & ChrW(&HE1) & "nh H" & ChrW(&H1B0) & "ng"
    c.Add "Kh" & ChrW(&HE1) & "nh H" & ChrW(&H1EAD) & "u"
    c.Add "Kh" & ChrW(&HE1) & "nh H" & ChrW(&H1ED9) & "i"
    c.Add "Kh" & ChrW(&HE1) & "nh Kh" & ChrW(&HEA)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart44", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart45(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Kh" & ChrW(&HE1) & "nh L" & ChrW(&HE2) & "m"
    c.Add "Kh" & ChrW(&HE1) & "nh Nh" & ChrW(&H1EA1) & "c"
    c.Add "Kh" & ChrW(&HE1) & "nh S" & ChrW(&H1A1) & "n"
    c.Add "Kh" & ChrW(&HE1) & "nh Thi" & ChrW(&H1EC7) & "n"
    c.Add "Kh" & ChrW(&HE1) & "nh Trung"
    c.Add "Kh" & ChrW(&HE1) & "nh V" & ChrW(&H129) & "nh"
    c.Add "Kh" & ChrW(&HE1) & "nh Xu" & ChrW(&HE2) & "n"
    c.Add "Kh" & ChrW(&HE1) & "nh Y" & ChrW(&HEA) & "n"
    c.Add "Kh" & ChrW(&HE2) & "m " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    c.Add "Kh" & ChrW(&HE2) & "u Vai"
    c.Add "Kh" & ChrW(&HFA) & "c Th" & ChrW(&H1EEB) & "a D" & ChrW(&H1EE5)
    c.Add "Kh" & ChrW(&H1B0) & ChrW(&H1A1) & "ng " & ChrW(&H110) & ChrW(&HEC) & "nh"
    c.Add "Kh" & ChrW(&H1EA3) & " C" & ChrW(&H1EED) & "u"
    c.Add "Kh" & ChrW(&H1ED5) & "ng L" & ChrW(&HE0) & "o"
    c.Add "Kim Anh"
    c.Add "Kim Bon"
    c.Add "Kim B" & ChrW(&HEC) & "nh"
    c.Add "Kim B" & ChrW(&HF4) & "i"
    c.Add "Kim B" & ChrW(&H1EA3) & "ng"
    c.Add "Kim Hoa"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart45", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart46(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Kim Li" & ChrW(&HEA) & "n"
    c.Add "Kim Long"
    c.Add "Kim Ng" & ChrW(&HE2) & "n"
    c.Add "Kim Ph" & ChrW(&HFA)
    c.Add "Kim Ph" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng"
    c.Add "Kim S" & ChrW(&H1A1) & "n"
    c.Add "Kim Thanh"
    c.Add "Kim Th" & ChrW(&HE0) & "nh"
    c.Add "Kim Tr" & ChrW(&HE0)
    c.Add "Kim T" & ChrW(&HE2) & "n"
    c.Add "Kim " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "n"
    c.Add "Kim " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "Kim " & ChrW(&H110) & ChrW(&H1ED3) & "ng"
    c.Add "Kinh B" & ChrW(&H1EAF) & "c"
    c.Add "Kinh M" & ChrW(&HF4) & "n"
    c.Add "Ki" & ChrW(&HEA) & "n H" & ChrW(&H1EA3) & "i"
    c.Add "Ki" & ChrW(&HEA) & "n Lao"
    c.Add "Ki" & ChrW(&HEA) & "n L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Ki" & ChrW(&HEA) & "n M" & ChrW(&H1ED9) & "c"
    c.Add "Ki" & ChrW(&HEA) & "n Th" & ChrW(&H1ECD)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart46", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart47(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ki" & ChrW(&HEA) & "n " & ChrW(&H110) & ChrW(&HE0) & "i"
    c.Add "Ki" & ChrW(&H1EBF) & "n An"
    c.Add "Ki" & ChrW(&H1EBF) & "n H" & ChrW(&H1B0) & "ng"
    c.Add "Ki" & ChrW(&H1EBF) & "n H" & ChrW(&H1EA3) & "i"
    c.Add "Ki" & ChrW(&H1EBF) & "n Minh"
    c.Add "Ki" & ChrW(&H1EBF) & "n Thi" & ChrW(&H1EBF) & "t"
    c.Add "Ki" & ChrW(&H1EBF) & "n Th" & ChrW(&H1EE5) & "y"
    c.Add "Ki" & ChrW(&H1EBF) & "n T" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "Ki" & ChrW(&H1EBF) & "n X" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Ki" & ChrW(&H1EBF) & "n " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    c.Add "Ki" & ChrW(&H1EC1) & "u Ph" & ChrW(&HFA)
    c.Add "Kon Braih"
    c.Add "Kon Chi" & ChrW(&HEA) & "ng"
    c.Add "Kon Gang"
    c.Add "Kon Pl" & ChrW(&HF4) & "ng"
    c.Add "Kon Tum"
    c.Add "Kon " & ChrW(&H110) & ChrW(&HE0) & "o"
    c.Add "Krong"
    c.Add "Kr" & ChrW(&HF4) & "ng Ana"
    c.Add "Kr" & ChrW(&HF4) & "ng B" & ChrW(&HF4) & "ng"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart47", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart48(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Kr" & ChrW(&HF4) & "ng B" & ChrW(&HFA) & "k"
    c.Add "Kr" & ChrW(&HF4) & "ng N" & ChrW(&HF4)
    c.Add "Kr" & ChrW(&HF4) & "ng N" & ChrW(&H103) & "ng"
    c.Add "Kr" & ChrW(&HF4) & "ng P" & ChrW(&H1EAF) & "c"
    c.Add "Kr" & ChrW(&HF4) & "ng " & ChrW(&HC1)
    c.Add "K" & ChrW(&HE9) & "p"
    c.Add "K" & ChrW(&HF4) & "ng B" & ChrW(&H1A1) & " La"
    c.Add "K" & ChrW(&HF4) & "ng Chro"
    c.Add "K" & ChrW(&H1EBB) & " S" & ChrW(&H1EB7) & "t"
    c.Add "K" & ChrW(&H1EBF) & " S" & ChrW(&HE1) & "ch"
    c.Add "K" & ChrW(&H1EF3) & " Anh"
    c.Add "K" & ChrW(&H1EF3) & " Hoa"
    c.Add "K" & ChrW(&H1EF3) & " Khang"
    c.Add "K" & ChrW(&H1EF3) & " L" & ChrW(&H1EA1) & "c"
    c.Add "K" & ChrW(&H1EF3) & " L" & ChrW(&H1EEB) & "a"
    c.Add "K" & ChrW(&H1EF3) & " S" & ChrW(&H1A1) & "n"
    c.Add "K" & ChrW(&H1EF3) & " Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng"
    c.Add "K" & ChrW(&H1EF3) & " V" & ChrW(&H103) & "n"
    c.Add "K" & ChrW(&H1EF3) & " Xu" & ChrW(&HE2) & "n"
    c.Add "La B" & ChrW(&H1EB1) & "ng"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart48", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart49(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "La D" & ChrW(&HEA) & ChrW(&HEA)
    c.Add "La D" & ChrW(&H1EA1)
    c.Add "La Gi"
    c.Add "La Hi" & ChrW(&HEA) & "n"
    c.Add "La Lay"
    c.Add "La Ng" & ChrW(&HE0)
    c.Add "La " & ChrW(&HCA) & ChrW(&HEA)
    c.Add "Lai H" & ChrW(&HF2) & "a"
    c.Add "Lai Kh" & ChrW(&HEA)
    c.Add "Lai Th" & ChrW(&HE0) & "nh"
    c.Add "Lai Vung"
    c.Add "Lai " & ChrW(&H110) & ChrW(&H1ED3) & "ng"
    c.Add "Lam S" & ChrW(&H1A1) & "n"
    c.Add "Lam Th" & ChrW(&HE0) & "nh"
    c.Add "Lam V" & ChrW(&H1EF9)
    c.Add "Lang Biang - " & ChrW(&H110) & ChrW(&HE0) & " L" & ChrW(&H1EA1) & "t"
    c.Add "Lao B" & ChrW(&H1EA3) & "o"
    c.Add "Lao Ch" & ChrW(&H1EA3) & "i"
    c.Add "Linh H" & ChrW(&H1ED3)
    c.Add "Linh S" & ChrW(&H1A1) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart49", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart50(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Linh Xu" & ChrW(&HE2) & "n"
    c.Add "Li" & ChrW(&HEA) & "m H" & ChrW(&HE0)
    c.Add "Li" & ChrW(&HEA) & "m Tuy" & ChrW(&H1EC1) & "n"
    c.Add "Li" & ChrW(&HEA) & "n B" & ChrW(&HE3) & "o"
    c.Add "Li" & ChrW(&HEA) & "n Chi" & ChrW(&H1EC3) & "u"
    c.Add "Li" & ChrW(&HEA) & "n Ch" & ChrW(&HE2) & "u"
    c.Add "Li" & ChrW(&HEA) & "n Hi" & ChrW(&H1EC7) & "p"
    c.Add "Li" & ChrW(&HEA) & "n H" & ChrW(&HF2) & "a"
    c.Add "Li" & ChrW(&HEA) & "n H" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Li" & ChrW(&HEA) & "n Minh"
    c.Add "Li" & ChrW(&HEA) & "n S" & ChrW(&H1A1) & "n"
    c.Add "Li" & ChrW(&HEA) & "n S" & ChrW(&H1A1) & "n L" & ChrW(&H1EAF) & "k"
    c.Add "Li" & ChrW(&HEA) & "u T" & ChrW(&HFA)
    c.Add "Long An"
    c.Add "Long Bi" & ChrW(&HEA) & "n"
    c.Add "Long B" & ChrW(&HEC) & "nh"
    c.Add "Long Cang"
    c.Add "Long Ch" & ChrW(&HE2) & "u"
    c.Add "Long Ch" & ChrW(&H1EEF)
    c.Add "Long C" & ChrW(&H1ED1) & "c"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart50", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart51(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Long Hi" & ChrW(&H1EC7) & "p"
    c.Add "Long Hoa"
    c.Add "Long H" & ChrW(&HE0)
    c.Add "Long H" & ChrW(&HF2) & "a"
    c.Add "Long H" & ChrW(&H1B0) & "ng"
    c.Add "Long H" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Long H" & ChrW(&H1EA3) & "i"
    c.Add "Long H" & ChrW(&H1EB9)
    c.Add "Long H" & ChrW(&H1ED3)
    c.Add "Long H" & ChrW(&H1EEF) & "u"
    c.Add "Long H" & ChrW(&H1EF1) & "u"
    c.Add "Long Kh" & ChrW(&HE1) & "nh"
    c.Add "Long Ki" & ChrW(&H1EBF) & "n"
    c.Add "Long M" & ChrW(&H1EF9)
    c.Add "Long Nguy" & ChrW(&HEA) & "n"
    c.Add "Long Ph" & ChrW(&HFA)
    c.Add "Long Ph" & ChrW(&HFA) & " 1"
    c.Add "Long Ph" & ChrW(&HFA) & " Thu" & ChrW(&H1EAD) & "n"
    c.Add "Long Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add "Long Ph" & ChrW(&H1EE5) & "ng"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart51", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart52(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Long Qu" & ChrW(&H1EA3) & "ng"
    c.Add "Long S" & ChrW(&H1A1) & "n"
    c.Add "Long Thu" & ChrW(&H1EAD) & "n"
    c.Add "Long Th" & ChrW(&HE0) & "nh"
    c.Add "Long Th" & ChrW(&H1EA1) & "nh"
    c.Add "Long Ti" & ChrW(&HEA) & "n"
    c.Add "Long Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "Long Tuy" & ChrW(&H1EC1) & "n"
    c.Add "Long V" & ChrW(&H129) & "nh"
    c.Add "Long Xuy" & ChrW(&HEA) & "n"
    c.Add "Long " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "n"
    c.Add "Long " & ChrW(&H110) & ChrW(&H1ECB) & "nh"
    c.Add "Long " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    c.Add "Lu" & ChrW(&H1EAD) & "n Th" & ChrW(&HE0) & "nh"
    c.Add "L" & ChrW(&HE0) & "o Cai"
    c.Add "L" & ChrW(&HE1) & "i Thi" & ChrW(&HEA) & "u"
    c.Add "L" & ChrW(&HE1) & "ng"
    c.Add "L" & ChrW(&HE1) & "ng Tr" & ChrW(&HF2) & "n"
    c.Add "L" & ChrW(&HE2) & "m B" & ChrW(&HEC) & "nh"
    c.Add "L" & ChrW(&HE2) & "m Giang"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart52", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart53(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "L" & ChrW(&HE2) & "m S" & ChrW(&H1A1) & "n"
    c.Add "L" & ChrW(&HE2) & "m Thao"
    c.Add "L" & ChrW(&HE2) & "m Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng"
    c.Add "L" & ChrW(&HE2) & "m T" & ChrW(&HE2) & "n"
    c.Add "L" & ChrW(&HE2) & "m Vi" & ChrW(&HEA) & "n - " & ChrW(&H110) & ChrW(&HE0) & " L" & ChrW(&H1EA1) & "t"
    c.Add "L" & ChrW(&HE2) & "n Phong"
    c.Add "L" & ChrW(&HE3) & "nh Ng" & ChrW(&H1ECD) & "c"
    c.Add "L" & ChrW(&HEA) & " Ch" & ChrW(&HE2) & "n"
    c.Add "L" & ChrW(&HEA) & " H" & ChrW(&H1ED3)
    c.Add "L" & ChrW(&HEA) & " L" & ChrW(&H1EE3) & "i"
    c.Add "L" & ChrW(&HEA) & " Qu" & ChrW(&HFD) & " " & ChrW(&H110) & ChrW(&HF4) & "n"
    c.Add "L" & ChrW(&HEA) & " Thanh Ngh" & ChrW(&H1ECB)
    c.Add "L" & ChrW(&HEA) & " " & ChrW(&HCD) & "ch M" & ChrW(&H1ED9) & "c"
    c.Add "L" & ChrW(&HEA) & " " & ChrW(&H110) & ChrW(&H1EA1) & "i H" & ChrW(&HE0) & "nh"
    c.Add "L" & ChrW(&HEC) & "a"
    c.Add "L" & ChrW(&HF3) & "ng Phi" & ChrW(&HEA) & "ng"
    c.Add "L" & ChrW(&HF3) & "ng S" & ChrW(&H1EAD) & "p"
    c.Add "L" & ChrW(&HF9) & "ng Ph" & ChrW(&HEC) & "nh"
    c.Add "L" & ChrW(&HF9) & "ng T" & ChrW(&HE1) & "m"
    c.Add "L" & ChrW(&HFD) & " B" & ChrW(&HF4) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart53", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart54(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "L" & ChrW(&HFD) & " Nh" & ChrW(&HE2) & "n"
    c.Add "L" & ChrW(&HFD) & " Qu" & ChrW(&H1ED1) & "c"
    c.Add "L" & ChrW(&HFD) & " S" & ChrW(&H1A1) & "n"
    c.Add "L" & ChrW(&HFD) & " Th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Ki" & ChrW(&H1EC7) & "t"
    c.Add "L" & ChrW(&HFD) & " V" & ChrW(&H103) & "n L" & ChrW(&HE2) & "m"
    c.Add "L" & ChrW(&H129) & "nh Nam"
    c.Add "L" & ChrW(&H129) & "nh To" & ChrW(&H1EA1) & "i"
    c.Add "L" & ChrW(&H169) & "ng C" & ChrW(&HFA)
    c.Add "L" & ChrW(&H169) & "ng N" & ChrW(&H1EB7) & "m"
    c.Add "L" & ChrW(&H169) & "ng Ph" & ChrW(&HEC) & "n"
    c.Add "L" & ChrW(&H1A1) & " Pang"
    c.Add "L" & ChrW(&H1B0) & "u Ki" & ChrW(&H1EBF) & "m"
    c.Add "L" & ChrW(&H1B0) & "u Nghi" & ChrW(&H1EC7) & "p Anh"
    c.Add "L" & ChrW(&H1B0) & "u V" & ChrW(&H1EC7)
    c.Add "L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng B" & ChrW(&H1EB1) & "ng"
    c.Add "L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng H" & ChrW(&HF2) & "a"
    c.Add "L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng H" & ChrW(&HF2) & "a L" & ChrW(&H1EA1) & "c"
    c.Add "L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng Minh"
    c.Add "L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng Ph" & ChrW(&HFA)
    c.Add "L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng S" & ChrW(&H1A1) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart54", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart55(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng Th" & ChrW(&H1EBF) & " Tr" & ChrW(&HE2) & "n"
    c.Add "L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng Th" & ChrW(&H1ECB) & "nh"
    c.Add "L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng T" & ChrW(&HE0) & "i"
    c.Add "L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng T" & ChrW(&HE2) & "m"
    c.Add "L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng V" & ChrW(&H103) & "n Tri"
    c.Add "L" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng Minh"
    c.Add "L" & ChrW(&H1EA1) & "c D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "L" & ChrW(&H1EA1) & "c L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "L" & ChrW(&H1EA1) & "c Ph" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng"
    c.Add "L" & ChrW(&H1EA1) & "c S" & ChrW(&H1A1) & "n"
    c.Add "L" & ChrW(&H1EA1) & "c Th" & ChrW(&H1EE7) & "y"
    c.Add "L" & ChrW(&H1EA1) & "c " & ChrW(&H110) & ChrW(&H1EA1) & "o"
    c.Add "L" & ChrW(&H1EA1) & "ng Giang"
    c.Add "L" & ChrW(&H1EA5) & "p V" & ChrW(&HF2)
    c.Add "L" & ChrW(&H1EAD) & "p Th" & ChrW(&H1EA1) & "ch"
    c.Add "L" & ChrW(&H1EC7) & " Ninh"
    c.Add "L" & ChrW(&H1EC7) & " Th" & ChrW(&H1EE7) & "y"
    c.Add "L" & ChrW(&H1ECB) & "ch H" & ChrW(&H1ED9) & "i Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng"
    c.Add "L" & ChrW(&H1ED9) & "c An"
    c.Add "L" & ChrW(&H1ED9) & "c B" & ChrW(&HEC) & "nh"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart55", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart56(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "L" & ChrW(&H1ED9) & "c H" & ChrW(&HE0)
    c.Add "L" & ChrW(&H1ED9) & "c H" & ChrW(&H1B0) & "ng"
    c.Add "L" & ChrW(&H1ED9) & "c Ninh"
    c.Add "L" & ChrW(&H1ED9) & "c Quang"
    c.Add "L" & ChrW(&H1ED9) & "c Thu" & ChrW(&H1EAD) & "n"
    c.Add "L" & ChrW(&H1ED9) & "c Th" & ChrW(&HE0) & "nh"
    c.Add "L" & ChrW(&H1ED9) & "c Th" & ChrW(&H1EA1) & "nh"
    c.Add "L" & ChrW(&H1ED9) & "c T" & ChrW(&H1EA5) & "n"
    c.Add "L" & ChrW(&H1EE3) & "i B" & ChrW(&HE1) & "c"
    c.Add "L" & ChrW(&H1EE5) & "c H" & ChrW(&H1ED3) & "n"
    c.Add "L" & ChrW(&H1EE5) & "c Nam"
    c.Add "L" & ChrW(&H1EE5) & "c Ng" & ChrW(&H1EA1) & "n"
    c.Add "L" & ChrW(&H1EE5) & "c S" & ChrW(&H129) & " Th" & ChrW(&HE0) & "nh"
    c.Add "L" & ChrW(&H1EE5) & "c S" & ChrW(&H1A1) & "n"
    c.Add "L" & ChrW(&H1EE5) & "c Y" & ChrW(&HEA) & "n"
    c.Add "L" & ChrW(&H1EF1) & "c H" & ChrW(&HE0) & "nh"
    c.Add "Mai Ch" & ChrW(&HE2) & "u"
    c.Add "Mai Hoa"
    c.Add "Mai H" & ChrW(&H1EA1)
    c.Add "Mai Ph" & ChrW(&H1EE5)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart56", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart57(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Mai S" & ChrW(&H1A1) & "n"
    c.Add "Mang Yang"
    c.Add "Mao " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "n"
    c.Add "Minh Ch" & ChrW(&HE2) & "u"
    c.Add "Minh H" & ChrW(&HF2) & "a"
    c.Add "Minh H" & ChrW(&HF3) & "a"
    c.Add "Minh H" & ChrW(&H1B0) & "ng"
    c.Add "Minh H" & ChrW(&H1EE3) & "p"
    c.Add "Minh Khai"
    c.Add "Minh Long"
    c.Add "Minh L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Minh Ng" & ChrW(&H1ECD) & "c"
    c.Add "Minh Ph" & ChrW(&H1EE5) & "ng"
    c.Add "Minh Quang"
    c.Add "Minh S" & ChrW(&H1A1) & "n"
    c.Add "Minh Thanh"
    c.Add "Minh Th" & ChrW(&HE1) & "i"
    c.Add "Minh Th" & ChrW(&H1EA1) & "nh"
    c.Add "Minh Th" & ChrW(&H1ECD)
    c.Add "Minh T" & ChrW(&HE2) & "m"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart57", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart58(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Minh T" & ChrW(&HE2) & "n"
    c.Add "Minh Xu" & ChrW(&HE2) & "n"
    c.Add "Minh " & ChrW(&H110) & ChrW(&HE0) & "i"
    c.Add "Minh " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    c.Add "Mu" & ChrW(&H1ED5) & "i N" & ChrW(&H1ECD) & "i"
    c.Add "M" & ChrW(&HE3) & "o " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "n"
    c.Add "M" & ChrW(&HE8) & "o V" & ChrW(&H1EA1) & "c"
    c.Add "M" & ChrW(&HEA) & " Linh"
    c.Add "M" & ChrW(&HF3) & "ng C" & ChrW(&HE1) & "i 1"
    c.Add "M" & ChrW(&HF3) & "ng C" & ChrW(&HE1) & "i 2"
    c.Add "M" & ChrW(&HF3) & "ng C" & ChrW(&HE1) & "i 3"
    c.Add "M" & ChrW(&HF4) & " Rai"
    c.Add "M" & ChrW(&HF4) & "n S" & ChrW(&H1A1) & "n"
    c.Add "M" & ChrW(&HF4) & "ng D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "M" & ChrW(&HF9) & " Cang Ch" & ChrW(&H1EA3) & "i"
    c.Add "M" & ChrW(&HF9) & " C" & ChrW(&H1EA3)
    c.Add "M" & ChrW(&H103) & "ng B" & ChrW(&HFA) & "t"
    c.Add "M" & ChrW(&H103) & "ng Ri"
    c.Add "M" & ChrW(&H103) & "ng " & ChrW(&H110) & "en"
    c.Add "M" & ChrW(&H169) & "i N" & ChrW(&HE9)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart58", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart59(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Bang"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Bi"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Bo"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng B" & ChrW(&HE1) & "m"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng B" & ChrW(&HFA)
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Chanh"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Chi" & ChrW(&HEA) & "n"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Ch" & ChrW(&HE0)
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Ch" & ChrW(&H1ECD) & "ng"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng C" & ChrW(&H1A1) & "i"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Gi" & ChrW(&HF4) & "n"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Ham"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Hoa"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Hum"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Hung"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Khi" & ChrW(&HEA) & "ng"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Khoa"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Kh" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Kim"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng La"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart59", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart60(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Lai"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Lay"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Lu" & ChrW(&HE2) & "n"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng L" & ChrW(&HE1) & "t"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng L" & ChrW(&HE8) & "o"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng L" & ChrW(&HFD)
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng L" & ChrW(&H1EA1) & "n"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng L" & ChrW(&H1EA7) & "m"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng L" & ChrW(&H1ED1) & "ng"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng M" & ChrW(&HEC) & "n"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng M" & ChrW(&HF4)
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng M" & ChrW(&HF9) & "n"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Nh" & ChrW(&HE0)
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Nh" & ChrW(&HE9)
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Ph" & ChrW(&H103) & "ng"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng P" & ChrW(&H1ED3) & "n"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Qu" & ChrW(&HE0) & "ng"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng S" & ChrW(&H1EA1) & "i"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Than"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Thanh"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart60", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart61(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Th" & ChrW(&HE0) & "ng"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Toong"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng T" & ChrW(&HE8)
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng T" & ChrW(&HED) & "p"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng T" & ChrW(&HF9) & "ng"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Vang"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng X" & ChrW(&HE9) & "n"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng " & ChrW(&HC9)
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng " & ChrW(&H110) & ChrW(&H1ED9) & "ng"
    c.Add "M" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng " & ChrW(&H1EA2) & "ng"
    c.Add "M" & ChrW(&H1EA1) & "o Kh" & ChrW(&HEA)
    c.Add "M" & ChrW(&H1EAB) & "u S" & ChrW(&H1A1) & "n"
    c.Add "M" & ChrW(&H1EAD) & "u A"
    c.Add "M" & ChrW(&H1EAD) & "u Du" & ChrW(&H1EC7)
    c.Add "M" & ChrW(&H1EAD) & "u L" & ChrW(&HE2) & "m"
    c.Add "M" & ChrW(&H1EAD) & "u Th" & ChrW(&H1EA1) & "ch"
    c.Add "M" & ChrW(&H1EC5) & " S" & ChrW(&H1EDF)
    c.Add "M" & ChrW(&H1ECF) & " C" & ChrW(&HE0) & "y"
    c.Add "M" & ChrW(&H1ECF) & " V" & ChrW(&HE0) & "ng"
    c.Add "M" & ChrW(&H1ED9) & " " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart61", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart62(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "M" & ChrW(&H1ED9) & "c Ch" & ChrW(&HE2) & "u"
    c.Add "M" & ChrW(&H1ED9) & "c H" & ChrW(&HF3) & "a"
    c.Add "M" & ChrW(&H1ED9) & "c S" & ChrW(&H1A1) & "n"
    c.Add "M" & ChrW(&H1EF9) & " An"
    c.Add "M" & ChrW(&H1EF9) & " An H" & ChrW(&H1B0) & "ng"
    c.Add "M" & ChrW(&H1EF9) & " Ch" & ChrW(&HE1) & "nh H" & ChrW(&HF2) & "a"
    c.Add "M" & ChrW(&H1EF9) & " Hi" & ChrW(&H1EC7) & "p"
    c.Add "M" & ChrW(&H1EF9) & " H" & ChrW(&HE0) & "o"
    c.Add "M" & ChrW(&H1EF9) & " H" & ChrW(&HF2) & "a H" & ChrW(&H1B0) & "ng"
    c.Add "M" & ChrW(&H1EF9) & " H" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "M" & ChrW(&H1EF9) & " H" & ChrW(&H1EA1) & "nh"
    c.Add "M" & ChrW(&H1EF9) & " Long"
    c.Add "M" & ChrW(&H1EF9) & " L" & ChrW(&HE2) & "m"
    c.Add "M" & ChrW(&H1EF9) & " L" & ChrW(&HFD)
    c.Add "M" & ChrW(&H1EF9) & " L" & ChrW(&H1EC7)
    c.Add "M" & ChrW(&H1EF9) & " L" & ChrW(&H1ED9) & "c"
    c.Add "M" & ChrW(&H1EF9) & " L" & ChrW(&H1EE3) & "i"
    c.Add "M" & ChrW(&H1EF9) & " Ng" & ChrW(&HE3) & "i"
    c.Add "M" & ChrW(&H1EF9) & " Phong"
    c.Add "M" & ChrW(&H1EF9) & " Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart62", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart63(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "M" & ChrW(&H1EF9) & " Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c T" & ChrW(&HE2) & "y"
    c.Add "M" & ChrW(&H1EF9) & " Qu" & ChrW(&HED)
    c.Add "M" & ChrW(&H1EF9) & " Qu" & ChrW(&HFD)
    c.Add "M" & ChrW(&H1EF9) & " Qu" & ChrW(&H1EDB) & "i"
    c.Add "M" & ChrW(&H1EF9) & " S" & ChrW(&H1A1) & "n"
    c.Add "M" & ChrW(&H1EF9) & " Thi" & ChrW(&H1EC7) & "n"
    c.Add "M" & ChrW(&H1EF9) & " Tho"
    c.Add "M" & ChrW(&H1EF9) & " Thu" & ChrW(&H1EAD) & "n"
    c.Add "M" & ChrW(&H1EF9) & " Th" & ChrW(&HE0) & "nh"
    c.Add "M" & ChrW(&H1EF9) & " Th" & ChrW(&HE1) & "i"
    c.Add "M" & ChrW(&H1EF9) & " Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng"
    c.Add "M" & ChrW(&H1EF9) & " Th" & ChrW(&H1EA1) & "nh"
    c.Add "M" & ChrW(&H1EF9) & " Th" & ChrW(&H1ECD)
    c.Add "M" & ChrW(&H1EF9) & " Th" & ChrW(&H1EDB) & "i"
    c.Add "M" & ChrW(&H1EF9) & " Th" & ChrW(&H1EE7) & "y"
    c.Add "M" & ChrW(&H1EF9) & " Tr" & ChrW(&HE0)
    c.Add "M" & ChrW(&H1EF9) & " T" & ChrW(&HFA)
    c.Add "M" & ChrW(&H1EF9) & " T" & ChrW(&H1ECB) & "nh An"
    c.Add "M" & ChrW(&H1EF9) & " Xuy" & ChrW(&HEA) & "n"
    c.Add "M" & ChrW(&H1EF9) & " Y" & ChrW(&HEA) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart63", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart64(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "M" & ChrW(&H1EF9) & " " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    c.Add "M" & ChrW(&H1EF9) & " " & ChrW(&H110) & ChrW(&H1EE9) & "c T" & ChrW(&HE2) & "y"
    c.Add "M" & ChrW(&H2019) & "Dr" & ChrW(&H1EAF) & "k"
    c.Add "Na D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Na Loi"
    c.Add "Na M" & ChrW(&HE8) & "o"
    c.Add "Na Ngoi"
    c.Add "Na R" & ChrW(&HEC)
    c.Add "Na Sang"
    c.Add "Na Son"
    c.Add "Na S" & ChrW(&H1EA7) & "m"
    c.Add "Nam An Ph" & ChrW(&H1EE5)
    c.Add "Nam Ba " & ChrW(&H110) & ChrW(&H1ED3) & "n"
    c.Add "Nam Ban L" & ChrW(&HE2) & "m H" & ChrW(&HE0)
    c.Add "Nam Cam Ranh"
    c.Add "Nam C" & ChrW(&HE1) & "t Ti" & ChrW(&HEA) & "n"
    c.Add "Nam C" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "Nam C" & ChrW(&H1EED) & "a Vi" & ChrW(&H1EC7) & "t"
    c.Add "Nam Dong"
    c.Add "Nam D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart64", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart65(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Nam Gia Ngh" & ChrW(&H129) & "a"
    c.Add "Nam Giang"
    c.Add "Nam Gianh"
    c.Add "Nam Hoa L" & ChrW(&H1B0)
    c.Add "Nam H" & ChrW(&HE0) & " L" & ChrW(&HE2) & "m H" & ChrW(&HE0)
    c.Add "Nam H" & ChrW(&HF2) & "a"
    c.Add "Nam H" & ChrW(&H1EA3) & "i L" & ChrW(&H103) & "ng"
    c.Add "Nam H" & ChrW(&H1ED3) & "ng"
    c.Add "Nam H" & ChrW(&H1ED3) & "ng L" & ChrW(&H129) & "nh"
    c.Add "Nam Ka"
    c.Add "Nam Kh" & ChrW(&HE1) & "nh V" & ChrW(&H129) & "nh"
    c.Add "Nam L" & ChrW(&HFD)
    c.Add "Nam Minh"
    c.Add "Nam Nha Trang"
    c.Add "Nam Ninh"
    c.Add "Nam Ninh H" & ChrW(&HF2) & "a"
    c.Add "Nam Ph" & ChrW(&HF9)
    c.Add "Nam Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add "Nam Quang"
    c.Add "Nam S" & ChrW(&HE1) & "ch"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart65", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart66(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Nam S" & ChrW(&H1A1) & "n"
    c.Add "Nam S" & ChrW(&H1EA7) & "m S" & ChrW(&H1A1) & "n"
    c.Add "Nam Thanh Mi" & ChrW(&H1EC7) & "n"
    c.Add "Nam Th" & ChrW(&HE0) & "nh"
    c.Add "Nam Th" & ChrW(&HE1) & "i Ninh"
    c.Add "Nam Th" & ChrW(&H1EE5) & "y Anh"
    c.Add "Nam Ti" & ChrW(&HEA) & "n H" & ChrW(&H1B0) & "ng"
    c.Add "Nam Ti" & ChrW(&H1EC1) & "n H" & ChrW(&H1EA3) & "i"
    c.Add "Nam Tri" & ChrW(&H1EC7) & "u"
    c.Add "Nam Tr" & ChrW(&HE0) & " My"
    c.Add "Nam Tr" & ChrW(&H1EA1) & "ch"
    c.Add "Nam Tr" & ChrW(&H1EF1) & "c"
    c.Add "Nam Tu" & ChrW(&H1EA5) & "n"
    c.Add "Nam Xang"
    c.Add "Nam Xu" & ChrW(&HE2) & "n"
    c.Add "Nam " & ChrW(&H110) & ChrW(&HE0)
    c.Add "Nam " & ChrW(&H110) & ChrW(&HE0) & "n"
    c.Add "Nam " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "Nam " & ChrW(&H110) & ChrW(&HF4) & "ng H" & ChrW(&HE0)
    c.Add "Nam " & ChrW(&H110) & ChrW(&HF4) & "ng H" & ChrW(&H1B0) & "ng"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart66", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart67(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Nam " & ChrW(&H110) & ChrW(&H1ECB) & "nh"
    c.Add "Nam " & ChrW(&H110) & ChrW(&H1ED3) & " S" & ChrW(&H1A1) & "n"
    c.Add "Nam " & ChrW(&H110) & ChrW(&H1ED3) & "ng"
    c.Add "Nga An"
    c.Add "Nga My"
    c.Add "Nga S" & ChrW(&H1A1) & "n"
    c.Add "Nga Th" & ChrW(&H1EAF) & "ng"
    c.Add "Nghi D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Nghi L" & ChrW(&H1ED9) & "c"
    c.Add "Nghi S" & ChrW(&H1A1) & "n"
    c.Add "Nghi Xu" & ChrW(&HE2) & "n"
    c.Add "Nghinh T" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "Nghi" & ChrW(&HEA) & "n Loan"
    c.Add "Ngh" & ChrW(&H129) & "a D" & ChrW(&HE2) & "n"
    c.Add "Ngh" & ChrW(&H129) & "a Giang"
    c.Add "Ngh" & ChrW(&H129) & "a H" & ChrW(&HE0) & "nh"
    c.Add "Ngh" & ChrW(&H129) & "a H" & ChrW(&H1B0) & "ng"
    c.Add "Ngh" & ChrW(&H129) & "a Kh" & ChrW(&HE1) & "nh"
    c.Add "Ngh" & ChrW(&H129) & "a L" & ChrW(&HE2) & "m"
    c.Add "Ngh" & ChrW(&H129) & "a L" & ChrW(&H1ED9)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart67", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart68(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ngh" & ChrW(&H129) & "a L" & ChrW(&H1ED9) & "c"
    c.Add "Ngh" & ChrW(&H129) & "a Mai"
    c.Add "Ngh" & ChrW(&H129) & "a Ph" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Ngh" & ChrW(&H129) & "a S" & ChrW(&H1A1) & "n"
    c.Add "Ngh" & ChrW(&H129) & "a Thu" & ChrW(&H1EAD) & "n"
    c.Add "Ngh" & ChrW(&H129) & "a Th" & ChrW(&HE0) & "nh"
    c.Add "Ngh" & ChrW(&H129) & "a Th" & ChrW(&H1ECD)
    c.Add "Ngh" & ChrW(&H129) & "a Trung"
    c.Add "Ngh" & ChrW(&H129) & "a Tr" & ChrW(&H1EE5)
    c.Add "Ngh" & ChrW(&H129) & "a T" & ChrW(&HE1)
    c.Add "Ngh" & ChrW(&H129) & "a T" & ChrW(&HE2) & "m"
    c.Add "Ngh" & ChrW(&H129) & "a " & ChrW(&H110) & ChrW(&HE0) & "n"
    c.Add "Ngh" & ChrW(&H129) & "a " & ChrW(&H110) & ChrW(&HF4)
    c.Add "Ngh" & ChrW(&H129) & "a " & ChrW(&H110) & ChrW(&H1ED3) & "ng"
    c.Add "Ngh" & ChrW(&H1ECB) & " " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    c.Add "Nguy" & ChrW(&HEA) & "n B" & ChrW(&HEC) & "nh"
    c.Add "Nguy" & ChrW(&HEA) & "n Gi" & ChrW(&HE1) & "p"
    c.Add "Nguy" & ChrW(&H1EC5) & "n B" & ChrW(&H1EC9) & "nh Khi" & ChrW(&HEA) & "m"
    c.Add "Nguy" & ChrW(&H1EC5) & "n Du"
    c.Add "Nguy" & ChrW(&H1EC5) & "n Hu" & ChrW(&H1EC7)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart68", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart69(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Nguy" & ChrW(&H1EC5) & "n L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng B" & ChrW(&H1EB1) & "ng"
    c.Add "Nguy" & ChrW(&H1EC5) & "n Nghi" & ChrW(&HEA) & "m"
    c.Add "Nguy" & ChrW(&H1EC5) & "n Ph" & ChrW(&HED) & "ch"
    c.Add "Nguy" & ChrW(&H1EC5) & "n Tr" & ChrW(&HE3) & "i"
    c.Add "Nguy" & ChrW(&H1EC5) & "n U" & ChrW(&HFD)
    c.Add "Nguy" & ChrW(&H1EC5) & "n Vi" & ChrW(&H1EC7) & "t Kh" & ChrW(&HE1) & "i"
    c.Add "Nguy" & ChrW(&H1EC5) & "n V" & ChrW(&H103) & "n Linh"
    c.Add "Nguy" & ChrW(&H1EC5) & "n " & ChrW(&H110) & ChrW(&H1EA1) & "i N" & ChrW(&H103) & "ng"
    c.Add "Nguy" & ChrW(&H1EC7) & "t H" & ChrW(&HF3) & "a"
    c.Add "Nguy" & ChrW(&H1EC7) & "t Vi" & ChrW(&HEA) & "n"
    c.Add "Nguy" & ChrW(&H1EC7) & "t " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    c.Add "Nguy" & ChrW(&H1EC7) & "t " & ChrW(&H1EA4) & "n"
    c.Add "Ng" & ChrW(&HE2) & "n S" & ChrW(&H1A1) & "n"
    c.Add "Ng" & ChrW(&HE3) & " B" & ChrW(&H1EA3) & "y"
    c.Add "Ng" & ChrW(&HE3) & " N" & ChrW(&H103) & "m"
    c.Add "Ng" & ChrW(&HE3) & "i Giao"
    c.Add "Ng" & ChrW(&HE3) & "i T" & ChrW(&H1EE9)
    c.Add "Ng" & ChrW(&HF4) & " M" & ChrW(&HE2) & "y"
    c.Add "Ng" & ChrW(&HF4) & " Quy" & ChrW(&H1EC1) & "n"
    c.Add "Ng" & ChrW(&H169) & " Ch" & ChrW(&H1EC9) & " S" & ChrW(&H1A1) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart69", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart70(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ng" & ChrW(&H169) & " Hi" & ChrW(&H1EC7) & "p"
    c.Add "Ng" & ChrW(&H169) & " H" & ChrW(&HE0) & "nh S" & ChrW(&H1A1) & "n"
    c.Add "Ng" & ChrW(&H169) & " L" & ChrW(&H1EA1) & "c"
    c.Add "Ng" & ChrW(&H1ECD) & "c Chi" & ChrW(&H1EBF) & "n"
    c.Add "Ng" & ChrW(&H1ECD) & "c Ch" & ChrW(&HFA) & "c"
    c.Add "Ng" & ChrW(&H1ECD) & "c H" & ChrW(&HE0)
    c.Add "Ng" & ChrW(&H1ECD) & "c H" & ChrW(&H1ED3) & "i"
    c.Add "Ng" & ChrW(&H1ECD) & "c Linh"
    c.Add "Ng" & ChrW(&H1ECD) & "c Li" & ChrW(&HEA) & "n"
    c.Add "Ng" & ChrW(&H1ECD) & "c Long"
    c.Add "Ng" & ChrW(&H1ECD) & "c L" & ChrW(&HE2) & "m"
    c.Add "Ng" & ChrW(&H1ECD) & "c L" & ChrW(&H1EB7) & "c"
    c.Add "Ng" & ChrW(&H1ECD) & "c S" & ChrW(&H1A1) & "n"
    c.Add "Ng" & ChrW(&H1ECD) & "c Thi" & ChrW(&H1EC7) & "n"
    c.Add "Ng" & ChrW(&H1ECD) & "c Tr" & ChrW(&H1EA1) & "o"
    c.Add "Ng" & ChrW(&H1ECD) & "c T" & ChrW(&H1ED1)
    c.Add "Ng" & ChrW(&H1ECD) & "c " & ChrW(&H110) & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "Ng" & ChrW(&H1ECD) & "k Bay"
    c.Add "Ng" & ChrW(&H1ECD) & "k R" & ChrW(&HE9) & "o"
    c.Add "Ng" & ChrW(&H1ECD) & "k T" & ChrW(&H1EE5)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart70", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart71(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ng" & ChrW(&H1EF1) & " Thi" & ChrW(&HEA) & "n"
    c.Add "Nha B" & ChrW(&HED) & "ch"
    c.Add "Nha Trang"
    c.Add "Nhi S" & ChrW(&H1A1) & "n"
    c.Add "Nhi" & ChrW(&HEA) & "u L" & ChrW(&H1ED9) & "c"
    c.Add "Nho Quan"
    c.Add "Nhu Gia"
    c.Add "Nhu" & ChrW(&H1EAD) & "n Ph" & ChrW(&HFA) & " T" & ChrW(&HE2) & "n"
    c.Add "Nhu" & ChrW(&H1EAD) & "n " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    c.Add "Nh" & ChrW(&HE0) & " B" & ChrW(&HE8)
    c.Add "Nh" & ChrW(&HE2) & "n C" & ChrW(&H1A1)
    c.Add "Nh" & ChrW(&HE2) & "n H" & ChrW(&HE0)
    c.Add "Nh" & ChrW(&HE2) & "n H" & ChrW(&HF2) & "a"
    c.Add "Nh" & ChrW(&HE2) & "n L" & ChrW(&HFD)
    c.Add "Nh" & ChrW(&HE2) & "n Ngh" & ChrW(&H129) & "a"
    c.Add "Nh" & ChrW(&HE2) & "n Th" & ChrW(&H1EAF) & "ng"
    c.Add "Nh" & ChrW(&HE3) & " Nam"
    c.Add "Nh" & ChrW(&HF4) & "n Mai"
    c.Add "Nh" & ChrW(&H1A1) & "n Ch" & ChrW(&HE2) & "u"
    c.Add "Nh" & ChrW(&H1A1) & "n H" & ChrW(&HF2) & "a L" & ChrW(&H1EAD) & "p"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart71", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart72(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Nh" & ChrW(&H1A1) & "n H" & ChrW(&H1ED9) & "i"
    c.Add "Nh" & ChrW(&H1A1) & "n M" & ChrW(&H1EF9)
    c.Add "Nh" & ChrW(&H1A1) & "n Ninh"
    c.Add "Nh" & ChrW(&H1A1) & "n Ph" & ChrW(&HFA)
    c.Add "Nh" & ChrW(&H1A1) & "n Tr" & ChrW(&H1EA1) & "ch"
    c.Add "Nh" & ChrW(&H1A1) & "n " & ChrW(&HC1) & "i"
    c.Add "Nh" & ChrW(&H1B0) & " Qu" & ChrW(&H1EF3) & "nh"
    c.Add "Nh" & ChrW(&H1B0) & " Thanh"
    c.Add "Nh" & ChrW(&H1B0) & " Xu" & ChrW(&HE2) & "n"
    c.Add "Nh" & ChrW(&H1EA5) & "t H" & ChrW(&HF2) & "a"
    c.Add "Nh" & ChrW(&H1ECB) & " Chi" & ChrW(&H1EC3) & "u"
    c.Add "Nh" & ChrW(&H1ECB) & " Long"
    c.Add "Nh" & ChrW(&H1ECB) & " Qu" & ChrW(&HFD)
    c.Add "Nh" & ChrW(&H1ECB) & " Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "Nh" & ChrW(&H1EEF) & " Kh" & ChrW(&HEA)
    c.Add "Nh" & ChrW(&H1EF1) & "t T" & ChrW(&H1EA3) & "o"
    c.Add "Ninh Ch" & ChrW(&HE2) & "u"
    c.Add "Ninh Ch" & ChrW(&H1EED)
    c.Add "Ninh C" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "Ninh Gia"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart72", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart73(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ninh Giang"
    c.Add "Ninh H" & ChrW(&HF2) & "a"
    c.Add "Ninh H" & ChrW(&H1EA3) & "i"
    c.Add "Ninh Ki" & ChrW(&H1EC1) & "u"
    c.Add "Ninh Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add "Ninh Qu" & ChrW(&H1EDB) & "i"
    c.Add "Ninh S" & ChrW(&H1A1) & "n"
    c.Add "Ninh Th" & ChrW(&H1EA1) & "nh"
    c.Add "Ninh Th" & ChrW(&H1EA1) & "nh L" & ChrW(&H1EE3) & "i"
    c.Add "Ninh X" & ChrW(&HE1)
    c.Add "Ninh " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "n"
    c.Add "Ni" & ChrW(&HEA) & "m S" & ChrW(&H1A1) & "n"
    c.Add "N" & ChrW(&HE0) & " B" & ChrW(&H1EE7) & "ng"
    c.Add "N" & ChrW(&HE0) & " Hang"
    c.Add "N" & ChrW(&HE0) & " H" & ChrW(&H1EF3)
    c.Add "N" & ChrW(&HE0) & " Ph" & ChrW(&H1EB7) & "c"
    c.Add "N" & ChrW(&HE0) & " T" & ChrW(&H1EA5) & "u"
    c.Add "N" & ChrW(&HE2) & "m Nung"
    c.Add "N" & ChrW(&HF4) & "ng C" & ChrW(&H1ED1) & "ng"
    c.Add "N" & ChrW(&HF4) & "ng S" & ChrW(&H1A1) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart73", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart74(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "N" & ChrW(&HF4) & "ng Ti" & ChrW(&H1EBF) & "n"
    c.Add "N" & ChrW(&HF4) & "ng Trang"
    c.Add "N" & ChrW(&HF9) & "ng Tr" & ChrW(&HED) & " Cao"
    c.Add "N" & ChrW(&HFA) & "a Ngam"
    c.Add "N" & ChrW(&HFA) & "i C" & ChrW(&H1EA5) & "m"
    c.Add "N" & ChrW(&HFA) & "i Th" & ChrW(&HE0) & "nh"
    c.Add "N" & ChrW(&H103) & "m C" & ChrW(&H103) & "n"
    c.Add "N" & ChrW(&H1EA5) & "m D" & ChrW(&H1EA9) & "n"
    c.Add "N" & ChrW(&H1EAD) & "m Ch" & ChrW(&HE0) & "y"
    c.Add "N" & ChrW(&H1EAD) & "m Cu" & ChrW(&H1ED5) & "i"
    c.Add "N" & ChrW(&H1EAD) & "m C" & ChrW(&HF3)
    c.Add "N" & ChrW(&H1EAD) & "m C" & ChrW(&H1EAF) & "n"
    c.Add "N" & ChrW(&H1EAD) & "m D" & ChrW(&H1ECB) & "ch"
    c.Add "N" & ChrW(&H1EAD) & "m H" & ChrW(&HE0) & "ng"
    c.Add "N" & ChrW(&H1EAD) & "m K" & ChrW(&HE8)
    c.Add "N" & ChrW(&H1EAD) & "m L" & ChrW(&H1EA7) & "u"
    c.Add "N" & ChrW(&H1EAD) & "m M" & ChrW(&H1EA1)
    c.Add "N" & ChrW(&H1EAD) & "m N" & ChrW(&HE8) & "n"
    c.Add "N" & ChrW(&H1EAD) & "m S" & ChrW(&H1ECF)
    c.Add "N" & ChrW(&H1EAD) & "m Ty"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart74", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart75(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "N" & ChrW(&H1EAD) & "m T" & ChrW(&H103) & "m"
    c.Add "N" & ChrW(&H1EAD) & "m X" & ChrW(&HE9)
    c.Add "N" & ChrW(&H1EAD) & "t S" & ChrW(&H1A1) & "n"
    c.Add "N" & ChrW(&H1EBF) & "nh"
    c.Add "N" & ChrW(&H1ED9) & "i B" & ChrW(&HE0) & "i"
    c.Add "Pa Ham"
    c.Add "Pa T" & ChrW(&H1EA7) & "n"
    c.Add "Pa " & ChrW(&H1EE6)
    c.Add "Pha Long"
    c.Add "Phan Ng" & ChrW(&H1ECD) & "c Hi" & ChrW(&H1EC3) & "n"
    c.Add "Phan Rang"
    c.Add "Phan R" & ChrW(&HED) & " C" & ChrW(&H1EED) & "a"
    c.Add "Phan S" & ChrW(&H1A1) & "n"
    c.Add "Phan Thanh"
    c.Add "Phan Thi" & ChrW(&H1EBF) & "t"
    c.Add "Phan " & ChrW(&H110) & ChrW(&HEC) & "nh Ph" & ChrW(&HF9) & "ng"
    c.Add "Phi" & ChrW(&HEA) & "ng C" & ChrW(&H1EB1) & "m"
    c.Add "Phi" & ChrW(&HEA) & "ng Kho" & ChrW(&HE0) & "i"
    c.Add "Phi" & ChrW(&HEA) & "ng P" & ChrW(&H1EB1) & "n"
    c.Add "Phong Ch" & ChrW(&HE2) & "u"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart75", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart76(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Phong C" & ChrW(&H1ED1) & "c"
    c.Add "Phong Dinh"
    c.Add "Phong Doanh"
    c.Add "Phong D" & ChrW(&H1EE5) & " H" & ChrW(&H1EA1)
    c.Add "Phong D" & ChrW(&H1EE5) & " Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng"
    c.Add "Phong Hi" & ChrW(&H1EC7) & "p"
    c.Add "Phong H" & ChrW(&HF2) & "a"
    c.Add "Phong H" & ChrW(&H1EA3) & "i"
    c.Add "Phong M" & ChrW(&H1EF9)
    c.Add "Phong Nha"
    c.Add "Phong N" & ChrW(&H1EAB) & "m"
    c.Add "Phong Ph" & ChrW(&HFA)
    c.Add "Phong Quang"
    c.Add "Phong Qu" & ChrW(&H1EA3) & "ng"
    c.Add "Phong Th" & ChrW(&HE1) & "i"
    c.Add "Phong Th" & ChrW(&H1EA1) & "nh"
    c.Add "Phong Th" & ChrW(&H1ED5)
    c.Add "Phong " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "n"
    c.Add "Ph" & ChrW(&HE1) & "t Di" & ChrW(&H1EC7) & "m"
    c.Add "Ph" & ChrW(&HEC) & "nh Gi" & ChrW(&HE0) & "ng"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart76", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart77(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ph" & ChrW(&HEC) & "nh H" & ChrW(&H1ED3)
    c.Add "Ph" & ChrW(&HF9) & " C" & ChrW(&HE1) & "t"
    c.Add "Ph" & ChrW(&HF9) & " Kh" & ChrW(&HEA)
    c.Add "Ph" & ChrW(&HF9) & " Li" & ChrW(&H1EC5) & "n"
    c.Add "Ph" & ChrW(&HF9) & " L" & ChrW(&HE3) & "ng"
    c.Add "Ph" & ChrW(&HF9) & " L" & ChrW(&H1B0) & "u"
    c.Add "Ph" & ChrW(&HF9) & " M" & ChrW(&H1EF9)
    c.Add "Ph" & ChrW(&HF9) & " M" & ChrW(&H1EF9) & " B" & ChrW(&H1EAF) & "c"
    c.Add "Ph" & ChrW(&HF9) & " M" & ChrW(&H1EF9) & " Nam"
    c.Add "Ph" & ChrW(&HF9) & " M" & ChrW(&H1EF9) & " T" & ChrW(&HE2) & "y"
    c.Add "Ph" & ChrW(&HF9) & " M" & ChrW(&H1EF9) & " " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "Ph" & ChrW(&HF9) & " Ninh"
    c.Add "Ph" & ChrW(&HF9) & " V" & ChrW(&HE2) & "n"
    c.Add "Ph" & ChrW(&HF9) & " Y" & ChrW(&HEA) & "n"
    c.Add "Ph" & ChrW(&HF9) & " " & ChrW(&H110) & ChrW(&H1ED5) & "ng"
    c.Add "Ph" & ChrW(&HF9) & "ng Nguy" & ChrW(&HEA) & "n"
    c.Add "Ph" & ChrW(&HFA) & " An"
    c.Add "Ph" & ChrW(&HFA) & " B" & ChrW(&HE0) & "i"
    c.Add "Ph" & ChrW(&HFA) & " B" & ChrW(&HEC) & "nh"
    c.Add "Ph" & ChrW(&HFA) & " C" & ChrW(&HE1) & "t"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart77", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart78(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ph" & ChrW(&HFA) & " C" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "Ph" & ChrW(&HFA) & " Di" & ChrW(&H1EC5) & "n"
    c.Add "Ph" & ChrW(&HFA) & " Gi" & ChrW(&HE1) & "o"
    c.Add "Ph" & ChrW(&HFA) & " H" & ChrW(&HF2) & "a"
    c.Add "Ph" & ChrW(&HFA) & " H" & ChrW(&HF2) & "a 1"
    c.Add "Ph" & ChrW(&HFA) & " H" & ChrW(&HF2) & "a 2"
    c.Add "Ph" & ChrW(&HFA) & " H" & ChrW(&HF2) & "a " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "Ph" & ChrW(&HFA) & " H" & ChrW(&H1ED3)
    c.Add "Ph" & ChrW(&HFA) & " H" & ChrW(&H1EEF) & "u"
    c.Add "Ph" & ChrW(&HFA) & " H" & ChrW(&H1EF1) & "u"
    c.Add "Ph" & ChrW(&HFA) & " Kh" & ChrW(&HEA)
    c.Add "Ph" & ChrW(&HFA) & " Kh" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Ph" & ChrW(&HFA) & " Linh"
    c.Add "Ph" & ChrW(&HFA) & " Long"
    c.Add "Ph" & ChrW(&HFA) & " L" & ChrW(&HE2) & "m"
    c.Add "Ph" & ChrW(&HFA) & " L" & ChrW(&HFD)
    c.Add "Ph" & ChrW(&HFA) & " L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Ph" & ChrW(&HFA) & " L" & ChrW(&H1EA1) & "c"
    c.Add "Ph" & ChrW(&HFA) & " L" & ChrW(&H1EC7)
    c.Add "Ph" & ChrW(&HFA) & " L" & ChrW(&H1ED9) & "c"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart78", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart79(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ph" & ChrW(&HFA) & " L" & ChrW(&H1EE3) & "i"
    c.Add "Ph" & ChrW(&HFA) & " M" & ChrW(&H1EE1)
    c.Add "Ph" & ChrW(&HFA) & " M" & ChrW(&H1EF9)
    c.Add "Ph" & ChrW(&HFA) & " Ngh" & ChrW(&H129) & "a"
    c.Add "Ph" & ChrW(&HFA) & " Nhu" & ChrW(&H1EAD) & "n"
    c.Add "Ph" & ChrW(&HFA) & " Ninh"
    c.Add "Ph" & ChrW(&HFA) & " Ph" & ChrW(&H1EE5) & "ng"
    c.Add "Ph" & ChrW(&HFA) & " Qu" & ChrW(&HFD)
    c.Add "Ph" & ChrW(&HFA) & " Qu" & ChrW(&H1ED1) & "c"
    c.Add "Ph" & ChrW(&HFA) & " Qu" & ChrW(&H1EDB) & "i"
    c.Add "Ph" & ChrW(&HFA) & " Ri" & ChrW(&H1EC1) & "ng"
    c.Add "Ph" & ChrW(&HFA) & " S" & ChrW(&H1A1) & "n"
    c.Add "Ph" & ChrW(&HFA) & " S" & ChrW(&H1A1) & "n L" & ChrW(&HE2) & "m H" & ChrW(&HE0)
    c.Add "Ph" & ChrW(&HFA) & " Thi" & ChrW(&H1EC7) & "n"
    c.Add "Ph" & ChrW(&HFA) & " Thu" & ChrW(&H1EAD) & "n"
    c.Add "Ph" & ChrW(&HFA) & " Thu" & ChrW(&H1EF7)
    c.Add "Ph" & ChrW(&HFA) & " Th" & ChrW(&HE0) & "nh"
    c.Add "Ph" & ChrW(&HFA) & " Th" & ChrW(&HE1) & "i"
    c.Add "Ph" & ChrW(&HFA) & " Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng"
    c.Add "Ph" & ChrW(&HFA) & " Th" & ChrW(&H1EA1) & "nh"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart79", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart80(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ph" & ChrW(&HFA) & " Th" & ChrW(&H1ECB) & "nh"
    c.Add "Ph" & ChrW(&HFA) & " Th" & ChrW(&H1ECD)
    c.Add "Ph" & ChrW(&HFA) & " Th" & ChrW(&H1ECD) & " H" & ChrW(&HF2) & "a"
    c.Add "Ph" & ChrW(&HFA) & " Trung"
    c.Add "Ph" & ChrW(&HFA) & " Tr" & ChrW(&H1EA1) & "ch"
    c.Add "Ph" & ChrW(&HFA) & " T" & ChrW(&HE2) & "m"
    c.Add "Ph" & ChrW(&HFA) & " T" & ChrW(&HE2) & "n"
    c.Add "Ph" & ChrW(&HFA) & " T" & ChrW(&HFA) & "c"
    c.Add "Ph" & ChrW(&HFA) & " Vang"
    c.Add "Ph" & ChrW(&HFA) & " Vinh"
    c.Add "Ph" & ChrW(&HFA) & " Xuy" & ChrW(&HEA) & "n"
    c.Add "Ph" & ChrW(&HFA) & " Xu" & ChrW(&HE2) & "n"
    c.Add "Ph" & ChrW(&HFA) & " Y" & ChrW(&HEA) & "n"
    c.Add "Ph" & ChrW(&HFA) & " " & ChrW(&H110) & ChrW(&HEC) & "nh"
    c.Add "Ph" & ChrW(&HFA) & " " & ChrW(&H110) & ChrW(&H1ECB) & "nh"
    c.Add "Ph" & ChrW(&HFA) & "c H" & ChrW(&HF2) & "a"
    c.Add "Ph" & ChrW(&HFA) & "c Kh" & ChrW(&HE1) & "nh"
    c.Add "Ph" & ChrW(&HFA) & "c L" & ChrW(&H1ED9) & "c"
    c.Add "Ph" & ChrW(&HFA) & "c L" & ChrW(&H1EE3) & "i"
    c.Add "Ph" & ChrW(&HFA) & "c S" & ChrW(&H1A1) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart80", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart81(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ph" & ChrW(&HFA) & "c Thu" & ChrW(&H1EAD) & "n"
    c.Add "Ph" & ChrW(&HFA) & "c Th" & ChrW(&H1ECB) & "nh"
    c.Add "Ph" & ChrW(&HFA) & "c Th" & ChrW(&H1ECD)
    c.Add "Ph" & ChrW(&HFA) & "c Th" & ChrW(&H1ECD) & " L" & ChrW(&HE2) & "m H" & ChrW(&HE0)
    c.Add "Ph" & ChrW(&HFA) & "c Tr" & ChrW(&H1EA1) & "ch"
    c.Add "Ph" & ChrW(&HFA) & "c Y" & ChrW(&HEA) & "n"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1A1) & "ng B" & ChrW(&HEC) & "nh"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1A1) & "ng Li" & ChrW(&H1EC5) & "u"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1A1) & "ng Li" & ChrW(&H1EC7) & "t"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1A1) & "ng Th" & ChrW(&H1ECB) & "nh"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c An"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c B" & ChrW(&HEC) & "nh"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c Ch" & ChrW(&HE1) & "nh"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c Ch" & ChrW(&H1EC9)
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c Dinh"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c Giang"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c Hi" & ChrW(&H1EC7) & "p"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c H" & ChrW(&HE0)
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c H" & ChrW(&HF2) & "a"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c H" & ChrW(&H1EA3) & "i"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart81", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart82(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c H" & ChrW(&H1EAD) & "u"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c H" & ChrW(&H1ED9) & "i"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c H" & ChrW(&H1EEF) & "u"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c Long"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c L" & ChrW(&HFD)
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c M" & ChrW(&H1EF9) & " Trung"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c N" & ChrW(&H103) & "ng"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c S" & ChrW(&H1A1) & "n"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c Th" & ChrW(&HE0) & "nh"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c Th" & ChrW(&HE1) & "i"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c Th" & ChrW(&H1EA1) & "nh"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c Th" & ChrW(&H1EAF) & "ng"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c Th" & ChrW(&H1EDB) & "i"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c Tr" & ChrW(&HE0)
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c T" & ChrW(&HE2) & "n"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c Vinh"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c V" & ChrW(&H129) & "nh T" & ChrW(&HE2) & "y"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng D" & ChrW(&H1EF1) & "c"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng S" & ChrW(&H1A1) & "n"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng Ti" & ChrW(&H1EBF) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart82", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart83(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ph" & ChrW(&H1EA1) & "m Ng" & ChrW(&H169) & " L" & ChrW(&HE3) & "o"
    c.Add "Ph" & ChrW(&H1EA1) & "m S" & ChrW(&H1B0) & " M" & ChrW(&H1EA1) & "nh"
    c.Add "Ph" & ChrW(&H1EAD) & "t T" & ChrW(&HED) & "ch"
    c.Add "Ph" & ChrW(&H1ED1) & " B" & ChrW(&H1EA3) & "ng"
    c.Add "Ph" & ChrW(&H1ED1) & " Hi" & ChrW(&H1EBF) & "n"
    c.Add "Ph" & ChrW(&H1ED5) & " Y" & ChrW(&HEA) & "n"
    c.Add "Ph" & ChrW(&H1EE5) & " D" & ChrW(&H1EF1) & "c"
    c.Add "Ph" & ChrW(&H1EE5) & "c H" & ChrW(&HF2) & "a"
    c.Add "Ph" & ChrW(&H1EE5) & "ng C" & ChrW(&HF4) & "ng"
    c.Add "Ph" & ChrW(&H1EE5) & "ng Hi" & ChrW(&H1EC7) & "p"
    c.Add "Ph" & ChrW(&H1EE7) & " L" & ChrW(&HFD)
    c.Add "Ph" & ChrW(&H1EE7) & " Th" & ChrW(&HF4) & "ng"
    c.Add "Pleiku"
    c.Add "Pu Nhi"
    c.Add "Pu Sam C" & ChrW(&HE1) & "p"
    c.Add "P" & ChrW(&HE0) & " C" & ChrW(&HF2)
    c.Add "P" & ChrW(&HE0) & " V" & ChrW(&H1EA7) & "y S" & ChrW(&H1EE7)
    c.Add "P" & ChrW(&HF9) & " Lu" & ChrW(&HF4) & "ng"
    c.Add "P" & ChrW(&HF9) & " Nhi"
    c.Add "P" & ChrW(&HFA) & " Nhung"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart83", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart84(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "P" & ChrW(&HFA) & "ng B" & ChrW(&HE1) & "nh"
    c.Add "P" & ChrW(&HFA) & "ng Lu" & ChrW(&HF4) & "ng"
    c.Add "P" & ChrW(&H1A1) & "ng Drang"
    c.Add "P" & ChrW(&H1EAF) & "c Ng" & ChrW(&HE0)
    c.Add "P" & ChrW(&H1EAF) & "c Ta"
    c.Add "P" & ChrW(&H1EDD) & " Ly Ng" & ChrW(&HE0) & "i"
    c.Add "P" & ChrW(&H1EDD) & " T" & ChrW(&HF3)
    c.Add "Quan S" & ChrW(&H1A1) & "n"
    c.Add "Quan Th" & ChrW(&HE0) & "nh"
    c.Add "Quan Tri" & ChrW(&H1EC1) & "u"
    c.Add "Quang B" & ChrW(&HEC) & "nh"
    c.Add "Quang Chi" & ChrW(&H1EC3) & "u"
    c.Add "Quang Hanh"
    c.Add "Quang H" & ChrW(&HE1) & "n"
    c.Add "Quang H" & ChrW(&H1B0) & "ng"
    c.Add "Quang Long"
    c.Add "Quang L" & ChrW(&H1ECB) & "ch"
    c.Add "Quang Minh"
    c.Add "Quang S" & ChrW(&H1A1) & "n"
    c.Add "Quang Thi" & ChrW(&H1EC7) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart84", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart85(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Quang Trung"
    c.Add "Quang " & ChrW(&H110) & ChrW(&H1ED3) & "ng"
    c.Add "Quy M" & ChrW(&HF4) & "ng"
    c.Add "Quy Nh" & ChrW(&H1A1) & "n"
    c.Add "Quy Nh" & ChrW(&H1A1) & "n B" & ChrW(&H1EAF) & "c"
    c.Add "Quy Nh" & ChrW(&H1A1) & "n Nam"
    c.Add "Quy Nh" & ChrW(&H1A1) & "n T" & ChrW(&HE2) & "y"
    c.Add "Quy Nh" & ChrW(&H1A1) & "n " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "Quy " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    c.Add "Quy" & ChrW(&H1EBF) & "t Th" & ChrW(&H1EAF) & "ng"
    c.Add "Qu" & ChrW(&HE0) & "i T" & ChrW(&H1EDF)
    c.Add "Qu" & ChrW(&HE1) & "ch Ph" & ChrW(&H1EA9) & "m"
    c.Add "Qu" & ChrW(&HE2) & "n Chu"
    c.Add "Qu" & ChrW(&HFD) & " H" & ChrW(&HF2) & "a"
    c.Add "Qu" & ChrW(&HFD) & " L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Qu" & ChrW(&HFD) & " L" & ChrW(&H1ED9) & "c"
    c.Add "Qu" & ChrW(&H1EA3) & "n B" & ChrW(&H1EA1)
    c.Add "Qu" & ChrW(&H1EA3) & "ng B" & ChrW(&HEC) & "nh"
    c.Add "Qu" & ChrW(&H1EA3) & "ng B" & ChrW(&H1EA1) & "ch"
    c.Add "Qu" & ChrW(&H1EA3) & "ng B" & ChrW(&H1ECB)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart85", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart86(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Qu" & ChrW(&H1EA3) & "ng Ch" & ChrW(&HE2) & "u"
    c.Add "Qu" & ChrW(&H1EA3) & "ng Ch" & ChrW(&HED) & "nh"
    c.Add "Qu" & ChrW(&H1EA3) & "ng H" & ChrW(&HE0)
    c.Add "Qu" & ChrW(&H1EA3) & "ng H" & ChrW(&HF2) & "a"
    c.Add "Qu" & ChrW(&H1EA3) & "ng Kh" & ChrW(&HEA)
    c.Add "Qu" & ChrW(&H1EA3) & "ng La"
    c.Add "Qu" & ChrW(&H1EA3) & "ng L" & ChrW(&HE2) & "m"
    c.Add "Qu" & ChrW(&H1EA3) & "ng L" & ChrW(&H1EAD) & "p"
    c.Add "Qu" & ChrW(&H1EA3) & "ng Nguy" & ChrW(&HEA) & "n"
    c.Add "Qu" & ChrW(&H1EA3) & "ng Ng" & ChrW(&H1ECD) & "c"
    c.Add "Qu" & ChrW(&H1EA3) & "ng Ninh"
    c.Add "Qu" & ChrW(&H1EA3) & "ng Oai"
    c.Add "Qu" & ChrW(&H1EA3) & "ng Ph" & ChrW(&HFA)
    c.Add "Qu" & ChrW(&H1EA3) & "ng S" & ChrW(&H1A1) & "n"
    c.Add "Qu" & ChrW(&H1EA3) & "ng Tr" & ChrW(&H1EA1) & "ch"
    c.Add "Qu" & ChrW(&H1EA3) & "ng Tr" & ChrW(&H1ECB)
    c.Add "Qu" & ChrW(&H1EA3) & "ng Tr" & ChrW(&H1EF1) & "c"
    c.Add "Qu" & ChrW(&H1EA3) & "ng T" & ChrW(&HE2) & "n"
    c.Add "Qu" & ChrW(&H1EA3) & "ng T" & ChrW(&HED) & "n"
    c.Add "Qu" & ChrW(&H1EA3) & "ng Uy" & ChrW(&HEA) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart86", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart87(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Qu" & ChrW(&H1EA3) & "ng Y" & ChrW(&HEA) & "n"
    c.Add "Qu" & ChrW(&H1EA3) & "ng " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "n"
    c.Add "Qu" & ChrW(&H1EA3) & "ng " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    c.Add "Qu" & ChrW(&H1EBF) & " Phong"
    c.Add "Qu" & ChrW(&H1EBF) & " Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add "Qu" & ChrW(&H1EBF) & " S" & ChrW(&H1A1) & "n"
    c.Add "Qu" & ChrW(&H1EBF) & " S" & ChrW(&H1A1) & "n Trung"
    c.Add "Qu" & ChrW(&H1EBF) & " V" & ChrW(&HF5)
    c.Add "Qu" & ChrW(&H1ED1) & "c Kh" & ChrW(&HE1) & "nh"
    c.Add "Qu" & ChrW(&H1ED1) & "c Oai"
    c.Add "Qu" & ChrW(&H1ED1) & "c Vi" & ChrW(&H1EC7) & "t"
    c.Add "Qu" & ChrW(&H1EDB) & "i An"
    c.Add "Qu" & ChrW(&H1EDB) & "i Thi" & ChrW(&H1EC7) & "n"
    c.Add "Qu" & ChrW(&H1EDB) & "i " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "n"
    c.Add "Qu" & ChrW(&H1EF3) & " Ch" & ChrW(&HE2) & "u"
    c.Add "Qu" & ChrW(&H1EF3) & " H" & ChrW(&H1EE3) & "p"
    c.Add "Qu" & ChrW(&H1EF3) & "nh An"
    c.Add "Qu" & ChrW(&H1EF3) & "nh Anh"
    c.Add "Qu" & ChrW(&H1EF3) & "nh L" & ChrW(&H1B0) & "u"
    c.Add "Qu" & ChrW(&H1EF3) & "nh Mai"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart87", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart88(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Qu" & ChrW(&H1EF3) & "nh Nhai"
    c.Add "Qu" & ChrW(&H1EF3) & "nh Ph" & ChrW(&HFA)
    c.Add "Qu" & ChrW(&H1EF3) & "nh Ph" & ChrW(&H1EE5)
    c.Add "Qu" & ChrW(&H1EF3) & "nh S" & ChrW(&H1A1) & "n"
    c.Add "Qu" & ChrW(&H1EF3) & "nh Tam"
    c.Add "Qu" & ChrW(&H1EF3) & "nh Th" & ChrW(&H1EAF) & "ng"
    c.Add "Qu" & ChrW(&H1EF3) & "nh V" & ChrW(&H103) & "n"
    c.Add "Qu" & ChrW(&H1EF9) & " Nh" & ChrW(&H1EA5) & "t"
    c.Add "R" & ChrW(&H1EA1) & "ch D" & ChrW(&H1EEB) & "a"
    c.Add "R" & ChrW(&H1EA1) & "ch Gi" & ChrW(&HE1)
    c.Add "R" & ChrW(&H1EA1) & "ch Ki" & ChrW(&H1EBF) & "n"
    c.Add "R" & ChrW(&H1EA1) & "ng " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "R" & ChrW(&H1EDD) & " K" & ChrW(&H1A1) & "i"
    c.Add "SR" & ChrW(&HF3)
    c.Add "Sa B" & ChrW(&HEC) & "nh"
    c.Add "Sa Hu" & ChrW(&H1EF3) & "nh"
    c.Add "Sa Loong"
    c.Add "Sa L" & ChrW(&HFD)
    c.Add "Sa Pa"
    c.Add "Sa Th" & ChrW(&H1EA7) & "y"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart88", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart89(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Sa " & ChrW(&H110) & ChrW(&HE9) & "c"
    c.Add "Sam M" & ChrW(&H1EE9) & "n"
    c.Add "Sao V" & ChrW(&HE0) & "ng"
    c.Add "Sen Ng" & ChrW(&H1B0)
    c.Add "Si Ma Cai"
    c.Add "Si Pa Ph" & ChrW(&HEC) & "n"
    c.Add "Sin Su" & ChrW(&H1ED1) & "i H" & ChrW(&H1ED3)
    c.Add "Song Kh" & ChrW(&H1EE7) & "a"
    c.Add "Song Li" & ChrW(&H1EC5) & "u"
    c.Add "Song L" & ChrW(&H1ED9) & "c"
    c.Add "Song Ph" & ChrW(&HFA)
    c.Add "Su" & ChrW(&H1ED1) & "i D" & ChrW(&H1EA7) & "u"
    c.Add "Su" & ChrW(&H1ED1) & "i Hai"
    c.Add "Su" & ChrW(&H1ED1) & "i Hi" & ChrW(&H1EC7) & "p"
    c.Add "Su" & ChrW(&H1ED1) & "i Ki" & ChrW(&H1EBF) & "t"
    c.Add "Su" & ChrW(&H1ED1) & "i Trai"
    c.Add "Su" & ChrW(&H1ED1) & "i T" & ChrW(&H1ECD)
    c.Add "S" & ChrW(&HE0) & " Ph" & ChrW(&HEC) & "n"
    c.Add "S" & ChrW(&HE0) & "i G" & ChrW(&HF2) & "n"
    c.Add "S" & ChrW(&HE1) & "ng Nh" & ChrW(&HE8)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart89", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart90(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "S" & ChrW(&HEC) & " L" & ChrW(&H1EDF) & " L" & ChrW(&H1EA7) & "u"
    c.Add "S" & ChrW(&HEC) & "n H" & ChrW(&H1ED3)
    c.Add "S" & ChrW(&HED) & "n Ch" & ChrW(&HE9) & "ng"
    c.Add "S" & ChrW(&HED) & "n Ch" & ChrW(&H1EA3) & "i"
    c.Add "S" & ChrW(&HED) & "n Th" & ChrW(&H1EA7) & "u"
    c.Add "S" & ChrW(&HED) & "nh Ph" & ChrW(&HEC) & "nh"
    c.Add "S" & ChrW(&HF3) & "c S" & ChrW(&H1A1) & "n"
    c.Add "S" & ChrW(&HF3) & "c Tr" & ChrW(&H103) & "ng"
    c.Add "S" & ChrW(&HF4) & "ng C" & ChrW(&HF4) & "ng"
    c.Add "S" & ChrW(&HF4) & "ng C" & ChrW(&H1EA7) & "u"
    c.Add "S" & ChrW(&HF4) & "ng Hinh"
    c.Add "S" & ChrW(&HF4) & "ng K" & ChrW(&HF4) & "n"
    c.Add "S" & ChrW(&HF4) & "ng L" & ChrW(&HF4)
    c.Add "S" & ChrW(&HF4) & "ng L" & ChrW(&H169) & "y"
    c.Add "S" & ChrW(&HF4) & "ng M" & ChrW(&HE3)
    c.Add "S" & ChrW(&HF4) & "ng Ray"
    c.Add "S" & ChrW(&HF4) & "ng Tr" & ChrW(&HED)
    c.Add "S" & ChrW(&HF4) & "ng V" & ChrW(&HE0) & "ng"
    c.Add "S" & ChrW(&HF4) & "ng " & ChrW(&H110) & ChrW(&H1ED1) & "c"
    c.Add "S" & ChrW(&H1A1) & "n C" & ChrW(&H1EA9) & "m H" & ChrW(&HE0)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart90", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart91(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "S" & ChrW(&H1A1) & "n D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "S" & ChrW(&H1A1) & "n Giang"
    c.Add "S" & ChrW(&H1A1) & "n H" & ChrW(&HE0)
    c.Add "S" & ChrW(&H1A1) & "n H" & ChrW(&HF2) & "a"
    c.Add "S" & ChrW(&H1A1) & "n H" & ChrW(&H1EA1)
    c.Add "S" & ChrW(&H1A1) & "n H" & ChrW(&H1EA3) & "i"
    c.Add "S" & ChrW(&H1A1) & "n H" & ChrW(&H1ED3) & "ng"
    c.Add "S" & ChrW(&H1A1) & "n Kim 1"
    c.Add "S" & ChrW(&H1A1) & "n Kim 2"
    c.Add "S" & ChrW(&H1A1) & "n Ki" & ChrW(&HEA) & "n"
    c.Add "S" & ChrW(&H1A1) & "n K" & ChrW(&H1EF3)
    c.Add "S" & ChrW(&H1A1) & "n Lang"
    c.Add "S" & ChrW(&H1A1) & "n Linh"
    c.Add "S" & ChrW(&H1A1) & "n L" & ChrW(&HE2) & "m"
    c.Add "S" & ChrW(&H1A1) & "n L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "S" & ChrW(&H1A1) & "n L" & ChrW(&H1ED9)
    c.Add "S" & ChrW(&H1A1) & "n Mai"
    c.Add "S" & ChrW(&H1A1) & "n M" & ChrW(&H1EF9)
    c.Add "S" & ChrW(&H1A1) & "n Nam"
    c.Add "S" & ChrW(&H1A1) & "n Qui"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart91", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart92(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "S" & ChrW(&H1A1) & "n Thu" & ChrW(&H1EF7)
    c.Add "S" & ChrW(&H1A1) & "n Th" & ChrW(&HE0) & "nh"
    c.Add "S" & ChrW(&H1A1) & "n Th" & ChrW(&H1EE7) & "y"
    c.Add "S" & ChrW(&H1A1) & "n Ti" & ChrW(&H1EBF) & "n"
    c.Add "S" & ChrW(&H1A1) & "n Tr" & ChrW(&HE0)
    c.Add "S" & ChrW(&H1A1) & "n T" & ChrW(&HE2) & "y"
    c.Add "S" & ChrW(&H1A1) & "n T" & ChrW(&HE2) & "y H" & ChrW(&H1EA1)
    c.Add "S" & ChrW(&H1A1) & "n T" & ChrW(&HE2) & "y Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng"
    c.Add "S" & ChrW(&H1A1) & "n T" & ChrW(&H1ECB) & "nh"
    c.Add "S" & ChrW(&H1A1) & "n V" & ChrW(&H129)
    c.Add "S" & ChrW(&H1A1) & "n " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "n"
    c.Add "S" & ChrW(&H1A1) & "n " & ChrW(&H110) & "i" & ChrW(&H1EC7) & "n"
    c.Add "S" & ChrW(&H1A1) & "n " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "S" & ChrW(&H1A1) & "n " & ChrW(&H110) & ChrW(&H1ED3) & "ng"
    c.Add "S" & ChrW(&H1A1) & "n " & ChrW(&H110) & ChrW(&H1ED9) & "ng"
    c.Add "S" & ChrW(&H1EA3) & "ng M" & ChrW(&H1ED9) & "c"
    c.Add "S" & ChrW(&H1EA7) & "m S" & ChrW(&H1A1) & "n"
    c.Add "S" & ChrW(&H1ED1) & "p C" & ChrW(&H1ED9) & "p"
    c.Add "S" & ChrW(&H1EE7) & "ng M" & ChrW(&HE1) & "ng"
    c.Add "Tam Anh"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart92", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart93(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Tam B" & ChrW(&HEC) & "nh"
    c.Add "Tam Chung"
    c.Add "Tam Ch" & ChrW(&HFA) & "c"
    c.Add "Tam D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Tam D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng B" & ChrW(&H1EAF) & "c"
    c.Add "Tam Giang"
    c.Add "Tam Hi" & ChrW(&H1EC7) & "p"
    c.Add "Tam H" & ChrW(&H1B0) & "ng"
    c.Add "Tam H" & ChrW(&H1EA3) & "i"
    c.Add "Tam H" & ChrW(&H1ED3) & "ng"
    c.Add "Tam H" & ChrW(&H1EE3) & "p"
    c.Add "Tam Kim"
    c.Add "Tam K" & ChrW(&H1EF3)
    c.Add "Tam Long"
    c.Add "Tam L" & ChrW(&H1B0)
    c.Add "Tam M" & ChrW(&H1EF9)
    c.Add "Tam Ng" & ChrW(&HE3) & "i"
    c.Add "Tam N" & ChrW(&HF4) & "ng"
    c.Add "Tam Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add "Tam Quan"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart93", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart94(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Tam Quang"
    c.Add "Tam S" & ChrW(&H1A1) & "n"
    c.Add "Tam Thanh"
    c.Add "Tam Th" & ChrW(&HE1) & "i"
    c.Add "Tam Th" & ChrW(&H1EAF) & "ng"
    c.Add "Tam Ti" & ChrW(&H1EBF) & "n"
    c.Add "Tam Xu" & ChrW(&HE2) & "n"
    c.Add "Tam " & ChrW(&H110) & "a"
    c.Add "Tam " & ChrW(&H110) & "i" & ChrW(&H1EC7) & "p"
    c.Add "Tam " & ChrW(&H110) & ChrW(&H1EA3) & "o"
    c.Add "Tam " & ChrW(&H110) & ChrW(&H1ED3) & "ng"
    c.Add "Than Uy" & ChrW(&HEA) & "n"
    c.Add "Thanh An"
    c.Add "Thanh Ba"
    c.Add "Thanh B" & ChrW(&HEC) & "nh"
    c.Add "Thanh B" & ChrW(&H1ED3) & "ng"
    c.Add "Thanh H" & ChrW(&HE0)
    c.Add "Thanh H" & ChrW(&HF2) & "a"
    c.Add "Thanh H" & ChrW(&H1B0) & "ng"
    c.Add "Thanh Kh" & ChrW(&HEA)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart94", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart95(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Thanh K" & ChrW(&H1EF3)
    c.Add "Thanh Li" & ChrW(&HEA) & "m"
    c.Add "Thanh Li" & ChrW(&H1EC7) & "t"
    c.Add "Thanh Long"
    c.Add "Thanh L" & ChrW(&HE2) & "m"
    c.Add "Thanh Mai"
    c.Add "Thanh Mi" & ChrW(&H1EBF) & "u"
    c.Add "Thanh Mi" & ChrW(&H1EC7) & "n"
    c.Add "Thanh M" & ChrW(&H1EF9)
    c.Add "Thanh N" & ChrW(&H1B0) & "a"
    c.Add "Thanh Oai"
    c.Add "Thanh Phong"
    c.Add "Thanh Qu" & ChrW(&HE2) & "n"
    c.Add "Thanh S" & ChrW(&H1A1) & "n"
    c.Add "Thanh Th" & ChrW(&H1ECB) & "nh"
    c.Add "Thanh Th" & ChrW(&H1EE7) & "y"
    c.Add "Thanh Tr" & ChrW(&HEC)
    c.Add "Thanh T" & ChrW(&HF9) & "ng"
    c.Add "Thanh Xu" & ChrW(&HE2) & "n"
    c.Add "Thanh Y" & ChrW(&HEA) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart95", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart96(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Thanh " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "n"
    c.Add "Thanh " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    c.Add "Thi" & ChrW(&HEA) & "n C" & ChrW(&H1EA7) & "m"
    c.Add "Thi" & ChrW(&HEA) & "n H" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Thi" & ChrW(&HEA) & "n L" & ChrW(&H1ED9) & "c"
    c.Add "Thi" & ChrW(&HEA) & "n Nh" & ChrW(&H1EAB) & "n"
    c.Add "Thi" & ChrW(&HEA) & "n Ph" & ChrW(&H1EE7)
    c.Add "Thi" & ChrW(&HEA) & "n Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "Thi" & ChrW(&H1EBF) & "t " & ChrW(&H1ED0) & "ng"
    c.Add "Thi" & ChrW(&H1EC7) & "n H" & ChrW(&HF2) & "a"
    c.Add "Thi" & ChrW(&H1EC7) & "n H" & ChrW(&H1B0) & "ng"
    c.Add "Thi" & ChrW(&H1EC7) & "n Long"
    c.Add "Thi" & ChrW(&H1EC7) & "n Thu" & ChrW(&H1EAD) & "t"
    c.Add "Thi" & ChrW(&H1EC7) & "n T" & ChrW(&HE2) & "n"
    c.Add "Thi" & ChrW(&H1EC7) & "n T" & ChrW(&HED) & "n"
    c.Add "Thi" & ChrW(&H1EC7) & "u H" & ChrW(&HF3) & "a"
    c.Add "Thi" & ChrW(&H1EC7) & "u Quang"
    c.Add "Thi" & ChrW(&H1EC7) & "u Ti" & ChrW(&H1EBF) & "n"
    c.Add "Thi" & ChrW(&H1EC7) & "u To" & ChrW(&HE1) & "n"
    c.Add "Thi" & ChrW(&H1EC7) & "u Trung"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart96", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart97(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Tho" & ChrW(&H1EA1) & "i S" & ChrW(&H1A1) & "n"
    c.Add "Thu B" & ChrW(&H1ED3) & "n"
    c.Add "Thu C" & ChrW(&HFA) & "c"
    c.Add "Thu L" & ChrW(&H169) & "m"
    c.Add "Thung Nai"
    c.Add "Thu" & ChrW(&H1EA7) & "n Trung"
    c.Add "Thu" & ChrW(&H1EAD) & "n An"
    c.Add "Thu" & ChrW(&H1EAD) & "n B" & ChrW(&H1EAF) & "c"
    c.Add "Thu" & ChrW(&H1EAD) & "n Ch" & ChrW(&HE2) & "u"
    c.Add "Thu" & ChrW(&H1EAD) & "n Giao"
    c.Add "Thu" & ChrW(&H1EAD) & "n Ho" & ChrW(&HE0)
    c.Add "Thu" & ChrW(&H1EAD) & "n H" & ChrW(&HF2) & "a"
    c.Add "Thu" & ChrW(&H1EAD) & "n H" & ChrW(&HF3) & "a"
    c.Add "Thu" & ChrW(&H1EAD) & "n H" & ChrW(&H1B0) & "ng"
    c.Add "Thu" & ChrW(&H1EAD) & "n H" & ChrW(&H1EA1) & "nh"
    c.Add "Thu" & ChrW(&H1EAD) & "n L" & ChrW(&H1EE3) & "i"
    c.Add "Thu" & ChrW(&H1EAD) & "n M" & ChrW(&H1EF9)
    c.Add "Thu" & ChrW(&H1EAD) & "n Nam"
    c.Add "Thu" & ChrW(&H1EAD) & "n Th" & ChrW(&HE0) & "nh"
    c.Add "Th" & ChrW(&HE0) & "ng T" & ChrW(&HED) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart97", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart98(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Th" & ChrW(&HE0) & "nh B" & ChrW(&HEC) & "nh Th" & ChrW(&H1ECD)
    c.Add "Th" & ChrW(&HE0) & "nh C" & ChrW(&HF4) & "ng"
    c.Add "Th" & ChrW(&HE0) & "nh Nam"
    c.Add "Th" & ChrW(&HE0) & "nh Nh" & ChrW(&H1EA5) & "t"
    c.Add "Th" & ChrW(&HE0) & "nh Sen"
    c.Add "Th" & ChrW(&HE0) & "nh Th" & ChrW(&H1EDB) & "i"
    c.Add "Th" & ChrW(&HE0) & "nh Vinh"
    c.Add "Th" & ChrW(&HE0) & "nh " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "Th" & ChrW(&HE1) & "c B" & ChrW(&HE0)
    c.Add "Th" & ChrW(&HE1) & "i B" & ChrW(&HEC) & "nh"
    c.Add "Th" & ChrW(&HE1) & "i Ho" & ChrW(&HE0)
    c.Add "Th" & ChrW(&HE1) & "i H" & ChrW(&HF2) & "a"
    c.Add "Th" & ChrW(&HE1) & "i M" & ChrW(&H1EF9)
    c.Add "Th" & ChrW(&HE1) & "i Ninh"
    c.Add "Th" & ChrW(&HE1) & "i S" & ChrW(&H1A1) & "n"
    c.Add "Th" & ChrW(&HE1) & "i Th" & ChrW(&H1EE5) & "y"
    c.Add "Th" & ChrW(&HE1) & "i T" & ChrW(&HE2) & "n"
    c.Add "Th" & ChrW(&HE1) & "p M" & ChrW(&H1B0) & ChrW(&H1EDD) & "i"
    c.Add "Th" & ChrW(&HF4) & "ng Nguy" & ChrW(&HEA) & "n"
    c.Add "Th" & ChrW(&HF4) & "ng N" & ChrW(&HF4) & "ng"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart98", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart99(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Th" & ChrW(&HF4) & "ng Th" & ChrW(&H1EE5)
    c.Add "Th" & ChrW(&HF4) & "ng T" & ChrW(&HE2) & "y H" & ChrW(&H1ED9) & "i"
    c.Add "Th" & ChrW(&H103) & "ng An"
    c.Add "Th" & ChrW(&H103) & "ng B" & ChrW(&HEC) & "nh"
    c.Add "Th" & ChrW(&H103) & "ng Ph" & ChrW(&HFA)
    c.Add "Th" & ChrW(&H103) & "ng Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "Th" & ChrW(&H103) & "ng " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "n"
    c.Add "Th" & ChrW(&H1B0) & " L" & ChrW(&HE2) & "m"
    c.Add "Th" & ChrW(&H1B0) & " Tr" & ChrW(&HEC)
    c.Add "Th" & ChrW(&H1B0) & " V" & ChrW(&H169)
    c.Add "Th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng L" & ChrW(&H1EA1) & "c"
    c.Add "Th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add "Th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng T" & ChrW(&HE2) & "n"
    c.Add "Th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng T" & ChrW(&HED) & "n"
    c.Add "Th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Xu" & ChrW(&HE2) & "n"
    c.Add "Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng B" & ChrW(&H1EB1) & "ng La"
    c.Add "Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng C" & ChrW(&HE1) & "t"
    c.Add "Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng C" & ChrW(&H1ED1) & "c"
    c.Add "Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng H" & ChrW(&HE0)
    c.Add "Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng H" & ChrW(&H1ED3) & "ng"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart99", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart100(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng Long"
    c.Add "Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng L" & ChrW(&HE2) & "m"
    c.Add "Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng Minh"
    c.Add "Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng Ninh"
    c.Add "Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng N" & ChrW(&HF4) & "ng"
    c.Add "Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng Ph" & ChrW(&HFA) & "c"
    c.Add "Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng Quan"
    c.Add "Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng S" & ChrW(&H1A1) & "n"
    c.Add "Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng Tr" & ChrW(&H1EA1) & "ch"
    c.Add "Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    c.Add "Th" & ChrW(&H1EA1) & "ch An"
    c.Add "Th" & ChrW(&H1EA1) & "ch B" & ChrW(&HEC) & "nh"
    c.Add "Th" & ChrW(&H1EA1) & "ch H" & ChrW(&HE0)
    c.Add "Th" & ChrW(&H1EA1) & "ch Kh" & ChrW(&HEA)
    c.Add "Th" & ChrW(&H1EA1) & "ch Kh" & ChrW(&HF4) & "i"
    c.Add "Th" & ChrW(&H1EA1) & "ch L" & ChrW(&H1EA1) & "c"
    c.Add "Th" & ChrW(&H1EA1) & "ch L" & ChrW(&H1EAD) & "p"
    c.Add "Th" & ChrW(&H1EA1) & "ch Qu" & ChrW(&H1EA3) & "ng"
    c.Add "Th" & ChrW(&H1EA1) & "ch Th" & ChrW(&H1EA5) & "t"
    c.Add "Th" & ChrW(&H1EA1) & "ch Xu" & ChrW(&HE2) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart100", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart101(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Th" & ChrW(&H1EA1) & "nh An"
    c.Add "Th" & ChrW(&H1EA1) & "nh B" & ChrW(&HEC) & "nh"
    c.Add "Th" & ChrW(&H1EA1) & "nh H" & ChrW(&HF2) & "a"
    c.Add "Th" & ChrW(&H1EA1) & "nh H" & ChrW(&HF3) & "a"
    c.Add "Th" & ChrW(&H1EA1) & "nh H" & ChrW(&H1B0) & "ng"
    c.Add "Th" & ChrW(&H1EA1) & "nh H" & ChrW(&H1EA3) & "i"
    c.Add "Th" & ChrW(&H1EA1) & "nh L" & ChrW(&H1ED9) & "c"
    c.Add "Th" & ChrW(&H1EA1) & "nh L" & ChrW(&H1EE3) & "i"
    c.Add "Th" & ChrW(&H1EA1) & "nh M" & ChrW(&H1EF9)
    c.Add "Th" & ChrW(&H1EA1) & "nh M" & ChrW(&H1EF9) & " T" & ChrW(&HE2) & "y"
    c.Add "Th" & ChrW(&H1EA1) & "nh Phong"
    c.Add "Th" & ChrW(&H1EA1) & "nh Ph" & ChrW(&HFA)
    c.Add "Th" & ChrW(&H1EA1) & "nh Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add "Th" & ChrW(&H1EA1) & "nh Qu" & ChrW(&H1EDB) & "i"
    c.Add "Th" & ChrW(&H1EA1) & "nh Th" & ChrW(&H1EDB) & "i An"
    c.Add "Th" & ChrW(&H1EA1) & "nh Tr" & ChrW(&H1ECB)
    c.Add "Th" & ChrW(&H1EA1) & "nh Xu" & ChrW(&HE2) & "n"
    c.Add "Th" & ChrW(&H1EA1) & "nh " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "Th" & ChrW(&H1EA1) & "nh " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    c.Add "Th" & ChrW(&H1EA3) & "o Nguy" & ChrW(&HEA) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart101", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart102(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Th" & ChrW(&H1EA5) & "t Kh" & ChrW(&HEA)
    c.Add "Th" & ChrW(&H1EA7) & "n Kh" & ChrW(&HEA)
    c.Add "Th" & ChrW(&H1EA7) & "n L" & ChrW(&H129) & "nh"
    c.Add "Th" & ChrW(&H1EA7) & "n Sa"
    c.Add "Th" & ChrW(&H1EAF) & "ng L" & ChrW(&H1ED9) & "c"
    c.Add "Th" & ChrW(&H1EAF) & "ng L" & ChrW(&H1EE3) & "i"
    c.Add "Th" & ChrW(&H1EAF) & "ng M" & ChrW(&H1ED1)
    c.Add "Th" & ChrW(&H1ECB) & "nh Minh"
    c.Add "Th" & ChrW(&H1ECD) & " B" & ChrW(&HEC) & "nh"
    c.Add "Th" & ChrW(&H1ECD) & " Long"
    c.Add "Th" & ChrW(&H1ECD) & " L" & ChrW(&H1EAD) & "p"
    c.Add "Th" & ChrW(&H1ECD) & " Ng" & ChrW(&H1ECD) & "c"
    c.Add "Th" & ChrW(&H1ECD) & " Phong"
    c.Add "Th" & ChrW(&H1ECD) & " Ph" & ChrW(&HFA)
    c.Add "Th" & ChrW(&H1ECD) & " S" & ChrW(&H1A1) & "n"
    c.Add "Th" & ChrW(&H1ECD) & " V" & ChrW(&H103) & "n"
    c.Add "Th" & ChrW(&H1ECD) & " Xu" & ChrW(&HE2) & "n"
    c.Add "Th" & ChrW(&H1ED1) & "ng Nh" & ChrW(&H1EA5) & "t"
    c.Add "Th" & ChrW(&H1ED1) & "t N" & ChrW(&H1ED1) & "t"
    c.Add "Th" & ChrW(&H1ED5) & " Ch" & ChrW(&HE2) & "u"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart102", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart103(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Th" & ChrW(&H1ED5) & " Tang"
    c.Add "Th" & ChrW(&H1EDB) & "i An"
    c.Add "Th" & ChrW(&H1EDB) & "i An H" & ChrW(&H1ED9) & "i"
    c.Add "Th" & ChrW(&H1EDB) & "i An " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "Th" & ChrW(&H1EDB) & "i B" & ChrW(&HEC) & "nh"
    c.Add "Th" & ChrW(&H1EDB) & "i H" & ChrW(&HF2) & "a"
    c.Add "Th" & ChrW(&H1EDB) & "i H" & ChrW(&H1B0) & "ng"
    c.Add "Th" & ChrW(&H1EDB) & "i Lai"
    c.Add "Th" & ChrW(&H1EDB) & "i Long"
    c.Add "Th" & ChrW(&H1EDB) & "i S" & ChrW(&H1A1) & "n"
    c.Add "Th" & ChrW(&H1EDB) & "i Thu" & ChrW(&H1EAD) & "n"
    c.Add "Th" & ChrW(&H1EE5) & "c Ph" & ChrW(&HE1) & "n"
    c.Add "Th" & ChrW(&H1EE5) & "y Anh"
    c.Add "Th" & ChrW(&H1EE5) & "y H" & ChrW(&HF9) & "ng"
    c.Add "Th" & ChrW(&H1EE7) & " D" & ChrW(&H1EA7) & "u M" & ChrW(&H1ED9) & "t"
    c.Add "Th" & ChrW(&H1EE7) & " Th" & ChrW(&H1EEB) & "a"
    c.Add "Th" & ChrW(&H1EE7) & " " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    c.Add "Th" & ChrW(&H1EE7) & "y Nguy" & ChrW(&HEA) & "n"
    c.Add "Th" & ChrW(&H1EE7) & "y Xu" & ChrW(&HE2) & "n"
    c.Add "Ti" & ChrW(&HEA) & "n Du"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart103", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart104(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ti" & ChrW(&HEA) & "n Hoa"
    c.Add "Ti" & ChrW(&HEA) & "n H" & ChrW(&H1B0) & "ng"
    c.Add "Ti" & ChrW(&HEA) & "n H" & ChrW(&H1EA3) & "i"
    c.Add "Ti" & ChrW(&HEA) & "n La"
    c.Add "Ti" & ChrW(&HEA) & "n L" & ChrW(&HE3) & "ng"
    c.Add "Ti" & ChrW(&HEA) & "n L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Ti" & ChrW(&HEA) & "n L" & ChrW(&H1EE5) & "c"
    c.Add "Ti" & ChrW(&HEA) & "n L" & ChrW(&H1EEF)
    c.Add "Ti" & ChrW(&HEA) & "n Minh"
    c.Add "Ti" & ChrW(&HEA) & "n Nguy" & ChrW(&HEA) & "n"
    c.Add "Ti" & ChrW(&HEA) & "n Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add "Ti" & ChrW(&HEA) & "n S" & ChrW(&H1A1) & "n"
    c.Add "Ti" & ChrW(&HEA) & "n Th" & ChrW(&H1EE7) & "y"
    c.Add "Ti" & ChrW(&HEA) & "n Ti" & ChrW(&H1EBF) & "n"
    c.Add "Ti" & ChrW(&HEA) & "n Trang"
    c.Add "Ti" & ChrW(&HEA) & "n Y" & ChrW(&HEA) & "n"
    c.Add "Ti" & ChrW(&HEA) & "n " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "n"
    c.Add "Ti" & ChrW(&HEA) & "n " & ChrW(&H110) & ChrW(&H1ED3) & "ng"
    c.Add "Ti" & ChrW(&H1EBF) & "n Th" & ChrW(&HE0) & "nh"
    c.Add "Ti" & ChrW(&H1EBF) & "n Th" & ChrW(&H1EAF) & "ng"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart104", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart105(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Ti" & ChrW(&H1EC1) & "n H" & ChrW(&H1EA3) & "i"
    c.Add "Ti" & ChrW(&H1EC1) & "n Phong"
    c.Add "Ti" & ChrW(&H1EC3) & "u C" & ChrW(&H1EA7) & "n"
    c.Add "To" & ChrW(&HE0) & "n L" & ChrW(&H1B0) & "u"
    c.Add "To" & ChrW(&HE0) & "n Th" & ChrW(&H1EAF) & "ng"
    c.Add "Tri L" & ChrW(&H1EC5)
    c.Add "Tri Ph" & ChrW(&HFA)
    c.Add "Tri T" & ChrW(&HF4) & "n"
    c.Add "Tri" & ChrW(&H1EC7) & "u B" & ChrW(&HEC) & "nh"
    c.Add "Tri" & ChrW(&H1EC7) & "u C" & ChrW(&H1A1)
    c.Add "Tri" & ChrW(&H1EC7) & "u L" & ChrW(&H1ED9) & "c"
    c.Add "Tri" & ChrW(&H1EC7) & "u Phong"
    c.Add "Tri" & ChrW(&H1EC7) & "u S" & ChrW(&H1A1) & "n"
    c.Add "Tri" & ChrW(&H1EC7) & "u Vi" & ChrW(&H1EC7) & "t V" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Trung An"
    c.Add "Trung Ch" & ChrW(&HED) & "nh"
    c.Add "Trung Hi" & ChrW(&H1EC7) & "p"
    c.Add "Trung H" & ChrW(&HE0)
    c.Add "Trung H" & ChrW(&H1B0) & "ng"
    c.Add "Trung H" & ChrW(&H1EA1)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart105", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart106(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Trung H" & ChrW(&H1ED9) & "i"
    c.Add "Trung Kh" & ChrW(&HE1) & "nh V" & ChrW(&H129) & "nh"
    c.Add "Trung K" & ChrW(&HEA) & "nh"
    c.Add "Trung L" & ChrW(&HFD)
    c.Add "Trung L" & ChrW(&H1ED9) & "c"
    c.Add "Trung M" & ChrW(&H1EF9) & " T" & ChrW(&HE2) & "y"
    c.Add "Trung Ng" & ChrW(&HE3) & "i"
    c.Add "Trung Nh" & ChrW(&H1EE9) & "t"
    c.Add "Trung S" & ChrW(&H1A1) & "n"
    c.Add "Trung Thu" & ChrW(&H1EA7) & "n"
    c.Add "Trung Th" & ChrW(&HE0) & "nh"
    c.Add "Trung Th" & ChrW(&H1ECB) & "nh"
    c.Add "Trung T" & ChrW(&HE2) & "m"
    c.Add "Tru" & ChrW(&HF4) & "ng M" & ChrW(&HED) & "t"
    c.Add "Tr" & ChrW(&HE0) & " B" & ChrW(&H1ED3) & "ng"
    c.Add "Tr" & ChrW(&HE0) & " C" & ChrW(&HE2) & "u"
    c.Add "Tr" & ChrW(&HE0) & " C" & ChrW(&HF4) & "n"
    c.Add "Tr" & ChrW(&HE0) & " C" & ChrW(&HFA)
    c.Add "Tr" & ChrW(&HE0) & " Giang"
    c.Add "Tr" & ChrW(&HE0) & " Gi" & ChrW(&HE1) & "p"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart106", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart107(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Tr" & ChrW(&HE0) & " Leng"
    c.Add "Tr" & ChrW(&HE0) & " Linh"
    c.Add "Tr" & ChrW(&HE0) & " Li" & ChrW(&HEA) & "n"
    c.Add "Tr" & ChrW(&HE0) & " L" & ChrW(&HFD)
    c.Add "Tr" & ChrW(&HE0) & " L" & ChrW(&H129) & "nh"
    c.Add "Tr" & ChrW(&HE0) & " My"
    c.Add "Tr" & ChrW(&HE0) & " T" & ChrW(&HE2) & "n"
    c.Add "Tr" & ChrW(&HE0) & " T" & ChrW(&H1EAD) & "p"
    c.Add "Tr" & ChrW(&HE0) & " Vinh"
    c.Add "Tr" & ChrW(&HE0) & " Vong"
    c.Add "Tr" & ChrW(&HE0) & " V" & ChrW(&HE2) & "n"
    c.Add "Tr" & ChrW(&HE0) & " " & ChrW(&HD4) & "n"
    c.Add "Tr" & ChrW(&HE0) & " " & ChrW(&H110) & ChrW(&H1ED1) & "c"
    c.Add "Tr" & ChrW(&HE0) & "m Chim"
    c.Add "Tr" & ChrW(&HE0) & "ng X" & ChrW(&HE1)
    c.Add "Tr" & ChrW(&HE0) & "ng " & ChrW(&H110) & ChrW(&H1ECB) & "nh"
    c.Add "Tr" & ChrW(&HED) & " Ph" & ChrW(&H1EA3) & "i"
    c.Add "Tr" & ChrW(&HED) & " Qu" & ChrW(&H1EA3)
    c.Add "Tr" & ChrW(&HF9) & "ng Kh" & ChrW(&HE1) & "nh"
    c.Add "Tr" & ChrW(&HFA) & "c L" & ChrW(&HE2) & "m"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart107", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart108(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Tr" & ChrW(&H1B0) & ChrW(&H1A1) & "ng Quang Tr" & ChrW(&H1ECD) & "ng"
    c.Add "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Giang"
    c.Add "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng H" & ChrW(&HE0)
    c.Add "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Kh" & ChrW(&HE1) & "nh"
    c.Add "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Long"
    c.Add "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Long H" & ChrW(&HF2) & "a"
    c.Add "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Long T" & ChrW(&HE2) & "y"
    c.Add "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng L" & ChrW(&HE2) & "m"
    c.Add "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng L" & ChrW(&H1B0) & "u"
    c.Add "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Ninh"
    c.Add "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Ph" & ChrW(&HFA)
    c.Add "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Sa"
    c.Add "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Sinh"
    c.Add "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng S" & ChrW(&H1A1) & "n"
    c.Add "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Thi"
    c.Add "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Th" & ChrW(&HE0) & "nh"
    c.Add "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng T" & ChrW(&HE2) & "n"
    c.Add "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Vinh"
    c.Add "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng V" & ChrW(&H103) & "n"
    c.Add "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Xu" & ChrW(&HE2) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart108", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart109(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Tr" & ChrW(&H1EA1) & "i Cau"
    c.Add "Tr" & ChrW(&H1EA1) & "m L" & ChrW(&H1ED9)
    c.Add "Tr" & ChrW(&H1EA1) & "m Th" & ChrW(&H1EA3) & "n"
    c.Add "Tr" & ChrW(&H1EA1) & "m T" & ChrW(&H1EA5) & "u"
    c.Add "Tr" & ChrW(&H1EA3) & "ng Bom"
    c.Add "Tr" & ChrW(&H1EA3) & "ng B" & ChrW(&HE0) & "ng"
    c.Add "Tr" & ChrW(&H1EA3) & "ng D" & ChrW(&HE0) & "i"
    c.Add "Tr" & ChrW(&H1EA5) & "n Bi" & ChrW(&HEA) & "n"
    c.Add "Tr" & ChrW(&H1EA5) & "n Y" & ChrW(&HEA) & "n"
    c.Add "Tr" & ChrW(&H1EA7) & "n H" & ChrW(&H1B0) & "ng " & ChrW(&H110) & ChrW(&H1EA1) & "o"
    c.Add "Tr" & ChrW(&H1EA7) & "n Li" & ChrW(&H1EC5) & "u"
    c.Add "Tr" & ChrW(&H1EA7) & "n L" & ChrW(&HE3) & "m"
    c.Add "Tr" & ChrW(&H1EA7) & "n Nh" & ChrW(&HE2) & "n T" & ChrW(&HF4) & "ng"
    c.Add "Tr" & ChrW(&H1EA7) & "n Ph" & ChrW(&HE1) & "n"
    c.Add "Tr" & ChrW(&H1EA7) & "n Ph" & ChrW(&HFA)
    c.Add "Tr" & ChrW(&H1EA7) & "n Th" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Tr" & ChrW(&H1EA7) & "n V" & ChrW(&H103) & "n Th" & ChrW(&H1EDD) & "i"
    c.Add "Tr" & ChrW(&H1EA7) & "n " & ChrW(&H110) & ChrW(&H1EC1)
    c.Add "Tr" & ChrW(&H1ECB) & " An"
    c.Add "Tr" & ChrW(&H1ECB) & "nh T" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart109", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart110(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Tr" & ChrW(&H1EEB) & " V" & ChrW(&H103) & "n Th" & ChrW(&H1ED1)
    c.Add "Tr" & ChrW(&H1EF1) & "c Ninh"
    c.Add "Tu B" & ChrW(&HF4) & "ng"
    c.Add "Tu M" & ChrW(&H1A1) & " R" & ChrW(&HF4) & "ng"
    c.Add "Tu V" & ChrW(&H169)
    c.Add "Tuy An B" & ChrW(&H1EAF) & "c"
    c.Add "Tuy An Nam"
    c.Add "Tuy An T" & ChrW(&HE2) & "y"
    c.Add "Tuy An " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "Tuy H" & ChrW(&HF2) & "a"
    c.Add "Tuy Phong"
    c.Add "Tuy Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add "Tuy Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c B" & ChrW(&H1EAF) & "c"
    c.Add "Tuy Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c T" & ChrW(&HE2) & "y"
    c.Add "Tuy Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "Tuy " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    c.Add "Tuy" & ChrW(&HEA) & "n B" & ChrW(&HEC) & "nh"
    c.Add "Tuy" & ChrW(&HEA) & "n H" & ChrW(&HF3) & "a"
    c.Add "Tuy" & ChrW(&HEA) & "n L" & ChrW(&HE2) & "m"
    c.Add "Tuy" & ChrW(&HEA) & "n Ph" & ChrW(&HFA)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart110", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart111(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Tuy" & ChrW(&HEA) & "n Quang"
    c.Add "Tuy" & ChrW(&HEA) & "n S" & ChrW(&H1A1) & "n"
    c.Add "Tuy" & ChrW(&HEA) & "n Th" & ChrW(&H1EA1) & "nh"
    c.Add "Tu" & ChrW(&H1EA5) & "n S" & ChrW(&H1A1) & "n"
    c.Add "Tu" & ChrW(&H1EA5) & "n " & ChrW(&H110) & ChrW(&H1EA1) & "o"
    c.Add "Tu" & ChrW(&H1EA7) & "n Ch" & ChrW(&HE2) & "u"
    c.Add "Tu" & ChrW(&H1EA7) & "n Gi" & ChrW(&HE1) & "o"
    c.Add "Tu" & ChrW(&H1EC7) & " T" & ChrW(&H129) & "nh"
    c.Add "T" & ChrW(&HE0) & " Hine"
    c.Add "T" & ChrW(&HE0) & " H" & ChrW(&H1ED9) & "c"
    c.Add "T" & ChrW(&HE0) & " L" & ChrW(&HE0) & "i"
    c.Add "T" & ChrW(&HE0) & " N" & ChrW(&H103) & "ng"
    c.Add "T" & ChrW(&HE0) & " R" & ChrW(&H1EE5) & "t"
    c.Add "T" & ChrW(&HE0) & " T" & ChrW(&H1ED5) & "ng"
    c.Add "T" & ChrW(&HE0) & " Xi L" & ChrW(&HE1) & "ng"
    c.Add "T" & ChrW(&HE0) & " X" & ChrW(&HF9) & "a"
    c.Add "T" & ChrW(&HE0) & " " & ChrW(&H110) & ChrW(&HF9) & "ng"
    c.Add "T" & ChrW(&HE0) & "i V" & ChrW(&H103) & "n"
    c.Add "T" & ChrW(&HE1) & "nh Linh"
    c.Add "T" & ChrW(&HE1) & "t Ng" & ChrW(&HE0)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart111", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart112(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "T" & ChrW(&HE2) & "n An"
    c.Add "T" & ChrW(&HE2) & "n An H" & ChrW(&H1ED9) & "i"
    c.Add "T" & ChrW(&HE2) & "n Bi" & ChrW(&HEA) & "n"
    c.Add "T" & ChrW(&HE2) & "n B" & ChrW(&HEC) & "nh"
    c.Add "T" & ChrW(&HE2) & "n Chi"
    c.Add "T" & ChrW(&HE2) & "n Ch" & ChrW(&HE2) & "u"
    c.Add "T" & ChrW(&HE2) & "n C" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "T" & ChrW(&HE2) & "n D" & ChrW(&HE2) & "n"
    c.Add "T" & ChrW(&HE2) & "n D" & ChrW(&H129) & "nh"
    c.Add "T" & ChrW(&HE2) & "n D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "T" & ChrW(&HE2) & "n Giang"
    c.Add "T" & ChrW(&HE2) & "n Gianh"
    c.Add "T" & ChrW(&HE2) & "n Hi" & ChrW(&H1EC7) & "p"
    c.Add "T" & ChrW(&HE2) & "n H" & ChrW(&HE0) & " L" & ChrW(&HE2) & "m H" & ChrW(&HE0)
    c.Add "T" & ChrW(&HE2) & "n H" & ChrW(&HE0) & "o"
    c.Add "T" & ChrW(&HE2) & "n H" & ChrW(&HF2) & "a"
    c.Add "T" & ChrW(&HE2) & "n H" & ChrW(&H1B0) & "ng"
    c.Add "T" & ChrW(&HE2) & "n H" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "T" & ChrW(&HE2) & "n H" & ChrW(&H1EA1) & "nh"
    c.Add "T" & ChrW(&HE2) & "n H" & ChrW(&H1EA3) & "i"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart112", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart113(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "T" & ChrW(&HE2) & "n H" & ChrW(&H1ED3) & "ng"
    c.Add "T" & ChrW(&HE2) & "n H" & ChrW(&H1ED9) & " C" & ChrW(&H1A1)
    c.Add "T" & ChrW(&HE2) & "n H" & ChrW(&H1ED9) & "i"
    c.Add "T" & ChrW(&HE2) & "n H" & ChrW(&H1EE3) & "p"
    c.Add "T" & ChrW(&HE2) & "n Khai"
    c.Add "T" & ChrW(&HE2) & "n Kh" & ChrW(&HE1) & "nh"
    c.Add "T" & ChrW(&HE2) & "n Kh" & ChrW(&HE1) & "nh Trung"
    c.Add "T" & ChrW(&HE2) & "n K" & ChrW(&H1EF3)
    c.Add "T" & ChrW(&HE2) & "n Long"
    c.Add "T" & ChrW(&HE2) & "n Long H" & ChrW(&H1ED9) & "i"
    c.Add "T" & ChrW(&HE2) & "n L" & ChrW(&HE2) & "n"
    c.Add "T" & ChrW(&HE2) & "n L" & ChrW(&H129) & "nh"
    c.Add "T" & ChrW(&HE2) & "n L" & ChrW(&H1B0) & ChrW(&H1EE3) & "c"
    c.Add "T" & ChrW(&HE2) & "n L" & ChrW(&H1EA1) & "c"
    c.Add "T" & ChrW(&HE2) & "n L" & ChrW(&H1EAD) & "p"
    c.Add "T" & ChrW(&HE2) & "n L" & ChrW(&H1ED9) & "c"
    c.Add "T" & ChrW(&HE2) & "n L" & ChrW(&H1EE3) & "i"
    c.Add "T" & ChrW(&HE2) & "n Mai"
    c.Add "T" & ChrW(&HE2) & "n Minh"
    c.Add "T" & ChrW(&HE2) & "n M" & ChrW(&H1EF9)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart113", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart114(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "T" & ChrW(&HE2) & "n Ng" & ChrW(&HE3) & "i"
    c.Add "T" & ChrW(&HE2) & "n Nhu" & ChrW(&H1EAD) & "n " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "T" & ChrW(&HE2) & "n Nh" & ChrW(&H1EF1) & "t"
    c.Add "T" & ChrW(&HE2) & "n Ninh"
    c.Add "T" & ChrW(&HE2) & "n Pheo"
    c.Add "T" & ChrW(&HE2) & "n Phong"
    c.Add "T" & ChrW(&HE2) & "n Ph" & ChrW(&HFA)
    c.Add "T" & ChrW(&HE2) & "n Ph" & ChrW(&HFA) & " Trung"
    c.Add "T" & ChrW(&HE2) & "n Ph" & ChrW(&HFA) & " " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "T" & ChrW(&HE2) & "n Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add "T" & ChrW(&HE2) & "n Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c 1"
    c.Add "T" & ChrW(&HE2) & "n Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c 2"
    c.Add "T" & ChrW(&HE2) & "n Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c 3"
    c.Add "T" & ChrW(&HE2) & "n Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c H" & ChrW(&H1B0) & "ng"
    c.Add "T" & ChrW(&HE2) & "n Quan"
    c.Add "T" & ChrW(&HE2) & "n Quang"
    c.Add "T" & ChrW(&HE2) & "n Qu" & ChrW(&H1EDB) & "i"
    c.Add "T" & ChrW(&HE2) & "n S" & ChrW(&H1A1) & "n"
    c.Add "T" & ChrW(&HE2) & "n S" & ChrW(&H1A1) & "n H" & ChrW(&HF2) & "a"
    c.Add "T" & ChrW(&HE2) & "n S" & ChrW(&H1A1) & "n Nh" & ChrW(&HEC)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart114", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart115(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "T" & ChrW(&HE2) & "n S" & ChrW(&H1A1) & "n Nh" & ChrW(&H1EA5) & "t"
    c.Add "T" & ChrW(&HE2) & "n Thanh"
    c.Add "T" & ChrW(&HE2) & "n Thu" & ChrW(&H1EAD) & "n"
    c.Add "T" & ChrW(&HE2) & "n Thu" & ChrW(&H1EAD) & "n B" & ChrW(&HEC) & "nh"
    c.Add "T" & ChrW(&HE2) & "n Th" & ChrW(&HE0) & "nh"
    c.Add "T" & ChrW(&HE2) & "n Th" & ChrW(&HE0) & "nh B" & ChrW(&HEC) & "nh"
    c.Add "T" & ChrW(&HE2) & "n Th" & ChrW(&H1EA1) & "nh"
    c.Add "T" & ChrW(&HE2) & "n Th" & ChrW(&H1EDB) & "i"
    c.Add "T" & ChrW(&HE2) & "n Th" & ChrW(&H1EDB) & "i Hi" & ChrW(&H1EC7) & "p"
    c.Add "T" & ChrW(&HE2) & "n Th" & ChrW(&H1EE7) & "y"
    c.Add "T" & ChrW(&HE2) & "n Ti" & ChrW(&H1EBF) & "n"
    c.Add "T" & ChrW(&HE2) & "n Tri"
    c.Add "T" & ChrW(&HE2) & "n Tri" & ChrW(&H1EC1) & "u"
    c.Add "T" & ChrW(&HE2) & "n Tr" & ChrW(&HE0) & "o"
    c.Add "T" & ChrW(&HE2) & "n Tr" & ChrW(&H1ECB) & "nh"
    c.Add "T" & ChrW(&HE2) & "n Tr" & ChrW(&H1EE5)
    c.Add "T" & ChrW(&HE2) & "n T" & ChrW(&HE2) & "y"
    c.Add "T" & ChrW(&HE2) & "n T" & ChrW(&H1EA1) & "o"
    c.Add "T" & ChrW(&HE2) & "n T" & ChrW(&H1EAD) & "p"
    c.Add "T" & ChrW(&HE2) & "n Uy" & ChrW(&HEA) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart115", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart116(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "T" & ChrW(&HE2) & "n V" & ChrW(&H103) & "n"
    c.Add "T" & ChrW(&HE2) & "n V" & ChrW(&H129) & "nh L" & ChrW(&H1ED9) & "c"
    c.Add "T" & ChrW(&HE2) & "n Xu" & ChrW(&HE2) & "n"
    c.Add "T" & ChrW(&HE2) & "n Y" & ChrW(&HEA) & "n"
    c.Add "T" & ChrW(&HE2) & "n " & ChrW(&HC2) & "n"
    c.Add "T" & ChrW(&HE2) & "n " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "n"
    c.Add "T" & ChrW(&HE2) & "n " & ChrW(&H110) & "o" & ChrW(&HE0) & "n"
    c.Add "T" & ChrW(&HE2) & "n " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "T" & ChrW(&HE2) & "n " & ChrW(&H110) & ChrW(&HF4) & "ng Hi" & ChrW(&H1EC7) & "p"
    c.Add "T" & ChrW(&HE2) & "n " & ChrW(&H110) & ChrW(&H1ECB) & "nh"
    c.Add "T" & ChrW(&HE2) & "y C" & ChrW(&H1ED1) & "c"
    c.Add "T" & ChrW(&HE2) & "y Giang"
    c.Add "T" & ChrW(&HE2) & "y Hi" & ChrW(&H1EBF) & "u"
    c.Add "T" & ChrW(&HE2) & "y Hoa L" & ChrW(&H1B0)
    c.Add "T" & ChrW(&HE2) & "y H" & ChrW(&HF2) & "a"
    c.Add "T" & ChrW(&HE2) & "y H" & ChrW(&H1ED3)
    c.Add "T" & ChrW(&HE2) & "y Kh" & ChrW(&HE1) & "nh S" & ChrW(&H1A1) & "n"
    c.Add "T" & ChrW(&HE2) & "y Kh" & ChrW(&HE1) & "nh V" & ChrW(&H129) & "nh"
    c.Add "T" & ChrW(&HE2) & "y M" & ChrW(&H1ED7)
    c.Add "T" & ChrW(&HE2) & "y Nam"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart116", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart117(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "T" & ChrW(&HE2) & "y Nha Trang"
    c.Add "T" & ChrW(&HE2) & "y Ninh H" & ChrW(&HF2) & "a"
    c.Add "T" & ChrW(&HE2) & "y Ph" & ChrW(&HFA)
    c.Add "T" & ChrW(&HE2) & "y Ph" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "T" & ChrW(&HE2) & "y S" & ChrW(&H1A1) & "n"
    c.Add "T" & ChrW(&HE2) & "y Th" & ChrW(&HE1) & "i Ninh"
    c.Add "T" & ChrW(&HE2) & "y Th" & ChrW(&H1EA1) & "nh"
    c.Add "T" & ChrW(&HE2) & "y Th" & ChrW(&H1EE5) & "y Anh"
    c.Add "T" & ChrW(&HE2) & "y Ti" & ChrW(&H1EC1) & "n H" & ChrW(&H1EA3) & "i"
    c.Add "T" & ChrW(&HE2) & "y Tr" & ChrW(&HE0)
    c.Add "T" & ChrW(&HE2) & "y Tr" & ChrW(&HE0) & " B" & ChrW(&H1ED3) & "ng"
    c.Add "T" & ChrW(&HE2) & "y T" & ChrW(&H1EF1) & "u"
    c.Add "T" & ChrW(&HE2) & "y Y" & ChrW(&HEA) & "n"
    c.Add "T" & ChrW(&HE2) & "y Y" & ChrW(&HEA) & "n T" & ChrW(&H1EED)
    c.Add "T" & ChrW(&HE2) & "y " & ChrW(&H110) & ChrW(&HF4)
    c.Add "T" & ChrW(&HEC) & "a D" & ChrW(&HEC) & "nh"
    c.Add "T" & ChrW(&HED) & "ch L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "T" & ChrW(&HF4) & " Ch" & ChrW(&HE2) & "u"
    c.Add "T" & ChrW(&HF4) & " Hi" & ChrW(&H1EC7) & "u"
    c.Add "T" & ChrW(&HF4) & " M" & ChrW(&HFA) & "a"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart117", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart118(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "T" & ChrW(&HF9) & "ng B" & ChrW(&HE1)
    c.Add "T" & ChrW(&HF9) & "ng L" & ChrW(&H1ED9) & "c"
    c.Add "T" & ChrW(&HF9) & "ng Thi" & ChrW(&H1EC7) & "n"
    c.Add "T" & ChrW(&HF9) & "ng V" & ChrW(&HE0) & "i"
    c.Add "T" & ChrW(&HFA) & " L" & ChrW(&H1EC7)
    c.Add "T" & ChrW(&H103) & "ng Nh" & ChrW(&H1A1) & "n Ph" & ChrW(&HFA)
    c.Add "T" & ChrW(&H129) & "nh Gia"
    c.Add "T" & ChrW(&H129) & "nh T" & ChrW(&HFA) & "c"
    c.Add "T" & ChrW(&H1A1) & " Tung"
    c.Add "T" & ChrW(&H1B0) & " Ngh" & ChrW(&H129) & "a"
    c.Add "T" & ChrW(&H1B0) & ChrW(&H1A1) & "ng D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "T" & ChrW(&H1B0) & ChrW(&H1A1) & "ng Mai"
    c.Add "T" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng H" & ChrW(&H1EA1)
    c.Add "T" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng L" & ChrW(&H129) & "nh"
    c.Add "T" & ChrW(&H1EA1) & " An Kh" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "T" & ChrW(&H1EA1) & " Khoa"
    c.Add "T" & ChrW(&H1EA3) & " C" & ChrW(&H1EE7) & " T" & ChrW(&H1EF7)
    c.Add "T" & ChrW(&H1EA3) & " L" & ChrW(&HE8) & "ng"
    c.Add "T" & ChrW(&H1EA3) & " Ph" & ChrW(&HEC) & "n"
    c.Add "T" & ChrW(&H1EA3) & " Van"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart118", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart119(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "T" & ChrW(&H1EA7) & "m Vu"
    c.Add "T" & ChrW(&H1EAD) & "p Ng" & ChrW(&HE3) & "i"
    c.Add "T" & ChrW(&H1EAD) & "p S" & ChrW(&H1A1) & "n"
    c.Add "T" & ChrW(&H1EB1) & "ng Lo" & ChrW(&H1ECF) & "ng"
    c.Add "T" & ChrW(&H1EC1) & " L" & ChrW(&H1ED7)
    c.Add "T" & ChrW(&H1ECB) & "nh Bi" & ChrW(&HEA) & "n"
    c.Add "T" & ChrW(&H1ECB) & "nh Kh" & ChrW(&HEA)
    c.Add "T" & ChrW(&H1ED1) & "ng S" & ChrW(&H1A1) & "n"
    c.Add "T" & ChrW(&H1ED1) & "ng Tr" & ChrW(&HE2) & "n"
    c.Add "T" & ChrW(&H1ED5) & "ng C" & ChrW(&H1ECD) & "t"
    c.Add "T" & ChrW(&H1EE7) & "a Ch" & ChrW(&HF9) & "a"
    c.Add "T" & ChrW(&H1EE7) & "a S" & ChrW(&HED) & "n Ch" & ChrW(&H1EA3) & "i"
    c.Add "T" & ChrW(&H1EE7) & "a Th" & ChrW(&HE0) & "ng"
    c.Add "T" & ChrW(&H1EE9) & " K" & ChrW(&H1EF3)
    c.Add "T" & ChrW(&H1EE9) & " Minh"
    c.Add "T" & ChrW(&H1EE9) & " M" & ChrW(&H1EF9)
    c.Add "T" & ChrW(&H1EEB) & " Li" & ChrW(&HEA) & "m"
    c.Add "T" & ChrW(&H1EEB) & " S" & ChrW(&H1A1) & "n"
    c.Add "T" & ChrW(&H1EF1) & " L" & ChrW(&H1EA1) & "n"
    c.Add "U Minh"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart119", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart120(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "U Minh Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng"
    c.Add "Uar"
    c.Add "U" & ChrW(&HF4) & "ng B" & ChrW(&HED)
    c.Add "Vinh H" & ChrW(&H1B0) & "ng"
    c.Add "Vinh Kim"
    c.Add "Vinh L" & ChrW(&H1ED9) & "c"
    c.Add "Vinh Ph" & ChrW(&HFA)
    c.Add "Vinh Qu" & ChrW(&HFD)
    c.Add "Vi" & ChrW(&H1EC7) & "t An"
    c.Add "Vi" & ChrW(&H1EC7) & "t H" & ChrW(&HF2) & "a"
    c.Add "Vi" & ChrW(&H1EC7) & "t H" & ChrW(&H1B0) & "ng"
    c.Add "Vi" & ChrW(&H1EC7) & "t H" & ChrW(&H1ED3) & "ng"
    c.Add "Vi" & ChrW(&H1EC7) & "t Kh" & ChrW(&HEA)
    c.Add "Vi" & ChrW(&H1EC7) & "t L" & ChrW(&HE2) & "m"
    c.Add "Vi" & ChrW(&H1EC7) & "t Ti" & ChrW(&H1EBF) & "n"
    c.Add "Vi" & ChrW(&H1EC7) & "t Tr" & ChrW(&HEC)
    c.Add "Vi" & ChrW(&H1EC7) & "t Xuy" & ChrW(&HEA) & "n"
    c.Add "Vi" & ChrW(&H1EC7) & "t Y" & ChrW(&HEA) & "n"
    c.Add "Vu Gia"
    c.Add "V" & ChrW(&HE0) & "m C" & ChrW(&H1ECF)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart120", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart121(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "V" & ChrW(&HE0) & "ng Danh"
    c.Add "V" & ChrW(&HE2) & "n B" & ChrW(&HE1) & "n"
    c.Add "V" & ChrW(&HE2) & "n Canh"
    c.Add "V" & ChrW(&HE2) & "n Du"
    c.Add "V" & ChrW(&HE2) & "n H" & ChrW(&HE0)
    c.Add "V" & ChrW(&HE2) & "n H" & ChrW(&HF2) & "a"
    c.Add "V" & ChrW(&HE2) & "n H" & ChrW(&H1ED3)
    c.Add "V" & ChrW(&HE2) & "n Kh" & ChrW(&HE1) & "nh"
    c.Add "V" & ChrW(&HE2) & "n Nham"
    c.Add "V" & ChrW(&HE2) & "n Ph" & ChrW(&HFA)
    c.Add "V" & ChrW(&HE2) & "n S" & ChrW(&H1A1) & "n"
    c.Add "V" & ChrW(&HE2) & "n T" & ChrW(&H1EE5)
    c.Add "V" & ChrW(&HE2) & "n " & ChrW(&H110) & ChrW(&HEC) & "nh"
    c.Add "V" & ChrW(&HE2) & "n " & ChrW(&H110) & ChrW(&H1ED3) & "n"
    c.Add "V" & ChrW(&HF4) & " Tranh"
    c.Add "V" & ChrW(&HF5) & " C" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "V" & ChrW(&HF5) & " Lao"
    c.Add "V" & ChrW(&HF5) & " Mi" & ChrW(&H1EBF) & "u"
    c.Add "V" & ChrW(&HF5) & " Nhai"
    c.Add "V" & ChrW(&H103) & "n B" & ChrW(&HE0) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart121", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart122(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "V" & ChrW(&H103) & "n Ch" & ChrW(&H1EA5) & "n"
    c.Add "V" & ChrW(&H103) & "n Giang"
    c.Add "V" & ChrW(&H103) & "n Hi" & ChrW(&H1EBF) & "n"
    c.Add "V" & ChrW(&H103) & "n H" & ChrW(&HE1) & "n"
    c.Add "V" & ChrW(&H103) & "n Ki" & ChrW(&H1EC1) & "u"
    c.Add "V" & ChrW(&H103) & "n Lang"
    c.Add "V" & ChrW(&H103) & "n L" & ChrW(&HE3) & "ng"
    c.Add "V" & ChrW(&H103) & "n L" & ChrW(&H103) & "ng"
    c.Add "V" & ChrW(&H103) & "n Mi" & ChrW(&H1EBF) & "u"
    c.Add "V" & ChrW(&H103) & "n Mi" & ChrW(&H1EBF) & "u - Qu" & ChrW(&H1ED1) & "c T" & ChrW(&H1EED) & " Gi" & ChrW(&HE1) & "m"
    c.Add "V" & ChrW(&H103) & "n M" & ChrW(&HF4) & "n"
    c.Add "V" & ChrW(&H103) & "n Nho"
    c.Add "V" & ChrW(&H103) & "n Ph" & ChrW(&HFA)
    c.Add "V" & ChrW(&H103) & "n Quan"
    c.Add "V" & ChrW(&H129) & "nh Am"
    c.Add "V" & ChrW(&H129) & "nh An"
    c.Add "V" & ChrW(&H129) & "nh B" & ChrW(&HEC) & "nh"
    c.Add "V" & ChrW(&H129) & "nh B" & ChrW(&H1EA3) & "o"
    c.Add "V" & ChrW(&H129) & "nh Ch" & ChrW(&HE2) & "n"
    c.Add "V" & ChrW(&H129) & "nh Ch" & ChrW(&HE2) & "u"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart122", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart123(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "V" & ChrW(&H129) & "nh C" & ChrW(&HF4) & "ng"
    c.Add "V" & ChrW(&H129) & "nh Gia"
    c.Add "V" & ChrW(&H129) & "nh Hanh"
    c.Add "V" & ChrW(&H129) & "nh Ho" & ChrW(&HE0) & "ng"
    c.Add "V" & ChrW(&H129) & "nh H" & ChrW(&HF2) & "a"
    c.Add "V" & ChrW(&H129) & "nh H" & ChrW(&HF2) & "a H" & ChrW(&H1B0) & "ng"
    c.Add "V" & ChrW(&H129) & "nh H" & ChrW(&H1B0) & "ng"
    c.Add "V" & ChrW(&H129) & "nh H" & ChrW(&H1EA3) & "i"
    c.Add "V" & ChrW(&H129) & "nh H" & ChrW(&H1EA3) & "o"
    c.Add "V" & ChrW(&H129) & "nh H" & ChrW(&H1EAD) & "u"
    c.Add "V" & ChrW(&H129) & "nh H" & ChrW(&H1ED9) & "i"
    c.Add "V" & ChrW(&H129) & "nh H" & ChrW(&H1EF1) & "u"
    c.Add "V" & ChrW(&H129) & "nh Kim"
    c.Add "V" & ChrW(&H129) & "nh Linh"
    c.Add "V" & ChrW(&H129) & "nh L" & ChrW(&H1EA1) & "i"
    c.Add "V" & ChrW(&H129) & "nh L" & ChrW(&H1ED9) & "c"
    c.Add "V" & ChrW(&H129) & "nh L" & ChrW(&H1EE3) & "i"
    c.Add "V" & ChrW(&H129) & "nh M" & ChrW(&H1EF9)
    c.Add "V" & ChrW(&H129) & "nh Phong"
    c.Add "V" & ChrW(&H129) & "nh Ph" & ChrW(&HFA)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart123", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart124(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "V" & ChrW(&H129) & "nh Ph" & ChrW(&HFA) & "c"
    c.Add "V" & ChrW(&H129) & "nh Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add "V" & ChrW(&H129) & "nh Quang"
    c.Add "V" & ChrW(&H129) & "nh S" & ChrW(&H1A1) & "n"
    c.Add "V" & ChrW(&H129) & "nh Thanh"
    c.Add "V" & ChrW(&H129) & "nh Thu" & ChrW(&H1EAD) & "n"
    c.Add "V" & ChrW(&H129) & "nh Thu" & ChrW(&H1EAD) & "n " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "V" & ChrW(&H129) & "nh Th" & ChrW(&HE0) & "nh"
    c.Add "V" & ChrW(&H129) & "nh Th" & ChrW(&HF4) & "ng"
    c.Add "V" & ChrW(&H129) & "nh Th" & ChrW(&H1EA1) & "nh"
    c.Add "V" & ChrW(&H129) & "nh Th" & ChrW(&H1EA1) & "nh Trung"
    c.Add "V" & ChrW(&H129) & "nh Th" & ChrW(&H1ECB) & "nh"
    c.Add "V" & ChrW(&H129) & "nh Th" & ChrW(&H1EE7) & "y"
    c.Add "V" & ChrW(&H129) & "nh Th" & ChrW(&H1EF1) & "c"
    c.Add "V" & ChrW(&H129) & "nh Trinh"
    c.Add "V" & ChrW(&H129) & "nh Tr" & ChrW(&H1EA1) & "ch"
    c.Add "V" & ChrW(&H129) & "nh Tr" & ChrW(&H1EE5)
    c.Add "V" & ChrW(&H129) & "nh Tuy"
    c.Add "V" & ChrW(&H129) & "nh T" & ChrW(&HE2) & "n"
    c.Add "V" & ChrW(&H129) & "nh T" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart124", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart125(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "V" & ChrW(&H129) & "nh T" & ChrW(&H1EBF)
    c.Add "V" & ChrW(&H129) & "nh Vi" & ChrW(&H1EC5) & "n"
    c.Add "V" & ChrW(&H129) & "nh Xu" & ChrW(&HE2) & "n"
    c.Add "V" & ChrW(&H129) & "nh X" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "V" & ChrW(&H129) & "nh Y" & ChrW(&HEA) & "n"
    c.Add "V" & ChrW(&H129) & "nh " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "u"
    c.Add "V" & ChrW(&H129) & "nh " & ChrW(&H110) & ChrW(&H1ECB) & "nh"
    c.Add "V" & ChrW(&H169) & " D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "V" & ChrW(&H169) & " L" & ChrW(&H103) & "ng"
    c.Add "V" & ChrW(&H169) & " L" & ChrW(&H1EC5)
    c.Add "V" & ChrW(&H169) & " Ninh"
    c.Add "V" & ChrW(&H169) & " Ph" & ChrW(&HFA) & "c"
    c.Add "V" & ChrW(&H169) & " Quang"
    c.Add "V" & ChrW(&H169) & " Qu" & ChrW(&HFD)
    c.Add "V" & ChrW(&H169) & " Th" & ChrW(&H1B0)
    c.Add "V" & ChrW(&H169) & " Ti" & ChrW(&HEA) & "n"
    c.Add "V" & ChrW(&H169) & "ng T" & ChrW(&HE0) & "u"
    c.Add "V" & ChrW(&H169) & "ng " & ChrW(&HC1) & "ng"
    c.Add "V" & ChrW(&H1B0) & ChrW(&H1EDD) & "n L" & ChrW(&HE0) & "i"
    c.Add "V" & ChrW(&H1EA1) & "n An"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart125", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart126(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "V" & ChrW(&H1EA1) & "n H" & ChrW(&H1B0) & "ng"
    c.Add "V" & ChrW(&H1EA1) & "n Linh"
    c.Add "V" & ChrW(&H1EA1) & "n L" & ChrW(&H1ED9) & "c"
    c.Add "V" & ChrW(&H1EA1) & "n Ninh"
    c.Add "V" & ChrW(&H1EA1) & "n Ph" & ChrW(&HFA)
    c.Add "V" & ChrW(&H1EA1) & "n Th" & ChrW(&H1EAF) & "ng"
    c.Add "V" & ChrW(&H1EA1) & "n T" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "V" & ChrW(&H1EA1) & "n Xu" & ChrW(&HE2) & "n"
    c.Add "V" & ChrW(&H1EA1) & "n " & ChrW(&H110) & ChrW(&H1EE9) & "c"
    c.Add "V" & ChrW(&H1EAD) & "t L" & ChrW(&H1EA1) & "i"
    c.Add "V" & ChrW(&H1EC7) & " Giang"
    c.Add "V" & ChrW(&H1ECB) & " Kh" & ChrW(&HEA)
    c.Add "V" & ChrW(&H1ECB) & " Thanh"
    c.Add "V" & ChrW(&H1ECB) & " Thanh 1"
    c.Add "V" & ChrW(&H1ECB) & " Th" & ChrW(&H1EE7) & "y"
    c.Add "V" & ChrW(&H1ECB) & " T" & ChrW(&HE2) & "n"
    c.Add "V" & ChrW(&H1ECB) & " Xuy" & ChrW(&HEA) & "n"
    c.Add "V" & ChrW(&H1EE5) & " B" & ChrW(&H1EA3) & "n"
    c.Add "V" & ChrW(&H1EE5) & " B" & ChrW(&H1ED5) & "n"
    c.Add "V" & ChrW(&H1EF9) & " D" & ChrW(&H1EA1)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart126", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart127(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Xa Dung"
    c.Add "Xa" & ChrW(&H303) & " H" & ChrW(&H1B0) & "ng " & ChrW(&H110) & ChrW(&H1EA1) & "o"
    c.Add "Xa" & ChrW(&H303) & " Trung Gi" & ChrW(&HE3)
    c.Add "Xa" & ChrW(&H303) & " " & ChrW(&H110) & "an Ph" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng"
    c.Add "Xuy" & ChrW(&HEA) & "n M" & ChrW(&H1ED9) & "c"
    c.Add "Xu" & ChrW(&HE2) & "n An"
    c.Add "Xu" & ChrW(&HE2) & "n B" & ChrW(&HEC) & "nh"
    c.Add "Xu" & ChrW(&HE2) & "n B" & ChrW(&H1EAF) & "c"
    c.Add "Xu" & ChrW(&HE2) & "n Chinh"
    c.Add "Xu" & ChrW(&HE2) & "n C" & ChrW(&H1EA3) & "nh"
    c.Add "Xu" & ChrW(&HE2) & "n C" & ChrW(&H1EA9) & "m"
    c.Add "Xu" & ChrW(&HE2) & "n Du"
    c.Add "Xu" & ChrW(&HE2) & "n D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Xu" & ChrW(&HE2) & "n Giang"
    c.Add "Xu" & ChrW(&HE2) & "n H" & ChrW(&HF2) & "a"
    c.Add "Xu" & ChrW(&HE2) & "n H" & ChrW(&H1B0) & "ng"
    c.Add "Xu" & ChrW(&HE2) & "n H" & ChrW(&H1B0) & ChrW(&H1A1) & "ng - " & ChrW(&H110) & ChrW(&HE0) & " L" & ChrW(&H1EA1) & "t"
    c.Add "Xu" & ChrW(&HE2) & "n H" & ChrW(&H1EA3) & "i"
    c.Add "Xu" & ChrW(&HE2) & "n H" & ChrW(&H1ED3) & "ng"
    c.Add "Xu" & ChrW(&HE2) & "n L" & ChrW(&HE2) & "m"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart127", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart128(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Xu" & ChrW(&HE2) & "n L" & ChrW(&HE3) & "ng"
    c.Add "Xu" & ChrW(&HE2) & "n L" & ChrW(&HE3) & "nh"
    c.Add "Xu" & ChrW(&HE2) & "n L" & ChrW(&H169) & "ng"
    c.Add "Xu" & ChrW(&HE2) & "n L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Xu" & ChrW(&HE2) & "n L" & ChrW(&H1EAD) & "p"
    c.Add "Xu" & ChrW(&HE2) & "n L" & ChrW(&H1ED9) & "c"
    c.Add "Xu" & ChrW(&HE2) & "n Mai"
    c.Add "Xu" & ChrW(&HE2) & "n Nha"
    c.Add "Xu" & ChrW(&HE2) & "n Ph" & ChrW(&HFA)
    c.Add "Xu" & ChrW(&HE2) & "n Ph" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Xu" & ChrW(&HE2) & "n Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add "Xu" & ChrW(&HE2) & "n Quang"
    c.Add "Xu" & ChrW(&HE2) & "n Qu" & ChrW(&H1EBF)
    c.Add "Xu" & ChrW(&HE2) & "n S" & ChrW(&H1A1) & "n"
    c.Add "Xu" & ChrW(&HE2) & "n Th" & ChrW(&HE0) & "nh"
    c.Add "Xu" & ChrW(&HE2) & "n Th" & ChrW(&HE1) & "i"
    c.Add "Xu" & ChrW(&HE2) & "n Th" & ChrW(&H1ECD)
    c.Add "Xu" & ChrW(&HE2) & "n Th" & ChrW(&H1EDB) & "i S" & ChrW(&H1A1) & "n"
    c.Add "Xu" & ChrW(&HE2) & "n Tr" & ChrW(&HFA) & "c"
    c.Add "Xu" & ChrW(&HE2) & "n Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart128", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart129(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Xu" & ChrW(&HE2) & "n Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng - " & ChrW(&H110) & ChrW(&HE0) & " L" & ChrW(&H1EA1) & "t"
    c.Add "Xu" & ChrW(&HE2) & "n T" & ChrW(&HED) & "n"
    c.Add "Xu" & ChrW(&HE2) & "n Vi" & ChrW(&HEA) & "n"
    c.Add "Xu" & ChrW(&HE2) & "n V" & ChrW(&HE2) & "n"
    c.Add "Xu" & ChrW(&HE2) & "n " & ChrW(&HC1) & "i"
    c.Add "Xu" & ChrW(&HE2) & "n " & ChrW(&H110) & ChrW(&HE0) & "i"
    c.Add "Xu" & ChrW(&HE2) & "n " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add "Xu" & ChrW(&HE2) & "n " & ChrW(&H110) & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "Xu" & ChrW(&HE2) & "n " & ChrW(&H110) & ChrW(&H1EC9) & "nh"
    c.Add "Xu" & ChrW(&HE2) & "n " & ChrW(&H110) & ChrW(&H1ECB) & "nh"
    c.Add "X" & ChrW(&HE0) & " Phi" & ChrW(&HEA) & "n"
    c.Add "X" & ChrW(&HED) & "m V" & ChrW(&HE0) & "ng"
    c.Add "X" & ChrW(&HED) & "n M" & ChrW(&H1EA7) & "n"
    c.Add "X" & ChrW(&HF3) & "m Chi" & ChrW(&H1EBF) & "u"
    c.Add "X" & ChrW(&H1ED1) & "p"
    c.Add "Y T" & ChrW(&HFD)
    c.Add "Ya H" & ChrW(&H1ED9) & "i"
    c.Add "Ya Ly"
    c.Add "Ya Ma"
    c.Add "Yang Mao"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart129", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart130(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Y" & ChrW(&HEA) & "n B" & ChrW(&HE0) & "i"
    c.Add "Y" & ChrW(&HEA) & "n B" & ChrW(&HE1) & "i"
    c.Add "Y" & ChrW(&HEA) & "n B" & ChrW(&HEC) & "nh"
    c.Add "Y" & ChrW(&HEA) & "n Ch" & ChrW(&HE2) & "u"
    c.Add "Y" & ChrW(&HEA) & "n C" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "Y" & ChrW(&HEA) & "n D" & ChrW(&H169) & "ng"
    c.Add "Y" & ChrW(&HEA) & "n Hoa"
    c.Add "Y" & ChrW(&HEA) & "n H" & ChrW(&HF2) & "a"
    c.Add "Y" & ChrW(&HEA) & "n Kh" & ChrW(&HE1) & "nh"
    c.Add "Y" & ChrW(&HEA) & "n Kh" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add "Y" & ChrW(&HEA) & "n K" & ChrW(&H1EF3)
    c.Add "Y" & ChrW(&HEA) & "n L" & ChrW(&HE3) & "ng"
    c.Add "Y" & ChrW(&HEA) & "n L" & ChrW(&H1EA1) & "c"
    c.Add "Y" & ChrW(&HEA) & "n L" & ChrW(&H1EAD) & "p"
    c.Add "Y" & ChrW(&HEA) & "n Minh"
    c.Add "Y" & ChrW(&HEA) & "n M" & ChrW(&HF4)
    c.Add "Y" & ChrW(&HEA) & "n M" & ChrW(&H1EA1) & "c"
    c.Add "Y" & ChrW(&HEA) & "n M" & ChrW(&H1EF9)
    c.Add "Y" & ChrW(&HEA) & "n Na"
    c.Add "Y" & ChrW(&HEA) & "n Ngh" & ChrW(&H129) & "a"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart130", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart131(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Y" & ChrW(&HEA) & "n Nguy" & ChrW(&HEA) & "n"
    c.Add "Y" & ChrW(&HEA) & "n Nh" & ChrW(&HE2) & "n"
    c.Add "Y" & ChrW(&HEA) & "n Ninh"
    c.Add "Y" & ChrW(&HEA) & "n Phong"
    c.Add "Y" & ChrW(&HEA) & "n Ph" & ChrW(&HFA)
    c.Add "Y" & ChrW(&HEA) & "n Ph" & ChrW(&HFA) & "c"
    c.Add "Y" & ChrW(&HEA) & "n S" & ChrW(&H1A1) & "n"
    c.Add "Y" & ChrW(&HEA) & "n S" & ChrW(&H1EDF)
    c.Add "Y" & ChrW(&HEA) & "n Th" & ChrW(&HE0) & "nh"
    c.Add "Y" & ChrW(&HEA) & "n Th" & ChrW(&H1EAF) & "ng"
    c.Add "Y" & ChrW(&HEA) & "n Th" & ChrW(&H1EBF)
    c.Add "Y" & ChrW(&HEA) & "n Th" & ChrW(&H1ECB) & "nh"
    c.Add "Y" & ChrW(&HEA) & "n Th" & ChrW(&H1ECD)
    c.Add "Y" & ChrW(&HEA) & "n Th" & ChrW(&H1ED5)
    c.Add "Y" & ChrW(&HEA) & "n Th" & ChrW(&H1EE7) & "y"
    c.Add "Y" & ChrW(&HEA) & "n Trung"
    c.Add "Y" & ChrW(&HEA) & "n Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "Y" & ChrW(&HEA) & "n Tr" & ChrW(&H1EA1) & "ch"
    c.Add "Y" & ChrW(&HEA) & "n Tr" & ChrW(&H1ECB)
    c.Add "Y" & ChrW(&HEA) & "n T" & ChrW(&H1EEB)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart131", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart132(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Y" & ChrW(&HEA) & "n T" & ChrW(&H1EED)
    c.Add "Y" & ChrW(&HEA) & "n Xu" & ChrW(&HE2) & "n"
    c.Add "Y" & ChrW(&HEA) & "n " & ChrW(&H110) & ChrW(&H1ECB) & "nh"
    c.Add "Y" & ChrW(&HEA) & "n " & ChrW(&H110) & ChrW(&H1ED3) & "ng"
    c.Add "Y" & ChrW(&H1EBF) & "t Ki" & ChrW(&HEA) & "u"
    c.Add "x" & ChrW(&HE3) & " B" & ChrW(&H1EAF) & "c S" & ChrW(&H1A1) & "n"
    c.Add ChrW(&HC1) & "i Qu" & ChrW(&H1ED1) & "c"
    c.Add ChrW(&HC1) & "i T" & ChrW(&H1EED)
    c.Add ChrW(&HC2) & "n H" & ChrW(&H1EA3) & "o"
    c.Add ChrW(&HC2) & "n Thi"
    c.Add ChrW(&HC2) & "n T" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add ChrW(&HC2) & "u C" & ChrW(&H1A1)
    c.Add ChrW(&HC2) & "u L" & ChrW(&HE2) & "u"
    c.Add ChrW(&HD3) & "c Eo"
    c.Add ChrW(&HD4) & " Ch" & ChrW(&H1EE3) & " D" & ChrW(&H1EEB) & "a"
    c.Add ChrW(&HD4) & " Di" & ChrW(&HEA) & "n"
    c.Add ChrW(&HD4) & " Loan"
    c.Add ChrW(&HD4) & " L" & ChrW(&HE2) & "m"
    c.Add ChrW(&HD4) & " M" & ChrW(&HF4) & "n"
    c.Add ChrW(&HDD) & " Y" & ChrW(&HEA) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart132", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart133(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add ChrW(&H110) & "a Kia"
    c.Add ChrW(&H110) & "a Mai"
    c.Add ChrW(&H110) & "a Ph" & ChrW(&HFA) & "c"
    c.Add ChrW(&H110) & "ak Lua"
    c.Add ChrW(&H110) & "ak Nhau"
    c.Add ChrW(&H110) & "ak P" & ChrW(&H1A1)
    c.Add ChrW(&H110) & "ak Rong"
    c.Add ChrW(&H110) & "ak S" & ChrW(&H1A1) & "mei"
    c.Add ChrW(&H110) & "ak " & ChrW(&H110) & "oa"
    c.Add ChrW(&H110) & "akr" & ChrW(&HF4) & "ng"
    c.Add ChrW(&H110) & "am R" & ChrW(&HF4) & "ng 1"
    c.Add ChrW(&H110) & "am R" & ChrW(&HF4) & "ng 2"
    c.Add ChrW(&H110) & "am R" & ChrW(&HF4) & "ng 3"
    c.Add ChrW(&H110) & "am R" & ChrW(&HF4) & "ng 4"
    c.Add ChrW(&H110) & "an H" & ChrW(&H1EA3) & "i"
    c.Add ChrW(&H110) & "an Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng"
    c.Add ChrW(&H110) & "an " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "n"
    c.Add ChrW(&H110) & "inh Trang Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng"
    c.Add ChrW(&H110) & "inh V" & ChrW(&H103) & "n L" & ChrW(&HE2) & "m H" & ChrW(&HE0)
    c.Add ChrW(&H110) & "i" & ChrW(&H1EC1) & "m He"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart133", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart134(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add ChrW(&H110) & "i" & ChrW(&H1EC1) & "m Th" & ChrW(&H1EE5) & "y"
    c.Add ChrW(&H110) & "i" & ChrW(&H1EC1) & "n L" & ChrW(&H1B0)
    c.Add ChrW(&H110) & "i" & ChrW(&H1EC1) & "n Quang"
    c.Add ChrW(&H110) & "i" & ChrW(&H1EC1) & "n X" & ChrW(&HE1)
    c.Add ChrW(&H110) & "i" & ChrW(&H1EC7) & "n Bi" & ChrW(&HEA) & "n Ph" & ChrW(&H1EE7)
    c.Add ChrW(&H110) & "i" & ChrW(&H1EC7) & "n B" & ChrW(&HE0) & "n"
    c.Add ChrW(&H110) & "i" & ChrW(&H1EC7) & "n B" & ChrW(&HE0) & "n B" & ChrW(&H1EAF) & "c"
    c.Add ChrW(&H110) & "i" & ChrW(&H1EC7) & "n B" & ChrW(&HE0) & "n T" & ChrW(&HE2) & "y"
    c.Add ChrW(&H110) & "i" & ChrW(&H1EC7) & "n B" & ChrW(&HE0) & "n " & ChrW(&H110) & ChrW(&HF4) & "ng"
    c.Add ChrW(&H110) & "oan H" & ChrW(&HF9) & "ng"
    c.Add ChrW(&H110) & "o" & ChrW(&HE0) & "i D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add ChrW(&H110) & "o" & ChrW(&HE0) & "i Ph" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add ChrW(&H110) & "o" & ChrW(&HE0) & "n K" & ChrW(&H1EBF) & "t"
    c.Add ChrW(&H110) & "o" & ChrW(&HE0) & "n " & ChrW(&H110) & ChrW(&HE0) & "o"
    c.Add ChrW(&H110) & ChrW(&HE0) & " B" & ChrW(&H1EAF) & "c"
    c.Add ChrW(&H110) & ChrW(&HE0) & "m Th" & ChrW(&H1EE7) & "y"
    c.Add ChrW(&H110) & ChrW(&HE0) & "o Duy T" & ChrW(&H1EEB)
    c.Add ChrW(&H110) & ChrW(&HE0) & "o Vi" & ChrW(&HEA) & "n"
    c.Add ChrW(&H110) & ChrW(&HE0) & "o X" & ChrW(&HE1)
    c.Add ChrW(&H110) & ChrW(&HE1) & " B" & ChrW(&H1EA1) & "c"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart134", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart135(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add ChrW(&H110) & ChrW(&HE8) & "o Gia"
    c.Add ChrW(&H110) & ChrW(&HEC) & "nh C" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add ChrW(&H110) & ChrW(&HEC) & "nh L" & ChrW(&H1EAD) & "p"
    c.Add ChrW(&H110) & ChrW(&HEC) & "nh Phong"
    c.Add ChrW(&H110) & ChrW(&HF4) & " L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add ChrW(&H110) & ChrW(&HF4) & " Vinh"
    c.Add ChrW(&H110) & ChrW(&HF4) & "n Ch" & ChrW(&HE2) & "u"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng A"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Anh"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Cu" & ChrW(&HF4) & "ng"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng C" & ChrW(&H1EE9) & "u"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Gia Ngh" & ChrW(&H129) & "a"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Giang"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Hi" & ChrW(&H1EBF) & "u"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Hi" & ChrW(&H1EC7) & "p"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Hoa L" & ChrW(&H1B0)
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng H" & ChrW(&HE0)
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng H" & ChrW(&HF2) & "a"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng H" & ChrW(&H1B0) & "ng"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng H" & ChrW(&H1B0) & "ng Thu" & ChrW(&H1EAD) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart135", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart136(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng H" & ChrW(&H1EA3) & "i"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Kh" & ChrW(&HE1) & "nh S" & ChrW(&H1A1) & "n"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Kh" & ChrW(&HEA)
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Kinh"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng L" & ChrW(&H1ED9) & "c"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Mai"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Ng" & ChrW(&H169)
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Ng" & ChrW(&H1EA1) & "c"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Ninh H" & ChrW(&HF2) & "a"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Ph" & ChrW(&HFA)
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Quan"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Quang"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng S" & ChrW(&H1A1) & "n"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Thu" & ChrW(&H1EAD) & "n"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Th" & ChrW(&HE0) & "nh"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Th" & ChrW(&HE1) & "i"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Th" & ChrW(&HE1) & "i Ninh"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Th" & ChrW(&H1EA1) & "nh"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Th" & ChrW(&H1ECD)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart136", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart137(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Th" & ChrW(&H1EE5) & "y Anh"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Ti" & ChrW(&HEA) & "n H" & ChrW(&H1B0) & "ng"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Ti" & ChrW(&H1EBF) & "n"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Ti" & ChrW(&H1EC1) & "n H" & ChrW(&H1EA3) & "i"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Tri" & ChrW(&H1EC1) & "u"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Tr" & ChrW(&HE0) & " B" & ChrW(&H1ED3) & "ng"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Tr" & ChrW(&H1EA1) & "ch"
    c.Add ChrW(&H110) & ChrW(&H103) & "k Bla"
    c.Add ChrW(&H110) & ChrW(&H103) & "k C" & ChrW(&H1EA5) & "m"
    c.Add ChrW(&H110) & ChrW(&H103) & "k H" & ChrW(&HE0)
    c.Add ChrW(&H110) & ChrW(&H103) & "k K" & ChrW(&HF4) & "i"
    c.Add ChrW(&H110) & ChrW(&H103) & "k Long"
    c.Add ChrW(&H110) & ChrW(&H103) & "k Mar"
    c.Add ChrW(&H110) & ChrW(&H103) & "k M" & ChrW(&HF4) & "n"
    c.Add ChrW(&H110) & ChrW(&H103) & "k Pl" & ChrW(&HF4)
    c.Add ChrW(&H110) & ChrW(&H103) & "k Pxi"
    c.Add ChrW(&H110) & ChrW(&H103) & "k P" & ChrW(&HE9) & "k"
    c.Add ChrW(&H110) & ChrW(&H103) & "k Rve"
    c.Add ChrW(&H110) & ChrW(&H103) & "k R" & ChrW(&H1A1) & " Wa"
    c.Add ChrW(&H110) & ChrW(&H103) & "k Sao"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart137", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart138(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add ChrW(&H110) & ChrW(&H103) & "k Song"
    c.Add ChrW(&H110) & ChrW(&H103) & "k T" & ChrW(&HF4)
    c.Add ChrW(&H110) & ChrW(&H103) & "k T" & ChrW(&H1EDD) & " Kan"
    c.Add ChrW(&H110) & ChrW(&H103) & "k Ui"
    c.Add ChrW(&H110) & ChrW(&H103) & "k " & ChrW(&H1A0)
    c.Add ChrW(&H110) & ChrW(&H103) & ChrW(&H323) & "c khu Ho" & ChrW(&HE0) & "ng Sa"
    c.Add ChrW(&H110) & ChrW(&H1A1) & "n D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add ChrW(&H110) & ChrW(&H1B0) & ChrW(&H1EDD) & "ng An"
    c.Add ChrW(&H110) & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Hoa"
    c.Add ChrW(&H110) & ChrW(&H1B0) & ChrW(&H1EDD) & "ng H" & ChrW(&HE0) & "o"
    c.Add ChrW(&H110) & ChrW(&H1B0) & ChrW(&H1EDD) & "ng H" & ChrW(&H1ED3) & "ng"
    c.Add ChrW(&H110) & ChrW(&H1B0) & ChrW(&H1EDD) & "ng Th" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & " Huoai"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & " Huoai 2"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & " Huoai 3"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & " T" & ChrW(&H1EBB) & "h"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & " T" & ChrW(&H1EBB) & "h 2"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & " T" & ChrW(&H1EBB) & "h 3"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & "i An"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & "i Ho" & ChrW(&HE0) & "ng"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart138", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart139(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add ChrW(&H110) & ChrW(&H1EA1) & "i Hu" & ChrW(&H1EC7)
    c.Add ChrW(&H110) & ChrW(&H1EA1) & "i H" & ChrW(&H1EA3) & "i"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & "i Lai"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & "i L" & ChrW(&HE3) & "nh"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & "i L" & ChrW(&H1ED9) & "c"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & "i M" & ChrW(&H1ED7)
    c.Add ChrW(&H110) & ChrW(&H1EA1) & "i Ng" & ChrW(&HE3) & "i"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & "i Ph" & ChrW(&HFA) & "c"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & "i Ph" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & "i S" & ChrW(&H1A1) & "n"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & "i Thanh"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & "i Th" & ChrW(&HE0) & "nh"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & "i T" & ChrW(&H1EEB)
    c.Add ChrW(&H110) & ChrW(&H1EA1) & "i Xuy" & ChrW(&HEA) & "n"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & "i " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "n"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & "i " & ChrW(&H110) & ChrW(&HEC) & "nh"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & "i " & ChrW(&H110) & ChrW(&H1ED3) & "ng"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & "o Th" & ChrW(&H1EA1) & "nh"
    c.Add ChrW(&H110) & ChrW(&H1EA1) & "o Tr" & ChrW(&HF9)
    c.Add ChrW(&H110) & ChrW(&H1EA5) & "t M" & ChrW(&H169) & "i"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart139", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart140(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add ChrW(&H110) & ChrW(&H1EA5) & "t M" & ChrW(&H1EDB) & "i"
    c.Add ChrW(&H110) & ChrW(&H1EA5) & "t " & ChrW(&H110) & ChrW(&H1ECF)
    c.Add ChrW(&H110) & ChrW(&H1EA7) & "m D" & ChrW(&H1A1) & "i"
    c.Add ChrW(&H110) & ChrW(&H1EA7) & "m H" & ChrW(&HE0)
    c.Add ChrW(&H110) & ChrW(&H1EAF) & "c Pring"
    c.Add ChrW(&H110) & ChrW(&H1EAF) & "k Li" & ChrW(&HEA) & "ng"
    c.Add ChrW(&H110) & ChrW(&H1EAF) & "k Mil"
    c.Add ChrW(&H110) & ChrW(&H1EAF) & "k Ph" & ChrW(&H1A1) & "i"
    c.Add ChrW(&H110) & ChrW(&H1EAF) & "k Song"
    c.Add ChrW(&H110) & ChrW(&H1EAF) & "k S" & ChrW(&H1EAF) & "k"
    c.Add ChrW(&H110) & ChrW(&H1EAF) & "k Wil"
    c.Add ChrW(&H110) & ChrW(&H1EB7) & "ng Th" & ChrW(&HF9) & "y Tr" & ChrW(&HE2) & "m"
    c.Add ChrW(&H110) & ChrW(&H1EC1) & " Gi"
    c.Add ChrW(&H110) & ChrW(&H1ECB) & "nh C" & ChrW(&HF4) & "ng"
    c.Add ChrW(&H110) & ChrW(&H1ECB) & "nh H" & ChrW(&HF2) & "a"
    c.Add ChrW(&H110) & ChrW(&H1ECB) & "nh H" & ChrW(&HF3) & "a"
    c.Add ChrW(&H110) & ChrW(&H1ECB) & "nh M" & ChrW(&H1EF9)
    c.Add ChrW(&H110) & ChrW(&H1ECB) & "nh Qu" & ChrW(&HE1) & "n"
    c.Add ChrW(&H110) & ChrW(&H1ECB) & "nh Th" & ChrW(&HE0) & "nh"
    c.Add ChrW(&H110) & ChrW(&H1ECB) & "nh T" & ChrW(&HE2) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart140", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart141(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add ChrW(&H110) & ChrW(&H1ED1) & "c Binh Ki" & ChrW(&H1EC1) & "u"
    c.Add ChrW(&H110) & ChrW(&H1ED1) & "ng " & ChrW(&H110) & "a"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & " S" & ChrW(&H1A1) & "n"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng B" & ChrW(&H1EB1) & "ng"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng Ch" & ChrW(&HE2) & "u"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng D" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng H" & ChrW(&H1EDB) & "i"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng H" & ChrW(&H1EF7)
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng Kho"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng Kh" & ChrW(&H1EDF) & "i"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng K" & ChrW(&H1EF3)
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng L" & ChrW(&HEA)
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng L" & ChrW(&H1ED9) & "c"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng Nguy" & ChrW(&HEA) & "n"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng Ph" & ChrW(&HFA)
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng Ph" & ChrW(&HFA) & "c"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng S" & ChrW(&H1A1) & "n"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng Thu" & ChrW(&H1EAD) & "n"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng Th" & ChrW(&HE1) & "i"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart141", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart142(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng Th" & ChrW(&H1ECB) & "nh"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng Ti" & ChrW(&H1EBF) & "n"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng T" & ChrW(&HE2) & "m"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng Vi" & ChrW(&H1EC7) & "t"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng V" & ChrW(&H103) & "n"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng Xo" & ChrW(&HE0) & "i"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng Xu" & ChrW(&HE2) & "n"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng Y" & ChrW(&HEA) & "n"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng " & ChrW(&H110) & ChrW(&H103) & "ng"
    c.Add ChrW(&H110) & ChrW(&H1ED9) & "c L" & ChrW(&H1EAD) & "p"
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c An"
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c B" & ChrW(&HEC) & "nh"
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c Ch" & ChrW(&HE2) & "u"
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c C" & ChrW(&H1A1)
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c Hu" & ChrW(&H1EC7)
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c H" & ChrW(&HF2) & "a"
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c H" & ChrW(&H1EE3) & "p"
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c Linh"
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c Long"
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c L" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart142", Err.description
End Sub

Private Sub LoadRawAdministrativeUnitNames_CommunesPart143(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c L" & ChrW(&H1EAD) & "p"
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c Minh"
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c Nhu" & ChrW(&H1EAD) & "n"
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c Nh" & ChrW(&HE0) & "n"
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c Ph" & ChrW(&HFA)
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c Ph" & ChrW(&H1ED5)
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c Quang"
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c Th" & ChrW(&H1ECB) & "nh"
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c Th" & ChrW(&H1ECD)
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c Tr" & ChrW(&H1ECD) & "ng"
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c Xu" & ChrW(&HE2) & "n"
    c.Add ChrW(&H110) & ChrW(&H1EE9) & "c " & ChrW(&H110) & ChrW(&H1ED3) & "ng"
    c.Add ChrW(&H1EE8) & "ng H" & ChrW(&HF2) & "a"
    c.Add ChrW(&H1EE8) & "ng Thi" & ChrW(&HEA) & "n"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_CommunesPart143", Err.description
End Sub

Private Function LoadRawAdministrativeUnitNames_Communes() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    LoadRawAdministrativeUnitNames_CommunesPart1 c
    LoadRawAdministrativeUnitNames_CommunesPart2 c
    LoadRawAdministrativeUnitNames_CommunesPart3 c
    LoadRawAdministrativeUnitNames_CommunesPart4 c
    LoadRawAdministrativeUnitNames_CommunesPart5 c
    LoadRawAdministrativeUnitNames_CommunesPart6 c
    LoadRawAdministrativeUnitNames_CommunesPart7 c
    LoadRawAdministrativeUnitNames_CommunesPart8 c
    LoadRawAdministrativeUnitNames_CommunesPart9 c
    LoadRawAdministrativeUnitNames_CommunesPart10 c
    LoadRawAdministrativeUnitNames_CommunesPart11 c
    LoadRawAdministrativeUnitNames_CommunesPart12 c
    LoadRawAdministrativeUnitNames_CommunesPart13 c
    LoadRawAdministrativeUnitNames_CommunesPart14 c
    LoadRawAdministrativeUnitNames_CommunesPart15 c
    LoadRawAdministrativeUnitNames_CommunesPart16 c
    LoadRawAdministrativeUnitNames_CommunesPart17 c
    LoadRawAdministrativeUnitNames_CommunesPart18 c
    LoadRawAdministrativeUnitNames_CommunesPart19 c
    LoadRawAdministrativeUnitNames_CommunesPart20 c
    LoadRawAdministrativeUnitNames_CommunesPart21 c
    LoadRawAdministrativeUnitNames_CommunesPart22 c
    LoadRawAdministrativeUnitNames_CommunesPart23 c
    LoadRawAdministrativeUnitNames_CommunesPart24 c
    LoadRawAdministrativeUnitNames_CommunesPart25 c
    LoadRawAdministrativeUnitNames_CommunesPart26 c
    LoadRawAdministrativeUnitNames_CommunesPart27 c
    LoadRawAdministrativeUnitNames_CommunesPart28 c
    LoadRawAdministrativeUnitNames_CommunesPart29 c
    LoadRawAdministrativeUnitNames_CommunesPart30 c
    LoadRawAdministrativeUnitNames_CommunesPart31 c
    LoadRawAdministrativeUnitNames_CommunesPart32 c
    LoadRawAdministrativeUnitNames_CommunesPart33 c
    LoadRawAdministrativeUnitNames_CommunesPart34 c
    LoadRawAdministrativeUnitNames_CommunesPart35 c
    LoadRawAdministrativeUnitNames_CommunesPart36 c
    LoadRawAdministrativeUnitNames_CommunesPart37 c
    LoadRawAdministrativeUnitNames_CommunesPart38 c
    LoadRawAdministrativeUnitNames_CommunesPart39 c
    LoadRawAdministrativeUnitNames_CommunesPart40 c
    LoadRawAdministrativeUnitNames_CommunesPart41 c
    LoadRawAdministrativeUnitNames_CommunesPart42 c
    LoadRawAdministrativeUnitNames_CommunesPart43 c
    LoadRawAdministrativeUnitNames_CommunesPart44 c
    LoadRawAdministrativeUnitNames_CommunesPart45 c
    LoadRawAdministrativeUnitNames_CommunesPart46 c
    LoadRawAdministrativeUnitNames_CommunesPart47 c
    LoadRawAdministrativeUnitNames_CommunesPart48 c
    LoadRawAdministrativeUnitNames_CommunesPart49 c
    LoadRawAdministrativeUnitNames_CommunesPart50 c
    LoadRawAdministrativeUnitNames_CommunesPart51 c
    LoadRawAdministrativeUnitNames_CommunesPart52 c
    LoadRawAdministrativeUnitNames_CommunesPart53 c
    LoadRawAdministrativeUnitNames_CommunesPart54 c
    LoadRawAdministrativeUnitNames_CommunesPart55 c
    LoadRawAdministrativeUnitNames_CommunesPart56 c
    LoadRawAdministrativeUnitNames_CommunesPart57 c
    LoadRawAdministrativeUnitNames_CommunesPart58 c
    LoadRawAdministrativeUnitNames_CommunesPart59 c
    LoadRawAdministrativeUnitNames_CommunesPart60 c
    LoadRawAdministrativeUnitNames_CommunesPart61 c
    LoadRawAdministrativeUnitNames_CommunesPart62 c
    LoadRawAdministrativeUnitNames_CommunesPart63 c
    LoadRawAdministrativeUnitNames_CommunesPart64 c
    LoadRawAdministrativeUnitNames_CommunesPart65 c
    LoadRawAdministrativeUnitNames_CommunesPart66 c
    LoadRawAdministrativeUnitNames_CommunesPart67 c
    LoadRawAdministrativeUnitNames_CommunesPart68 c
    LoadRawAdministrativeUnitNames_CommunesPart69 c
    LoadRawAdministrativeUnitNames_CommunesPart70 c
    LoadRawAdministrativeUnitNames_CommunesPart71 c
    LoadRawAdministrativeUnitNames_CommunesPart72 c
    LoadRawAdministrativeUnitNames_CommunesPart73 c
    LoadRawAdministrativeUnitNames_CommunesPart74 c
    LoadRawAdministrativeUnitNames_CommunesPart75 c
    LoadRawAdministrativeUnitNames_CommunesPart76 c
    LoadRawAdministrativeUnitNames_CommunesPart77 c
    LoadRawAdministrativeUnitNames_CommunesPart78 c
    LoadRawAdministrativeUnitNames_CommunesPart79 c
    LoadRawAdministrativeUnitNames_CommunesPart80 c
    LoadRawAdministrativeUnitNames_CommunesPart81 c
    LoadRawAdministrativeUnitNames_CommunesPart82 c
    LoadRawAdministrativeUnitNames_CommunesPart83 c
    LoadRawAdministrativeUnitNames_CommunesPart84 c
    LoadRawAdministrativeUnitNames_CommunesPart85 c
    LoadRawAdministrativeUnitNames_CommunesPart86 c
    LoadRawAdministrativeUnitNames_CommunesPart87 c
    LoadRawAdministrativeUnitNames_CommunesPart88 c
    LoadRawAdministrativeUnitNames_CommunesPart89 c
    LoadRawAdministrativeUnitNames_CommunesPart90 c
    LoadRawAdministrativeUnitNames_CommunesPart91 c
    LoadRawAdministrativeUnitNames_CommunesPart92 c
    LoadRawAdministrativeUnitNames_CommunesPart93 c
    LoadRawAdministrativeUnitNames_CommunesPart94 c
    LoadRawAdministrativeUnitNames_CommunesPart95 c
    LoadRawAdministrativeUnitNames_CommunesPart96 c
    LoadRawAdministrativeUnitNames_CommunesPart97 c
    LoadRawAdministrativeUnitNames_CommunesPart98 c
    LoadRawAdministrativeUnitNames_CommunesPart99 c
    LoadRawAdministrativeUnitNames_CommunesPart100 c
    LoadRawAdministrativeUnitNames_CommunesPart101 c
    LoadRawAdministrativeUnitNames_CommunesPart102 c
    LoadRawAdministrativeUnitNames_CommunesPart103 c
    LoadRawAdministrativeUnitNames_CommunesPart104 c
    LoadRawAdministrativeUnitNames_CommunesPart105 c
    LoadRawAdministrativeUnitNames_CommunesPart106 c
    LoadRawAdministrativeUnitNames_CommunesPart107 c
    LoadRawAdministrativeUnitNames_CommunesPart108 c
    LoadRawAdministrativeUnitNames_CommunesPart109 c
    LoadRawAdministrativeUnitNames_CommunesPart110 c
    LoadRawAdministrativeUnitNames_CommunesPart111 c
    LoadRawAdministrativeUnitNames_CommunesPart112 c
    LoadRawAdministrativeUnitNames_CommunesPart113 c
    LoadRawAdministrativeUnitNames_CommunesPart114 c
    LoadRawAdministrativeUnitNames_CommunesPart115 c
    LoadRawAdministrativeUnitNames_CommunesPart116 c
    LoadRawAdministrativeUnitNames_CommunesPart117 c
    LoadRawAdministrativeUnitNames_CommunesPart118 c
    LoadRawAdministrativeUnitNames_CommunesPart119 c
    LoadRawAdministrativeUnitNames_CommunesPart120 c
    LoadRawAdministrativeUnitNames_CommunesPart121 c
    LoadRawAdministrativeUnitNames_CommunesPart122 c
    LoadRawAdministrativeUnitNames_CommunesPart123 c
    LoadRawAdministrativeUnitNames_CommunesPart124 c
    LoadRawAdministrativeUnitNames_CommunesPart125 c
    LoadRawAdministrativeUnitNames_CommunesPart126 c
    LoadRawAdministrativeUnitNames_CommunesPart127 c
    LoadRawAdministrativeUnitNames_CommunesPart128 c
    LoadRawAdministrativeUnitNames_CommunesPart129 c
    LoadRawAdministrativeUnitNames_CommunesPart130 c
    LoadRawAdministrativeUnitNames_CommunesPart131 c
    LoadRawAdministrativeUnitNames_CommunesPart132 c
    LoadRawAdministrativeUnitNames_CommunesPart133 c
    LoadRawAdministrativeUnitNames_CommunesPart134 c
    LoadRawAdministrativeUnitNames_CommunesPart135 c
    LoadRawAdministrativeUnitNames_CommunesPart136 c
    LoadRawAdministrativeUnitNames_CommunesPart137 c
    LoadRawAdministrativeUnitNames_CommunesPart138 c
    LoadRawAdministrativeUnitNames_CommunesPart139 c
    LoadRawAdministrativeUnitNames_CommunesPart140 c
    LoadRawAdministrativeUnitNames_CommunesPart141 c
    LoadRawAdministrativeUnitNames_CommunesPart142 c
    LoadRawAdministrativeUnitNames_CommunesPart143 c
    Set LoadRawAdministrativeUnitNames_Communes = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames_Communes", Err.description
End Function

Public Function LoadRawAdministrativeUnitNames() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "1.0.0"
    d.Add "status", "Day du 34 tinh/thanh + 2854 ten xa/phuong rieng biet (mot so ten trung lap giua cac tinh, da gop lai), theo cau truc hanh chinh hai cap sau sap nhap thang 7/2025."
    d.Add "sourceLabel", "THONG LE"
    d.Add "provinces", LoadRawAdministrativeUnitNames_Provinces()
    d.Add "communes", LoadRawAdministrativeUnitNames_Communes()
    Set LoadRawAdministrativeUnitNames = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawAdministrativeUnitNames", Err.description
End Function

Private Function LoadRawFontPatterns_Patterns_Tcvn3Upper_Examples() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add ".VnTimeH"
    c.Add ".VnArialH"
    Set LoadRawFontPatterns_Patterns_Tcvn3Upper_Examples = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFontPatterns_Patterns_Tcvn3Upper_Examples", Err.description
End Function

Private Function LoadRawFontPatterns_Patterns_Tcvn3Upper() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "regex", "^\.Vn.*H$"
    d.Add "mapFile", "bang-ma-tcvn3-upper.json"
    d.Add "examples", LoadRawFontPatterns_Patterns_Tcvn3Upper_Examples()
    Set LoadRawFontPatterns_Patterns_Tcvn3Upper = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFontPatterns_Patterns_Tcvn3Upper", Err.description
End Function

Private Function LoadRawFontPatterns_Patterns_Tcvn3Lower_Examples() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add ".VnTime"
    c.Add ".VnArial"
    c.Add ".VnCentury Schoolbook"
    Set LoadRawFontPatterns_Patterns_Tcvn3Lower_Examples = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFontPatterns_Patterns_Tcvn3Lower_Examples", Err.description
End Function

Private Function LoadRawFontPatterns_Patterns_Tcvn3Lower() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "regex", "^\.Vn(?!.*H$).*$"
    d.Add "mapFile", "bang-ma-tcvn3-lower.json"
    d.Add "examples", LoadRawFontPatterns_Patterns_Tcvn3Lower_Examples()
    Set LoadRawFontPatterns_Patterns_Tcvn3Lower = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFontPatterns_Patterns_Tcvn3Lower", Err.description
End Function

Private Function LoadRawFontPatterns_Patterns_Vni_Examples() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "VNI-Times"
    c.Add "VNI-Helve"
    c.Add "VNI-Aptima"
    Set LoadRawFontPatterns_Patterns_Vni_Examples = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFontPatterns_Patterns_Vni_Examples", Err.description
End Function

Private Function LoadRawFontPatterns_Patterns_Vni() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "regex", "^VNI-"
    d.Add "mapFile", "bang-ma-vni.json"
    d.Add "examples", LoadRawFontPatterns_Patterns_Vni_Examples()
    Set LoadRawFontPatterns_Patterns_Vni = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFontPatterns_Patterns_Vni", Err.description
End Function

Private Function LoadRawFontPatterns_Patterns() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "tcvn3Upper", LoadRawFontPatterns_Patterns_Tcvn3Upper()
    d.Add "tcvn3Lower", LoadRawFontPatterns_Patterns_Tcvn3Lower()
    d.Add "vni", LoadRawFontPatterns_Patterns_Vni()
    Set LoadRawFontPatterns_Patterns = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFontPatterns_Patterns", Err.description
End Function

Public Function LoadRawFontPatterns() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "1.0.0"
    d.Add "status", "mau regex day du, danh sach loai tru con trong - can bo sung khi gap thuc te"
    d.Add "sourceLabel", "SUY RA"
    d.Add "patterns", LoadRawFontPatterns_Patterns()
    d.Add "excludeFromUpperMatch", New Collection
    Set LoadRawFontPatterns = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFontPatterns", Err.description
End Function

Private Function LoadRawIyMapping_Pairs() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "k" & ChrW(&H1EC9), "k" & ChrW(&H1EF7)
    d.Add "k" & ChrW(&H1EC9) & " ni" & ChrW(&H1EC7) & "m", "k" & ChrW(&H1EF7) & " ni" & ChrW(&H1EC7) & "m"
    d.Add "k" & ChrW(&H1EC9) & " lu" & ChrW(&H1EAD) & "t", "k" & ChrW(&H1EF7) & " lu" & ChrW(&H1EAD) & "t"
    d.Add "k" & ChrW(&H129), "k" & ChrW(&H1EF9)
    d.Add "k" & ChrW(&H129) & " n" & ChrW(&H103) & "ng", "k" & ChrW(&H1EF9) & " n" & ChrW(&H103) & "ng"
    d.Add "k" & ChrW(&H129) & " s" & ChrW(&H1B0), "k" & ChrW(&H1EF9) & " s" & ChrW(&H1B0)
    d.Add "l" & ChrW(&HED), "l" & ChrW(&HFD)
    d.Add "l" & ChrW(&HED) & " lu" & ChrW(&H1EAD) & "n", "l" & ChrW(&HFD) & " lu" & ChrW(&H1EAD) & "n"
    d.Add "l" & ChrW(&HED) & " do", "l" & ChrW(&HFD) & " do"
    d.Add "l" & ChrW(&HED) & " thuy" & ChrW(&H1EBF) & "t", "l" & ChrW(&HFD) & " thuy" & ChrW(&H1EBF) & "t"
    d.Add "m" & ChrW(&H129), "m" & ChrW(&H1EF9)
    d.Add "m" & ChrW(&H129) & " thu" & ChrW(&H1EAD) & "t", "m" & ChrW(&H1EF9) & " thu" & ChrW(&H1EAD) & "t"
    d.Add "t" & ChrW(&H1EC9), "t" & ChrW(&H1EF7)
    d.Add "t" & ChrW(&H1EC9) & " l" & ChrW(&H1EC7), "t" & ChrW(&H1EF7) & " l" & ChrW(&H1EC7)
    d.Add "s" & ChrW(&H129), "s" & ChrW(&H1EF9)
    d.Add "b" & ChrW(&HE1) & "c s" & ChrW(&H129), "b" & ChrW(&HE1) & "c s" & ChrW(&H1EF9)
    d.Add "hi", "hy"
    d.Add "hi v" & ChrW(&H1ECD) & "ng", "hy v" & ChrW(&H1ECD) & "ng"
    d.Add "vi", "vy"
    d.Add "k" & ChrW(&HEC), "k" & ChrW(&H1EF3)
    d.Add "k" & ChrW(&HEC) & " h" & ChrW(&H1EA1) & "n", "k" & ChrW(&H1EF3) & " h" & ChrW(&H1EA1) & "n"
    d.Add "k" & ChrW(&HEC) & " thi", "k" & ChrW(&H1EF3) & " thi"
    d.Add "chu k" & ChrW(&HEC), "chu k" & ChrW(&H1EF3)
    d.Add ChrW(&H111) & ChrW(&H1ECB) & "a l" & ChrW(&HED), ChrW(&H111) & ChrW(&H1ECB) & "a l" & ChrW(&HFD)
    d.Add "v" & ChrW(&H1EAD) & "t l" & ChrW(&HED), "v" & ChrW(&H1EAD) & "t l" & ChrW(&HFD)
    d.Add "qu" & ChrW(&H1EA3) & "n l" & ChrW(&HED), "qu" & ChrW(&H1EA3) & "n l" & ChrW(&HFD)
    d.Add "x" & ChrW(&H1EED) & " l" & ChrW(&HED), "x" & ChrW(&H1EED) & " l" & ChrW(&HFD)
    d.Add "h" & ChrW(&H1EE3) & "p l" & ChrW(&HED), "h" & ChrW(&H1EE3) & "p l" & ChrW(&HFD)
    d.Add "nguy" & ChrW(&HEA) & "n l" & ChrW(&HED), "nguy" & ChrW(&HEA) & "n l" & ChrW(&HFD)
    d.Add "c" & ChrW(&HF4) & "ng l" & ChrW(&HED), "c" & ChrW(&HF4) & "ng l" & ChrW(&HFD)
    d.Add ChrW(&H111) & ChrW(&H1EA1) & "i l" & ChrW(&HED), ChrW(&H111) & ChrW(&H1EA1) & "i l" & ChrW(&HFD)
    d.Add "l" & ChrW(&HED) & " l" & ChrW(&H1ECB) & "ch", "l" & ChrW(&HFD) & " l" & ChrW(&H1ECB) & "ch"
    d.Add "th" & ChrW(&H1EBF) & " k" & ChrW(&H1EC9), "th" & ChrW(&H1EBF) & " k" & ChrW(&H1EF7)
    d.Add "k" & ChrW(&H1EC9) & " y" & ChrW(&H1EBF) & "u", "k" & ChrW(&H1EF7) & " y" & ChrW(&H1EBF) & "u"
    d.Add "k" & ChrW(&H129) & " thu" & ChrW(&H1EAD) & "t", "k" & ChrW(&H1EF9) & " thu" & ChrW(&H1EAD) & "t"
    d.Add "k" & ChrW(&HED), "k" & ChrW(&HFD)
    d.Add "k" & ChrW(&HED) & " hi" & ChrW(&H1EC7) & "u", "k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u"
    d.Add "k" & ChrW(&HED) & " t" & ChrW(&HEA) & "n", "k" & ChrW(&HFD) & " t" & ChrW(&HEA) & "n"
    d.Add "ch" & ChrW(&H1EEF) & " k" & ChrW(&HED), "ch" & ChrW(&H1EEF) & " k" & ChrW(&HFD)
    d.Add "k" & ChrW(&HED) & " k" & ChrW(&H1EBF) & "t", "k" & ChrW(&HFD) & " k" & ChrW(&H1EBF) & "t"
    Set LoadRawIyMapping_Pairs = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawIyMapping_Pairs", Err.description
End Function

Private Function LoadRawIyMapping_ExcludeAbsolute_DocumentTypeNames_Terms() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "Quy ch" & ChrW(&H1EBF)
    c.Add "Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh"
    c.Add "Quy" & ChrW(&H1EBF) & "t " & ChrW(&H111) & ChrW(&H1ECB) & "nh"
    c.Add "Quy tr" & ChrW(&HEC) & "nh"
    c.Add "quy ch" & ChrW(&H1EBF)
    c.Add "quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh"
    c.Add "quy" & ChrW(&H1EBF) & "t " & ChrW(&H111) & ChrW(&H1ECB) & "nh"
    c.Add "quy tr" & ChrW(&HEC) & "nh"
    Set LoadRawIyMapping_ExcludeAbsolute_DocumentTypeNames_Terms = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawIyMapping_ExcludeAbsolute_DocumentTypeNames_Terms", Err.description
End Function

Private Function LoadRawIyMapping_ExcludeAbsolute_DocumentTypeNames() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "terms", LoadRawIyMapping_ExcludeAbsolute_DocumentTypeNames_Terms()
    Set LoadRawIyMapping_ExcludeAbsolute_DocumentTypeNames = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawIyMapping_ExcludeAbsolute_DocumentTypeNames", Err.description
End Function

Private Function LoadRawIyMapping_ExcludeAbsolute_TypeAbbreviations_Terms() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "QC"
    c.Add "Qy" & ChrW(&H110)
    c.Add "Q" & ChrW(&H110)
    Set LoadRawIyMapping_ExcludeAbsolute_TypeAbbreviations_Terms = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawIyMapping_ExcludeAbsolute_TypeAbbreviations_Terms", Err.description
End Function

Private Function LoadRawIyMapping_ExcludeAbsolute_TypeAbbreviations() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "terms", LoadRawIyMapping_ExcludeAbsolute_TypeAbbreviations_Terms()
    Set LoadRawIyMapping_ExcludeAbsolute_TypeAbbreviations = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawIyMapping_ExcludeAbsolute_TypeAbbreviations", Err.description
End Function

Private Function LoadRawIyMapping_ExcludeAbsolute_Nd30Terminology_Terms() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "k" & ChrW(&H1EF9) & " thu" & ChrW(&H1EAD) & "t tr" & ChrW(&HEC) & "nh b" & ChrW(&HE0) & "y"
    c.Add "k" & ChrW(&H1EF9) & " thu" & ChrW(&H1EAD) & "t"
    c.Add "k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u"
    c.Add "ch" & ChrW(&H1EEF) & " k" & ChrW(&HFD)
    c.Add "ch" & ChrW(&H1EEF) & " k" & ChrW(&HFD) & " s" & ChrW(&H1ED1)
    c.Add "k" & ChrW(&HFD) & " s" & ChrW(&H1ED1)
    c.Add "k" & ChrW(&HFD) & " ban h" & ChrW(&HE0) & "nh"
    c.Add "k" & ChrW(&HFD) & " thay"
    c.Add "k" & ChrW(&HFD) & " thay m" & ChrW(&H1EB7) & "t"
    c.Add "k" & ChrW(&HFD) & " th" & ChrW(&H1EEB) & "a l" & ChrW(&H1EC7) & "nh"
    c.Add "k" & ChrW(&HFD) & " th" & ChrW(&H1EEB) & "a " & ChrW(&H1EE7) & "y quy" & ChrW(&H1EC1) & "n"
    c.Add "ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i k" & ChrW(&HFD)
    c.Add "quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh"
    c.Add "quy ch" & ChrW(&H1EBF)
    c.Add "quy" & ChrW(&H1EBF) & "t " & ChrW(&H111) & ChrW(&H1ECB) & "nh"
    c.Add "quy" & ChrW(&H1EC1) & "n h" & ChrW(&H1EA1) & "n"
    Set LoadRawIyMapping_ExcludeAbsolute_Nd30Terminology_Terms = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawIyMapping_ExcludeAbsolute_Nd30Terminology_Terms", Err.description
End Function

Private Function LoadRawIyMapping_ExcludeAbsolute_Nd30Terminology() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "terms", LoadRawIyMapping_ExcludeAbsolute_Nd30Terminology_Terms()
    Set LoadRawIyMapping_ExcludeAbsolute_Nd30Terminology = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawIyMapping_ExcludeAbsolute_Nd30Terminology", Err.description
End Function

Private Function LoadRawIyMapping_ExcludeAbsolute_NationalTitle_Terms() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "C" & ChrW(&H1ED8) & "NG H" & ChrW(&HD2) & "A X" & ChrW(&HC3) & " H" & ChrW(&H1ED8) & "I CH" & ChrW(&H1EE6) & " NGH" & ChrW(&H128) & "A VI" & ChrW(&H1EC6) & "T NAM"
    c.Add "CH" & ChrW(&H1EE6) & " NGH" & ChrW(&H128) & "A"
    Set LoadRawIyMapping_ExcludeAbsolute_NationalTitle_Terms = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawIyMapping_ExcludeAbsolute_NationalTitle_Terms", Err.description
End Function

Private Function LoadRawIyMapping_ExcludeAbsolute_NationalTitle() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "terms", LoadRawIyMapping_ExcludeAbsolute_NationalTitle_Terms()
    Set LoadRawIyMapping_ExcludeAbsolute_NationalTitle = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawIyMapping_ExcludeAbsolute_NationalTitle", Err.description
End Function

Private Function LoadRawIyMapping_ExcludeAbsolute_StartsWithQu() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "enabled", True
    Set LoadRawIyMapping_ExcludeAbsolute_StartsWithQu = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawIyMapping_ExcludeAbsolute_StartsWithQu", Err.description
End Function

Private Function LoadRawIyMapping_ExcludeAbsolute_ProperNouns_Examples() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "L" & ChrW(&HFD)
    c.Add "M" & ChrW(&H1EF9)
    c.Add "Qu" & ChrW(&HFD)
    c.Add "K" & ChrW(&H1EF3)
    c.Add "Hy"
    c.Add "S" & ChrW(&H1EF9)
    Set LoadRawIyMapping_ExcludeAbsolute_ProperNouns_Examples = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawIyMapping_ExcludeAbsolute_ProperNouns_Examples", Err.description
End Function

Private Function LoadRawIyMapping_ExcludeAbsolute_ProperNouns() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "skipCapitalizedMidSentence", True
    d.Add "examples", LoadRawIyMapping_ExcludeAbsolute_ProperNouns_Examples()
    Set LoadRawIyMapping_ExcludeAbsolute_ProperNouns = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawIyMapping_ExcludeAbsolute_ProperNouns", Err.description
End Function

Private Function LoadRawIyMapping_ExcludeAbsolute_InsideQuotes() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "enabled", True
    Set LoadRawIyMapping_ExcludeAbsolute_InsideQuotes = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawIyMapping_ExcludeAbsolute_InsideQuotes", Err.description
End Function

Private Function LoadRawIyMapping_ExcludeAbsolute_ComponentRoles_Roles() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "nationalTitle"
    c.Add "nationalMotto"
    c.Add "superiorOrganName"
    c.Add "organName"
    c.Add "signerName"
    c.Add "placeName"
    c.Add "codeNumberNotation"
    Set LoadRawIyMapping_ExcludeAbsolute_ComponentRoles_Roles = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawIyMapping_ExcludeAbsolute_ComponentRoles_Roles", Err.description
End Function

Private Function LoadRawIyMapping_ExcludeAbsolute_ComponentRoles() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "roles", LoadRawIyMapping_ExcludeAbsolute_ComponentRoles_Roles()
    Set LoadRawIyMapping_ExcludeAbsolute_ComponentRoles = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawIyMapping_ExcludeAbsolute_ComponentRoles", Err.description
End Function

Private Function LoadRawIyMapping_ExcludeAbsolute() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "documentTypeNames", LoadRawIyMapping_ExcludeAbsolute_DocumentTypeNames()
    d.Add "typeAbbreviations", LoadRawIyMapping_ExcludeAbsolute_TypeAbbreviations()
    d.Add "nd30Terminology", LoadRawIyMapping_ExcludeAbsolute_Nd30Terminology()
    d.Add "nationalTitle", LoadRawIyMapping_ExcludeAbsolute_NationalTitle()
    d.Add "startsWithQu", LoadRawIyMapping_ExcludeAbsolute_StartsWithQu()
    d.Add "properNouns", LoadRawIyMapping_ExcludeAbsolute_ProperNouns()
    d.Add "insideQuotes", LoadRawIyMapping_ExcludeAbsolute_InsideQuotes()
    d.Add "componentRoles", LoadRawIyMapping_ExcludeAbsolute_ComponentRoles()
    Set LoadRawIyMapping_ExcludeAbsolute = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawIyMapping_ExcludeAbsolute", Err.description
End Function

Private Function LoadRawIyMapping_TestCases_MustConvertToI() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim c93 As New Collection
    c93.Add "k" & ChrW(&H1EF7) & " ni" & ChrW(&H1EC7) & "m"
    c93.Add "k" & ChrW(&H1EC9) & " ni" & ChrW(&H1EC7) & "m"
    c.Add c93
    Dim c94 As New Collection
    c94.Add "l" & ChrW(&HFD) & " lu" & ChrW(&H1EAD) & "n"
    c94.Add "l" & ChrW(&HED) & " lu" & ChrW(&H1EAD) & "n"
    c.Add c94
    Dim c95 As New Collection
    c95.Add "t" & ChrW(&H1EF7) & " l" & ChrW(&H1EC7)
    c95.Add "t" & ChrW(&H1EC9) & " l" & ChrW(&H1EC7)
    c.Add c95
    Set LoadRawIyMapping_TestCases_MustConvertToI = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawIyMapping_TestCases_MustConvertToI", Err.description
End Function

Private Function LoadRawIyMapping_TestCases_MustConvertToY() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim c96 As New Collection
    c96.Add "k" & ChrW(&H1EC9) & " ni" & ChrW(&H1EC7) & "m"
    c96.Add "k" & ChrW(&H1EF7) & " ni" & ChrW(&H1EC7) & "m"
    c.Add c96
    Dim c97 As New Collection
    c97.Add "l" & ChrW(&HED) & " lu" & ChrW(&H1EAD) & "n"
    c97.Add "l" & ChrW(&HFD) & " lu" & ChrW(&H1EAD) & "n"
    c.Add c97
    Dim c98 As New Collection
    c98.Add "t" & ChrW(&H1EC9) & " l" & ChrW(&H1EC7)
    c98.Add "t" & ChrW(&H1EF7) & " l" & ChrW(&H1EC7)
    c.Add c98
    Set LoadRawIyMapping_TestCases_MustConvertToY = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawIyMapping_TestCases_MustConvertToY", Err.description
End Function

Private Function LoadRawIyMapping_TestCases_MustNeverChange() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "Quy ch" & ChrW(&H1EBF)
    c.Add "Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh"
    c.Add "Quy" & ChrW(&H1EBF) & "t " & ChrW(&H111) & ChrW(&H1ECB) & "nh"
    c.Add "Q" & ChrW(&H110) & "-BNV"
    c.Add "k" & ChrW(&H1EF9) & " thu" & ChrW(&H1EAD) & "t tr" & ChrW(&HEC) & "nh b" & ChrW(&HE0) & "y"
    c.Add "k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n"
    c.Add "ch" & ChrW(&H1EEF) & " k" & ChrW(&HFD) & " s" & ChrW(&H1ED1)
    c.Add "C" & ChrW(&H1ED8) & "NG H" & ChrW(&HD2) & "A X" & ChrW(&HC3) & " H" & ChrW(&H1ED8) & "I CH" & ChrW(&H1EE6) & " NGH" & ChrW(&H128) & "A VI" & ChrW(&H1EC6) & "T NAM"
    c.Add "Nguy" & ChrW(&H1EC5) & "n V" & ChrW(&H103) & "n L" & ChrW(&HFD)
    c.Add "t" & ChrW(&H1EC9) & "nh Qu" & ChrW(&H1EA3) & "ng Ng" & ChrW(&HE3) & "i"
    c.Add "quy" & ChrW(&H1EC1) & "n h" & ChrW(&H1EA1) & "n c" & ChrW(&H1EE7) & "a ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i k" & ChrW(&HFD)
    Set LoadRawIyMapping_TestCases_MustNeverChange = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawIyMapping_TestCases_MustNeverChange", Err.description
End Function

Private Function LoadRawIyMapping_TestCases() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "mustConvertToI", LoadRawIyMapping_TestCases_MustConvertToI()
    d.Add "mustConvertToY", LoadRawIyMapping_TestCases_MustConvertToY()
    d.Add "mustNeverChange", LoadRawIyMapping_TestCases_MustNeverChange()
    Set LoadRawIyMapping_TestCases = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawIyMapping_TestCases", Err.description
End Function

Public Function LoadRawIyMapping() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "1.0.0"
    d.Add "sourceLabel", "QD1989"
    d.Add "ruleCode", "QD1989-IY-MIX"
    d.Add "actionType", "B"
    d.Add "pairs", LoadRawIyMapping_Pairs()
    d.Add "excludeAbsolute", LoadRawIyMapping_ExcludeAbsolute()
    d.Add "testCases", LoadRawIyMapping_TestCases()
    Set LoadRawIyMapping = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawIyMapping", Err.description
End Function

Private Sub LoadRawLunarYears_YearsPart1(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Gi" & ChrW(&HE1) & "p T" & ChrW(&HFD)
    c.Add ChrW(&H1EA4) & "t S" & ChrW(&H1EED) & "u"
    c.Add "B" & ChrW(&HED) & "nh D" & ChrW(&H1EA7) & "n"
    c.Add ChrW(&H110) & "inh M" & ChrW(&HE3) & "o"
    c.Add "M" & ChrW(&H1EAD) & "u Th" & ChrW(&HEC) & "n"
    c.Add "K" & ChrW(&H1EF7) & " T" & ChrW(&H1EF5)
    c.Add "Canh Ng" & ChrW(&H1ECD)
    c.Add "T" & ChrW(&HE2) & "n M" & ChrW(&HF9) & "i"
    c.Add "Nh" & ChrW(&HE2) & "m Th" & ChrW(&HE2) & "n"
    c.Add "Qu" & ChrW(&HFD) & " D" & ChrW(&H1EAD) & "u"
    c.Add "Gi" & ChrW(&HE1) & "p Tu" & ChrW(&H1EA5) & "t"
    c.Add ChrW(&H1EA4) & "t H" & ChrW(&H1EE3) & "i"
    c.Add "B" & ChrW(&HED) & "nh T" & ChrW(&HFD)
    c.Add ChrW(&H110) & "inh S" & ChrW(&H1EED) & "u"
    c.Add "M" & ChrW(&H1EAD) & "u D" & ChrW(&H1EA7) & "n"
    c.Add "K" & ChrW(&H1EF7) & " M" & ChrW(&HE3) & "o"
    c.Add "Canh Th" & ChrW(&HEC) & "n"
    c.Add "T" & ChrW(&HE2) & "n T" & ChrW(&H1EF5)
    c.Add "Nh" & ChrW(&HE2) & "m Ng" & ChrW(&H1ECD)
    c.Add "Qu" & ChrW(&HFD) & " M" & ChrW(&HF9) & "i"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawLunarYears_YearsPart1", Err.description
End Sub

Private Sub LoadRawLunarYears_YearsPart2(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Gi" & ChrW(&HE1) & "p Th" & ChrW(&HE2) & "n"
    c.Add ChrW(&H1EA4) & "t D" & ChrW(&H1EAD) & "u"
    c.Add "B" & ChrW(&HED) & "nh Tu" & ChrW(&H1EA5) & "t"
    c.Add ChrW(&H110) & "inh H" & ChrW(&H1EE3) & "i"
    c.Add "M" & ChrW(&H1EAD) & "u T" & ChrW(&HFD)
    c.Add "K" & ChrW(&H1EF7) & " S" & ChrW(&H1EED) & "u"
    c.Add "Canh D" & ChrW(&H1EA7) & "n"
    c.Add "T" & ChrW(&HE2) & "n M" & ChrW(&HE3) & "o"
    c.Add "Nh" & ChrW(&HE2) & "m Th" & ChrW(&HEC) & "n"
    c.Add "Qu" & ChrW(&HFD) & " T" & ChrW(&H1EF5)
    c.Add "Gi" & ChrW(&HE1) & "p Ng" & ChrW(&H1ECD)
    c.Add ChrW(&H1EA4) & "t M" & ChrW(&HF9) & "i"
    c.Add "B" & ChrW(&HED) & "nh Th" & ChrW(&HE2) & "n"
    c.Add ChrW(&H110) & "inh D" & ChrW(&H1EAD) & "u"
    c.Add "M" & ChrW(&H1EAD) & "u Tu" & ChrW(&H1EA5) & "t"
    c.Add "K" & ChrW(&H1EF7) & " H" & ChrW(&H1EE3) & "i"
    c.Add "Canh T" & ChrW(&HFD)
    c.Add "T" & ChrW(&HE2) & "n S" & ChrW(&H1EED) & "u"
    c.Add "Nh" & ChrW(&HE2) & "m D" & ChrW(&H1EA7) & "n"
    c.Add "Qu" & ChrW(&HFD) & " M" & ChrW(&HE3) & "o"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawLunarYears_YearsPart2", Err.description
End Sub

Private Sub LoadRawLunarYears_YearsPart3(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "Gi" & ChrW(&HE1) & "p Th" & ChrW(&HEC) & "n"
    c.Add ChrW(&H1EA4) & "t T" & ChrW(&H1EF5)
    c.Add "B" & ChrW(&HED) & "nh Ng" & ChrW(&H1ECD)
    c.Add ChrW(&H110) & "inh M" & ChrW(&HF9) & "i"
    c.Add "M" & ChrW(&H1EAD) & "u Th" & ChrW(&HE2) & "n"
    c.Add "K" & ChrW(&H1EF7) & " D" & ChrW(&H1EAD) & "u"
    c.Add "Canh Tu" & ChrW(&H1EA5) & "t"
    c.Add "T" & ChrW(&HE2) & "n H" & ChrW(&H1EE3) & "i"
    c.Add "Nh" & ChrW(&HE2) & "m T" & ChrW(&HFD)
    c.Add "Qu" & ChrW(&HFD) & " S" & ChrW(&H1EED) & "u"
    c.Add "Gi" & ChrW(&HE1) & "p D" & ChrW(&H1EA7) & "n"
    c.Add ChrW(&H1EA4) & "t M" & ChrW(&HE3) & "o"
    c.Add "B" & ChrW(&HED) & "nh Th" & ChrW(&HEC) & "n"
    c.Add ChrW(&H110) & "inh T" & ChrW(&H1EF5)
    c.Add "M" & ChrW(&H1EAD) & "u Ng" & ChrW(&H1ECD)
    c.Add "K" & ChrW(&H1EF7) & " M" & ChrW(&HF9) & "i"
    c.Add "Canh Th" & ChrW(&HE2) & "n"
    c.Add "T" & ChrW(&HE2) & "n D" & ChrW(&H1EAD) & "u"
    c.Add "Nh" & ChrW(&HE2) & "m Tu" & ChrW(&H1EA5) & "t"
    c.Add "Qu" & ChrW(&HFD) & " H" & ChrW(&H1EE3) & "i"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawLunarYears_YearsPart3", Err.description
End Sub

Private Function LoadRawLunarYears_Years() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    LoadRawLunarYears_YearsPart1 c
    LoadRawLunarYears_YearsPart2 c
    LoadRawLunarYears_YearsPart3 c
    Set LoadRawLunarYears_Years = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawLunarYears_Years", Err.description
End Function

Public Function LoadRawLunarYears() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "1.0.0"
    d.Add "status", "day du - chu ky 60 nam co dinh, khong doi"
    d.Add "sourceLabel", "ND30"
    d.Add "actionType", "TU SUA"
    d.Add "years", LoadRawLunarYears_Years()
    Set LoadRawLunarYears = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawLunarYears", Err.description
End Function

Private Function LoadRawHolidays_Entries() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim d99 As Object
    Set d99 = CreateObject("Scripting.Dictionary")
    d99.Add "name", "Qu" & ChrW(&H1ED1) & "c kh" & ChrW(&HE1) & "nh"
    d99.Add "date", "2-9"
    d99.Add "example", "ng" & ChrW(&HE0) & "y Qu" & ChrW(&H1ED1) & "c kh" & ChrW(&HE1) & "nh 2-9"
    c.Add d99
    Dim d100 As Object
    Set d100 = CreateObject("Scripting.Dictionary")
    d100.Add "name", "Qu" & ChrW(&H1ED1) & "c t" & ChrW(&H1EBF) & " Lao " & ChrW(&H111) & ChrW(&H1ED9) & "ng"
    d100.Add "date", "1-5"
    d100.Add "example", "ng" & ChrW(&HE0) & "y Qu" & ChrW(&H1ED1) & "c t" & ChrW(&H1EBF) & " Lao " & ChrW(&H111) & ChrW(&H1ED9) & "ng 1-5"
    c.Add d100
    Dim d101 As Object
    Set d101 = CreateObject("Scripting.Dictionary")
    d101.Add "name", "Qu" & ChrW(&H1ED1) & "c t" & ChrW(&H1EBF) & " Ph" & ChrW(&H1EE5) & " n" & ChrW(&H1EEF)
    d101.Add "date", "8-3"
    d101.Add "example", "ng" & ChrW(&HE0) & "y Qu" & ChrW(&H1ED1) & "c t" & ChrW(&H1EBF) & " Ph" & ChrW(&H1EE5) & " n" & ChrW(&H1EEF) & " 8-3"
    c.Add d101
    Dim d102 As Object
    Set d102 = CreateObject("Scripting.Dictionary")
    d102.Add "name", "Ph" & ChrW(&H1EE5) & " n" & ChrW(&H1EEF) & " Vi" & ChrW(&H1EC7) & "t Nam"
    d102.Add "date", "20-10"
    d102.Add "example", "ng" & ChrW(&HE0) & "y Ph" & ChrW(&H1EE5) & " n" & ChrW(&H1EEF) & " Vi" & ChrW(&H1EC7) & "t Nam 20-10"
    c.Add d102
    Dim d103 As Object
    Set d103 = CreateObject("Scripting.Dictionary")
    d103.Add "name", "Gi" & ChrW(&H1EA3) & "i ph" & ChrW(&HF3) & "ng mi" & ChrW(&H1EC1) & "n Nam"
    d103.Add "date", "30-4"
    d103.Add "example", "ng" & ChrW(&HE0) & "y Gi" & ChrW(&H1EA3) & "i ph" & ChrW(&HF3) & "ng mi" & ChrW(&H1EC1) & "n Nam 30-4"
    c.Add d103
    Dim d104 As Object
    Set d104 = CreateObject("Scripting.Dictionary")
    d104.Add "name", "Chi" & ChrW(&H1EBF) & "n th" & ChrW(&H1EAF) & "ng " & ChrW(&H110) & "i" & ChrW(&H1EC7) & "n Bi" & ChrW(&HEA) & "n Ph" & ChrW(&H1EE7)
    d104.Add "date", "7-5"
    d104.Add "example", "ng" & ChrW(&HE0) & "y Chi" & ChrW(&H1EBF) & "n th" & ChrW(&H1EAF) & "ng " & ChrW(&H110) & "i" & ChrW(&H1EC7) & "n Bi" & ChrW(&HEA) & "n Ph" & ChrW(&H1EE7) & " 7-5"
    c.Add d104
    Dim d105 As Object
    Set d105 = CreateObject("Scripting.Dictionary")
    d105.Add "name", "Th" & ChrW(&H1B0) & ChrW(&H1A1) & "ng binh Li" & ChrW(&H1EC7) & "t s" & ChrW(&H129)
    d105.Add "date", "27-7"
    d105.Add "example", "ng" & ChrW(&HE0) & "y Th" & ChrW(&H1B0) & ChrW(&H1A1) & "ng binh Li" & ChrW(&H1EC7) & "t s" & ChrW(&H129) & " 27-7"
    c.Add d105
    Dim d106 As Object
    Set d106 = CreateObject("Scripting.Dictionary")
    d106.Add "name", "Nh" & ChrW(&HE0) & " gi" & ChrW(&HE1) & "o Vi" & ChrW(&H1EC7) & "t Nam"
    d106.Add "date", "20-11"
    d106.Add "example", "ng" & ChrW(&HE0) & "y Nh" & ChrW(&HE0) & " gi" & ChrW(&HE1) & "o Vi" & ChrW(&H1EC7) & "t Nam 20-11"
    c.Add d106
    Dim d107 As Object
    Set d107 = CreateObject("Scripting.Dictionary")
    d107.Add "name", "Qu" & ChrW(&H1ED1) & "c t" & ChrW(&H1EBF) & " Thi" & ChrW(&H1EBF) & "u nhi"
    d107.Add "date", "1-6"
    d107.Add "example", "ng" & ChrW(&HE0) & "y Qu" & ChrW(&H1ED1) & "c t" & ChrW(&H1EBF) & " Thi" & ChrW(&H1EBF) & "u nhi 1-6"
    c.Add d107
    Dim d108 As Object
    Set d108 = CreateObject("Scripting.Dictionary")
    d108.Add "name", "B" & ChrW(&HE1) & "o ch" & ChrW(&HED) & " C" & ChrW(&HE1) & "ch m" & ChrW(&H1EA1) & "ng Vi" & ChrW(&H1EC7) & "t Nam"
    d108.Add "date", "21-6"
    d108.Add "example", "ng" & ChrW(&HE0) & "y B" & ChrW(&HE1) & "o ch" & ChrW(&HED) & " C" & ChrW(&HE1) & "ch m" & ChrW(&H1EA1) & "ng Vi" & ChrW(&H1EC7) & "t Nam 21-6"
    c.Add d108
    Dim d109 As Object
    Set d109 = CreateObject("Scripting.Dictionary")
    d109.Add "name", "Gia " & ChrW(&H111) & ChrW(&HEC) & "nh Vi" & ChrW(&H1EC7) & "t Nam"
    d109.Add "date", "28-6"
    d109.Add "example", "ng" & ChrW(&HE0) & "y Gia " & ChrW(&H111) & ChrW(&HEC) & "nh Vi" & ChrW(&H1EC7) & "t Nam 28-6"
    c.Add d109
    Dim d110 As Object
    Set d110 = CreateObject("Scripting.Dictionary")
    d110.Add "name", "Th" & ChrW(&H1EA7) & "y thu" & ChrW(&H1ED1) & "c Vi" & ChrW(&H1EC7) & "t Nam"
    d110.Add "date", "27-2"
    d110.Add "example", "ng" & ChrW(&HE0) & "y Th" & ChrW(&H1EA7) & "y thu" & ChrW(&H1ED1) & "c Vi" & ChrW(&H1EC7) & "t Nam 27-2"
    c.Add d110
    Dim d111 As Object
    Set d111 = CreateObject("Scripting.Dictionary")
    d111.Add "name", "Ph" & ChrW(&HE1) & "p lu" & ChrW(&H1EAD) & "t Vi" & ChrW(&H1EC7) & "t Nam"
    d111.Add "date", "9-11"
    d111.Add "example", "ng" & ChrW(&HE0) & "y Ph" & ChrW(&HE1) & "p lu" & ChrW(&H1EAD) & "t Vi" & ChrW(&H1EC7) & "t Nam 9-11"
    c.Add d111
    Dim d112 As Object
    Set d112 = CreateObject("Scripting.Dictionary")
    d112.Add "name", "Doanh nh" & ChrW(&HE2) & "n Vi" & ChrW(&H1EC7) & "t Nam"
    d112.Add "date", "13-10"
    d112.Add "example", "ng" & ChrW(&HE0) & "y Doanh nh" & ChrW(&HE2) & "n Vi" & ChrW(&H1EC7) & "t Nam 13-10"
    c.Add d112
    Set LoadRawHolidays_Entries = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawHolidays_Entries", Err.description
End Function

Public Function LoadRawHolidays() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "1.0.0"
    d.Add "status", "hat giong - cac ngay le pho bien nhat, chua day du"
    d.Add "sourceLabel", "ND30"
    d.Add "actionType", "TU SUA"
    d.Add "entries", LoadRawHolidays_Entries()
    Set LoadRawHolidays = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawHolidays", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_ND30_FontSizeSetKeys() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "set1"
    c.Add "set2"
    Set LoadRawRegimeConfig_Regimes_ND30_FontSizeSetKeys = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_ND30_FontSizeSetKeys", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_ND30_LineSpacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bodyZoneRule", "single"
    Set LoadRawRegimeConfig_Regimes_ND30_LineSpacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_ND30_LineSpacing", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_ND30() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "label", "VB h" & ChrW(&HE0) & "nh ch" & ChrW(&HED) & "nh (N" & ChrW(&H110) & " 30)"
    d.Add "fontSizeSetKeys", LoadRawRegimeConfig_Regimes_ND30_FontSizeSetKeys()
    d.Add "lineSpacing", LoadRawRegimeConfig_Regimes_ND30_LineSpacing()
    d.Add "contentEndMark", "./."
    Set LoadRawRegimeConfig_Regimes_ND30 = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_ND30", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_VIETTEL_FontSizeSetKeys() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "set1"
    Set LoadRawRegimeConfig_Regimes_VIETTEL_FontSizeSetKeys = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_VIETTEL_FontSizeSetKeys", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_VIETTEL_LineSpacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bodyZoneRule", "single"
    Set LoadRawRegimeConfig_Regimes_VIETTEL_LineSpacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_VIETTEL_LineSpacing", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_VIETTEL_FontSizeOverrides_Set1() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "nationalMotto", 13
    d.Add "placeAndIssuedDate", 13
    d.Add "subjectOfficialLetter", 12
    Set LoadRawRegimeConfig_Regimes_VIETTEL_FontSizeOverrides_Set1 = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_VIETTEL_FontSizeOverrides_Set1", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_VIETTEL_FontSizeOverrides() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "set1", LoadRawRegimeConfig_Regimes_VIETTEL_FontSizeOverrides_Set1()
    Set LoadRawRegimeConfig_Regimes_VIETTEL_FontSizeOverrides = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_VIETTEL_FontSizeOverrides", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis_StyleByDocumentTypeGroup_Decree_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", True
    d.Add "underline", False
    Set LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis_StyleByDocumentTypeGroup_Decree_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis_StyleByDocumentTypeGroup_Decree_Style", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis_StyleByDocumentTypeGroup_Decree() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "style", LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis_StyleByDocumentTypeGroup_Decree_Style()
    Set LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis_StyleByDocumentTypeGroup_Decree = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis_StyleByDocumentTypeGroup_Decree", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis_StyleByDocumentTypeGroup_Other_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis_StyleByDocumentTypeGroup_Other_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis_StyleByDocumentTypeGroup_Other_Style", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis_StyleByDocumentTypeGroup_Other() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "style", LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis_StyleByDocumentTypeGroup_Other_Style()
    d.Add "lastLineEndChar", ","
    Set LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis_StyleByDocumentTypeGroup_Other = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis_StyleByDocumentTypeGroup_Other", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis_StyleByDocumentTypeGroup() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "decree", LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis_StyleByDocumentTypeGroup_Decree()
    d.Add "other", LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis_StyleByDocumentTypeGroup_Other()
    Set LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis_StyleByDocumentTypeGroup = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis_StyleByDocumentTypeGroup", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "styleByDocumentTypeGroup", LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis_StyleByDocumentTypeGroup()
    Set LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_RecipientList() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "archiveLinePattern", "^-?\s*L" & ChrW(&H1B0) & "u\s*:\s*VT\s*,\s*[^\s.]+\.[^.]+\.$"
    Set LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_RecipientList = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_RecipientList", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "legalBasis", LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_LegalBasis()
    d.Add "recipientList", LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides_RecipientList()
    Set LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_VIETTEL() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "label", "Viettel (Q" & ChrW(&H110) & " 11095)"
    d.Add "fontSizeSetKeys", LoadRawRegimeConfig_Regimes_VIETTEL_FontSizeSetKeys()
    d.Add "lineSpacing", LoadRawRegimeConfig_Regimes_VIETTEL_LineSpacing()
    d.Add "contentEndMark", "./."
    d.Add "fontSizeOverrides", LoadRawRegimeConfig_Regimes_VIETTEL_FontSizeOverrides()
    d.Add "componentOverrides", LoadRawRegimeConfig_Regimes_VIETTEL_ComponentOverrides()
    Set LoadRawRegimeConfig_Regimes_VIETTEL = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_VIETTEL", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_FontSizeSetKeys() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "set1"
    c.Add "set3"
    Set LoadRawRegimeConfig_Regimes_DANG_FontSizeSetKeys = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_FontSizeSetKeys", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_Margins_TopMm() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 20
    d.Add "max", 20
    d.Add "default", 20
    Set LoadRawRegimeConfig_Regimes_DANG_Margins_TopMm = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_Margins_TopMm", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_Margins_BottomMm() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 20
    d.Add "max", 20
    d.Add "default", 20
    Set LoadRawRegimeConfig_Regimes_DANG_Margins_BottomMm = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_Margins_BottomMm", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_Margins_LeftMm() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 30
    d.Add "max", 30
    d.Add "default", 30
    Set LoadRawRegimeConfig_Regimes_DANG_Margins_LeftMm = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_Margins_LeftMm", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_Margins_RightMm() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 15
    d.Add "max", 15
    d.Add "default", 15
    Set LoadRawRegimeConfig_Regimes_DANG_Margins_RightMm = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_Margins_RightMm", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_Margins() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "topMm", LoadRawRegimeConfig_Regimes_DANG_Margins_TopMm()
    d.Add "bottomMm", LoadRawRegimeConfig_Regimes_DANG_Margins_BottomMm()
    d.Add "leftMm", LoadRawRegimeConfig_Regimes_DANG_Margins_LeftMm()
    d.Add "rightMm", LoadRawRegimeConfig_Regimes_DANG_Margins_RightMm()
    Set LoadRawRegimeConfig_Regimes_DANG_Margins = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_Margins", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_LineSpacing_BodyZoneExactlyPtBySizeSet() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "set1", 18
    d.Add "set3", 22
    Set LoadRawRegimeConfig_Regimes_DANG_LineSpacing_BodyZoneExactlyPtBySizeSet = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_LineSpacing_BodyZoneExactlyPtBySizeSet", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_LineSpacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bodyZoneRule", "exactly"
    d.Add "bodyZoneExactlyPtBySizeSet", LoadRawRegimeConfig_Regimes_DANG_LineSpacing_BodyZoneExactlyPtBySizeSet()
    Set LoadRawRegimeConfig_Regimes_DANG_LineSpacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_LineSpacing", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_FontSizeOverrides_Set1() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "partyHeader", 15
    d.Add "superiorOrganName", 14
    d.Add "organName", 14
    d.Add "codeNumberNotation", 14
    d.Add "placeAndIssuedDate", 14
    d.Add "subjectOfficialLetter", 12
    d.Add "typeName", 15
    d.Add "subject", 14
    d.Add "recipientSalutation", 14
    d.Add "recipientSalutationList", 14
    d.Add "recipientSalutationInline", 14
    d.Add "recipientSalutationInlineContent", 14
    d.Add "legalBasis", 14
    d.Add "bodyText", 14
    d.Add "recipientLabel", 14
    d.Add "recipientList", 12
    d.Add "appendixLabel", 14
    d.Add "appendixTitle", 14
    d.Add "appendixReference", 14
    d.Add "pageNumber", 14
    Set LoadRawRegimeConfig_Regimes_DANG_FontSizeOverrides_Set1 = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_FontSizeOverrides_Set1", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_FontSizeOverrides() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "set1", LoadRawRegimeConfig_Regimes_DANG_FontSizeOverrides_Set1()
    Set LoadRawRegimeConfig_Regimes_DANG_FontSizeOverrides = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_FontSizeOverrides", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_TypeName() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "firstLineIndentCm", 1
    Set LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_TypeName = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_TypeName", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_Subject() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "firstLineIndentCm", 1
    Set LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_Subject = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_Subject", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutation_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", True
    d.Add "underline", False
    Set LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutation_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutation_Style", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutation() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "style", LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutation_Style()
    Set LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutation = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutation", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutationInline_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", True
    d.Add "underline", False
    Set LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutationInline_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutationInline_Style", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutationInline() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "style", LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutationInline_Style()
    Set LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutationInline = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutationInline", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutationInlineContent_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", True
    d.Add "underline", False
    Set LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutationInlineContent_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutationInlineContent_Style", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutationInlineContent() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "style", LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutationInlineContent_Style()
    Set LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutationInlineContent = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutationInlineContent", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_LegalBasis_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_LegalBasis_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_LegalBasis_Style", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_LegalBasis() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "style", LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_LegalBasis_Style()
    d.Add "bulletChar", "-"
    Set LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_LegalBasis = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_LegalBasis", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientLabel_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", False
    d.Add "underline", True
    Set LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientLabel_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientLabel_Style", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientLabel() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "style", LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientLabel_Style()
    Set LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientLabel = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientLabel", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_SignerAuthorityTitle_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_SignerAuthorityTitle_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_SignerAuthorityTitle_Style", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_SignerAuthorityTitle() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "style", LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_SignerAuthorityTitle_Style()
    Set LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_SignerAuthorityTitle = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_SignerAuthorityTitle", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientList() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "archiveLinePattern", "^-?\s*L" & ChrW(&H1B0) & "u\s+\S+\.\S+\-\d+\.?$"
    Set LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientList = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientList", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "typeName", LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_TypeName()
    d.Add "subject", LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_Subject()
    d.Add "recipientSalutation", LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutation()
    d.Add "recipientSalutationInline", LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutationInline()
    d.Add "recipientSalutationInlineContent", LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientSalutationInlineContent()
    d.Add "legalBasis", LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_LegalBasis()
    d.Add "recipientLabel", LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientLabel()
    d.Add "signerAuthorityTitle", LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_SignerAuthorityTitle()
    d.Add "recipientList", LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides_RecipientList()
    Set LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes_DANG() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "label", "VB c" & ChrW(&H1EE7) & "a " & ChrW(&H110) & ChrW(&H1EA3) & "ng (HD 05)"
    d.Add "fontSizeSetKeys", LoadRawRegimeConfig_Regimes_DANG_FontSizeSetKeys()
    d.Add "contentEndMark", "."
    d.Add "margins", LoadRawRegimeConfig_Regimes_DANG_Margins()
    d.Add "lineSpacing", LoadRawRegimeConfig_Regimes_DANG_LineSpacing()
    d.Add "fontSizeOverrides", LoadRawRegimeConfig_Regimes_DANG_FontSizeOverrides()
    d.Add "componentOverrides", LoadRawRegimeConfig_Regimes_DANG_ComponentOverrides()
    Set LoadRawRegimeConfig_Regimes_DANG = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes_DANG", Err.description
End Function

Private Function LoadRawRegimeConfig_Regimes() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "ND30", LoadRawRegimeConfig_Regimes_ND30()
    d.Add "VIETTEL", LoadRawRegimeConfig_Regimes_VIETTEL()
    d.Add "DANG", LoadRawRegimeConfig_Regimes_DANG()
    Set LoadRawRegimeConfig_Regimes = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Regimes", Err.description
End Function

Private Function LoadRawRegimeConfig_DecreeDocumentTypeNames() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "Ngh" & ChrW(&H1ECB) & " quy" & ChrW(&H1EBF) & "t"
    c.Add "Quy" & ChrW(&H1EBF) & "t " & ChrW(&H111) & ChrW(&H1ECB) & "nh"
    c.Add "Quy ch" & ChrW(&H1EBF)
    c.Add "Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh"
    Set LoadRawRegimeConfig_DecreeDocumentTypeNames = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_DecreeDocumentTypeNames", Err.description
End Function

Private Function LoadRawRegimeConfig_Detection_ViettelMarkers() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "T" & ChrW(&H1EAC) & "P " & ChrW(&H110) & "O" & ChrW(&HC0) & "N C" & ChrW(&HD4) & "NG NGHI" & ChrW(&H1EC6) & "P"
    c.Add "VI" & ChrW(&H1EC4) & "N TH" & ChrW(&HD4) & "NG QU" & ChrW(&HC2) & "N " & ChrW(&H110) & ChrW(&H1ED8) & "I"
    c.Add "VIETTEL"
    Set LoadRawRegimeConfig_Detection_ViettelMarkers = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Detection_ViettelMarkers", Err.description
End Function

Private Function LoadRawRegimeConfig_Detection() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "nationalTitleTarget", "C" & ChrW(&H1ED8) & "NG H" & ChrW(&HD2) & "A X" & ChrW(&HC3) & " H" & ChrW(&H1ED8) & "I CH" & ChrW(&H1EE6) & " NGH" & ChrW(&H128) & "A VI" & ChrW(&H1EC6) & "T NAM"
    d.Add "partyHeaderTarget", ChrW(&H110) & ChrW(&H1EA2) & "NG C" & ChrW(&H1ED8) & "NG S" & ChrW(&H1EA2) & "N VI" & ChrW(&H1EC6) & "T NAM"
    d.Add "viettelMarkers", LoadRawRegimeConfig_Detection_ViettelMarkers()
    Set LoadRawRegimeConfig_Detection = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig_Detection", Err.description
End Function

Public Function LoadRawRegimeConfig() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "2.0.0"
    d.Add "sourceLabel", "SUY RA"
    d.Add "regimes", LoadRawRegimeConfig_Regimes()
    d.Add "decreeDocumentTypeNames", LoadRawRegimeConfig_DecreeDocumentTypeNames()
    d.Add "detection", LoadRawRegimeConfig_Detection()
    d.Add "headerWindowChars", 2000
    Set LoadRawRegimeConfig = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegimeConfig", Err.description
End Function

Private Function LoadRawCheckRules_ChecklistGroups() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim d113 As Object
    Set d113 = CreateObject("Scripting.Dictionary")
    d113.Add "id", 1
    d113.Add "label", "Kh" & ChrW(&H1ED5) & " gi" & ChrW(&H1EA5) & "y, l" & ChrW(&H1EC1) & ", h" & ChrW(&H1B0) & ChrW(&H1EDB) & "ng gi" & ChrW(&H1EA5) & "y, ph" & ChrW(&HF4) & "ng, m" & ChrW(&HE0) & "u ch" & ChrW(&H1EEF) & ", s" & ChrW(&H1ED1) & " trang"
    c.Add d113
    Dim d114 As Object
    Set d114 = CreateObject("Scripting.Dictionary")
    d114.Add "id", 2
    d114.Add "label", "Qu" & ChrW(&H1ED1) & "c hi" & ChrW(&H1EC7) & "u v" & ChrW(&HE0) & " Ti" & ChrW(&HEA) & "u ng" & ChrW(&H1EEF)
    c.Add d114
    Dim d115 As Object
    Set d115 = CreateObject("Scripting.Dictionary")
    d115.Add "id", 3
    d115.Add "label", "T" & ChrW(&HEA) & "n c" & ChrW(&H1A1) & " quan, t" & ChrW(&H1ED5) & " ch" & ChrW(&H1EE9) & "c ban h" & ChrW(&HE0) & "nh"
    c.Add d115
    Dim d116 As Object
    Set d116 = CreateObject("Scripting.Dictionary")
    d116.Add "id", 4
    d116.Add "label", "S" & ChrW(&H1ED1) & ", k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u c" & ChrW(&H1EE7) & "a v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n"
    c.Add d116
    Dim d117 As Object
    Set d117 = CreateObject("Scripting.Dictionary")
    d117.Add "id", 5
    d117.Add "label", ChrW(&H110) & ChrW(&H1ECB) & "a danh v" & ChrW(&HE0) & " th" & ChrW(&H1EDD) & "i gian ban h" & ChrW(&HE0) & "nh"
    c.Add d117
    Dim d118 As Object
    Set d118 = CreateObject("Scripting.Dictionary")
    d118.Add "id", 6
    d118.Add "label", "T" & ChrW(&HEA) & "n lo" & ChrW(&H1EA1) & "i v" & ChrW(&HE0) & " tr" & ChrW(&HED) & "ch y" & ChrW(&H1EBF) & "u n" & ChrW(&H1ED9) & "i dung"
    c.Add d118
    Dim d119 As Object
    Set d119 = CreateObject("Scripting.Dictionary")
    d119.Add "id", 7
    d119.Add "label", "N" & ChrW(&H1ED9) & "i dung v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n " & ChrW(&H2014) & " c" & ChrW(&H103) & "n c" & ChrW(&H1EE9) & ", b" & ChrW(&H1ED1) & " c" & ChrW(&H1EE5) & "c, th" & ChrW(&H1EE9) & " t" & ChrW(&H1EF1) & " " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m"
    c.Add d119
    Dim d120 As Object
    Set d120 = CreateObject("Scripting.Dictionary")
    d120.Add "id", 8
    d120.Add "label", "Quy" & ChrW(&H1EC1) & "n h" & ChrW(&H1EA1) & "n, ch" & ChrW(&H1EE9) & "c v" & ChrW(&H1EE5) & ", h" & ChrW(&H1ECD) & " t" & ChrW(&HEA) & "n ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i k" & ChrW(&HFD)
    c.Add d120
    Dim d121 As Object
    Set d121 = CreateObject("Scripting.Dictionary")
    d121.Add "id", 9
    d121.Add "label", "K" & ChrW(&HED) & "nh g" & ChrW(&H1EED) & "i v" & ChrW(&HE0) & " N" & ChrW(&H1A1) & "i nh" & ChrW(&H1EAD) & "n"
    c.Add d121
    Dim d122 As Object
    Set d122 = CreateObject("Scripting.Dictionary")
    d122.Add "id", 10
    d122.Add "label", "D" & ChrW(&H1EA5) & "u ch" & ChrW(&H1EC9) & " d" & ChrW(&H1EAB) & "n, k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u so" & ChrW(&H1EA1) & "n th" & ChrW(&H1EA3) & "o, th" & ChrW(&HF4) & "ng tin li" & ChrW(&HEA) & "n h" & ChrW(&H1EC7)
    c.Add d122
    Dim d123 As Object
    Set d123 = CreateObject("Scripting.Dictionary")
    d123.Add "id", 11
    d123.Add "label", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n"
    c.Add d123
    Dim d124 As Object
    Set d124 = CreateObject("Scripting.Dictionary")
    d124.Add "id", 12
    d124.Add "label", "B" & ChrW(&H1EA3) & "ng bi" & ChrW(&H1EC3) & "u v" & ChrW(&HE0) & " h" & ChrW(&HEC) & "nh " & ChrW(&H1EA3) & "nh"
    c.Add d124
    Dim d125 As Object
    Set d125 = CreateObject("Scripting.Dictionary")
    d125.Add "id", 13
    d125.Add "label", "Ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3) & " " & ChrW(&H2014) & " chuy" & ChrW(&H1EC3) & "n " & ChrW(&H111) & ChrW(&H1ED5) & "i to" & ChrW(&HE0) & "n v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n"
    c.Add d125
    Dim d126 As Object
    Set d126 = CreateObject("Scripting.Dictionary")
    d126.Add "id", 14
    d126.Add "label", "Ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3) & " " & ChrW(&H2014) & " s" & ChrW(&H1EED) & "a t" & ChrW(&H1EEB) & "ng ch" & ChrW(&H1ED7)
    c.Add d126
    Set LoadRawCheckRules_ChecklistGroups = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCheckRules_ChecklistGroups", Err.description
End Function

Private Sub LoadRawCheckRules_RulesPart1(ByRef c As Collection)
    On Error GoTo ErrHandler
    Dim d127 As Object
    Set d127 = CreateObject("Scripting.Dictionary")
    d127.Add "ruleCode", "ND30-PL1-M1-K1"
    d127.Add "checklistGroup", 1
    d127.Add "group", "pageSetup"
    d127.Add "severity", "error"
    d127.Add "sourceLabel", "ND30"
    d127.Add "actionType", "A"
    d127.Add "checkability", "full"
    d127.Add "autoFixable", True
    d127.Add "riskLevel", "low"
    d127.Add "title", "Kh" & ChrW(&H1ED5) & " gi" & ChrW(&H1EA5) & "y kh" & ChrW(&HF4) & "ng ph" & ChrW(&H1EA3) & "i A4"
    d127.Add "message", "Kh" & ChrW(&H1ED5) & " gi" & ChrW(&H1EA5) & "y hi" & ChrW(&H1EC7) & "n t" & ChrW(&H1EA1) & "i {actual}, quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh l" & ChrW(&HE0) & " A4 (210 x 297 mm)."
    d127.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c I, kho" & ChrW(&H1EA3) & "n 1"
    d127.Add "toolHint", "Kh" & ChrW(&H1ED5) & " A4"
    c.Add d127
    Dim d128 As Object
    Set d128 = CreateObject("Scripting.Dictionary")
    d128.Add "ruleCode", "ND30-PL1-M1-K2"
    d128.Add "checklistGroup", 1
    d128.Add "group", "pageSetup"
    d128.Add "severity", "warning"
    d128.Add "sourceLabel", "ND30"
    d128.Add "actionType", "A"
    d128.Add "checkability", "full"
    d128.Add "autoFixable", True
    d128.Add "riskLevel", "low"
    d128.Add "title", "H" & ChrW(&H1B0) & ChrW(&H1EDB) & "ng gi" & ChrW(&H1EA5) & "y n" & ChrW(&H1EB1) & "m ngang"
    d128.Add "message", "Section {index} " & ChrW(&H111) & ChrW(&H1EB7) & "t ngang. Ch" & ChrW(&H1EC9) & " " & ChrW(&H111) & ChrW(&H1B0) & ChrW(&H1EE3) & "c " & ChrW(&H111) & ChrW(&H1EB7) & "t ngang khi n" & ChrW(&H1ED9) & "i dung c" & ChrW(&HF3) & " b" & ChrW(&H1EA3) & "ng, bi" & ChrW(&H1EC3) & "u kh" & ChrW(&HF4) & "ng l" & ChrW(&HE0) & "m th" & ChrW(&HE0) & "nh ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c ri" & ChrW(&HEA) & "ng."
    d128.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c I, kho" & ChrW(&H1EA3) & "n 2"
    d128.Add "toolHint", "Gi" & ChrW(&H1EA5) & "y d" & ChrW(&H1ECD) & "c / ngang"
    c.Add d128
    Dim d129 As Object
    Set d129 = CreateObject("Scripting.Dictionary")
    d129.Add "ruleCode", "ND30-PL1-M1-K3"
    d129.Add "checklistGroup", 1
    d129.Add "group", "pageSetup"
    d129.Add "severity", "error"
    d129.Add "sourceLabel", "ND30"
    d129.Add "actionType", "A"
    d129.Add "checkability", "full"
    d129.Add "autoFixable", True
    d129.Add "riskLevel", "low"
    d129.Add "title", "L" & ChrW(&H1EC1) & " trang ngo" & ChrW(&HE0) & "i d" & ChrW(&H1EA3) & "i cho ph" & ChrW(&HE9) & "p"
    d129.Add "message", "L" & ChrW(&H1EC1) & " {side} l" & ChrW(&HE0) & " {actual} mm, d" & ChrW(&H1EA3) & "i cho ph" & ChrW(&HE9) & "p {min}-{max} mm."
    d129.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c I, kho" & ChrW(&H1EA3) & "n 3"
    d129.Add "toolHint", ChrW(&H110) & ChrW(&H1ECB) & "nh l" & ChrW(&H1EC1)
    c.Add d129
    Dim d130 As Object
    Set d130 = CreateObject("Scripting.Dictionary")
    d130.Add "ruleCode", "ND30-PL1-M1-K4-FONT"
    d130.Add "checklistGroup", 1
    d130.Add "group", "bodyText"
    d130.Add "severity", "error"
    d130.Add "sourceLabel", "ND30"
    d130.Add "actionType", "A"
    d130.Add "checkability", "full"
    d130.Add "autoFixable", True
    d130.Add "riskLevel", "low"
    d130.Add "title", "Ph" & ChrW(&HF4) & "ng ch" & ChrW(&H1EEF) & " kh" & ChrW(&HF4) & "ng ph" & ChrW(&H1EA3) & "i Times New Roman"
    d130.Add "message", "Ph" & ChrW(&HE1) & "t hi" & ChrW(&H1EC7) & "n {count} " & ChrW(&H111) & "o" & ChrW(&H1EA1) & "n d" & ChrW(&HF9) & "ng ph" & ChrW(&HF4) & "ng {fontName}."
    d130.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c I, kho" & ChrW(&H1EA3) & "n 4"
    d130.Add "toolHint", "Ph" & ChrW(&HF4) & "ng chu" & ChrW(&H1EA9) & "n"
    c.Add d130
    Dim d131 As Object
    Set d131 = CreateObject("Scripting.Dictionary")
    d131.Add "ruleCode", "ND30-PL1-M1-K4-ENC"
    d131.Add "checklistGroup", 1
    d131.Add "group", "encoding"
    d131.Add "severity", "error"
    d131.Add "sourceLabel", "ND30"
    d131.Add "actionType", "B"
    d131.Add "checkability", "full"
    d131.Add "autoFixable", True
    d131.Add "riskLevel", "high"
    d131.Add "title", "V" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n d" & ChrW(&HF9) & "ng b" & ChrW(&H1EA3) & "ng m" & ChrW(&HE3) & " kh" & ChrW(&HF4) & "ng ph" & ChrW(&H1EA3) & "i Unicode"
    d131.Add "message", "Ph" & ChrW(&HE1) & "t hi" & ChrW(&H1EC7) & "n {count} " & ChrW(&H111) & "o" & ChrW(&H1EA1) & "n d" & ChrW(&HF9) & "ng ph" & ChrW(&HF4) & "ng {fontName} thu" & ChrW(&H1ED9) & "c b" & ChrW(&H1EA3) & "ng m" & ChrW(&HE3) & " {encoding}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh d" & ChrW(&HF9) & "ng Unicode theo TCVN 6909:2001."
    d131.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c I, kho" & ChrW(&H1EA3) & "n 4"
    d131.Add "toolHint", "Chuy" & ChrW(&H1EC3) & "n b" & ChrW(&H1EA3) & "ng m" & ChrW(&HE3)
    c.Add d131
    Dim d132 As Object
    Set d132 = CreateObject("Scripting.Dictionary")
    d132.Add "ruleCode", "ND30-PL1-M1-K4-NFC"
    d132.Add "checklistGroup", 1
    d132.Add "group", "encoding"
    d132.Add "severity", "warning"
    d132.Add "sourceLabel", "ND30"
    d132.Add "actionType", "B"
    d132.Add "checkability", "full"
    d132.Add "autoFixable", True
    d132.Add "riskLevel", "low"
    d132.Add "title", "K" & ChrW(&HFD) & " t" & ChrW(&H1EF1) & " ch" & ChrW(&H1B0) & "a " & ChrW(&H1EDF) & " d" & ChrW(&H1EA1) & "ng d" & ChrW(&H1EF1) & "ng s" & ChrW(&H1EB5) & "n NFC"
    d132.Add "message", "Ph" & ChrW(&HE1) & "t hi" & ChrW(&H1EC7) & "n {count} k" & ChrW(&HFD) & " t" & ChrW(&H1EF1) & " d" & ChrW(&H1EA1) & "ng t" & ChrW(&H1ED5) & " h" & ChrW(&H1EE3) & "p d" & ChrW(&H1EA5) & "u r" & ChrW(&H1EDD) & "i."
    d132.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c I, kho" & ChrW(&H1EA3) & "n 4"
    d132.Add "toolHint", "Chu" & ChrW(&H1EA9) & "n Unicode"
    c.Add d132
    Dim d133 As Object
    Set d133 = CreateObject("Scripting.Dictionary")
    d133.Add "ruleCode", "ND30-PL1-M1-K4-COLOR"
    d133.Add "checklistGroup", 1
    d133.Add "group", "bodyText"
    d133.Add "severity", "error"
    d133.Add "sourceLabel", "ND30"
    d133.Add "actionType", "A"
    d133.Add "checkability", "full"
    d133.Add "autoFixable", True
    d133.Add "riskLevel", "low"
    d133.Add "title", "M" & ChrW(&HE0) & "u ch" & ChrW(&H1EEF) & " kh" & ChrW(&HF4) & "ng ph" & ChrW(&H1EA3) & "i m" & ChrW(&HE0) & "u " & ChrW(&H111) & "en"
    d133.Add "message", "Ph" & ChrW(&HE1) & "t hi" & ChrW(&H1EC7) & "n {count} " & ChrW(&H111) & "o" & ChrW(&H1EA1) & "n c" & ChrW(&HF3) & " m" & ChrW(&HE0) & "u ch" & ChrW(&H1EEF) & " {color}."
    d133.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c I, kho" & ChrW(&H1EA3) & "n 4"
    d133.Add "toolHint", "Ph" & ChrW(&HF4) & "ng chu" & ChrW(&H1EA9) & "n"
    c.Add d133
    Dim d134 As Object
    Set d134 = CreateObject("Scripting.Dictionary")
    d134.Add "ruleCode", "ND30-PL1-M1-K7"
    d134.Add "checklistGroup", 1
    d134.Add "group", "pageSetup"
    d134.Add "severity", "error"
    d134.Add "sourceLabel", "ND30"
    d134.Add "actionType", "A"
    d134.Add "checkability", "full"
    d134.Add "autoFixable", True
    d134.Add "riskLevel", "low"
    d134.Add "title", "S" & ChrW(&H1ED1) & " trang ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh"
    d134.Add "message", "{detail}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: ch" & ChrW(&H1EEF) & " s" & ChrW(&H1ED1) & " " & ChrW(&H1EA2) & " R" & ChrW(&H1EAD) & "p, b" & ChrW(&H1EAF) & "t " & ChrW(&H111) & ChrW(&H1EA7) & "u t" & ChrW(&H1EEB) & " 1, c" & ChrW(&H1EE1) & " 13-14, ki" & ChrW(&H1EC3) & "u " & ChrW(&H111) & ChrW(&H1EE9) & "ng, canh gi" & ChrW(&H1EEF) & "a trong l" & ChrW(&H1EC1) & " tr" & ChrW(&HEA) & "n, kh" & ChrW(&HF4) & "ng hi" & ChrW(&H1EC3) & "n th" & ChrW(&H1ECB) & " s" & ChrW(&H1ED1) & " trang th" & ChrW(&H1EE9) & " nh" & ChrW(&H1EA5) & "t."
    d134.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c I, kho" & ChrW(&H1EA3) & "n 7"
    d134.Add "toolHint", "S" & ChrW(&H1ED1) & " trang"
    c.Add d134
    Dim d135 As Object
    Set d135 = CreateObject("Scripting.Dictionary")
    d135.Add "ruleCode", "ND30-PL1-MV-CT1"
    d135.Add "checklistGroup", 1
    d135.Add "group", "bodyText"
    d135.Add "severity", "error"
    d135.Add "sourceLabel", "ND30"
    d135.Add "actionType", "A"
    d135.Add "checkability", "partial"
    d135.Add "autoFixable", True
    d135.Add "riskLevel", "low"
    d135.Add "title", "C" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & " gi" & ChrW(&H1EEF) & "a c" & ChrW(&HE1) & "c th" & ChrW(&HE0) & "nh ph" & ChrW(&H1EA7) & "n kh" & ChrW(&HF4) & "ng th" & ChrW(&H1ED1) & "ng nh" & ChrW(&H1EA5) & "t"
    d135.Add "message", "C" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & " hi" & ChrW(&H1EC7) & "n t" & ChrW(&H1EA1) & "i kh" & ChrW(&HF4) & "ng kh" & ChrW(&H1EDB) & "p tr" & ChrW(&H1ECD) & "n v" & ChrW(&H1EB9) & "n B" & ChrW(&H1ED9) & " 1 hay B" & ChrW(&H1ED9) & " 2. L" & ChrW(&H1EC7) & "ch t" & ChrW(&H1EA1) & "i: {detail}."
    d135.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c V, ch" & ChrW(&HFA) & " th" & ChrW(&HED) & "ch 1"
    d135.Add "toolHint", "B" & ChrW(&H1ED9) & " c" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF)
    c.Add d135
    Dim d136 As Object
    Set d136 = CreateObject("Scripting.Dictionary")
    d136.Add "ruleCode", "ND30-PL1-M2-K1-QH"
    d136.Add "checklistGroup", 2
    d136.Add "group", "component"
    d136.Add "severity", "error"
    d136.Add "sourceLabel", "ND30"
    d136.Add "actionType", "A"
    d136.Add "checkability", "partial"
    d136.Add "autoFixable", True
    d136.Add "riskLevel", "low"
    d136.Add "title", "Qu" & ChrW(&H1ED1) & "c hi" & ChrW(&H1EC7) & "u ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    d136.Add "message", "{detail}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: ch" & ChrW(&H1EEF) & " in hoa, c" & ChrW(&H1EE1) & " {expectedSize}, ki" & ChrW(&H1EC3) & "u " & ChrW(&H111) & ChrW(&H1EE9) & "ng " & ChrW(&H111) & ChrW(&H1EAD) & "m, " & ChrW(&H1EDF) & " ph" & ChrW(&HED) & "a tr" & ChrW(&HEA) & "n c" & ChrW(&HF9) & "ng b" & ChrW(&HEA) & "n ph" & ChrW(&H1EA3) & "i trang " & ChrW(&H111) & ChrW(&H1EA7) & "u."
    d136.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 1, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m a"
    d136.Add "toolHint", "Qu" & ChrW(&H1ED1) & "c hi" & ChrW(&H1EC7) & "u " & ChrW(&H2013) & " Ti" & ChrW(&HEA) & "u ng" & ChrW(&H1EEF)
    c.Add d136
    Dim d137 As Object
    Set d137 = CreateObject("Scripting.Dictionary")
    d137.Add "ruleCode", "ND30-PL1-M2-K1-TN"
    d137.Add "checklistGroup", 2
    d137.Add "group", "component"
    d137.Add "severity", "error"
    d137.Add "sourceLabel", "ND30"
    d137.Add "actionType", "A"
    d137.Add "checkability", "partial"
    d137.Add "autoFixable", True
    d137.Add "riskLevel", "low"
    d137.Add "title", "Ti" & ChrW(&HEA) & "u ng" & ChrW(&H1EEF) & " ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    d137.Add "message", "{detail}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: ch" & ChrW(&H1EEF) & " in th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng, c" & ChrW(&H1EE1) & " {expectedSize}, ki" & ChrW(&H1EC3) & "u " & ChrW(&H111) & ChrW(&H1EE9) & "ng " & ChrW(&H111) & ChrW(&H1EAD) & "m, canh gi" & ChrW(&H1EEF) & "a d" & ChrW(&H1B0) & ChrW(&H1EDB) & "i Qu" & ChrW(&H1ED1) & "c hi" & ChrW(&H1EC7) & "u."
    d137.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 1, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    d137.Add "toolHint", "Qu" & ChrW(&H1ED1) & "c hi" & ChrW(&H1EC7) & "u " & ChrW(&H2013) & " Ti" & ChrW(&HEA) & "u ng" & ChrW(&H1EEF)
    c.Add d137
    Dim d138 As Object
    Set d138 = CreateObject("Scripting.Dictionary")
    d138.Add "ruleCode", "ND30-PL1-M2-K1-TN-SEP"
    d138.Add "checklistGroup", 2
    d138.Add "group", "component"
    d138.Add "severity", "error"
    d138.Add "sourceLabel", "ND30"
    d138.Add "actionType", "C"
    d138.Add "checkability", "full"
    d138.Add "autoFixable", True
    d138.Add "riskLevel", "low"
    d138.Add "title", "D" & ChrW(&H1EA5) & "u n" & ChrW(&H1ED1) & "i trong Ti" & ChrW(&HEA) & "u ng" & ChrW(&H1EEF) & " ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng"
    d138.Add "message", "Ti" & ChrW(&HEA) & "u ng" & ChrW(&H1EEF) & " ph" & ChrW(&H1EA3) & "i l" & ChrW(&HE0) & " '" & ChrW(&H110) & ChrW(&H1ED9) & "c l" & ChrW(&H1EAD) & "p - T" & ChrW(&H1EF1) & " do - H" & ChrW(&H1EA1) & "nh ph" & ChrW(&HFA) & "c' v" & ChrW(&H1EDB) & "i g" & ChrW(&H1EA1) & "ch n" & ChrW(&H1ED1) & "i v" & ChrW(&HE0) & " c" & ChrW(&HF3) & " c" & ChrW(&HE1) & "ch ch" & ChrW(&H1EEF) & " hai b" & ChrW(&HEA) & "n."
    d138.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 1, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    d138.Add "toolHint", "Qu" & ChrW(&H1ED1) & "c hi" & ChrW(&H1EC7) & "u " & ChrW(&H2013) & " Ti" & ChrW(&HEA) & "u ng" & ChrW(&H1EEF)
    c.Add d138
    Dim d139 As Object
    Set d139 = CreateObject("Scripting.Dictionary")
    d139.Add "ruleCode", "ND30-PL1-M2-K1-TN-LINE"
    d139.Add "checklistGroup", 2
    d139.Add "group", "component"
    d139.Add "severity", "error"
    d139.Add "sourceLabel", "ND30"
    d139.Add "actionType", "A"
    d139.Add "checkability", "partial"
    d139.Add "autoFixable", True
    d139.Add "riskLevel", "low"
    d139.Add "title", "Thi" & ChrW(&H1EBF) & "u " & ChrW(&H111) & ChrW(&H1B0) & ChrW(&H1EDD) & "ng k" & ChrW(&H1EBB) & " d" & ChrW(&H1B0) & ChrW(&H1EDB) & "i Ti" & ChrW(&HEA) & "u ng" & ChrW(&H1EEF)
    d139.Add "message", "D" & ChrW(&H1B0) & ChrW(&H1EDB) & "i Ti" & ChrW(&HEA) & "u ng" & ChrW(&H1EEF) & " ph" & ChrW(&H1EA3) & "i c" & ChrW(&HF3) & " " & ChrW(&H111) & ChrW(&H1B0) & ChrW(&H1EDD) & "ng k" & ChrW(&H1EBB) & " ngang, n" & ChrW(&HE9) & "t li" & ChrW(&H1EC1) & "n, " & ChrW(&H111) & ChrW(&H1ED9) & " d" & ChrW(&HE0) & "i b" & ChrW(&H1EB1) & "ng " & ChrW(&H111) & ChrW(&H1ED9) & " d" & ChrW(&HE0) & "i d" & ChrW(&HF2) & "ng ch" & ChrW(&H1EEF) & "."
    d139.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 1, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    d139.Add "toolHint", "Qu" & ChrW(&H1ED1) & "c hi" & ChrW(&H1EC7) & "u " & ChrW(&H2013) & " Ti" & ChrW(&HEA) & "u ng" & ChrW(&H1EEF)
    c.Add d139
    Dim d140 As Object
    Set d140 = CreateObject("Scripting.Dictionary")
    d140.Add "ruleCode", "ND30-PL1-M2-K1-C"
    d140.Add "checklistGroup", 2
    d140.Add "group", "component"
    d140.Add "severity", "warning"
    d140.Add "sourceLabel", "ND30"
    d140.Add "actionType", "A"
    d140.Add "checkability", "partial"
    d140.Add "autoFixable", True
    d140.Add "riskLevel", "low"
    d140.Add "title", "Qu" & ChrW(&H1ED1) & "c hi" & ChrW(&H1EC7) & "u v" & ChrW(&HE0) & " Ti" & ChrW(&HEA) & "u ng" & ChrW(&H1EEF) & " kh" & ChrW(&HF4) & "ng c" & ChrW(&HE1) & "ch nhau d" & ChrW(&HF2) & "ng " & ChrW(&H111) & ChrW(&H1A1) & "n"
    d140.Add "message", "Hai d" & ChrW(&HF2) & "ng Qu" & ChrW(&H1ED1) & "c hi" & ChrW(&H1EC7) & "u v" & ChrW(&HE0) & " Ti" & ChrW(&HEA) & "u ng" & ChrW(&H1EEF) & " ph" & ChrW(&H1EA3) & "i tr" & ChrW(&HEC) & "nh b" & ChrW(&HE0) & "y c" & ChrW(&HE1) & "ch nhau d" & ChrW(&HF2) & "ng " & ChrW(&H111) & ChrW(&H1A1) & "n."
    d140.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 1, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m c"
    d140.Add "toolHint", "Qu" & ChrW(&H1ED1) & "c hi" & ChrW(&H1EC7) & "u " & ChrW(&H2013) & " Ti" & ChrW(&HEA) & "u ng" & ChrW(&H1EEF)
    c.Add d140
    Dim d141 As Object
    Set d141 = CreateObject("Scripting.Dictionary")
    d141.Add "ruleCode", "ND30-PL1-M2-K2-SUP"
    d141.Add "checklistGroup", 3
    d141.Add "group", "component"
    d141.Add "severity", "error"
    d141.Add "sourceLabel", "ND30"
    d141.Add "actionType", "A"
    d141.Add "checkability", "partial"
    d141.Add "autoFixable", True
    d141.Add "riskLevel", "low"
    d141.Add "title", "T" & ChrW(&HEA) & "n c" & ChrW(&H1A1) & " quan ch" & ChrW(&H1EE7) & " qu" & ChrW(&H1EA3) & "n ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    d141.Add "message", "{detail}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: ch" & ChrW(&H1EEF) & " in hoa, c" & ChrW(&H1EE1) & " {expectedSize}, ki" & ChrW(&H1EC3) & "u " & ChrW(&H111) & ChrW(&H1EE9) & "ng, kh" & ChrW(&HF4) & "ng " & ChrW(&H111) & ChrW(&H1EAD) & "m."
    d141.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 2, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    d141.Add "toolHint", "T" & ChrW(&HEA) & "n c" & ChrW(&H1A1) & " quan"
    c.Add d141
    Dim d142 As Object
    Set d142 = CreateObject("Scripting.Dictionary")
    d142.Add "ruleCode", "ND30-PL1-M2-K2-ORG"
    d142.Add "checklistGroup", 3
    d142.Add "group", "component"
    d142.Add "severity", "error"
    d142.Add "sourceLabel", "ND30"
    d142.Add "actionType", "A"
    d142.Add "checkability", "partial"
    d142.Add "autoFixable", True
    d142.Add "riskLevel", "low"
    d142.Add "title", "T" & ChrW(&HEA) & "n c" & ChrW(&H1A1) & " quan ban h" & ChrW(&HE0) & "nh ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    d142.Add "message", "{detail}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: ch" & ChrW(&H1EEF) & " in hoa, c" & ChrW(&H1EE1) & " {expectedSize}, ki" & ChrW(&H1EC3) & "u " & ChrW(&H111) & ChrW(&H1EE9) & "ng " & ChrW(&H111) & ChrW(&H1EAD) & "m, canh gi" & ChrW(&H1EEF) & "a d" & ChrW(&H1B0) & ChrW(&H1EDB) & "i t" & ChrW(&HEA) & "n c" & ChrW(&H1A1) & " quan ch" & ChrW(&H1EE7) & " qu" & ChrW(&H1EA3) & "n."
    d142.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 2, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    d142.Add "toolHint", "T" & ChrW(&HEA) & "n c" & ChrW(&H1A1) & " quan"
    c.Add d142
    Dim d143 As Object
    Set d143 = CreateObject("Scripting.Dictionary")
    d143.Add "ruleCode", "ND30-PL1-M2-K2-LINE"
    d143.Add "checklistGroup", 3
    d143.Add "group", "component"
    d143.Add "severity", "error"
    d143.Add "sourceLabel", "ND30"
    d143.Add "actionType", "A"
    d143.Add "checkability", "partial"
    d143.Add "autoFixable", True
    d143.Add "riskLevel", "low"
    d143.Add "title", ChrW(&H110) & ChrW(&H1B0) & ChrW(&H1EDD) & "ng k" & ChrW(&H1EBB) & " d" & ChrW(&H1B0) & ChrW(&H1EDB) & "i t" & ChrW(&HEA) & "n c" & ChrW(&H1A1) & " quan ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng " & ChrW(&H111) & ChrW(&H1ED9) & " d" & ChrW(&HE0) & "i"
    d143.Add "message", ChrW(&H110) & ChrW(&H1B0) & ChrW(&H1EDD) & "ng k" & ChrW(&H1EBB) & " ph" & ChrW(&H1EA3) & "i n" & ChrW(&HE9) & "t li" & ChrW(&H1EC1) & "n, " & ChrW(&H111) & ChrW(&H1ED9) & " d" & ChrW(&HE0) & "i b" & ChrW(&H1EB1) & "ng 1/3 " & ChrW(&H111) & ChrW(&H1EBF) & "n 1/2 " & ChrW(&H111) & ChrW(&H1ED9) & " d" & ChrW(&HE0) & "i d" & ChrW(&HF2) & "ng ch" & ChrW(&H1EEF) & ", " & ChrW(&H111) & ChrW(&H1EB7) & "t c" & ChrW(&HE2) & "n " & ChrW(&H111) & ChrW(&H1ED1) & "i so v" & ChrW(&H1EDB) & "i d" & ChrW(&HF2) & "ng ch" & ChrW(&H1EEF) & "."
    d143.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 2, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    d143.Add "toolHint", "T" & ChrW(&HEA) & "n c" & ChrW(&H1A1) & " quan"
    c.Add d143
    Dim d144 As Object
    Set d144 = CreateObject("Scripting.Dictionary")
    d144.Add "ruleCode", "ND30-PL1-M2-K3-PREFIX"
    d144.Add "checklistGroup", 4
    d144.Add "group", "component"
    d144.Add "severity", "error"
    d144.Add "sourceLabel", "ND30"
    d144.Add "actionType", "C"
    d144.Add "checkability", "full"
    d144.Add "autoFixable", True
    d144.Add "riskLevel", "low"
    d144.Add "title", "Thi" & ChrW(&H1EBF) & "u d" & ChrW(&H1EA5) & "u hai ch" & ChrW(&H1EA5) & "m sau t" & ChrW(&H1EEB) & " 'S" & ChrW(&H1ED1) & "'"
    d144.Add "message", "Sau t" & ChrW(&H1EEB) & " 'S" & ChrW(&H1ED1) & "' ph" & ChrW(&H1EA3) & "i c" & ChrW(&HF3) & " d" & ChrW(&H1EA5) & "u hai ch" & ChrW(&H1EA5) & "m."
    d144.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 3, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m c"
    d144.Add "toolHint", "S" & ChrW(&H1ED1) & ", k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u"
    c.Add d144
    Dim d145 As Object
    Set d145 = CreateObject("Scripting.Dictionary")
    d145.Add "ruleCode", "ND30-PL1-M2-K3-PAD"
    d145.Add "checklistGroup", 4
    d145.Add "group", "component"
    d145.Add "severity", "error"
    d145.Add "sourceLabel", "ND30"
    d145.Add "actionType", "C"
    d145.Add "checkability", "full"
    d145.Add "autoFixable", True
    d145.Add "riskLevel", "low"
    d145.Add "title", "S" & ChrW(&H1ED1) & " v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n nh" & ChrW(&H1ECF) & " h" & ChrW(&H1A1) & "n 10 ch" & ChrW(&H1B0) & "a th" & ChrW(&HEA) & "m s" & ChrW(&H1ED1) & " 0"
    d145.Add "message", "'{actual}' ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '{expected}'. V" & ChrW(&H1EDB) & "i nh" & ChrW(&H1EEF) & "ng s" & ChrW(&H1ED1) & " nh" & ChrW(&H1ECF) & " h" & ChrW(&H1A1) & "n 10 ph" & ChrW(&H1EA3) & "i ghi th" & ChrW(&HEA) & "m s" & ChrW(&H1ED1) & " 0 ph" & ChrW(&HED) & "a tr" & ChrW(&H1B0) & ChrW(&H1EDB) & "c."
    d145.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 3, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m c"
    d145.Add "toolHint", "S" & ChrW(&H1ED1) & ", k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u"
    c.Add d145
    Dim d146 As Object
    Set d146 = CreateObject("Scripting.Dictionary")
    d146.Add "ruleCode", "ND30-PL1-M2-K3-SPACE"
    d146.Add "checklistGroup", 4
    d146.Add "group", "component"
    d146.Add "severity", "error"
    d146.Add "sourceLabel", "ND30"
    d146.Add "actionType", "C"
    d146.Add "checkability", "full"
    d146.Add "autoFixable", True
    d146.Add "riskLevel", "low"
    d146.Add "title", "S" & ChrW(&H1ED1) & " v" & ChrW(&HE0) & " k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n c" & ChrW(&HF3) & " d" & ChrW(&H1EA5) & "u c" & ChrW(&HE1) & "ch"
    d146.Add "message", "'{actual}' ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '{expected}'. Gi" & ChrW(&H1EEF) & "a c" & ChrW(&HE1) & "c nh" & ChrW(&HF3) & "m ch" & ChrW(&H1EEF) & " vi" & ChrW(&H1EBF) & "t t" & ChrW(&H1EAF) & "t kh" & ChrW(&HF4) & "ng c" & ChrW(&HE1) & "ch ch" & ChrW(&H1EEF) & "."
    d146.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 3, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m c"
    d146.Add "toolHint", "S" & ChrW(&H1ED1) & ", k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u"
    c.Add d146
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCheckRules_RulesPart1", Err.description
End Sub

Private Sub LoadRawCheckRules_RulesPart2(ByRef c As Collection)
    On Error GoTo ErrHandler
    Dim d147 As Object
    Set d147 = CreateObject("Scripting.Dictionary")
    d147.Add "ruleCode", "ND30-PL1-M2-K3-SEP"
    d147.Add "checklistGroup", 4
    d147.Add "group", "component"
    d147.Add "severity", "error"
    d147.Add "sourceLabel", "ND30"
    d147.Add "actionType", "C"
    d147.Add "checkability", "full"
    d147.Add "autoFixable", True
    d147.Add "riskLevel", "low"
    d147.Add "title", "D" & ChrW(&H1EA5) & "u ph" & ChrW(&HE2) & "n c" & ChrW(&HE1) & "ch trong s" & ChrW(&H1ED1) & ", k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng"
    d147.Add "message", "Gi" & ChrW(&H1EEF) & "a s" & ChrW(&H1ED1) & " v" & ChrW(&HE0) & " k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u ph" & ChrW(&H1EA3) & "i l" & ChrW(&HE0) & " d" & ChrW(&H1EA5) & "u g" & ChrW(&H1EA1) & "ch ch" & ChrW(&HE9) & "o, gi" & ChrW(&H1EEF) & "a c" & ChrW(&HE1) & "c nh" & ChrW(&HF3) & "m ch" & ChrW(&H1EEF) & " vi" & ChrW(&H1EBF) & "t t" & ChrW(&H1EAF) & "t ph" & ChrW(&H1EA3) & "i l" & ChrW(&HE0) & " d" & ChrW(&H1EA5) & "u g" & ChrW(&H1EA1) & "ch n" & ChrW(&H1ED1) & "i."
    d147.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 3, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m c"
    d147.Add "toolHint", "S" & ChrW(&H1ED1) & ", k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u"
    c.Add d147
    Dim d148 As Object
    Set d148 = CreateObject("Scripting.Dictionary")
    d148.Add "ruleCode", "ND30-PL1-M2-K3-CASE"
    d148.Add "checklistGroup", 4
    d148.Add "group", "component"
    d148.Add "severity", "error"
    d148.Add "sourceLabel", "ND30"
    d148.Add "actionType", "A"
    d148.Add "checkability", "partial"
    d148.Add "autoFixable", True
    d148.Add "riskLevel", "low"
    d148.Add "title", "K" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n ch" & ChrW(&H1B0) & "a vi" & ChrW(&H1EBF) & "t in hoa"
    d148.Add "message", "K" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u c" & ChrW(&H1EE7) & "a v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n tr" & ChrW(&HEC) & "nh b" & ChrW(&HE0) & "y b" & ChrW(&H1EB1) & "ng ch" & ChrW(&H1EEF) & " in hoa, c" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & " 13, ki" & ChrW(&H1EC3) & "u ch" & ChrW(&H1EEF) & " " & ChrW(&H111) & ChrW(&H1EE9) & "ng."
    d148.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 3, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m c"
    d148.Add "toolHint", "S" & ChrW(&H1ED1) & ", k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u"
    c.Add d148
    Dim d149 As Object
    Set d149 = CreateObject("Scripting.Dictionary")
    d149.Add "ruleCode", "ND30-PL1-M2-K3-ABBR"
    d149.Add "checklistGroup", 4
    d149.Add "group", "component"
    d149.Add "severity", "warning"
    d149.Add "sourceLabel", "ND30"
    d149.Add "actionType", "C"
    d149.Add "checkability", "partial"
    d149.Add "autoFixable", False
    d149.Add "riskLevel", "low"
    d149.Add "title", "Ch" & ChrW(&H1EEF) & " vi" & ChrW(&H1EBF) & "t t" & ChrW(&H1EAF) & "t t" & ChrW(&HEA) & "n lo" & ChrW(&H1EA1) & "i v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n kh" & ChrW(&HF4) & "ng kh" & ChrW(&H1EDB) & "p b" & ChrW(&H1EA3) & "ng quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh"
    d149.Add "message", "K" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u ch" & ChrW(&H1EE9) & "a '{actual}'. B" & ChrW(&H1EA3) & "ng ch" & ChrW(&H1EEF) & " vi" & ChrW(&H1EBF) & "t t" & ChrW(&H1EAF) & "t t" & ChrW(&HEA) & "n lo" & ChrW(&H1EA1) & "i quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh '{expected}' cho {typeName}."
    d149.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c III, M" & ChrW(&H1EE5) & "c I"
    d149.Add "toolHint", "S" & ChrW(&H1ED1) & ", k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u"
    c.Add d149
    Dim d150 As Object
    Set d150 = CreateObject("Scripting.Dictionary")
    d150.Add "ruleCode", "ND30-PL1-M2-K4-STYLE"
    d150.Add "checklistGroup", 5
    d150.Add "group", "component"
    d150.Add "severity", "error"
    d150.Add "sourceLabel", "ND30"
    d150.Add "actionType", "A"
    d150.Add "checkability", "partial"
    d150.Add "autoFixable", True
    d150.Add "riskLevel", "low"
    d150.Add "title", ChrW(&H110) & ChrW(&H1ECB) & "a danh v" & ChrW(&HE0) & " th" & ChrW(&H1EDD) & "i gian ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    d150.Add "message", "{detail}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: ch" & ChrW(&H1EEF) & " in th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng, c" & ChrW(&H1EE1) & " {expectedSize}, ki" & ChrW(&H1EC3) & "u ch" & ChrW(&H1EEF) & " nghi" & ChrW(&HEA) & "ng."
    d150.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 4, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m c"
    d150.Add "toolHint", ChrW(&H110) & ChrW(&H1ECB) & "a danh, ng" & ChrW(&HE0) & "y th" & ChrW(&HE1) & "ng"
    c.Add d150
    Dim d151 As Object
    Set d151 = CreateObject("Scripting.Dictionary")
    d151.Add "ruleCode", "ND30-PL1-M2-K4-COMMA"
    d151.Add "checklistGroup", 5
    d151.Add "group", "component"
    d151.Add "severity", "error"
    d151.Add "sourceLabel", "ND30"
    d151.Add "actionType", "C"
    d151.Add "checkability", "full"
    d151.Add "autoFixable", True
    d151.Add "riskLevel", "low"
    d151.Add "title", "Thi" & ChrW(&H1EBF) & "u d" & ChrW(&H1EA5) & "u ph" & ChrW(&H1EA9) & "y sau " & ChrW(&H111) & ChrW(&H1ECB) & "a danh"
    d151.Add "message", "Sau " & ChrW(&H111) & ChrW(&H1ECB) & "a danh ph" & ChrW(&H1EA3) & "i c" & ChrW(&HF3) & " d" & ChrW(&H1EA5) & "u ph" & ChrW(&H1EA9) & "y."
    d151.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 4, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m c"
    d151.Add "toolHint", ChrW(&H110) & ChrW(&H1ECB) & "a danh, ng" & ChrW(&HE0) & "y th" & ChrW(&HE1) & "ng"
    c.Add d151
    Dim d152 As Object
    Set d152 = CreateObject("Scripting.Dictionary")
    d152.Add "ruleCode", "ND30-PL1-M2-K4-PAD"
    d152.Add "checklistGroup", 5
    d152.Add "group", "component"
    d152.Add "severity", "error"
    d152.Add "sourceLabel", "ND30"
    d152.Add "actionType", "C"
    d152.Add "checkability", "full"
    d152.Add "autoFixable", True
    d152.Add "riskLevel", "low"
    d152.Add "title", "Ng" & ChrW(&HE0) & "y th" & ChrW(&HE1) & "ng ch" & ChrW(&H1B0) & "a th" & ChrW(&HEA) & "m s" & ChrW(&H1ED1) & " 0 " & ChrW(&H111) & ChrW(&HFA) & "ng quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh"
    d152.Add "message", "'{actual}' ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '{expected}'. Ng" & ChrW(&HE0) & "y nh" & ChrW(&H1ECF) & " h" & ChrW(&H1A1) & "n 10 v" & ChrW(&HE0) & " th" & ChrW(&HE1) & "ng 1, 2 ph" & ChrW(&H1EA3) & "i ghi th" & ChrW(&HEA) & "m s" & ChrW(&H1ED1) & " 0 ph" & ChrW(&HED) & "a tr" & ChrW(&H1B0) & ChrW(&H1EDB) & "c; th" & ChrW(&HE1) & "ng 3 " & ChrW(&H111) & ChrW(&H1EBF) & "n 12 kh" & ChrW(&HF4) & "ng th" & ChrW(&HEA) & "m."
    d152.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 4, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    d152.Add "toolHint", ChrW(&H110) & ChrW(&H1ECB) & "a danh, ng" & ChrW(&HE0) & "y th" & ChrW(&HE1) & "ng"
    c.Add d152
    Dim d153 As Object
    Set d153 = CreateObject("Scripting.Dictionary")
    d153.Add "ruleCode", "ND30-PL1-M2-K4-CASE"
    d153.Add "checklistGroup", 5
    d153.Add "group", "capitalization"
    d153.Add "severity", "error"
    d153.Add "sourceLabel", "ND30"
    d153.Add "actionType", "C"
    d153.Add "checkability", "full"
    d153.Add "autoFixable", True
    d153.Add "riskLevel", "low"
    d153.Add "title", "Ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u " & ChrW(&H111) & ChrW(&H1ECB) & "a danh ch" & ChrW(&H1B0) & "a vi" & ChrW(&H1EBF) & "t hoa"
    d153.Add "message", "C" & ChrW(&HE1) & "c ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u c" & ChrW(&H1EE7) & "a " & ChrW(&H111) & ChrW(&H1ECB) & "a danh ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t hoa."
    d153.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 4, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m c"
    d153.Add "toolHint", ChrW(&H110) & ChrW(&H1ECB) & "a danh, ng" & ChrW(&HE0) & "y th" & ChrW(&HE1) & "ng"
    c.Add d153
    Dim d154 As Object
    Set d154 = CreateObject("Scripting.Dictionary")
    d154.Add "ruleCode", "ND30-PL1-M2-K5A-TYPE"
    d154.Add "checklistGroup", 6
    d154.Add "group", "component"
    d154.Add "severity", "error"
    d154.Add "sourceLabel", "ND30"
    d154.Add "actionType", "A"
    d154.Add "checkability", "partial"
    d154.Add "autoFixable", True
    d154.Add "riskLevel", "low"
    d154.Add "title", "T" & ChrW(&HEA) & "n lo" & ChrW(&H1EA1) & "i v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    d154.Add "message", "{detail}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: ch" & ChrW(&H1EEF) & " in hoa, c" & ChrW(&H1EE1) & " {expectedSize}, ki" & ChrW(&H1EC3) & "u " & ChrW(&H111) & ChrW(&H1EE9) & "ng " & ChrW(&H111) & ChrW(&H1EAD) & "m, canh gi" & ChrW(&H1EEF) & "a theo chi" & ChrW(&H1EC1) & "u ngang v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n."
    d154.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 5, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    d154.Add "toolHint", "T" & ChrW(&HEA) & "n lo" & ChrW(&H1EA1) & "i, tr" & ChrW(&HED) & "ch y" & ChrW(&H1EBF) & "u"
    c.Add d154
    Dim d155 As Object
    Set d155 = CreateObject("Scripting.Dictionary")
    d155.Add "ruleCode", "ND30-PL1-M2-K5A-SUBJ"
    d155.Add "checklistGroup", 6
    d155.Add "group", "component"
    d155.Add "severity", "error"
    d155.Add "sourceLabel", "ND30"
    d155.Add "actionType", "A"
    d155.Add "checkability", "partial"
    d155.Add "autoFixable", True
    d155.Add "riskLevel", "low"
    d155.Add "title", "Tr" & ChrW(&HED) & "ch y" & ChrW(&H1EBF) & "u n" & ChrW(&H1ED9) & "i dung ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    d155.Add "message", "{detail}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: ch" & ChrW(&H1EEF) & " in th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng, c" & ChrW(&H1EE1) & " {expectedSize}, ki" & ChrW(&H1EC3) & "u " & ChrW(&H111) & ChrW(&H1EE9) & "ng " & ChrW(&H111) & ChrW(&H1EAD) & "m, " & ChrW(&H111) & ChrW(&H1EB7) & "t ngay d" & ChrW(&H1B0) & ChrW(&H1EDB) & "i t" & ChrW(&HEA) & "n lo" & ChrW(&H1EA1) & "i."
    d155.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 5, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    d155.Add "toolHint", "T" & ChrW(&HEA) & "n lo" & ChrW(&H1EA1) & "i, tr" & ChrW(&HED) & "ch y" & ChrW(&H1EBF) & "u"
    c.Add d155
    Dim d156 As Object
    Set d156 = CreateObject("Scripting.Dictionary")
    d156.Add "ruleCode", "ND30-PL1-M2-K5A-LINE"
    d156.Add "checklistGroup", 6
    d156.Add "group", "component"
    d156.Add "severity", "error"
    d156.Add "sourceLabel", "ND30"
    d156.Add "actionType", "A"
    d156.Add "checkability", "partial"
    d156.Add "autoFixable", True
    d156.Add "riskLevel", "low"
    d156.Add "title", "Thi" & ChrW(&H1EBF) & "u " & ChrW(&H111) & ChrW(&H1B0) & ChrW(&H1EDD) & "ng k" & ChrW(&H1EBB) & " d" & ChrW(&H1B0) & ChrW(&H1EDB) & "i tr" & ChrW(&HED) & "ch y" & ChrW(&H1EBF) & "u"
    d156.Add "message", "B" & ChrW(&HEA) & "n d" & ChrW(&H1B0) & ChrW(&H1EDB) & "i tr" & ChrW(&HED) & "ch y" & ChrW(&H1EBF) & "u ph" & ChrW(&H1EA3) & "i c" & ChrW(&HF3) & " " & ChrW(&H111) & ChrW(&H1B0) & ChrW(&H1EDD) & "ng k" & ChrW(&H1EBB) & " ngang, n" & ChrW(&HE9) & "t li" & ChrW(&H1EC1) & "n, " & ChrW(&H111) & ChrW(&H1ED9) & " d" & ChrW(&HE0) & "i b" & ChrW(&H1EB1) & "ng 1/3 " & ChrW(&H111) & ChrW(&H1EBF) & "n 1/2 " & ChrW(&H111) & ChrW(&H1ED9) & " d" & ChrW(&HE0) & "i d" & ChrW(&HF2) & "ng ch" & ChrW(&H1EEF) & ", " & ChrW(&H111) & ChrW(&H1EB7) & "t c" & ChrW(&HE2) & "n " & ChrW(&H111) & ChrW(&H1ED1) & "i."
    d156.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 5, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    d156.Add "toolHint", "T" & ChrW(&HEA) & "n lo" & ChrW(&H1EA1) & "i, tr" & ChrW(&HED) & "ch y" & ChrW(&H1EBF) & "u"
    c.Add d156
    Dim d157 As Object
    Set d157 = CreateObject("Scripting.Dictionary")
    d157.Add "ruleCode", "ND30-PL1-M2-K5B-STYLE"
    d157.Add "checklistGroup", 6
    d157.Add "group", "component"
    d157.Add "severity", "error"
    d157.Add "sourceLabel", "ND30"
    d157.Add "actionType", "A"
    d157.Add "checkability", "partial"
    d157.Add "autoFixable", True
    d157.Add "riskLevel", "low"
    d157.Add "title", "Tr" & ChrW(&HED) & "ch y" & ChrW(&H1EBF) & "u c" & ChrW(&HF4) & "ng v" & ChrW(&H103) & "n ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    d157.Add "message", "{detail}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: sau ch" & ChrW(&H1EEF) & " 'V/v', ch" & ChrW(&H1EEF) & " in th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng, c" & ChrW(&H1EE1) & " {expectedSize}, ki" & ChrW(&H1EC3) & "u " & ChrW(&H111) & ChrW(&H1EE9) & "ng, canh gi" & ChrW(&H1EEF) & "a d" & ChrW(&H1B0) & ChrW(&H1EDB) & "i s" & ChrW(&H1ED1) & " v" & ChrW(&HE0) & " k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u."
    d157.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 5, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    d157.Add "toolHint", "T" & ChrW(&HEA) & "n lo" & ChrW(&H1EA1) & "i, tr" & ChrW(&HED) & "ch y" & ChrW(&H1EBF) & "u"
    c.Add d157
    Dim d158 As Object
    Set d158 = CreateObject("Scripting.Dictionary")
    d158.Add "ruleCode", "ND30-PL1-M2-K5B-SPACE"
    d158.Add "checklistGroup", 6
    d158.Add "group", "component"
    d158.Add "severity", "error"
    d158.Add "sourceLabel", "ND30"
    d158.Add "actionType", "A"
    d158.Add "checkability", "partial"
    d158.Add "autoFixable", True
    d158.Add "riskLevel", "low"
    d158.Add "title", "Kho" & ChrW(&H1EA3) & "ng c" & ChrW(&HE1) & "ch gi" & ChrW(&H1EEF) & "a tr" & ChrW(&HED) & "ch y" & ChrW(&H1EBF) & "u c" & ChrW(&HF4) & "ng v" & ChrW(&H103) & "n v" & ChrW(&HE0) & " s" & ChrW(&H1ED1) & " k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng"
    d158.Add "message", "Tr" & ChrW(&HED) & "ch y" & ChrW(&H1EBF) & "u c" & ChrW(&HF4) & "ng v" & ChrW(&H103) & "n ph" & ChrW(&H1EA3) & "i c" & ChrW(&HE1) & "ch d" & ChrW(&HF2) & "ng 6pt v" & ChrW(&H1EDB) & "i s" & ChrW(&H1ED1) & " v" & ChrW(&HE0) & " k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n."
    d158.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 5, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    d158.Add "toolHint", "T" & ChrW(&HEA) & "n lo" & ChrW(&H1EA1) & "i, tr" & ChrW(&HED) & "ch y" & ChrW(&H1EBF) & "u"
    c.Add d158
    Dim d159 As Object
    Set d159 = CreateObject("Scripting.Dictionary")
    d159.Add "ruleCode", "ND30-PL1-M2-K6A-STYLE"
    d159.Add "checklistGroup", 7
    d159.Add "group", "component"
    d159.Add "severity", "error"
    d159.Add "sourceLabel", "ND30"
    d159.Add "actionType", "A"
    d159.Add "checkability", "partial"
    d159.Add "autoFixable", True
    d159.Add "riskLevel", "low"
    d159.Add "title", "C" & ChrW(&H103) & "n c" & ChrW(&H1EE9) & " ban h" & ChrW(&HE0) & "nh ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    d159.Add "message", "{detail}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: ch" & ChrW(&H1EEF) & " in th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng, ki" & ChrW(&H1EC3) & "u ch" & ChrW(&H1EEF) & " nghi" & ChrW(&HEA) & "ng, c" & ChrW(&H1EE1) & " {expectedSize}."
    d159.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 6, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m a"
    d159.Add "toolHint", "C" & ChrW(&H103) & "n c" & ChrW(&H1EE9)
    c.Add d159
    Dim d160 As Object
    Set d160 = CreateObject("Scripting.Dictionary")
    d160.Add "ruleCode", "ND30-PL1-M2-K6A-PUNCT"
    d160.Add "checklistGroup", 7
    d160.Add "group", "component"
    d160.Add "severity", "error"
    d160.Add "sourceLabel", "ND30"
    d160.Add "actionType", "C"
    d160.Add "checkability", "full"
    d160.Add "autoFixable", True
    d160.Add "riskLevel", "low"
    d160.Add "title", "D" & ChrW(&H1EA5) & "u cu" & ChrW(&H1ED1) & "i d" & ChrW(&HF2) & "ng c" & ChrW(&H103) & "n c" & ChrW(&H1EE9) & " ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng"
    d160.Add "message", "Sau m" & ChrW(&H1ED7) & "i c" & ChrW(&H103) & "n c" & ChrW(&H1EE9) & " ph" & ChrW(&H1EA3) & "i xu" & ChrW(&H1ED1) & "ng d" & ChrW(&HF2) & "ng, cu" & ChrW(&H1ED1) & "i d" & ChrW(&HF2) & "ng c" & ChrW(&HF3) & " d" & ChrW(&H1EA5) & "u ch" & ChrW(&H1EA5) & "m ph" & ChrW(&H1EA9) & "y; d" & ChrW(&HF2) & "ng cu" & ChrW(&H1ED1) & "i c" & ChrW(&HF9) & "ng k" & ChrW(&H1EBF) & "t th" & ChrW(&HFA) & "c b" & ChrW(&H1EB1) & "ng d" & ChrW(&H1EA5) & "u ch" & ChrW(&H1EA5) & "m."
    d160.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 6, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m a"
    d160.Add "toolHint", "C" & ChrW(&H103) & "n c" & ChrW(&H1EE9)
    c.Add d160
    Dim d161 As Object
    Set d161 = CreateObject("Scripting.Dictionary")
    d161.Add "ruleCode", "ND30-PL1-M2-K6B-CITE"
    d161.Add "checklistGroup", 7
    d161.Add "group", "citation"
    d161.Add "severity", "error"
    d161.Add "sourceLabel", "ND30"
    d161.Add "actionType", "C"
    d161.Add "checkability", "partial"
    d161.Add "autoFixable", True
    d161.Add "riskLevel", "low"
    d161.Add "title", "Vi" & ChrW(&H1EC7) & "n d" & ChrW(&H1EAB) & "n v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&H1EA7) & "y " & ChrW(&H111) & ChrW(&H1EE7)
    d161.Add "message", "'{actual}' n" & ChrW(&HEA) & "n vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '{expected}'. Khi vi" & ChrW(&H1EC7) & "n d" & ChrW(&H1EAB) & "n l" & ChrW(&H1EA7) & "n " & ChrW(&H111) & ChrW(&H1EA7) & "u ph" & ChrW(&H1EA3) & "i ghi " & ChrW(&H111) & ChrW(&H1EA7) & "y " & ChrW(&H111) & ChrW(&H1EE7) & " t" & ChrW(&HEA) & "n lo" & ChrW(&H1EA1) & "i, s" & ChrW(&H1ED1) & ", k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u, th" & ChrW(&H1EDD) & "i gian ban h" & ChrW(&HE0) & "nh, t" & ChrW(&HEA) & "n c" & ChrW(&H1A1) & " quan ban h" & ChrW(&HE0) & "nh v" & ChrW(&HE0) & " tr" & ChrW(&HED) & "ch y" & ChrW(&H1EBF) & "u n" & ChrW(&H1ED9) & "i dung."
    d161.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 6, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    d161.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d161
    Dim d162 As Object
    Set d162 = CreateObject("Scripting.Dictionary")
    d162.Add "ruleCode", "ND30-PL1-M2-K6B-SO"
    d162.Add "checklistGroup", 7
    d162.Add "group", "citation"
    d162.Add "severity", "error"
    d162.Add "sourceLabel", "ND30"
    d162.Add "actionType", "C"
    d162.Add "checkability", "full"
    d162.Add "autoFixable", True
    d162.Add "riskLevel", "low"
    d162.Add "title", "Thi" & ChrW(&H1EBF) & "u ch" & ChrW(&H1EEF) & " 's" & ChrW(&H1ED1) & "' khi vi" & ChrW(&H1EC7) & "n d" & ChrW(&H1EAB) & "n v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n"
    d162.Add "message", "'{actual}' ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '{expected}'. Sau t" & ChrW(&HEA) & "n lo" & ChrW(&H1EA1) & "i v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n ph" & ChrW(&H1EA3) & "i c" & ChrW(&HF3) & " ch" & ChrW(&H1EEF) & " 's" & ChrW(&H1ED1) & "'."
    d162.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 6, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    d162.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d162
    Dim d163 As Object
    Set d163 = CreateObject("Scripting.Dictionary")
    d163.Add "ruleCode", "ND30-PL1-M2-K6B-DATE"
    d163.Add "checklistGroup", 7
    d163.Add "group", "citation"
    d163.Add "severity", "error"
    d163.Add "sourceLabel", "ND30"
    d163.Add "actionType", "C"
    d163.Add "checkability", "full"
    d163.Add "autoFixable", True
    d163.Add "riskLevel", "low"
    d163.Add "title", "Ng" & ChrW(&HE0) & "y th" & ChrW(&HE1) & "ng khi vi" & ChrW(&H1EC7) & "n d" & ChrW(&H1EAB) & "n vi" & ChrW(&H1EBF) & "t t" & ChrW(&H1EAF) & "t"
    d163.Add "message", "'{actual}' ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '{expected}'. Th" & ChrW(&H1EDD) & "i gian ban h" & ChrW(&HE0) & "nh ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t " & ChrW(&H111) & ChrW(&H1EA7) & "y " & ChrW(&H111) & ChrW(&H1EE7) & " d" & ChrW(&H1EA1) & "ng 'ng" & ChrW(&HE0) & "y ... th" & ChrW(&HE1) & "ng ... n" & ChrW(&H103) & "m ...'."
    d163.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 4, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b v" & ChrW(&HE0) & " kho" & ChrW(&H1EA3) & "n 6, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    d163.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d163
    Dim d164 As Object
    Set d164 = CreateObject("Scripting.Dictionary")
    d164.Add "ruleCode", "ND30-PL1-M2-K6D-TITLE"
    d164.Add "checklistGroup", 7
    d164.Add "group", "structure"
    d164.Add "severity", "warning"
    d164.Add "sourceLabel", "ND30"
    d164.Add "actionType", "C"
    d164.Add "checkability", "partial"
    d164.Add "autoFixable", False
    d164.Add "riskLevel", "low"
    d164.Add "title", ChrW(&H110) & "i" & ChrW(&H1EC1) & "u thi" & ChrW(&H1EBF) & "u ti" & ChrW(&HEA) & "u " & ChrW(&H111) & ChrW(&H1EC1)
    d164.Add "message", ChrW(&H110) & "i" & ChrW(&H1EC1) & "u ph" & ChrW(&H1EA3) & "i c" & ChrW(&HF3) & " ti" & ChrW(&HEA) & "u " & ChrW(&H111) & ChrW(&H1EC1) & "."
    d164.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 6, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m d"
    d164.Add "toolHint", "C" & ChrW(&H1EA5) & "p b" & ChrW(&H1ED1) & " c" & ChrW(&H1EE5) & "c"
    c.Add d164
    Dim d165 As Object
    Set d165 = CreateObject("Scripting.Dictionary")
    d165.Add "ruleCode", "ND30-PL1-M2-K6D-ARTICLE"
    d165.Add "checklistGroup", 7
    d165.Add "group", "structure"
    d165.Add "severity", "error"
    d165.Add "sourceLabel", "ND30"
    d165.Add "actionType", "A"
    d165.Add "checkability", "full"
    d165.Add "autoFixable", True
    d165.Add "riskLevel", "low"
    d165.Add "title", ChrW(&H110) & "i" & ChrW(&H1EC1) & "u ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    Dim tmp2 As String
    tmp2 = "{detail}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: ch" & ChrW(&H1EEF) & " in th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng, l" & ChrW(&HF9) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u d" & ChrW(&HF2) & "ng 1 cm ho" & ChrW(&H1EB7) & "c 1,27 cm, s" & ChrW(&H1ED1) & " " & ChrW(&H1EA2) & " R" & ChrW(&H1EAD) & "p, sau s" & ChrW(&H1ED1) & " th" & ChrW(&H1EE9) & " t" & ChrW(&H1EF1) & " c" & ChrW(&HF3) & " d" & ChrW(&H1EA5) & "u ch" & ChrW(&H1EA5) & "m, c" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & " b" & ChrW(&H1EB1) & "ng c" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & " ph" & ChrW(&H1EA7) & "n l" & ChrW(&H1EDD) & "i v" & ChrW(&H103) & "n, ki" & ChrW(&H1EC3) & "u " & ChrW(&H111) & ChrW(&H1EE9)
    tmp2 = tmp2 & "ng " & ChrW(&H111) & ChrW(&H1EAD) & "m."
    d165.Add "message", tmp2
    d165.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 6, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m " & ChrW(&H111)
    d165.Add "toolHint", "C" & ChrW(&H1EA5) & "p b" & ChrW(&H1ED1) & " c" & ChrW(&H1EE5) & "c"
    c.Add d165
    Dim d166 As Object
    Set d166 = CreateObject("Scripting.Dictionary")
    d166.Add "ruleCode", "ND30-PL1-M2-K6D-CLAUSE"
    d166.Add "checklistGroup", 7
    d166.Add "group", "structure"
    d166.Add "severity", "error"
    d166.Add "sourceLabel", "ND30"
    d166.Add "actionType", "A"
    d166.Add "checkability", "partial"
    d166.Add "autoFixable", True
    d166.Add "riskLevel", "low"
    d166.Add "title", "Kho" & ChrW(&H1EA3) & "n ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    Dim tmp3 As String
    tmp3 = "{detail}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: s" & ChrW(&H1ED1) & " " & ChrW(&H1EA2) & " R" & ChrW(&H1EAD) & "p, sau s" & ChrW(&H1ED1) & " th" & ChrW(&H1EE9) & " t" & ChrW(&H1EF1) & " c" & ChrW(&HF3) & " d" & ChrW(&H1EA5) & "u ch" & ChrW(&H1EA5) & "m, c" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & " b" & ChrW(&H1EB1) & "ng c" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & " ph" & ChrW(&H1EA7) & "n l" & ChrW(&H1EDD) & "i v" & ChrW(&H103) & "n, ki" & ChrW(&H1EC3) & "u " & ChrW(&H111) & ChrW(&H1EE9) & "ng; n" & ChrW(&H1EBF) & "u kho" & ChrW(&H1EA3) & "n c" & ChrW(&HF3) & " ti" & ChrW(&HEA) & "u " & ChrW(&H111) & ChrW(&H1EC1) & " th" & ChrW(&HEC) & " ti" & ChrW(&HEA) & "u "
    tmp3 = tmp3 & ChrW(&H111) & ChrW(&H1EC1) & " tr" & ChrW(&HEA) & "n d" & ChrW(&HF2) & "ng ri" & ChrW(&HEA) & "ng, ch" & ChrW(&H1EEF) & " in th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng " & ChrW(&H111) & ChrW(&H1EE9) & "ng " & ChrW(&H111) & ChrW(&H1EAD) & "m."
    d166.Add "message", tmp3
    d166.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 6, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m " & ChrW(&H111)
    d166.Add "toolHint", "C" & ChrW(&H1EA5) & "p b" & ChrW(&H1ED1) & " c" & ChrW(&H1EE5) & "c"
    c.Add d166
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCheckRules_RulesPart2", Err.description
End Sub

Private Sub LoadRawCheckRules_RulesPart3(ByRef c As Collection)
    On Error GoTo ErrHandler
    Dim d167 As Object
    Set d167 = CreateObject("Scripting.Dictionary")
    d167.Add "ruleCode", "ND30-PL1-M2-K6D-POINT"
    d167.Add "checklistGroup", 7
    d167.Add "group", "structure"
    d167.Add "severity", "error"
    d167.Add "sourceLabel", "ND30"
    d167.Add "actionType", "A"
    d167.Add "checkability", "full"
    d167.Add "autoFixable", True
    d167.Add "riskLevel", "low"
    d167.Add "title", ChrW(&H110) & "i" & ChrW(&H1EC3) & "m ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    Dim tmp4 As String
    tmp4 = "{detail}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i ti" & ChrW(&H1EBF) & "ng Vi" & ChrW(&H1EC7) & "t theo th" & ChrW(&H1EE9) & " t" & ChrW(&H1EF1) & " b" & ChrW(&H1EA3) & "ng ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i, sau c" & ChrW(&HF3) & " d" & ChrW(&H1EA5) & "u " & ChrW(&H111) & ChrW(&HF3) & "ng ngo" & ChrW(&H1EB7) & "c " & ChrW(&H111) & ChrW(&H1A1) & "n, ch" & ChrW(&H1EEF) & " in th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng, c" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & " b" & ChrW(&H1EB1) & "ng c" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & " ph" & ChrW(&H1EA7) & "n l" & ChrW(&H1EDD) & "i v" & ChrW(&H103) & "n, ki" & ChrW(&H1EC3) & "u "
    tmp4 = tmp4 & ChrW(&H111) & ChrW(&H1EE9) & "ng."
    d167.Add "message", tmp4
    d167.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 6, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m " & ChrW(&H111)
    d167.Add "toolHint", "C" & ChrW(&H1EA5) & "p b" & ChrW(&H1ED1) & " c" & ChrW(&H1EE5) & "c"
    c.Add d167
    Dim d168 As Object
    Set d168 = CreateObject("Scripting.Dictionary")
    d168.Add "ruleCode", "ND30-PL1-M2-K6D-ALPHABET"
    d168.Add "checklistGroup", 7
    d168.Add "group", "structure"
    d168.Add "severity", "error"
    d168.Add "sourceLabel", "ND30"
    d168.Add "actionType", "A"
    d168.Add "checkability", "full"
    d168.Add "autoFixable", True
    d168.Add "riskLevel", "low"
    d168.Add "title", "Th" & ChrW(&H1EE9) & " t" & ChrW(&H1EF1) & " " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m kh" & ChrW(&HF4) & "ng theo b" & ChrW(&H1EA3) & "ng ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i ti" & ChrW(&H1EBF) & "ng Vi" & ChrW(&H1EC7) & "t"
    d168.Add "message", "Ph" & ChrW(&HE1) & "t hi" & ChrW(&H1EC7) & "n " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m '{actual}'. B" & ChrW(&H1EA3) & "ng ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i ti" & ChrW(&H1EBF) & "ng Vi" & ChrW(&H1EC7) & "t kh" & ChrW(&HF4) & "ng c" & ChrW(&HF3) & " ch" & ChrW(&H1EEF) & " f, j, w, z v" & ChrW(&HE0) & " c" & ChrW(&HF3) & " ch" & ChrW(&H1EEF) & " " & ChrW(&H111) & "; sau e l" & ChrW(&HE0) & " g."
    d168.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 6, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m " & ChrW(&H111)
    d168.Add "toolHint", "S" & ChrW(&H1EED) & "a th" & ChrW(&H1EE9) & " t" & ChrW(&H1EF1) & " " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m"
    c.Add d168
    Dim d169 As Object
    Set d169 = CreateObject("Scripting.Dictionary")
    d169.Add "ruleCode", "ND30-PL1-M2-K6E-ALIGN"
    d169.Add "checklistGroup", 7
    d169.Add "group", "bodyText"
    d169.Add "severity", "error"
    d169.Add "sourceLabel", "ND30"
    d169.Add "actionType", "A"
    d169.Add "checkability", "full"
    d169.Add "autoFixable", True
    d169.Add "riskLevel", "low"
    d169.Add "title", "Ph" & ChrW(&H1EA7) & "n l" & ChrW(&H1EDD) & "i v" & ChrW(&H103) & "n ch" & ChrW(&H1B0) & "a canh " & ChrW(&H111) & ChrW(&H1EC1) & "u hai l" & ChrW(&H1EC1)
    d169.Add "message", "Ph" & ChrW(&HE1) & "t hi" & ChrW(&H1EC7) & "n {count} " & ChrW(&H111) & "o" & ChrW(&H1EA1) & "n ch" & ChrW(&H1B0) & "a canh " & ChrW(&H111) & ChrW(&H1EC1) & "u c" & ChrW(&H1EA3) & " hai l" & ChrW(&H1EC1) & "."
    d169.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 6, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m e"
    d169.Add "toolHint", "Canh " & ChrW(&H111) & ChrW(&H1EC1) & "u hai l" & ChrW(&H1EC1)
    c.Add d169
    Dim d170 As Object
    Set d170 = CreateObject("Scripting.Dictionary")
    d170.Add "ruleCode", "ND30-PL1-M2-K6E-INDENT"
    d170.Add "checklistGroup", 7
    d170.Add "group", "bodyText"
    d170.Add "severity", "error"
    d170.Add "sourceLabel", "ND30"
    d170.Add "actionType", "A"
    d170.Add "checkability", "full"
    d170.Add "autoFixable", True
    d170.Add "riskLevel", "low"
    d170.Add "title", "Th" & ChrW(&H1EE5) & "t " & ChrW(&H111) & ChrW(&H1EA7) & "u d" & ChrW(&HF2) & "ng ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng"
    d170.Add "message", "Th" & ChrW(&H1EE5) & "t " & ChrW(&H111) & ChrW(&H1EA7) & "u d" & ChrW(&HF2) & "ng hi" & ChrW(&H1EC7) & "n t" & ChrW(&H1EA1) & "i {actual} cm. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh l" & ChrW(&HF9) & "i v" & ChrW(&HE0) & "o 1 cm ho" & ChrW(&H1EB7) & "c 1,27 cm."
    d170.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 6, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m e"
    d170.Add "toolHint", "Th" & ChrW(&H1EE5) & "t " & ChrW(&H111) & ChrW(&H1EA7) & "u d" & ChrW(&HF2) & "ng"
    c.Add d170
    Dim d171 As Object
    Set d171 = CreateObject("Scripting.Dictionary")
    d171.Add "ruleCode", "ND30-PL1-M2-K6E-SPACEAFTER"
    d171.Add "checklistGroup", 7
    d171.Add "group", "bodyText"
    d171.Add "severity", "error"
    d171.Add "sourceLabel", "ND30"
    d171.Add "actionType", "A"
    d171.Add "checkability", "full"
    d171.Add "autoFixable", True
    d171.Add "riskLevel", "low"
    d171.Add "title", "Kho" & ChrW(&H1EA3) & "ng c" & ChrW(&HE1) & "ch gi" & ChrW(&H1EEF) & "a c" & ChrW(&HE1) & "c " & ChrW(&H111) & "o" & ChrW(&H1EA1) & "n nh" & ChrW(&H1ECF) & " h" & ChrW(&H1A1) & "n 6pt"
    d171.Add "message", "Kho" & ChrW(&H1EA3) & "ng c" & ChrW(&HE1) & "ch hi" & ChrW(&H1EC7) & "n t" & ChrW(&H1EA1) & "i {actual} pt. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh t" & ChrW(&H1ED1) & "i thi" & ChrW(&H1EC3) & "u 6pt."
    d171.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 6, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m e"
    d171.Add "toolHint", "Gi" & ChrW(&HE3) & "n " & ChrW(&H111) & "o" & ChrW(&H1EA1) & "n 6pt"
    c.Add d171
    Dim d172 As Object
    Set d172 = CreateObject("Scripting.Dictionary")
    d172.Add "ruleCode", "ND30-PL1-M2-K6E-LINESPACING"
    d172.Add "checklistGroup", 7
    d172.Add "group", "bodyText"
    d172.Add "severity", "error"
    d172.Add "sourceLabel", "ND30"
    d172.Add "actionType", "A"
    d172.Add "checkability", "full"
    d172.Add "autoFixable", True
    d172.Add "riskLevel", "low"
    d172.Add "title", "Kho" & ChrW(&H1EA3) & "ng c" & ChrW(&HE1) & "ch gi" & ChrW(&H1EEF) & "a c" & ChrW(&HE1) & "c d" & ChrW(&HF2) & "ng ngo" & ChrW(&HE0) & "i d" & ChrW(&H1EA3) & "i cho ph" & ChrW(&HE9) & "p"
    d172.Add "message", "Gi" & ChrW(&HE3) & "n d" & ChrW(&HF2) & "ng hi" & ChrW(&H1EC7) & "n t" & ChrW(&H1EA1) & "i {actual}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh t" & ChrW(&H1ED1) & "i thi" & ChrW(&H1EC3) & "u d" & ChrW(&HF2) & "ng " & ChrW(&H111) & ChrW(&H1A1) & "n, t" & ChrW(&H1ED1) & "i " & ChrW(&H111) & "a 1,5 lines."
    d172.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 6, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m e"
    d172.Add "toolHint", "Gi" & ChrW(&HE3) & "n d" & ChrW(&HF2) & "ng"
    c.Add d172
    Dim d173 As Object
    Set d173 = CreateObject("Scripting.Dictionary")
    d173.Add "ruleCode", "ND30-PL1-M2-K6E-DOTSLASH"
    d173.Add "checklistGroup", 7
    d173.Add "group", "bodyText"
    d173.Add "severity", "warning"
    d173.Add "sourceLabel", "SUY RA"
    d173.Add "actionType", "C"
    d173.Add "checkability", "partial"
    d173.Add "autoFixable", False
    d173.Add "riskLevel", "low"
    d173.Add "title", "D" & ChrW(&H1EA5) & "u k" & ChrW(&H1EBF) & "t th" & ChrW(&HFA) & "c n" & ChrW(&H1ED9) & "i dung v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng"
    d173.Add "message", ChrW(&H110) & "o" & ChrW(&H1EA1) & "n cu" & ChrW(&H1ED1) & "i n" & ChrW(&H1ED9) & "i dung v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n (tr" & ChrW(&H1B0) & ChrW(&H1EDB) & "c 'N" & ChrW(&H1A1) & "i nh" & ChrW(&H1EAD) & "n:') ch" & ChrW(&H1B0) & "a k" & ChrW(&H1EBF) & "t th" & ChrW(&HFA) & "c b" & ChrW(&H1EB1) & "ng {expected}."
    d173.Add "citation", "Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c Viettel, " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "u 12 kho" & ChrW(&H1EA3) & "n 2"
    d173.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d173
    Dim d174 As Object
    Set d174 = CreateObject("Scripting.Dictionary")
    d174.Add "ruleCode", "ND30-PL1-M2-K7B-AUTH"
    d174.Add "checklistGroup", 8
    d174.Add "group", "component"
    d174.Add "severity", "warning"
    d174.Add "sourceLabel", "ND30"
    d174.Add "actionType", "C"
    d174.Add "checkability", "partial"
    d174.Add "autoFixable", False
    d174.Add "riskLevel", "low"
    d174.Add "title", "Ch" & ChrW(&H1EEF) & " vi" & ChrW(&H1EBF) & "t t" & ChrW(&H1EAF) & "t quy" & ChrW(&H1EC1) & "n h" & ChrW(&H1EA1) & "n ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i k" & ChrW(&HFD) & " ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng"
    d174.Add "message", "Ph" & ChrW(&HE1) & "t hi" & ChrW(&H1EC7) & "n '{actual}'. C" & ChrW(&HE1) & "c ch" & ChrW(&H1EEF) & " vi" & ChrW(&H1EBF) & "t t" & ChrW(&H1EAF) & "t quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: TM., Q., KT., TL., TUQ."
    d174.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 7, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    d174.Add "toolHint", "Ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i k" & ChrW(&HFD)
    c.Add d174
    Dim d175 As Object
    Set d175 = CreateObject("Scripting.Dictionary")
    d175.Add "ruleCode", "ND30-PL1-M2-K7D-STYLE"
    d175.Add "checklistGroup", 8
    d175.Add "group", "component"
    d175.Add "severity", "error"
    d175.Add "sourceLabel", "ND30"
    d175.Add "actionType", "A"
    d175.Add "checkability", "partial"
    d175.Add "autoFixable", True
    d175.Add "riskLevel", "low"
    d175.Add "title", "Quy" & ChrW(&H1EC1) & "n h" & ChrW(&H1EA1) & "n, ch" & ChrW(&H1EE9) & "c v" & ChrW(&H1EE5) & " ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i k" & ChrW(&HFD) & " ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    d175.Add "message", "{detail}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: ch" & ChrW(&H1EEF) & " in hoa, c" & ChrW(&H1EE1) & " {expectedSize}, ki" & ChrW(&H1EC3) & "u " & ChrW(&H111) & ChrW(&H1EE9) & "ng " & ChrW(&H111) & ChrW(&H1EAD) & "m."
    d175.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 7, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m " & ChrW(&H111)
    d175.Add "toolHint", "Ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i k" & ChrW(&HFD)
    c.Add d175
    Dim d176 As Object
    Set d176 = CreateObject("Scripting.Dictionary")
    d176.Add "ruleCode", "ND30-PL1-M2-K9A-COLON"
    d176.Add "checklistGroup", 9
    d176.Add "group", "component"
    d176.Add "severity", "error"
    d176.Add "sourceLabel", "ND30"
    d176.Add "actionType", "C"
    d176.Add "checkability", "full"
    d176.Add "autoFixable", True
    d176.Add "riskLevel", "low"
    d176.Add "title", "Thi" & ChrW(&H1EBF) & "u d" & ChrW(&H1EA5) & "u hai ch" & ChrW(&H1EA5) & "m sau 'K" & ChrW(&HED) & "nh g" & ChrW(&H1EED) & "i'"
    d176.Add "message", "Sau t" & ChrW(&H1EEB) & " 'K" & ChrW(&HED) & "nh g" & ChrW(&H1EED) & "i' ph" & ChrW(&H1EA3) & "i c" & ChrW(&HF3) & " d" & ChrW(&H1EA5) & "u hai ch" & ChrW(&H1EA5) & "m."
    d176.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 9, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m d"
    d176.Add "toolHint", "K" & ChrW(&HED) & "nh g" & ChrW(&H1EED) & "i"
    c.Add d176
    Dim d177 As Object
    Set d177 = CreateObject("Scripting.Dictionary")
    d177.Add "ruleCode", "ND30-PL1-M2-K9A-LAYOUT"
    d177.Add "checklistGroup", 9
    d177.Add "group", "component"
    d177.Add "severity", "error"
    d177.Add "sourceLabel", "ND30"
    d177.Add "actionType", "A"
    d177.Add "checkability", "partial"
    d177.Add "autoFixable", True
    d177.Add "riskLevel", "low"
    d177.Add "title", "B" & ChrW(&H1ED1) & " tr" & ChrW(&HED) & " ph" & ChrW(&H1EA7) & "n 'K" & ChrW(&HED) & "nh g" & ChrW(&H1EED) & "i' ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng"
    Dim tmp5 As String
    tmp5 = "{detail}. G" & ChrW(&H1EED) & "i m" & ChrW(&H1ED9) & "t n" & ChrW(&H1A1) & "i th" & ChrW(&HEC) & " tr" & ChrW(&HEC) & "nh b" & ChrW(&HE0) & "y tr" & ChrW(&HEA) & "n c" & ChrW(&HF9) & "ng m" & ChrW(&H1ED9) & "t d" & ChrW(&HF2) & "ng; g" & ChrW(&H1EED) & "i t" & ChrW(&H1EEB) & " hai n" & ChrW(&H1A1) & "i tr" & ChrW(&H1EDF) & " l" & ChrW(&HEA) & "n th" & ChrW(&HEC) & " xu" & ChrW(&H1ED1) & "ng d" & ChrW(&HF2) & "ng, m" & ChrW(&H1ED7) & "i n" & ChrW(&H1A1) & "i m" & ChrW(&H1ED9) & "t d" & ChrW(&HF2) & "ng ri" & ChrW(&HEA) & "ng, " & ChrW(&H111) & ChrW(&H1EA7) & "u d" & ChrW(&HF2) & "ng c" & ChrW(&HF3) & " g" & ChrW(&H1EA1) & "ch " & ChrW(&H111) & ChrW(&H1EA7)
    tmp5 = tmp5 & "u d" & ChrW(&HF2) & "ng th" & ChrW(&H1EB3) & "ng h" & ChrW(&HE0) & "ng d" & ChrW(&H1B0) & ChrW(&H1EDB) & "i d" & ChrW(&H1EA5) & "u hai ch" & ChrW(&H1EA5) & "m."
    d177.Add "message", tmp5
    d177.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 9, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m d"
    d177.Add "toolHint", "K" & ChrW(&HED) & "nh g" & ChrW(&H1EED) & "i"
    c.Add d177
    Dim d178 As Object
    Set d178 = CreateObject("Scripting.Dictionary")
    d178.Add "ruleCode", "ND30-PL1-M2-K9A-PUNCT"
    d178.Add "checklistGroup", 9
    d178.Add "group", "component"
    d178.Add "severity", "error"
    d178.Add "sourceLabel", "ND30"
    d178.Add "actionType", "C"
    d178.Add "checkability", "full"
    d178.Add "autoFixable", True
    d178.Add "riskLevel", "low"
    d178.Add "title", "D" & ChrW(&H1EA5) & "u cu" & ChrW(&H1ED1) & "i d" & ChrW(&HF2) & "ng trong 'K" & ChrW(&HED) & "nh g" & ChrW(&H1EED) & "i' ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng"
    d178.Add "message", "Cu" & ChrW(&H1ED1) & "i m" & ChrW(&H1ED7) & "i d" & ChrW(&HF2) & "ng c" & ChrW(&HF3) & " d" & ChrW(&H1EA5) & "u ch" & ChrW(&H1EA5) & "m ph" & ChrW(&H1EA9) & "y, d" & ChrW(&HF2) & "ng cu" & ChrW(&H1ED1) & "i c" & ChrW(&HF9) & "ng c" & ChrW(&HF3) & " d" & ChrW(&H1EA5) & "u ch" & ChrW(&H1EA5) & "m."
    d178.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 9, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m d"
    d178.Add "toolHint", "K" & ChrW(&HED) & "nh g" & ChrW(&H1EED) & "i"
    c.Add d178
    Dim d179 As Object
    Set d179 = CreateObject("Scripting.Dictionary")
    d179.Add "ruleCode", "ND30-PL1-M2-K9A-INLINE-END"
    d179.Add "checklistGroup", 9
    d179.Add "group", "component"
    d179.Add "severity", "error"
    d179.Add "sourceLabel", "SUY RA"
    d179.Add "actionType", "C"
    d179.Add "checkability", "full"
    d179.Add "autoFixable", False
    d179.Add "riskLevel", "low"
    d179.Add "title", "N" & ChrW(&H1A1) & "i k" & ChrW(&HED) & "nh g" & ChrW(&H1EED) & "i vi" & ChrW(&H1EBF) & "t c" & ChrW(&HF9) & "ng d" & ChrW(&HF2) & "ng ch" & ChrW(&H1B0) & "a k" & ChrW(&H1EBF) & "t th" & ChrW(&HFA) & "c b" & ChrW(&H1EB1) & "ng d" & ChrW(&H1EA5) & "u ch" & ChrW(&H1EA5) & "m"
    d179.Add "message", "Khi n" & ChrW(&H1A1) & "i nh" & ChrW(&H1EAD) & "n " & ChrW(&H111) & ChrW(&H1B0) & ChrW(&H1EE3) & "c vi" & ChrW(&H1EBF) & "t ngay sau 'K" & ChrW(&HED) & "nh g" & ChrW(&H1EED) & "i:' tr" & ChrW(&HEA) & "n c" & ChrW(&HF9) & "ng m" & ChrW(&H1ED9) & "t d" & ChrW(&HF2) & "ng th" & ChrW(&HEC) & " cu" & ChrW(&H1ED1) & "i d" & ChrW(&HF2) & "ng ph" & ChrW(&H1EA3) & "i c" & ChrW(&HF3) & " d" & ChrW(&H1EA5) & "u ch" & ChrW(&H1EA5) & "m."
    d179.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 9, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m d"
    d179.Add "toolHint", "K" & ChrW(&HED) & "nh g" & ChrW(&H1EED) & "i"
    c.Add d179
    Dim d180 As Object
    Set d180 = CreateObject("Scripting.Dictionary")
    d180.Add "ruleCode", "ND30-PL1-M2-K9B-LABEL"
    d180.Add "checklistGroup", 9
    d180.Add "group", "component"
    d180.Add "severity", "error"
    d180.Add "sourceLabel", "ND30"
    d180.Add "actionType", "A"
    d180.Add "checkability", "partial"
    d180.Add "autoFixable", True
    d180.Add "riskLevel", "low"
    d180.Add "title", "T" & ChrW(&H1EEB) & " 'N" & ChrW(&H1A1) & "i nh" & ChrW(&H1EAD) & "n' ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    Dim tmp6 As String
    tmp6 = "{detail}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: tr" & ChrW(&HEA) & "n m" & ChrW(&H1ED9) & "t d" & ChrW(&HF2) & "ng ri" & ChrW(&HEA) & "ng, ngang h" & ChrW(&HE0) & "ng v" & ChrW(&H1EDB) & "i d" & ChrW(&HF2) & "ng quy" & ChrW(&H1EC1) & "n h" & ChrW(&H1EA1) & "n ch" & ChrW(&H1EE9) & "c v" & ChrW(&H1EE5) & " c" & ChrW(&H1EE7) & "a ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i k" & ChrW(&HFD) & ", s" & ChrW(&HE1) & "t l" & ChrW(&H1EC1) & " tr" & ChrW(&HE1) & "i, sau c" & ChrW(&HF3) & " d" & ChrW(&H1EA5) & "u hai ch" & ChrW(&H1EA5) & "m, ch" & ChrW(&H1EEF) & " in th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng c" & ChrW(&H1EE1) & " 12, ki" & ChrW(&H1EC3) & "u nghi" & ChrW(&HEA)
    tmp6 = tmp6 & "ng " & ChrW(&H111) & ChrW(&H1EAD) & "m."
    d180.Add "message", tmp6
    d180.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 9, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m d"
    d180.Add "toolHint", "N" & ChrW(&H1A1) & "i nh" & ChrW(&H1EAD) & "n"
    c.Add d180
    Dim d181 As Object
    Set d181 = CreateObject("Scripting.Dictionary")
    d181.Add "ruleCode", "ND30-PL1-M2-K9B-LIST"
    d181.Add "checklistGroup", 9
    d181.Add "group", "component"
    d181.Add "severity", "error"
    d181.Add "sourceLabel", "ND30"
    d181.Add "actionType", "A"
    d181.Add "checkability", "partial"
    d181.Add "autoFixable", True
    d181.Add "riskLevel", "low"
    d181.Add "title", "Danh s" & ChrW(&HE1) & "ch n" & ChrW(&H1A1) & "i nh" & ChrW(&H1EAD) & "n ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    Dim tmp7 As String
    tmp7 = "{detail}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: ch" & ChrW(&H1EEF) & " in th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng c" & ChrW(&H1EE1) & " 11, ki" & ChrW(&H1EC3) & "u " & ChrW(&H111) & ChrW(&H1EE9) & "ng, m" & ChrW(&H1ED7) & "i n" & ChrW(&H1A1) & "i nh" & ChrW(&H1EAD) & "n tr" & ChrW(&HEA) & "n m" & ChrW(&H1ED9) & "t d" & ChrW(&HF2) & "ng ri" & ChrW(&HEA) & "ng, " & ChrW(&H111) & ChrW(&H1EA7) & "u d" & ChrW(&HF2) & "ng c" & ChrW(&HF3) & " g" & ChrW(&H1EA1) & "ch " & ChrW(&H111) & ChrW(&H1EA7) & "u d" & ChrW(&HF2) & "ng s" & ChrW(&HE1) & "t l" & ChrW(&H1EC1) & " tr" & ChrW(&HE1) & "i, cu" & ChrW(&H1ED1) & "i d" & ChrW(&HF2) & "ng c" & ChrW(&HF3) & " d"
    tmp7 = tmp7 & ChrW(&H1EA5) & "u ch" & ChrW(&H1EA5) & "m ph" & ChrW(&H1EA9) & "y."
    d181.Add "message", tmp7
    d181.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 9, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m d"
    d181.Add "toolHint", "N" & ChrW(&H1A1) & "i nh" & ChrW(&H1EAD) & "n"
    c.Add d181
    Dim d182 As Object
    Set d182 = CreateObject("Scripting.Dictionary")
    d182.Add "ruleCode", "ND30-PL1-M2-K9B-LUU"
    d182.Add "checklistGroup", 9
    d182.Add "group", "component"
    d182.Add "severity", "error"
    d182.Add "sourceLabel", "ND30"
    d182.Add "actionType", "C"
    d182.Add "checkability", "full"
    d182.Add "autoFixable", True
    d182.Add "riskLevel", "low"
    d182.Add "title", "D" & ChrW(&HF2) & "ng 'L" & ChrW(&H1B0) & "u' ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh"
    Dim tmp8 As String
    tmp8 = "D" & ChrW(&HF2) & "ng cu" & ChrW(&H1ED1) & "i c" & ChrW(&HF9) & "ng ph" & ChrW(&H1EA3) & "i g" & ChrW(&H1ED3) & "m ch" & ChrW(&H1EEF) & " 'L" & ChrW(&H1B0) & "u', d" & ChrW(&H1EA5) & "u hai ch" & ChrW(&H1EA5) & "m, ch" & ChrW(&H1EEF) & " vi" & ChrW(&H1EBF) & "t t" & ChrW(&H1EAF) & "t 'VT', d" & ChrW(&H1EA5) & "u ph" & ChrW(&H1EA9) & "y, ch" & ChrW(&H1EEF) & " vi" & ChrW(&H1EBF) & "t t" & ChrW(&H1EAF) & "t t" & ChrW(&HEA) & "n " & ChrW(&H111) & ChrW(&H1A1) & "n v" & ChrW(&H1ECB) & " so" & ChrW(&H1EA1) & "n th" & ChrW(&H1EA3) & "o v" & ChrW(&HE0) & " s" & ChrW(&H1ED1) & " l" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng b" & ChrW(&H1EA3) & "n l" & ChrW(&H1B0) & "u, cu" & ChrW(&H1ED1)
    tmp8 = tmp8 & "i c" & ChrW(&HF9) & "ng l" & ChrW(&HE0) & " d" & ChrW(&H1EA5) & "u ch" & ChrW(&H1EA5) & "m."
    d182.Add "message", tmp8
    d182.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 9, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m d"
    d182.Add "toolHint", "N" & ChrW(&H1A1) & "i nh" & ChrW(&H1EAD) & "n"
    c.Add d182
    Dim d183 As Object
    Set d183 = CreateObject("Scripting.Dictionary")
    d183.Add "ruleCode", "ND30-PL1-M3-K1A-REF"
    d183.Add "checklistGroup", 11
    d183.Add "group", "component"
    d183.Add "severity", "warning"
    d183.Add "sourceLabel", "ND30"
    d183.Add "actionType", "C"
    d183.Add "checkability", "partial"
    d183.Add "autoFixable", False
    d183.Add "riskLevel", "low"
    d183.Add "title", "V" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n c" & ChrW(&HF3) & " ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c nh" & ChrW(&H1B0) & "ng thi" & ChrW(&H1EBF) & "u ch" & ChrW(&H1EC9) & " d" & ChrW(&H1EAB) & "n"
    d183.Add "message", "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng h" & ChrW(&H1EE3) & "p v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n c" & ChrW(&HF3) & " ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c k" & ChrW(&HE8) & "m theo th" & ChrW(&HEC) & " trong v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n ph" & ChrW(&H1EA3) & "i c" & ChrW(&HF3) & " ch" & ChrW(&H1EC9) & " d" & ChrW(&H1EAB) & "n v" & ChrW(&H1EC1) & " ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c " & ChrW(&H111) & ChrW(&HF3) & "."
    d183.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c III, kho" & ChrW(&H1EA3) & "n 1, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m a"
    d183.Add "toolHint", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c"
    c.Add d183
    Dim d184 As Object
    Set d184 = CreateObject("Scripting.Dictionary")
    d184.Add "ruleCode", "ND30-PL1-M3-K1A-NUM"
    d184.Add "checklistGroup", 11
    d184.Add "group", "component"
    d184.Add "severity", "error"
    d184.Add "sourceLabel", "ND30"
    d184.Add "actionType", "A"
    d184.Add "checkability", "partial"
    d184.Add "autoFixable", True
    d184.Add "riskLevel", "low"
    d184.Add "title", "S" & ChrW(&H1ED1) & " th" & ChrW(&H1EE9) & " t" & ChrW(&H1EF1) & " ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c ch" & ChrW(&H1B0) & "a d" & ChrW(&HF9) & "ng ch" & ChrW(&H1EEF) & " s" & ChrW(&H1ED1) & " La M" & ChrW(&HE3)
    d184.Add "message", "V" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n c" & ChrW(&HF3) & " t" & ChrW(&H1EEB) & " hai ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c tr" & ChrW(&H1EDF) & " l" & ChrW(&HEA) & "n th" & ChrW(&HEC) & " c" & ChrW(&HE1) & "c ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c ph" & ChrW(&H1EA3) & "i " & ChrW(&H111) & ChrW(&H1B0) & ChrW(&H1EE3) & "c " & ChrW(&H111) & ChrW(&HE1) & "nh s" & ChrW(&H1ED1) & " th" & ChrW(&H1EE9) & " t" & ChrW(&H1EF1) & " b" & ChrW(&H1EB1) & "ng ch" & ChrW(&H1EEF) & " s" & ChrW(&H1ED1) & " La M" & ChrW(&HE3) & "."
    d184.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c III, kho" & ChrW(&H1EA3) & "n 1, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m a"
    d184.Add "toolHint", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c"
    c.Add d184
    Dim d185 As Object
    Set d185 = CreateObject("Scripting.Dictionary")
    d185.Add "ruleCode", "ND30-PL1-M3-K1B"
    d185.Add "checklistGroup", 11
    d185.Add "group", "component"
    d185.Add "severity", "error"
    d185.Add "sourceLabel", "ND30"
    d185.Add "actionType", "A"
    d185.Add "checkability", "partial"
    d185.Add "autoFixable", True
    d185.Add "riskLevel", "low"
    d185.Add "title", "Ti" & ChrW(&HEA) & "u " & ChrW(&H111) & ChrW(&H1EC1) & " ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    Dim tmp9 As String
    tmp9 = "{detail}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: t" & ChrW(&H1EEB) & " 'Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c' v" & ChrW(&HE0) & " s" & ChrW(&H1ED1) & " th" & ChrW(&H1EE9) & " t" & ChrW(&H1EF1) & " tr" & ChrW(&HEA) & "n d" & ChrW(&HF2) & "ng ri" & ChrW(&HEA) & "ng, canh gi" & ChrW(&H1EEF) & "a, ch" & ChrW(&H1EEF) & " in th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng c" & ChrW(&H1EE1) & " 14 " & ChrW(&H111) & ChrW(&H1EE9) & "ng " & ChrW(&H111) & ChrW(&H1EAD) & "m; t" & ChrW(&HEA) & "n ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c canh gi" & ChrW(&H1EEF) & "a, ch" & ChrW(&H1EEF) & " in hoa c" & ChrW(&H1EE1) & " 13-14 " & ChrW(&H111) & ChrW(&H1EE9) & "ng " & ChrW(&H111)
    tmp9 = tmp9 & ChrW(&H1EAD) & "m."
    d185.Add "message", tmp9
    d185.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c III, kho" & ChrW(&H1EA3) & "n 1, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    d185.Add "toolHint", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c"
    c.Add d185
    Dim d186 As Object
    Set d186 = CreateObject("Scripting.Dictionary")
    d186.Add "ruleCode", "ND30-PL1-M3-K1C"
    d186.Add "checklistGroup", 11
    d186.Add "group", "component"
    d186.Add "severity", "error"
    d186.Add "sourceLabel", "ND30"
    d186.Add "actionType", "A"
    d186.Add "checkability", "partial"
    d186.Add "autoFixable", True
    d186.Add "riskLevel", "low"
    d186.Add "title", "Th" & ChrW(&HF4) & "ng tin ch" & ChrW(&H1EC9) & " d" & ChrW(&H1EAB) & "n k" & ChrW(&HE8) & "m theo ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    d186.Add "message", "{detail}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: canh gi" & ChrW(&H1EEF) & "a ph" & ChrW(&HED) & "a d" & ChrW(&H1B0) & ChrW(&H1EDB) & "i t" & ChrW(&HEA) & "n ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c, ch" & ChrW(&H1EEF) & " in th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng c" & ChrW(&H1EE1) & " 13-14, ki" & ChrW(&H1EC3) & "u nghi" & ChrW(&HEA) & "ng, c" & ChrW(&HF9) & "ng ph" & ChrW(&HF4) & "ng ch" & ChrW(&H1EEF) & " v" & ChrW(&H1EDB) & "i n" & ChrW(&H1ED9) & "i dung v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n, m" & ChrW(&HE0) & "u " & ChrW(&H111) & "en."
    d186.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c III, kho" & ChrW(&H1EA3) & "n 1, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m c"
    d186.Add "toolHint", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c"
    c.Add d186
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCheckRules_RulesPart3", Err.description
End Sub

Private Sub LoadRawCheckRules_RulesPart4(ByRef c As Collection)
    On Error GoTo ErrHandler
    Dim d187 As Object
    Set d187 = CreateObject("Scripting.Dictionary")
    d187.Add "ruleCode", "ND30-PL1-M3-K1D"
    d187.Add "checklistGroup", 11
    d187.Add "group", "pageSetup"
    d187.Add "severity", "error"
    d187.Add "sourceLabel", "ND30"
    d187.Add "actionType", "A"
    d187.Add "checkability", "partial"
    d187.Add "autoFixable", True
    d187.Add "riskLevel", "low"
    d187.Add "title", "S" & ChrW(&H1ED1) & " trang ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HE1) & "nh ri" & ChrW(&HEA) & "ng"
    d187.Add "message", "S" & ChrW(&H1ED1) & " trang c" & ChrW(&H1EE7) & "a ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c " & ChrW(&H111) & ChrW(&H1B0) & ChrW(&H1EE3) & "c " & ChrW(&H111) & ChrW(&HE1) & "nh s" & ChrW(&H1ED1) & " ri" & ChrW(&HEA) & "ng theo t" & ChrW(&H1EEB) & "ng ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c."
    d187.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c III, kho" & ChrW(&H1EA3) & "n 1, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m " & ChrW(&H111)
    d187.Add "toolHint", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c"
    c.Add d187
    Dim d188 As Object
    Set d188 = CreateObject("Scripting.Dictionary")
    d188.Add "ruleCode", "ND30-PL1-M3-K2B"
    d188.Add "checklistGroup", 10
    d188.Add "group", "component"
    d188.Add "severity", "error"
    d188.Add "sourceLabel", "ND30"
    d188.Add "actionType", "A"
    d188.Add "checkability", "partial"
    d188.Add "autoFixable", True
    d188.Add "riskLevel", "low"
    d188.Add "title", "D" & ChrW(&H1EA5) & "u ch" & ChrW(&H1EC9) & " m" & ChrW(&H1EE9) & "c " & ChrW(&H111) & ChrW(&H1ED9) & " kh" & ChrW(&H1EA9) & "n ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    Dim tmp10 As String
    tmp10 = "{detail}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: ch" & ChrW(&H1EEF) & " in hoa, ph" & ChrW(&HF4) & "ng Times New Roman, c" & ChrW(&H1EE1) & " 13-14, ki" & ChrW(&H1EC3) & "u " & ChrW(&H111) & ChrW(&H1EE9) & "ng " & ChrW(&H111) & ChrW(&H1EAD) & "m, " & ChrW(&H111) & ChrW(&H1EB7) & "t c" & ChrW(&HE2) & "n " & ChrW(&H111) & ChrW(&H1ED1) & "i trong khung h" & ChrW(&HEC) & "nh ch" & ChrW(&H1EEF) & " nh" & ChrW(&H1EAD) & "t vi" & ChrW(&H1EC1) & "n " & ChrW(&H111) & ChrW(&H1A1) & "n; k" & ChrW(&HED) & "ch th" & ChrW(&H1B0) & ChrW(&H1EDB) & "c H" & ChrW(&H1ECE) & "A T" & ChrW(&H1ED0) & "C 30x8 mm, TH" & ChrW(&H1AF) & ChrW(&H1EE2) & "NG KH" & ChrW(&H1EA8) & "N 40x8 mm, KH"
    tmp10 = tmp10 & ChrW(&H1EA8) & "N 20x8 mm."
    d188.Add "message", tmp10
    d188.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c III, kho" & ChrW(&H1EA3) & "n 2, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    d188.Add "toolHint", "M" & ChrW(&H1EE9) & "c " & ChrW(&H111) & ChrW(&H1ED9) & " kh" & ChrW(&H1EA9) & "n"
    c.Add d188
    Dim d189 As Object
    Set d189 = CreateObject("Scripting.Dictionary")
    d189.Add "ruleCode", "ND30-PL1-M3-K2C"
    d189.Add "checklistGroup", 10
    d189.Add "group", "component"
    d189.Add "severity", "error"
    d189.Add "sourceLabel", "ND30"
    d189.Add "actionType", "A"
    d189.Add "checkability", "partial"
    d189.Add "autoFixable", True
    d189.Add "riskLevel", "low"
    d189.Add "title", "Ch" & ChrW(&H1EC9) & " d" & ChrW(&H1EAB) & "n ph" & ChrW(&H1EA1) & "m vi l" & ChrW(&H1B0) & "u h" & ChrW(&HE0) & "nh ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    d189.Add "message", "{detail}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: ch" & ChrW(&H1EEF) & " in hoa, ph" & ChrW(&HF4) & "ng Times New Roman, c" & ChrW(&H1EE1) & " 13-14, ki" & ChrW(&H1EC3) & "u " & ChrW(&H111) & ChrW(&H1EE9) & "ng " & ChrW(&H111) & ChrW(&H1EAD) & "m, tr" & ChrW(&HEC) & "nh b" & ChrW(&HE0) & "y c" & ChrW(&HE2) & "n " & ChrW(&H111) & ChrW(&H1ED1) & "i trong khung h" & ChrW(&HEC) & "nh ch" & ChrW(&H1EEF) & " nh" & ChrW(&H1EAD) & "t vi" & ChrW(&H1EC1) & "n " & ChrW(&H111) & ChrW(&H1A1) & "n."
    d189.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c III, kho" & ChrW(&H1EA3) & "n 2, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m c"
    d189.Add "toolHint", "Ph" & ChrW(&H1EA1) & "m vi l" & ChrW(&H1B0) & "u h" & ChrW(&HE0) & "nh"
    c.Add d189
    Dim d190 As Object
    Set d190 = CreateObject("Scripting.Dictionary")
    d190.Add "ruleCode", "ND30-PL1-M3-K3"
    d190.Add "checklistGroup", 10
    d190.Add "group", "component"
    d190.Add "severity", "error"
    d190.Add "sourceLabel", "ND30"
    d190.Add "actionType", "A"
    d190.Add "checkability", "partial"
    d190.Add "autoFixable", True
    d190.Add "riskLevel", "low"
    d190.Add "title", "K" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i so" & ChrW(&H1EA1) & "n th" & ChrW(&H1EA3) & "o ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    d190.Add "message", "{detail}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u b" & ChrW(&H1EB1) & "ng ch" & ChrW(&H1EEF) & " in hoa, s" & ChrW(&H1ED1) & " l" & ChrW(&H1B0) & ChrW(&H1EE3) & "ng b" & ChrW(&H1EA3) & "n b" & ChrW(&H1EB1) & "ng ch" & ChrW(&H1EEF) & " s" & ChrW(&H1ED1) & " " & ChrW(&H1EA2) & " R" & ChrW(&H1EAD) & "p, c" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & " 11, ki" & ChrW(&H1EC3) & "u " & ChrW(&H111) & ChrW(&H1EE9) & "ng."
    d190.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c III, kho" & ChrW(&H1EA3) & "n 3"
    d190.Add "toolHint", "K" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u so" & ChrW(&H1EA1) & "n th" & ChrW(&H1EA3) & "o"
    c.Add d190
    Dim d191 As Object
    Set d191 = CreateObject("Scripting.Dictionary")
    d191.Add "ruleCode", "ND30-PL1-M3-K4"
    d191.Add "checklistGroup", 10
    d191.Add "group", "component"
    d191.Add "severity", "error"
    d191.Add "sourceLabel", "ND30"
    d191.Add "actionType", "A"
    d191.Add "checkability", "partial"
    d191.Add "autoFixable", True
    d191.Add "riskLevel", "low"
    d191.Add "title", "Th" & ChrW(&HF4) & "ng tin li" & ChrW(&HEA) & "n h" & ChrW(&H1EC7) & " c" & ChrW(&H1EE7) & "a c" & ChrW(&H1A1) & " quan ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    d191.Add "message", "{detail}. Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh: " & ChrW(&H1EDF) & " trang th" & ChrW(&H1EE9) & " nh" & ChrW(&H1EA5) & "t, ch" & ChrW(&H1EEF) & " in th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng c" & ChrW(&H1EE1) & " 11-12, ki" & ChrW(&H1EC3) & "u " & ChrW(&H111) & ChrW(&H1EE9) & "ng, d" & ChrW(&H1B0) & ChrW(&H1EDB) & "i m" & ChrW(&H1ED9) & "t " & ChrW(&H111) & ChrW(&H1B0) & ChrW(&H1EDD) & "ng k" & ChrW(&H1EBB) & " n" & ChrW(&HE9) & "t li" & ChrW(&H1EC1) & "n k" & ChrW(&HE9) & "o d" & ChrW(&HE0) & "i h" & ChrW(&H1EBF) & "t chi" & ChrW(&H1EC1) & "u ngang v" & ChrW(&HF9) & "ng tr" & ChrW(&HEC) & "nh b" & ChrW(&HE0) & "y v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n."
    d191.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c III, kho" & ChrW(&H1EA3) & "n 4"
    d191.Add "toolHint", "Th" & ChrW(&HF4) & "ng tin li" & ChrW(&HEA) & "n h" & ChrW(&H1EC7)
    c.Add d191
    Dim d192 As Object
    Set d192 = CreateObject("Scripting.Dictionary")
    d192.Add "ruleCode", "ND30-PL1-M4-POS"
    d192.Add "checklistGroup", 1
    d192.Add "group", "component"
    d192.Add "severity", "warning"
    d192.Add "sourceLabel", "ND30"
    d192.Add "actionType", "C"
    d192.Add "checkability", "partial"
    d192.Add "autoFixable", False
    d192.Add "riskLevel", "low"
    d192.Add "title", "V" & ChrW(&H1ECB) & " tr" & ChrW(&HED) & " th" & ChrW(&HE0) & "nh ph" & ChrW(&H1EA7) & "n th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng s" & ChrW(&H1A1) & " " & ChrW(&H111) & ChrW(&H1ED3)
    d192.Add "message", "Th" & ChrW(&HE0) & "nh ph" & ChrW(&H1EA7) & "n {componentName} (" & ChrW(&HF4) & " s" & ChrW(&H1ED1) & " {cellNumber}) kh" & ChrW(&HF4) & "ng n" & ChrW(&H1EB1) & "m " & ChrW(&H111) & ChrW(&HFA) & "ng v" & ChrW(&HF9) & "ng quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh trong s" & ChrW(&H1A1) & " " & ChrW(&H111) & ChrW(&H1ED3) & " b" & ChrW(&H1ED1) & " tr" & ChrW(&HED) & "."
    d192.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c IV"
    d192.Add "toolHint", "Ki" & ChrW(&H1EC3) & "m tra th" & ChrW(&H1EC3) & " th" & ChrW(&H1EE9) & "c"
    c.Add d192
    Dim d193 As Object
    Set d193 = CreateObject("Scripting.Dictionary")
    d193.Add "ruleCode", "ND30-PL2-M1"
    d193.Add "checklistGroup", 14
    d193.Add "group", "capitalization"
    d193.Add "severity", "error"
    d193.Add "sourceLabel", "ND30"
    d193.Add "actionType", "C"
    d193.Add "checkability", "full"
    d193.Add "autoFixable", True
    d193.Add "riskLevel", "low"
    d193.Add "title", "Ch" & ChrW(&H1B0) & "a vi" & ChrW(&H1EBF) & "t hoa " & ChrW(&H111) & ChrW(&H1EA7) & "u c" & ChrW(&HE2) & "u"
    d193.Add "message", "'{actual}' n" & ChrW(&HEA) & "n vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '{expected}'. Vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u " & ChrW(&HE2) & "m ti" & ChrW(&H1EBF) & "t th" & ChrW(&H1EE9) & " nh" & ChrW(&H1EA5) & "t c" & ChrW(&H1EE7) & "a m" & ChrW(&H1ED9) & "t c" & ChrW(&HE2) & "u ho" & ChrW(&HE0) & "n ch" & ChrW(&H1EC9) & "nh."
    d193.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c I"
    d193.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d193
    Dim d194 As Object
    Set d194 = CreateObject("Scripting.Dictionary")
    d194.Add "ruleCode", "ND30-PL2-M2-K1"
    d194.Add "checklistGroup", 14
    d194.Add "group", "capitalization"
    d194.Add "severity", "warning"
    d194.Add "sourceLabel", "ND30"
    d194.Add "actionType", "C"
    d194.Add "checkability", "warnOnly"
    d194.Add "autoFixable", False
    d194.Add "riskLevel", "low"
    d194.Add "title", "T" & ChrW(&HEA) & "n ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i c" & ChrW(&HF3) & " th" & ChrW(&H1EC3) & " ch" & ChrW(&H1B0) & "a vi" & ChrW(&H1EBF) & "t hoa " & ChrW(&H111) & ChrW(&HFA) & "ng"
    d194.Add "message", "'{actual}' " & ChrW(&H2014) & " vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u t" & ChrW(&H1EA5) & "t c" & ChrW(&H1EA3) & " c" & ChrW(&HE1) & "c " & ChrW(&HE2) & "m ti" & ChrW(&H1EBF) & "t c" & ChrW(&H1EE7) & "a danh t" & ChrW(&H1EEB) & " ri" & ChrW(&HEA) & "ng ch" & ChrW(&H1EC9) & " t" & ChrW(&HEA) & "n ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i."
    d194.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 1"
    d194.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d194
    Dim d195 As Object
    Set d195 = CreateObject("Scripting.Dictionary")
    d195.Add "ruleCode", "ND30-PL2-M2-K2"
    d195.Add "checklistGroup", 14
    d195.Add "group", "capitalization"
    d195.Add "severity", "info"
    d195.Add "sourceLabel", "ND30"
    d195.Add "actionType", "C"
    d195.Add "checkability", "warnOnly"
    d195.Add "autoFixable", False
    d195.Add "riskLevel", "low"
    d195.Add "title", "T" & ChrW(&HEA) & "n ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i n" & ChrW(&H1B0) & ChrW(&H1EDB) & "c ngo" & ChrW(&HE0) & "i phi" & ChrW(&HEA) & "n " & ChrW(&HE2) & "m"
    Dim tmp11 As String
    tmp11 = "Phi" & ChrW(&HEA) & "n " & ChrW(&HE2) & "m H" & ChrW(&HE1) & "n-Vi" & ChrW(&H1EC7) & "t vi" & ChrW(&H1EBF) & "t theo quy t" & ChrW(&H1EAF) & "c t" & ChrW(&HEA) & "n ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i Vi" & ChrW(&H1EC7) & "t Nam; phi" & ChrW(&HEA) & "n " & ChrW(&HE2) & "m tr" & ChrW(&H1EF1) & "c ti" & ChrW(&H1EBF) & "p vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u " & ChrW(&HE2) & "m ti" & ChrW(&H1EBF) & "t th" & ChrW(&H1EE9) & " nh" & ChrW(&H1EA5) & "t c" & ChrW(&H1EE7) & "a m" & ChrW(&H1ED7) & "i th" & ChrW(&HE0) & "nh ph" & ChrW(&H1EA7) & "n, c" & ChrW(&HE1) & "c " & ChrW(&HE2) & "m ti"
    tmp11 = tmp11 & ChrW(&H1EBF) & "t trong c" & ChrW(&HF9) & "ng th" & ChrW(&HE0) & "nh ph" & ChrW(&H1EA7) & "n n" & ChrW(&H1ED1) & "i b" & ChrW(&H1EB1) & "ng g" & ChrW(&H1EA1) & "ch n" & ChrW(&H1ED1) & "i."
    d195.Add "message", tmp11
    d195.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 2"
    d195.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d195
    Dim d196 As Object
    Set d196 = CreateObject("Scripting.Dictionary")
    d196.Add "ruleCode", "ND30-PL2-M3-K1A"
    d196.Add "checklistGroup", 14
    d196.Add "group", "capitalization"
    d196.Add "severity", "warning"
    d196.Add "sourceLabel", "ND30"
    d196.Add "actionType", "C"
    d196.Add "checkability", "warnOnly"
    d196.Add "autoFixable", False
    d196.Add "riskLevel", "low"
    d196.Add "title", "T" & ChrW(&HEA) & "n " & ChrW(&H111) & ChrW(&H1A1) & "n v" & ChrW(&H1ECB) & " h" & ChrW(&HE0) & "nh ch" & ChrW(&HED) & "nh c" & ChrW(&HF3) & " th" & ChrW(&H1EC3) & " vi" & ChrW(&H1EBF) & "t hoa sai"
    d196.Add "message", "'{actual}' " & ChrW(&H2014) & " vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u c" & ChrW(&HE1) & "c " & ChrW(&HE2) & "m ti" & ChrW(&H1EBF) & "t t" & ChrW(&H1EA1) & "o th" & ChrW(&HE0) & "nh t" & ChrW(&HEA) & "n ri" & ChrW(&HEA) & "ng, kh" & ChrW(&HF4) & "ng vi" & ChrW(&H1EBF) & "t hoa danh t" & ChrW(&H1EEB) & " chung ch" & ChrW(&H1EC9) & " " & ChrW(&H111) & ChrW(&H1A1) & "n v" & ChrW(&H1ECB) & " h" & ChrW(&HE0) & "nh ch" & ChrW(&HED) & "nh, kh" & ChrW(&HF4) & "ng d" & ChrW(&HF9) & "ng g" & ChrW(&H1EA1) & "ch n" & ChrW(&H1ED1) & "i."
    d196.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c III, kho" & ChrW(&H1EA3) & "n 1, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m a"
    d196.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d196
    Dim d197 As Object
    Set d197 = CreateObject("Scripting.Dictionary")
    d197.Add "ruleCode", "ND30-PL2-M3-K1B"
    d197.Add "checklistGroup", 14
    d197.Add "group", "capitalization"
    d197.Add "severity", "warning"
    d197.Add "sourceLabel", "ND30"
    d197.Add "actionType", "C"
    d197.Add "checkability", "partial"
    d197.Add "autoFixable", True
    d197.Add "riskLevel", "low"
    d197.Add "title", ChrW(&H110) & ChrW(&H1A1) & "n v" & ChrW(&H1ECB) & " h" & ChrW(&HE0) & "nh ch" & ChrW(&HED) & "nh " & ChrW(&H111) & ChrW(&H1EB7) & "t theo ch" & ChrW(&H1EEF) & " s" & ChrW(&H1ED1) & " ho" & ChrW(&H1EB7) & "c t" & ChrW(&HEA) & "n ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i"
    d197.Add "message", "'{actual}' n" & ChrW(&HEA) & "n vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '{expected}'. Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng h" & ChrW(&H1EE3) & "p t" & ChrW(&HEA) & "n " & ChrW(&H111) & ChrW(&H1A1) & "n v" & ChrW(&H1ECB) & " h" & ChrW(&HE0) & "nh ch" & ChrW(&HED) & "nh c" & ChrW(&H1EA5) & "u t" & ChrW(&H1EA1) & "o v" & ChrW(&H1EDB) & "i ch" & ChrW(&H1EEF) & " s" & ChrW(&H1ED1) & ", t" & ChrW(&HEA) & "n ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i ho" & ChrW(&H1EB7) & "c s" & ChrW(&H1EF1) & " ki" & ChrW(&H1EC7) & "n l" & ChrW(&H1ECB) & "ch s" & ChrW(&H1EED) & " th" & ChrW(&HEC) & " vi" & ChrW(&H1EBF) & "t hoa c" & ChrW(&H1EA3) & " danh t" & ChrW(&H1EEB) & " chung."
    d197.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c III, kho" & ChrW(&H1EA3) & "n 1, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    d197.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d197
    Dim d198 As Object
    Set d198 = CreateObject("Scripting.Dictionary")
    d198.Add "ruleCode", "ND30-PL2-M3-K1C"
    d198.Add "checklistGroup", 14
    d198.Add "group", "capitalization"
    d198.Add "severity", "error"
    d198.Add "sourceLabel", "ND30"
    d198.Add "actionType", "C"
    d198.Add "checkability", "full"
    d198.Add "autoFixable", True
    d198.Add "riskLevel", "low"
    d198.Add "title", "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng h" & ChrW(&H1EE3) & "p vi" & ChrW(&H1EBF) & "t hoa " & ChrW(&H111) & ChrW(&H1EB7) & "c bi" & ChrW(&H1EC7) & "t v" & ChrW(&H1EC1) & " " & ChrW(&H111) & ChrW(&H1ECB) & "a l" & ChrW(&HFD)
    d198.Add "message", "'{actual}' ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '{expected}'."
    d198.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c III, kho" & ChrW(&H1EA3) & "n 1, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m c"
    d198.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d198
    Dim d199 As Object
    Set d199 = CreateObject("Scripting.Dictionary")
    d199.Add "ruleCode", "ND30-PL2-M3-K1D"
    d199.Add "checklistGroup", 14
    d199.Add "group", "capitalization"
    d199.Add "severity", "warning"
    d199.Add "sourceLabel", "ND30"
    d199.Add "actionType", "C"
    d199.Add "checkability", "partial"
    d199.Add "autoFixable", True
    d199.Add "riskLevel", "low"
    d199.Add "title", "Danh t" & ChrW(&H1EEB) & " chung ch" & ChrW(&H1EC9) & " " & ChrW(&H111) & ChrW(&H1ECB) & "a h" & ChrW(&HEC) & "nh vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng"
    d199.Add "message", "'{actual}' n" & ChrW(&HEA) & "n vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '{expected}'. Danh t" & ChrW(&H1EEB) & " chung ch" & ChrW(&H1EC9) & " " & ChrW(&H111) & ChrW(&H1ECB) & "a h" & ChrW(&HEC) & "nh " & ChrW(&H111) & ChrW(&HE3) & " tr" & ChrW(&H1EDF) & " th" & ChrW(&HE0) & "nh t" & ChrW(&HEA) & "n ri" & ChrW(&HEA) & "ng th" & ChrW(&HEC) & " vi" & ChrW(&H1EBF) & "t hoa t" & ChrW(&H1EA5) & "t c" & ChrW(&H1EA3) & "; " & ChrW(&H111) & "i li" & ChrW(&H1EC1) & "n danh t" & ChrW(&H1EEB) & " ri" & ChrW(&HEA) & "ng th" & ChrW(&HEC) & " kh" & ChrW(&HF4) & "ng vi" & ChrW(&H1EBF) & "t hoa danh t" & ChrW(&H1EEB) & " chung."
    d199.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c III, kho" & ChrW(&H1EA3) & "n 1, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m d"
    d199.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d199
    Dim d200 As Object
    Set d200 = CreateObject("Scripting.Dictionary")
    d200.Add "ruleCode", "ND30-PL2-M3-K1E"
    d200.Add "checklistGroup", 14
    d200.Add "group", "capitalization"
    d200.Add "severity", "warning"
    d200.Add "sourceLabel", "ND30"
    d200.Add "actionType", "C"
    d200.Add "checkability", "partial"
    d200.Add "autoFixable", True
    d200.Add "riskLevel", "low"
    d200.Add "title", "T" & ChrW(&HEA) & "n v" & ChrW(&HF9) & "ng, mi" & ChrW(&H1EC1) & "n, khu v" & ChrW(&H1EF1) & "c vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng"
    d200.Add "message", "'{actual}' n" & ChrW(&HEA) & "n vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '{expected}'. Vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u c" & ChrW(&H1EE7) & "a t" & ChrW(&H1EA5) & "t c" & ChrW(&H1EA3) & " c" & ChrW(&HE1) & "c " & ChrW(&HE2) & "m ti" & ChrW(&H1EBF) & "t t" & ChrW(&H1EA1) & "o th" & ChrW(&HE0) & "nh t" & ChrW(&HEA) & "n g" & ChrW(&H1ECD) & "i."
    d200.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c III, kho" & ChrW(&H1EA3) & "n 1, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m " & ChrW(&H111)
    d200.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d200
    Dim d201 As Object
    Set d201 = CreateObject("Scripting.Dictionary")
    d201.Add "ruleCode", "ND30-PL2-M3-K2"
    d201.Add "checklistGroup", 14
    d201.Add "group", "capitalization"
    d201.Add "severity", "info"
    d201.Add "sourceLabel", "ND30"
    d201.Add "actionType", "C"
    d201.Add "checkability", "warnOnly"
    d201.Add "autoFixable", False
    d201.Add "riskLevel", "low"
    d201.Add "title", "T" & ChrW(&HEA) & "n " & ChrW(&H111) & ChrW(&H1ECB) & "a l" & ChrW(&HFD) & " n" & ChrW(&H1B0) & ChrW(&H1EDB) & "c ngo" & ChrW(&HE0) & "i phi" & ChrW(&HEA) & "n " & ChrW(&HE2) & "m"
    d201.Add "message", "Phi" & ChrW(&HEA) & "n " & ChrW(&HE2) & "m H" & ChrW(&HE1) & "n-Vi" & ChrW(&H1EC7) & "t vi" & ChrW(&H1EBF) & "t theo quy t" & ChrW(&H1EAF) & "c t" & ChrW(&HEA) & "n " & ChrW(&H111) & ChrW(&H1ECB) & "a l" & ChrW(&HFD) & " Vi" & ChrW(&H1EC7) & "t Nam; phi" & ChrW(&HEA) & "n " & ChrW(&HE2) & "m tr" & ChrW(&H1EF1) & "c ti" & ChrW(&H1EBF) & "p vi" & ChrW(&H1EBF) & "t theo quy t" & ChrW(&H1EAF) & "c t" & ChrW(&HEA) & "n ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i n" & ChrW(&H1B0) & ChrW(&H1EDB) & "c ngo" & ChrW(&HE0) & "i."
    d201.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c III, kho" & ChrW(&H1EA3) & "n 2"
    d201.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d201
    Dim d202 As Object
    Set d202 = CreateObject("Scripting.Dictionary")
    d202.Add "ruleCode", "ND30-PL2-M4-K1A"
    d202.Add "checklistGroup", 14
    d202.Add "group", "capitalization"
    d202.Add "severity", "warning"
    d202.Add "sourceLabel", "ND30"
    d202.Add "actionType", "C"
    d202.Add "checkability", "warnOnly"
    d202.Add "autoFixable", False
    d202.Add "riskLevel", "low"
    d202.Add "title", "T" & ChrW(&HEA) & "n c" & ChrW(&H1A1) & " quan, t" & ChrW(&H1ED5) & " ch" & ChrW(&H1EE9) & "c c" & ChrW(&HF3) & " th" & ChrW(&H1EC3) & " vi" & ChrW(&H1EBF) & "t hoa sai"
    d202.Add "message", "'{actual}' " & ChrW(&H2014) & " vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u c" & ChrW(&H1EE7) & "a c" & ChrW(&HE1) & "c t" & ChrW(&H1EEB) & ", c" & ChrW(&H1EE5) & "m t" & ChrW(&H1EEB) & " ch" & ChrW(&H1EC9) & " lo" & ChrW(&H1EA1) & "i h" & ChrW(&HEC) & "nh c" & ChrW(&H1A1) & " quan t" & ChrW(&H1ED5) & " ch" & ChrW(&H1EE9) & "c v" & ChrW(&HE0) & " ch" & ChrW(&H1EE9) & "c n" & ChrW(&H103) & "ng, l" & ChrW(&H129) & "nh v" & ChrW(&H1EF1) & "c ho" & ChrW(&H1EA1) & "t " & ChrW(&H111) & ChrW(&H1ED9) & "ng."
    d202.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c IV, kho" & ChrW(&H1EA3) & "n 1, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m a"
    d202.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d202
    Dim d203 As Object
    Set d203 = CreateObject("Scripting.Dictionary")
    d203.Add "ruleCode", "ND30-PL2-M4-K1B"
    d203.Add "checklistGroup", 14
    d203.Add "group", "capitalization"
    d203.Add "severity", "error"
    d203.Add "sourceLabel", "ND30"
    d203.Add "actionType", "C"
    d203.Add "checkability", "full"
    d203.Add "autoFixable", True
    d203.Add "riskLevel", "low"
    d203.Add "title", "Tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng h" & ChrW(&H1EE3) & "p vi" & ChrW(&H1EBF) & "t hoa " & ChrW(&H111) & ChrW(&H1EB7) & "c bi" & ChrW(&H1EC7) & "t v" & ChrW(&H1EC1) & " t" & ChrW(&HEA) & "n c" & ChrW(&H1A1) & " quan"
    d203.Add "message", "'{actual}' ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '{expected}'."
    d203.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c IV, kho" & ChrW(&H1EA3) & "n 1, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    d203.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d203
    Dim d204 As Object
    Set d204 = CreateObject("Scripting.Dictionary")
    d204.Add "ruleCode", "ND30-PL2-M4-K2"
    d204.Add "checklistGroup", 14
    d204.Add "group", "capitalization"
    d204.Add "severity", "info"
    d204.Add "sourceLabel", "ND30"
    d204.Add "actionType", "C"
    d204.Add "checkability", "warnOnly"
    d204.Add "autoFixable", False
    d204.Add "riskLevel", "low"
    d204.Add "title", "T" & ChrW(&HEA) & "n c" & ChrW(&H1A1) & " quan, t" & ChrW(&H1ED5) & " ch" & ChrW(&H1EE9) & "c n" & ChrW(&H1B0) & ChrW(&H1EDB) & "c ngo" & ChrW(&HE0) & "i"
    d204.Add "message", "T" & ChrW(&HEA) & "n " & ChrW(&H111) & ChrW(&HE3) & " d" & ChrW(&H1ECB) & "ch ngh" & ChrW(&H129) & "a vi" & ChrW(&H1EBF) & "t theo quy t" & ChrW(&H1EAF) & "c t" & ChrW(&HEA) & "n c" & ChrW(&H1A1) & " quan Vi" & ChrW(&H1EC7) & "t Nam; d" & ChrW(&H1EA1) & "ng vi" & ChrW(&H1EBF) & "t t" & ChrW(&H1EAF) & "t vi" & ChrW(&H1EBF) & "t b" & ChrW(&H1EB1) & "ng ch" & ChrW(&H1EEF) & " in hoa nh" & ChrW(&H1B0) & " nguy" & ChrW(&HEA) & "n ng" & ChrW(&H1EEF) & " ho" & ChrW(&H1EB7) & "c chuy" & ChrW(&H1EC3) & "n t" & ChrW(&H1EF1) & " La-tinh."
    d204.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c IV, kho" & ChrW(&H1EA3) & "n 2"
    d204.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d204
    Dim d205 As Object
    Set d205 = CreateObject("Scripting.Dictionary")
    d205.Add "ruleCode", "ND30-PL2-M5-K2"
    d205.Add "checklistGroup", 14
    d205.Add "group", "capitalization"
    d205.Add "severity", "warning"
    d205.Add "sourceLabel", "ND30"
    d205.Add "actionType", "C"
    d205.Add "checkability", "partial"
    d205.Add "autoFixable", True
    d205.Add "riskLevel", "low"
    d205.Add "title", "T" & ChrW(&HEA) & "n hu" & ChrW(&HE2) & "n ch" & ChrW(&H1B0) & ChrW(&H1A1) & "ng, huy ch" & ChrW(&H1B0) & ChrW(&H1A1) & "ng, danh hi" & ChrW(&H1EC7) & "u vinh d" & ChrW(&H1EF1)
    d205.Add "message", "'{actual}' n" & ChrW(&HEA) & "n vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '{expected}'. Vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u c" & ChrW(&H1EE7) & "a c" & ChrW(&HE1) & "c " & ChrW(&HE2) & "m ti" & ChrW(&H1EBF) & "t c" & ChrW(&H1EE7) & "a c" & ChrW(&HE1) & "c th" & ChrW(&HE0) & "nh ph" & ChrW(&H1EA7) & "n t" & ChrW(&H1EA1) & "o th" & ChrW(&HE0) & "nh t" & ChrW(&HEA) & "n ri" & ChrW(&HEA) & "ng v" & ChrW(&HE0) & " c" & ChrW(&HE1) & "c t" & ChrW(&H1EEB) & " ch" & ChrW(&H1EC9) & " th" & ChrW(&H1EE9) & ", h" & ChrW(&H1EA1) & "ng."
    d205.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c V, kho" & ChrW(&H1EA3) & "n 2"
    d205.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d205
    Dim d206 As Object
    Set d206 = CreateObject("Scripting.Dictionary")
    d206.Add "ruleCode", "ND30-PL2-M5-K3"
    d206.Add "checklistGroup", 14
    d206.Add "group", "capitalization"
    d206.Add "severity", "warning"
    d206.Add "sourceLabel", "ND30"
    d206.Add "actionType", "C"
    d206.Add "checkability", "warnOnly"
    d206.Add "autoFixable", False
    d206.Add "riskLevel", "low"
    d206.Add "title", "T" & ChrW(&HEA) & "n ch" & ChrW(&H1EE9) & "c v" & ChrW(&H1EE5) & ", h" & ChrW(&H1ECD) & "c v" & ChrW(&H1ECB) & ", danh hi" & ChrW(&H1EC7) & "u"
    d206.Add "message", "'{actual}' " & ChrW(&H2014) & " vi" & ChrW(&H1EBF) & "t hoa t" & ChrW(&HEA) & "n ch" & ChrW(&H1EE9) & "c v" & ChrW(&H1EE5) & ", h" & ChrW(&H1ECD) & "c v" & ChrW(&H1ECB) & " n" & ChrW(&H1EBF) & "u " & ChrW(&H111) & "i li" & ChrW(&H1EC1) & "n v" & ChrW(&H1EDB) & "i t" & ChrW(&HEA) & "n ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i c" & ChrW(&H1EE5) & " th" & ChrW(&H1EC3) & "."
    d206.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c V, kho" & ChrW(&H1EA3) & "n 3"
    d206.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d206
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCheckRules_RulesPart4", Err.description
End Sub

Private Sub LoadRawCheckRules_RulesPart5(ByRef c As Collection)
    On Error GoTo ErrHandler
    Dim d207 As Object
    Set d207 = CreateObject("Scripting.Dictionary")
    d207.Add "ruleCode", "ND30-PL2-M5-K4"
    d207.Add "checklistGroup", 14
    d207.Add "group", "capitalization"
    d207.Add "severity", "info"
    d207.Add "sourceLabel", "ND30"
    d207.Add "actionType", "C"
    d207.Add "checkability", "warnOnly"
    d207.Add "autoFixable", False
    d207.Add "riskLevel", "low"
    d207.Add "title", "Danh t" & ChrW(&H1EEB) & " chung " & ChrW(&H111) & ChrW(&HE3) & " ri" & ChrW(&HEA) & "ng h" & ChrW(&HF3) & "a"
    Dim tmp12 As String
    tmp12 = "Vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u khi d" & ChrW(&HF9) & "ng trong m" & ChrW(&H1ED9) & "t nh" & ChrW(&HE2) & "n x" & ChrW(&H1B0) & "ng, " & ChrW(&H111) & ChrW(&H1EE9) & "ng " & ChrW(&H111) & ChrW(&H1ED9) & "c l" & ChrW(&H1EAD) & "p v" & ChrW(&HE0) & " th" & ChrW(&H1EC3) & " hi" & ChrW(&H1EC7) & "n s" & ChrW(&H1EF1) & " tr" & ChrW(&HE2) & "n tr" & ChrW(&H1ECD) & "ng. Ph" & ChrW(&H1EE5) & " thu" & ChrW(&H1ED9) & "c ng" & ChrW(&H1EEF) & " c" & ChrW(&H1EA3) & "nh, s" & ChrW(&H1EA3) & "n ph" & ChrW(&H1EA9) & "m kh" & ChrW(&HF4) & "ng t" & ChrW(&H1EF1) & " quy" & ChrW(&H1EBF) & "t " & ChrW(&H111)
    tmp12 = tmp12 & ChrW(&H1ECB) & "nh."
    d207.Add "message", tmp12
    d207.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c V, kho" & ChrW(&H1EA3) & "n 4"
    d207.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d207
    Dim d208 As Object
    Set d208 = CreateObject("Scripting.Dictionary")
    d208.Add "ruleCode", "ND30-PL2-M5-K5"
    d208.Add "checklistGroup", 14
    d208.Add "group", "capitalization"
    d208.Add "severity", "warning"
    d208.Add "sourceLabel", "ND30"
    d208.Add "actionType", "C"
    d208.Add "checkability", "partial"
    d208.Add "autoFixable", True
    d208.Add "riskLevel", "low"
    d208.Add "title", "T" & ChrW(&HEA) & "n ng" & ChrW(&HE0) & "y l" & ChrW(&H1EC5) & ", ng" & ChrW(&HE0) & "y k" & ChrW(&H1EF7) & " ni" & ChrW(&H1EC7) & "m ch" & ChrW(&H1B0) & "a vi" & ChrW(&H1EBF) & "t hoa " & ChrW(&H111) & ChrW(&HFA) & "ng"
    d208.Add "message", "'{actual}' n" & ChrW(&HEA) & "n vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '{expected}'. Vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u c" & ChrW(&H1EE7) & "a " & ChrW(&HE2) & "m ti" & ChrW(&H1EBF) & "t t" & ChrW(&H1EA1) & "o th" & ChrW(&HE0) & "nh t" & ChrW(&HEA) & "n g" & ChrW(&H1ECD) & "i ng" & ChrW(&HE0) & "y l" & ChrW(&H1EC5) & ", ng" & ChrW(&HE0) & "y k" & ChrW(&H1EF7) & " ni" & ChrW(&H1EC7) & "m."
    d208.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c V, kho" & ChrW(&H1EA3) & "n 5"
    d208.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d208
    Dim d209 As Object
    Set d209 = CreateObject("Scripting.Dictionary")
    d209.Add "ruleCode", "ND30-PL2-M5-K6"
    d209.Add "checklistGroup", 14
    d209.Add "group", "capitalization"
    d209.Add "severity", "warning"
    d209.Add "sourceLabel", "ND30"
    d209.Add "actionType", "C"
    d209.Add "checkability", "partial"
    d209.Add "autoFixable", True
    d209.Add "riskLevel", "low"
    d209.Add "title", "T" & ChrW(&HEA) & "n lo" & ChrW(&H1EA1) & "i v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n c" & ChrW(&H1EE5) & " th" & ChrW(&H1EC3) & " ch" & ChrW(&H1B0) & "a vi" & ChrW(&H1EBF) & "t hoa " & ChrW(&H111) & ChrW(&HFA) & "ng"
    Dim tmp13 As String
    tmp13 = "'{actual}' n" & ChrW(&HEA) & "n vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '{expected}'. Vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u c" & ChrW(&H1EE7) & "a t" & ChrW(&HEA) & "n lo" & ChrW(&H1EA1) & "i v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n v" & ChrW(&HE0) & " ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u c" & ChrW(&H1EE7) & "a " & ChrW(&HE2) & "m ti" & ChrW(&H1EBF) & "t th" & ChrW(&H1EE9) & " nh" & ChrW(&H1EA5) & "t t" & ChrW(&H1EA1) & "o th" & ChrW(&HE0) & "nh t" & ChrW(&HEA) & "n g" & ChrW(&H1ECD) & "i khi n" & ChrW(&HF3) & "i " & ChrW(&H111) & ChrW(&H1EBF)
    tmp13 = tmp13 & "n m" & ChrW(&H1ED9) & "t v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n c" & ChrW(&H1EE5) & " th" & ChrW(&H1EC3) & "."
    d209.Add "message", tmp13
    d209.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c V, kho" & ChrW(&H1EA3) & "n 6"
    d209.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d209
    Dim d210 As Object
    Set d210 = CreateObject("Scripting.Dictionary")
    d210.Add "ruleCode", "ND30-PL2-M5-K7"
    d210.Add "checklistGroup", 14
    d210.Add "group", "capitalization"
    d210.Add "severity", "error"
    d210.Add "sourceLabel", "ND30"
    d210.Add "actionType", "C"
    d210.Add "checkability", "full"
    d210.Add "autoFixable", True
    d210.Add "riskLevel", "low"
    d210.Add "title", "Vi" & ChrW(&H1EC7) & "n d" & ChrW(&H1EAB) & "n " & ChrW(&H111) & "i" & ChrW(&H1EC1) & "u kho" & ChrW(&H1EA3) & "n vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng"
    d210.Add "message", "'{actual}' ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '{expected}'. Vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u c" & ChrW(&H1EE7) & "a " & ChrW(&H111) & "i" & ChrW(&H1EC1) & "u; vi" & ChrW(&H1EBF) & "t th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng kho" & ChrW(&H1EA3) & "n, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m."
    d210.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c V, kho" & ChrW(&H1EA3) & "n 7"
    d210.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d210
    Dim d211 As Object
    Set d211 = CreateObject("Scripting.Dictionary")
    d211.Add "ruleCode", "ND30-PL2-M5-K8A"
    d211.Add "checklistGroup", 14
    d211.Add "group", "capitalization"
    d211.Add "severity", "warning"
    d211.Add "sourceLabel", "ND30"
    d211.Add "actionType", "C"
    d211.Add "checkability", "partial"
    d211.Add "autoFixable", True
    d211.Add "riskLevel", "low"
    d211.Add "title", "T" & ChrW(&HEA) & "n n" & ChrW(&H103) & "m " & ChrW(&HE2) & "m l" & ChrW(&H1ECB) & "ch ch" & ChrW(&H1B0) & "a vi" & ChrW(&H1EBF) & "t hoa " & ChrW(&H111) & ChrW(&HFA) & "ng"
    d211.Add "message", "'{actual}' ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '{expected}'. Vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u c" & ChrW(&H1EE7) & "a t" & ChrW(&H1EA5) & "t c" & ChrW(&H1EA3) & " c" & ChrW(&HE1) & "c " & ChrW(&HE2) & "m ti" & ChrW(&H1EBF) & "t t" & ChrW(&H1EA1) & "o th" & ChrW(&HE0) & "nh t" & ChrW(&HEA) & "n g" & ChrW(&H1ECD) & "i."
    d211.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c V, kho" & ChrW(&H1EA3) & "n 8, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m a"
    d211.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d211
    Dim d212 As Object
    Set d212 = CreateObject("Scripting.Dictionary")
    d212.Add "ruleCode", "ND30-PL2-M5-K8B"
    d212.Add "checklistGroup", 14
    d212.Add "group", "capitalization"
    d212.Add "severity", "warning"
    d212.Add "sourceLabel", "ND30"
    d212.Add "actionType", "C"
    d212.Add "checkability", "partial"
    d212.Add "autoFixable", True
    d212.Add "riskLevel", "low"
    d212.Add "title", "T" & ChrW(&HEA) & "n ng" & ChrW(&HE0) & "y t" & ChrW(&H1EBF) & "t ch" & ChrW(&H1B0) & "a vi" & ChrW(&H1EBF) & "t hoa " & ChrW(&H111) & ChrW(&HFA) & "ng"
    Dim tmp14 As String
    tmp14 = "'{actual}' ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '{expected}'. Vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u c" & ChrW(&H1EE7) & "a " & ChrW(&HE2) & "m ti" & ChrW(&H1EBF) & "t th" & ChrW(&H1EE9) & " nh" & ChrW(&H1EA5) & "t t" & ChrW(&H1EA1) & "o th" & ChrW(&HE0) & "nh t" & ChrW(&HEA) & "n g" & ChrW(&H1ECD) & "i; ch" & ChrW(&H1EEF) & " 't" & ChrW(&H1EBF) & "t' vi" & ChrW(&H1EBF) & "t th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng, tr" & ChrW(&H1EEB) & " khi 'T" & ChrW(&H1EBF) & "t' d" & ChrW(&HF9) & "ng thay cho t" & ChrW(&H1EBF) & "t Nguy" & ChrW(&HEA) & "n " & ChrW(&H111) & ChrW(&HE1)
    tmp14 = tmp14 & "n."
    d212.Add "message", tmp14
    d212.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c V, kho" & ChrW(&H1EA3) & "n 8, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    d212.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d212
    Dim d213 As Object
    Set d213 = CreateObject("Scripting.Dictionary")
    d213.Add "ruleCode", "ND30-PL2-M5-K8C"
    d213.Add "checklistGroup", 14
    d213.Add "group", "capitalization"
    d213.Add "severity", "warning"
    d213.Add "sourceLabel", "ND30"
    d213.Add "actionType", "C"
    d213.Add "checkability", "partial"
    d213.Add "autoFixable", True
    d213.Add "riskLevel", "low"
    d213.Add "title", "T" & ChrW(&HEA) & "n ng" & ChrW(&HE0) & "y trong tu" & ChrW(&H1EA7) & "n, th" & ChrW(&HE1) & "ng trong n" & ChrW(&H103) & "m ch" & ChrW(&H1B0) & "a vi" & ChrW(&H1EBF) & "t hoa " & ChrW(&H111) & ChrW(&HFA) & "ng"
    d213.Add "message", "'{actual}' n" & ChrW(&HEA) & "n vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '{expected}'. Vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u c" & ChrW(&H1EE7) & "a " & ChrW(&HE2) & "m ti" & ChrW(&H1EBF) & "t ch" & ChrW(&H1EC9) & " ng" & ChrW(&HE0) & "y v" & ChrW(&HE0) & " th" & ChrW(&HE1) & "ng trong tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng h" & ChrW(&H1EE3) & "p kh" & ChrW(&HF4) & "ng d" & ChrW(&HF9) & "ng ch" & ChrW(&H1EEF) & " s" & ChrW(&H1ED1) & "."
    d213.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c V, kho" & ChrW(&H1EA3) & "n 8, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m c"
    d213.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d213
    Dim d214 As Object
    Set d214 = CreateObject("Scripting.Dictionary")
    d214.Add "ruleCode", "ND30-PL2-M5-K9"
    d214.Add "checklistGroup", 14
    d214.Add "group", "capitalization"
    d214.Add "severity", "warning"
    d214.Add "sourceLabel", "ND30"
    d214.Add "actionType", "C"
    d214.Add "checkability", "partial"
    d214.Add "autoFixable", True
    d214.Add "riskLevel", "low"
    d214.Add "title", "T" & ChrW(&HEA) & "n s" & ChrW(&H1EF1) & " ki" & ChrW(&H1EC7) & "n l" & ChrW(&H1ECB) & "ch s" & ChrW(&H1EED) & ", tri" & ChrW(&H1EC1) & "u " & ChrW(&H111) & ChrW(&H1EA1) & "i ch" & ChrW(&H1B0) & "a vi" & ChrW(&H1EBF) & "t hoa " & ChrW(&H111) & ChrW(&HFA) & "ng"
    Dim tmp15 As String
    tmp15 = "'{actual}' n" & ChrW(&HEA) & "n vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '{expected}'. Vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u c" & ChrW(&H1EE7) & "a c" & ChrW(&HE1) & "c " & ChrW(&HE2) & "m ti" & ChrW(&H1EBF) & "t t" & ChrW(&H1EA1) & "o th" & ChrW(&HE0) & "nh s" & ChrW(&H1EF1) & " ki" & ChrW(&H1EC7) & "n v" & ChrW(&HE0) & " t" & ChrW(&HEA) & "n s" & ChrW(&H1EF1) & " ki" & ChrW(&H1EC7) & "n; con s" & ChrW(&H1ED1) & " ch" & ChrW(&H1EC9) & " m" & ChrW(&H1ED1) & "c th" & ChrW(&H1EDD) & "i gian ghi b" & ChrW(&H1EB1) & "ng ch" & ChrW(&H1EEF) & " v" & ChrW(&HE0) & " vi" & ChrW(&H1EBF) & "t hoa ch"
    tmp15 = tmp15 & ChrW(&H1EEF) & " " & ChrW(&H111) & ChrW(&HF3) & "."
    d214.Add "message", tmp15
    d214.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c V, kho" & ChrW(&H1EA3) & "n 9"
    d214.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d214
    Dim d215 As Object
    Set d215 = CreateObject("Scripting.Dictionary")
    d215.Add "ruleCode", "ND30-PL2-M5-K10"
    d215.Add "checklistGroup", 14
    d215.Add "group", "capitalization"
    d215.Add "severity", "info"
    d215.Add "sourceLabel", "ND30"
    d215.Add "actionType", "C"
    d215.Add "checkability", "warnOnly"
    d215.Add "autoFixable", False
    d215.Add "riskLevel", "low"
    d215.Add "title", "T" & ChrW(&HEA) & "n t" & ChrW(&HE1) & "c ph" & ChrW(&H1EA9) & "m, s" & ChrW(&HE1) & "ch b" & ChrW(&HE1) & "o, t" & ChrW(&H1EA1) & "p ch" & ChrW(&HED)
    d215.Add "message", "Vi" & ChrW(&H1EBF) & "t hoa ch" & ChrW(&H1EEF) & " c" & ChrW(&HE1) & "i " & ChrW(&H111) & ChrW(&H1EA7) & "u c" & ChrW(&H1EE7) & "a " & ChrW(&HE2) & "m ti" & ChrW(&H1EBF) & "t th" & ChrW(&H1EE9) & " nh" & ChrW(&H1EA5) & "t t" & ChrW(&H1EA1) & "o th" & ChrW(&HE0) & "nh t" & ChrW(&HEA) & "n t" & ChrW(&HE1) & "c ph" & ChrW(&H1EA9) & "m, s" & ChrW(&HE1) & "ch b" & ChrW(&HE1) & "o. S" & ChrW(&H1EA3) & "n ph" & ChrW(&H1EA9) & "m kh" & ChrW(&HF4) & "ng nh" & ChrW(&H1EAD) & "n di" & ChrW(&H1EC7) & "n " & ChrW(&H111) & ChrW(&H1B0) & ChrW(&H1EE3) & "c b" & ChrW(&H1EB1) & "ng lu" & ChrW(&H1EAD) & "t."
    d215.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c V, kho" & ChrW(&H1EA3) & "n 10"
    d215.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d215
    Dim d216 As Object
    Set d216 = CreateObject("Scripting.Dictionary")
    d216.Add "ruleCode", "QD1989-D8-TONE-MIX"
    d216.Add "checklistGroup", 13
    d216.Add "group", "spelling"
    d216.Add "severity", "warning"
    d216.Add "sourceLabel", "QD1989"
    d216.Add "actionType", "B"
    d216.Add "checkability", "full"
    d216.Add "autoFixable", True
    d216.Add "riskLevel", "low"
    d216.Add "title", "V" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n d" & ChrW(&HF9) & "ng l" & ChrW(&H1EAB) & "n hai ki" & ChrW(&H1EC3) & "u " & ChrW(&H111) & ChrW(&H1EB7) & "t d" & ChrW(&H1EA5) & "u thanh"
    d216.Add "message", "Ph" & ChrW(&HE1) & "t hi" & ChrW(&H1EC7) & "n {countA} ch" & ChrW(&H1ED7) & " ki" & ChrW(&H1EC3) & "u '" & ChrW(&HF2) & "a, " & ChrW(&HFA) & "y' v" & ChrW(&HE0) & " {countB} ch" & ChrW(&H1ED7) & " ki" & ChrW(&H1EC3) & "u 'o" & ChrW(&HE0) & ", u" & ChrW(&HFD) & "' trong c" & ChrW(&HF9) & "ng v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n."
    d216.Add "citation", "Quy" & ChrW(&H1EBF) & "t " & ChrW(&H111) & ChrW(&H1ECB) & "nh s" & ChrW(&H1ED1) & " 1989/Q" & ChrW(&H110) & "-BGD" & ChrW(&H110) & "T ng" & ChrW(&HE0) & "y 25 th" & ChrW(&HE1) & "ng 5 n" & ChrW(&H103) & "m 2018, " & ChrW(&H110) & "i" & ChrW(&H1EC1) & "u 8"
    d216.Add "toolHint", "Ki" & ChrW(&H1EC3) & "u o" & ChrW(&HE0) & ", u" & ChrW(&HFD)
    c.Add d216
    Dim d217 As Object
    Set d217 = CreateObject("Scripting.Dictionary")
    d217.Add "ruleCode", "QD1989-IY-MIX"
    d217.Add "checklistGroup", 13
    d217.Add "group", "spelling"
    d217.Add "severity", "warning"
    d217.Add "sourceLabel", "QD1989"
    d217.Add "actionType", "B"
    d217.Add "checkability", "full"
    d217.Add "autoFixable", True
    d217.Add "riskLevel", "low"
    d217.Add "title", "V" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n d" & ChrW(&HF9) & "ng l" & ChrW(&H1EAB) & "n 'i' v" & ChrW(&HE0) & " 'y' cho c" & ChrW(&HF9) & "ng m" & ChrW(&H1ED9) & "t t" & ChrW(&H1EEB)
    d217.Add "message", "Ph" & ChrW(&HE1) & "t hi" & ChrW(&H1EC7) & "n c" & ChrW(&H1EA3) & " '{formI}' v" & ChrW(&HE0) & " '{formY}' trong c" & ChrW(&HF9) & "ng v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n."
    d217.Add "citation", "Quy" & ChrW(&H1EBF) & "t " & ChrW(&H111) & ChrW(&H1ECB) & "nh s" & ChrW(&H1ED1) & " 1989/Q" & ChrW(&H110) & "-BGD" & ChrW(&H110) & "T ng" & ChrW(&HE0) & "y 25 th" & ChrW(&HE1) & "ng 5 n" & ChrW(&H103) & "m 2018"
    d217.Add "toolHint", "Ki" & ChrW(&H1EC3) & "u i / Ki" & ChrW(&H1EC3) & "u y"
    c.Add d217
    Dim d218 As Object
    Set d218 = CreateObject("Scripting.Dictionary")
    d218.Add "ruleCode", "LOCAL-TYPO-SPACE"
    d218.Add "checklistGroup", 14
    d218.Add "group", "spelling"
    d218.Add "severity", "info"
    d218.Add "sourceLabel", "THONG LE"
    d218.Add "actionType", "C"
    d218.Add "checkability", "full"
    d218.Add "autoFixable", True
    d218.Add "riskLevel", "low"
    d218.Add "title", "D" & ChrW(&H1EA5) & "u c" & ChrW(&HE1) & "ch ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng"
    d218.Add "message", "'{actual}' n" & ChrW(&HEA) & "n vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '{expected}'."
    d218.Add "citation", "Kh" & ChrW(&HF4) & "ng ph" & ChrW(&H1EA3) & "i quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh c" & ChrW(&H1EE7) & "a N" & ChrW(&H110) & " 30 " & ChrW(&H2014) & " chu" & ChrW(&H1EA9) & "n so" & ChrW(&H1EA1) & "n th" & ChrW(&H1EA3) & "o ph" & ChrW(&H1ED5) & " bi" & ChrW(&H1EBF) & "n"
    d218.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d218
    Dim d219 As Object
    Set d219 = CreateObject("Scripting.Dictionary")
    d219.Add "ruleCode", "LOCAL-TYPO-PUNCT"
    d219.Add "checklistGroup", 14
    d219.Add "group", "spelling"
    d219.Add "severity", "info"
    d219.Add "sourceLabel", "THONG LE"
    d219.Add "actionType", "C"
    d219.Add "checkability", "full"
    d219.Add "autoFixable", True
    d219.Add "riskLevel", "low"
    d219.Add "title", "Kho" & ChrW(&H1EA3) & "ng tr" & ChrW(&H1EAF) & "ng quanh d" & ChrW(&H1EA5) & "u c" & ChrW(&HE2) & "u ch" & ChrW(&H1B0) & "a " & ChrW(&H111) & ChrW(&HFA) & "ng"
    d219.Add "message", "'{actual}' n" & ChrW(&HEA) & "n vi" & ChrW(&H1EBF) & "t l" & ChrW(&HE0) & " '{expected}'."
    d219.Add "citation", "Kh" & ChrW(&HF4) & "ng ph" & ChrW(&H1EA3) & "i quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh c" & ChrW(&H1EE7) & "a N" & ChrW(&H110) & " 30 " & ChrW(&H2014) & " chu" & ChrW(&H1EA9) & "n so" & ChrW(&H1EA1) & "n th" & ChrW(&H1EA3) & "o ph" & ChrW(&H1ED5) & " bi" & ChrW(&H1EBF) & "n"
    d219.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d219
    Dim d220 As Object
    Set d220 = CreateObject("Scripting.Dictionary")
    d220.Add "ruleCode", "LOCAL-TYPO-HIDDEN"
    d220.Add "checklistGroup", 14
    d220.Add "group", "spelling"
    d220.Add "severity", "info"
    d220.Add "sourceLabel", "THONG LE"
    d220.Add "actionType", "C"
    d220.Add "checkability", "full"
    d220.Add "autoFixable", True
    d220.Add "riskLevel", "low"
    d220.Add "title", "C" & ChrW(&HF3) & " k" & ChrW(&HFD) & " t" & ChrW(&H1EF1) & " " & ChrW(&H1EA9) & "n"
    d220.Add "message", "Ph" & ChrW(&HE1) & "t hi" & ChrW(&H1EC7) & "n {count} k" & ChrW(&HFD) & " t" & ChrW(&H1EF1) & " " & ChrW(&H1EA9) & "n lo" & ChrW(&H1EA1) & "i {charName}."
    d220.Add "citation", "Kh" & ChrW(&HF4) & "ng ph" & ChrW(&H1EA3) & "i quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh c" & ChrW(&H1EE7) & "a N" & ChrW(&H110) & " 30 " & ChrW(&H2014) & " chu" & ChrW(&H1EA9) & "n so" & ChrW(&H1EA1) & "n th" & ChrW(&H1EA3) & "o ph" & ChrW(&H1ED5) & " bi" & ChrW(&H1EBF) & "n"
    d220.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d220
    Dim d221 As Object
    Set d221 = CreateObject("Scripting.Dictionary")
    d221.Add "ruleCode", "LOCAL-TYPO-DICT"
    d221.Add "checklistGroup", 14
    d221.Add "group", "spelling"
    d221.Add "severity", "info"
    d221.Add "sourceLabel", "THONG LE"
    d221.Add "actionType", "C"
    d221.Add "checkability", "partial"
    d221.Add "autoFixable", True
    d221.Add "riskLevel", "low"
    d221.Add "title", "C" & ChrW(&HF3) & " th" & ChrW(&H1EC3) & " sai ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    d221.Add "message", "'{actual}' " & ChrW(&H2014) & " d" & ChrW(&H1EA1) & "ng " & ChrW(&H111) & ChrW(&HFA) & "ng th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng d" & ChrW(&HF9) & "ng l" & ChrW(&HE0) & " '{expected}'."
    d221.Add "citation", "Kh" & ChrW(&HF4) & "ng ph" & ChrW(&H1EA3) & "i quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh c" & ChrW(&H1EE7) & "a N" & ChrW(&H110) & " 30 " & ChrW(&H2014) & " t" & ChrW(&H1EEB) & " " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "n l" & ChrW(&H1ED7) & "i ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3) & " ph" & ChrW(&H1ED5) & " bi" & ChrW(&H1EBF) & "n"
    d221.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d221
    Dim d222 As Object
    Set d222 = CreateObject("Scripting.Dictionary")
    d222.Add "ruleCode", "LOCAL-TYPO-TELEX"
    d222.Add "checklistGroup", 14
    d222.Add "group", "spelling"
    d222.Add "severity", "info"
    d222.Add "sourceLabel", "THONG LE"
    d222.Add "actionType", "C"
    d222.Add "checkability", "partial"
    d222.Add "autoFixable", True
    d222.Add "riskLevel", "low"
    d222.Add "title", "C" & ChrW(&HF3) & " th" & ChrW(&H1EC3) & " l" & ChrW(&HE0) & " l" & ChrW(&H1ED7) & "i telex ch" & ChrW(&H1B0) & "a chuy" & ChrW(&H1EC3) & "n"
    d222.Add "message", "'{actual}' c" & ChrW(&HF3) & " th" & ChrW(&H1EC3) & " l" & ChrW(&HE0) & " '{expected}' do b" & ChrW(&H1ED9) & " g" & ChrW(&HF5) & " ch" & ChrW(&H1B0) & "a chuy" & ChrW(&H1EC3) & "n xong."
    d222.Add "citation", "Kh" & ChrW(&HF4) & "ng ph" & ChrW(&H1EA3) & "i quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh c" & ChrW(&H1EE7) & "a N" & ChrW(&H110) & " 30"
    d222.Add "toolHint", "R" & ChrW(&HE0) & " so" & ChrW(&HE1) & "t ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3)
    c.Add d222
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCheckRules_RulesPart5", Err.description
End Sub

Private Function LoadRawCheckRules_Rules() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    LoadRawCheckRules_RulesPart1 c
    LoadRawCheckRules_RulesPart2 c
    LoadRawCheckRules_RulesPart3 c
    LoadRawCheckRules_RulesPart4 c
    LoadRawCheckRules_RulesPart5 c
    Set LoadRawCheckRules_Rules = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCheckRules_Rules", Err.description
End Function

Public Function LoadRawCheckRules() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "1.1.0"
    d.Add "checklistGroups", LoadRawCheckRules_ChecklistGroups()
    d.Add "rules", LoadRawCheckRules_Rules()
    Set LoadRawCheckRules = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCheckRules", Err.description
End Function

Private Function LoadRawTelexWhitelist_Patterns() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim d223 As Object
    Set d223 = CreateObject("Scripting.Dictionary")
    d223.Add "pattern", "\bddaa\w*"
    d223.Add "replacement", ChrW(&H111) & ChrW(&HE2) & "..."
    d223.Add "note", "vi du: ddaay -> day (con thieu dau, chi doi phan dau)"
    c.Add d223
    Dim d224 As Object
    Set d224 = CreateObject("Scripting.Dictionary")
    d224.Add "pattern", "\bddoo\w*"
    d224.Add "replacement", ChrW(&H111) & ChrW(&HF4) & "..."
    d224.Add "note", "vi du: ddoong -> dong"
    c.Add d224
    Dim d225 As Object
    Set d225 = CreateObject("Scripting.Dictionary")
    d225.Add "pattern", "\bddee\w*"
    d225.Add "replacement", ChrW(&H111) & ChrW(&HEA) & "..."
    d225.Add "note", "vi du: ddeen -> den"
    c.Add d225
    Dim d226 As Object
    Set d226 = CreateObject("Scripting.Dictionary")
    d226.Add "pattern", "\bddoongs\b"
    d226.Add "replacement", ChrW(&H111) & ChrW(&HF4) & "ng"
    d226.Add "wordExample", True
    c.Add d226
    Dim d227 As Object
    Set d227 = CreateObject("Scripting.Dictionary")
    d227.Add "pattern", "\bnhuwng\b"
    d227.Add "replacement", "nh" & ChrW(&H1B0) & "ng"
    d227.Add "wordExample", True
    c.Add d227
    Dim d228 As Object
    Set d228 = CreateObject("Scripting.Dictionary")
    d228.Add "pattern", "\bcuwr\w*"
    d228.Add "replacement", "c" & ChrW(&H1EED) & "..."
    d228.Add "note", "vi du: cuwr tri -> c" & ChrW(&H1EED) & " tri"
    c.Add d228
    Dim d229 As Object
    Set d229 = CreateObject("Scripting.Dictionary")
    d229.Add "pattern", "\bnawm\b"
    d229.Add "replacement", "n" & ChrW(&H103) & "m"
    d229.Add "wordExample", True
    c.Add d229
    Dim d230 As Object
    Set d230 = CreateObject("Scripting.Dictionary")
    d230.Add "pattern", "\bthaangs\b"
    d230.Add "replacement", "th" & ChrW(&HE1) & "ng"
    d230.Add "wordExample", True
    c.Add d230
    Set LoadRawTelexWhitelist_Patterns = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTelexWhitelist_Patterns", Err.description
End Function

Public Function LoadRawTelexWhitelist() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "1.0.0"
    d.Add "status", "hat giong - rui ro cao, uu tien thap nhat trong nhom chinh ta"
    d.Add "sourceLabel", "TH" & ChrW(&HD4) & "NG L" & ChrW(&H1EC6)
    d.Add "actionType", "C"
    d.Add "riskLevel", "high"
    d.Add "patterns", LoadRawTelexWhitelist_Patterns()
    Set LoadRawTelexWhitelist = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTelexWhitelist", Err.description
End Function

Private Function LoadRawCommonOrganNames_Organs() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "Qu" & ChrW(&H1ED1) & "c h" & ChrW(&H1ED9) & "i"
    c.Add ChrW(&H1EE6) & "y ban Th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng v" & ChrW(&H1EE5) & " Qu" & ChrW(&H1ED1) & "c h" & ChrW(&H1ED9) & "i"
    c.Add "Ch" & ChrW(&H1EE7) & " t" & ChrW(&H1ECB) & "ch n" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add "Ch" & ChrW(&HED) & "nh ph" & ChrW(&H1EE7)
    c.Add "Th" & ChrW(&H1EE7) & " t" & ChrW(&H1B0) & ChrW(&H1EDB) & "ng Ch" & ChrW(&HED) & "nh ph" & ChrW(&H1EE7)
    c.Add "V" & ChrW(&H103) & "n ph" & ChrW(&HF2) & "ng Ch" & ChrW(&HED) & "nh ph" & ChrW(&H1EE7)
    c.Add "V" & ChrW(&H103) & "n ph" & ChrW(&HF2) & "ng Qu" & ChrW(&H1ED1) & "c h" & ChrW(&H1ED9) & "i"
    c.Add "V" & ChrW(&H103) & "n ph" & ChrW(&HF2) & "ng Ch" & ChrW(&H1EE7) & " t" & ChrW(&H1ECB) & "ch n" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add "T" & ChrW(&HF2) & "a " & ChrW(&HE1) & "n nh" & ChrW(&HE2) & "n d" & ChrW(&HE2) & "n t" & ChrW(&H1ED1) & "i cao"
    c.Add "Vi" & ChrW(&H1EC7) & "n ki" & ChrW(&H1EC3) & "m s" & ChrW(&HE1) & "t nh" & ChrW(&HE2) & "n d" & ChrW(&HE2) & "n t" & ChrW(&H1ED1) & "i cao"
    c.Add "Ki" & ChrW(&H1EC3) & "m to" & ChrW(&HE1) & "n nh" & ChrW(&HE0) & " n" & ChrW(&H1B0) & ChrW(&H1EDB) & "c"
    c.Add ChrW(&H1EE6) & "y ban nh" & ChrW(&HE2) & "n d" & ChrW(&HE2) & "n"
    c.Add "H" & ChrW(&H1ED9) & "i " & ChrW(&H111) & ChrW(&H1ED3) & "ng nh" & ChrW(&HE2) & "n d" & ChrW(&HE2) & "n"
    c.Add "M" & ChrW(&H1EB7) & "t tr" & ChrW(&H1EAD) & "n T" & ChrW(&H1ED5) & " qu" & ChrW(&H1ED1) & "c Vi" & ChrW(&H1EC7) & "t Nam"
    c.Add "Ban Ch" & ChrW(&H1EA5) & "p h" & ChrW(&HE0) & "nh Trung " & ChrW(&H1B0) & ChrW(&H1A1) & "ng " & ChrW(&H110) & ChrW(&H1EA3) & "ng C" & ChrW(&H1ED9) & "ng s" & ChrW(&H1EA3) & "n Vi" & ChrW(&H1EC7) & "t Nam"
    c.Add "V" & ChrW(&H103) & "n ph" & ChrW(&HF2) & "ng Trung " & ChrW(&H1B0) & ChrW(&H1A1) & "ng " & ChrW(&H110) & ChrW(&H1EA3) & "ng"
    Set LoadRawCommonOrganNames_Organs = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCommonOrganNames_Organs", Err.description
End Function

Public Function LoadRawCommonOrganNames() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "1.0.0"
    d.Add "status", "SEED - chi gom co quan on dinh lau dai, CHUA gom du danh sach bo/nganh hien hanh"
    d.Add "sourceLabel", "SUY RA"
    d.Add "actionType", "C"
    d.Add "organs", LoadRawCommonOrganNames_Organs()
    Set LoadRawCommonOrganNames = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCommonOrganNames", Err.description
End Function

Private Function LoadRawFormatSpec_LegalBasis() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "document", "Ngh" & ChrW(&H1ECB) & " " & ChrW(&H111) & ChrW(&H1ECB) & "nh s" & ChrW(&H1ED1) & " 30/2020/N" & ChrW(&H110) & "-CP"
    d.Add "issuedDate", "2020-03-05"
    d.Add "issuer", "Ch" & ChrW(&HED) & "nh ph" & ChrW(&H1EE7)
    d.Add "appendix", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I"
    Set LoadRawFormatSpec_LegalBasis = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_LegalBasis", Err.description
End Function

Private Function LoadRawFormatSpec_Units() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "mmToPoint", 2.834645669
    d.Add "cmToPoint", 28.34645669
    Set LoadRawFormatSpec_Units = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Units", Err.description
End Function

Private Function LoadRawFormatSpec_PageSetup_Margins_TopMm() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 20
    d.Add "max", 25
    d.Add "default", 25
    Set LoadRawFormatSpec_PageSetup_Margins_TopMm = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_PageSetup_Margins_TopMm", Err.description
End Function

Private Function LoadRawFormatSpec_PageSetup_Margins_BottomMm() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 20
    d.Add "max", 25
    d.Add "default", 20
    Set LoadRawFormatSpec_PageSetup_Margins_BottomMm = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_PageSetup_Margins_BottomMm", Err.description
End Function

Private Function LoadRawFormatSpec_PageSetup_Margins_LeftMm() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 30
    d.Add "max", 35
    d.Add "default", 35
    Set LoadRawFormatSpec_PageSetup_Margins_LeftMm = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_PageSetup_Margins_LeftMm", Err.description
End Function

Private Function LoadRawFormatSpec_PageSetup_Margins_RightMm() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 15
    d.Add "max", 20
    d.Add "default", 20
    Set LoadRawFormatSpec_PageSetup_Margins_RightMm = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_PageSetup_Margins_RightMm", Err.description
End Function

Private Function LoadRawFormatSpec_PageSetup_Margins() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "topMm", LoadRawFormatSpec_PageSetup_Margins_TopMm()
    d.Add "bottomMm", LoadRawFormatSpec_PageSetup_Margins_BottomMm()
    d.Add "leftMm", LoadRawFormatSpec_PageSetup_Margins_LeftMm()
    d.Add "rightMm", LoadRawFormatSpec_PageSetup_Margins_RightMm()
    Set LoadRawFormatSpec_PageSetup_Margins = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_PageSetup_Margins", Err.description
End Function

Private Function LoadRawFormatSpec_PageSetup() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "paperSize", "A4"
    d.Add "pageWidthMm", 210
    d.Add "pageHeightMm", 297
    d.Add "orientation", "portrait"
    d.Add "orientationExceptionAllowed", True
    d.Add "margins", LoadRawFormatSpec_PageSetup_Margins()
    d.Add "headerDistanceMm", 10
    Set LoadRawFormatSpec_PageSetup = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_PageSetup", Err.description
End Function

Private Function LoadRawFormatSpec_Font() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "name", "Times New Roman"
    d.Add "charset", "Unicode TCVN 6909:2001"
    d.Add "color", "#000000"
    Set LoadRawFormatSpec_Font = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Font", Err.description
End Function

Private Function LoadRawFormatSpec_PageNumber() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "enabled", True
    d.Add "startAt", 1
    d.Add "numeralType", "arabic"
    d.Add "position", "topMargin"
    d.Add "alignment", "center"
    d.Add "hideOnFirstPage", True
    d.Add "sizeSetKey", "pageNumber"
    Set LoadRawFormatSpec_PageNumber = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_PageNumber", Err.description
End Function

Private Function LoadRawFormatSpec_BodyText_FirstLineIndentCm_Allowed() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add 1
    c.Add 1.27
    Set LoadRawFormatSpec_BodyText_FirstLineIndentCm_Allowed = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_BodyText_FirstLineIndentCm_Allowed", Err.description
End Function

Private Function LoadRawFormatSpec_BodyText_FirstLineIndentCm() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "allowed", LoadRawFormatSpec_BodyText_FirstLineIndentCm_Allowed()
    d.Add "default", 1
    Set LoadRawFormatSpec_BodyText_FirstLineIndentCm = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_BodyText_FirstLineIndentCm", Err.description
End Function

Private Function LoadRawFormatSpec_BodyText_SpaceAfterPt() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 6
    d.Add "default", 6
    Set LoadRawFormatSpec_BodyText_SpaceAfterPt = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_BodyText_SpaceAfterPt", Err.description
End Function

Private Function LoadRawFormatSpec_BodyText_LineSpacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", "single"
    d.Add "max", "1.5lines"
    d.Add "default", "single"
    Set LoadRawFormatSpec_BodyText_LineSpacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_BodyText_LineSpacing", Err.description
End Function

Private Function LoadRawFormatSpec_BodyText() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "alignment", "justify"
    d.Add "firstLineIndentCm", LoadRawFormatSpec_BodyText_FirstLineIndentCm()
    d.Add "spaceAfterPt", LoadRawFormatSpec_BodyText_SpaceAfterPt()
    d.Add "lineSpacing", LoadRawFormatSpec_BodyText_LineSpacing()
    Set LoadRawFormatSpec_BodyText = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_BodyText", Err.description
End Function

Private Function LoadRawFormatSpec_CharSpacing() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "stepPt", 0.1
    d.Add "minPt", -7
    d.Add "maxPt", 7
    Set LoadRawFormatSpec_CharSpacing = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_CharSpacing", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeSets_Set1() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "label", "C" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & " 14"
    d.Add "nationalTitle", 13
    d.Add "nationalMotto", 14
    d.Add "partyHeader", 15
    d.Add "starSeparator", 14
    d.Add "superiorOrganName", 13
    d.Add "organName", 13
    d.Add "codeNumberNotation", 13
    d.Add "placeAndIssuedDate", 14
    d.Add "typeName", 14
    d.Add "subject", 14
    d.Add "subjectOfficialLetter", 13
    d.Add "recipientSalutation", 14
    d.Add "recipientSalutationList", 14
    d.Add "recipientSalutationInline", 14
    d.Add "recipientSalutationInlineContent", 14
    d.Add "legalBasis", 14
    d.Add "bodyText", 14
    d.Add "signerAuthority", 14
    d.Add "signerAuthorityTitle", 14
    d.Add "recipientLabel", 12
    d.Add "recipientList", 11
    d.Add "appendixLabel", 14
    d.Add "appendixTitle", 14
    d.Add "appendixReference", 14
    d.Add "pageNumber", 13
    Set LoadRawFormatSpec_FontSizeSets_Set1 = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeSets_Set1", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeSets_Set2() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "label", "C" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & " 13"
    d.Add "nationalTitle", 12
    d.Add "nationalMotto", 13
    d.Add "partyHeader", 15
    d.Add "starSeparator", 13
    d.Add "superiorOrganName", 12
    d.Add "organName", 12
    d.Add "codeNumberNotation", 13
    d.Add "placeAndIssuedDate", 13
    d.Add "typeName", 13
    d.Add "subject", 13
    d.Add "subjectOfficialLetter", 12
    d.Add "recipientSalutation", 13
    d.Add "recipientSalutationList", 13
    d.Add "recipientSalutationInline", 13
    d.Add "recipientSalutationInlineContent", 13
    d.Add "legalBasis", 13
    d.Add "bodyText", 13
    d.Add "signerAuthority", 13
    d.Add "signerAuthorityTitle", 13
    d.Add "recipientLabel", 12
    d.Add "recipientList", 11
    d.Add "appendixLabel", 14
    d.Add "appendixTitle", 13
    d.Add "appendixReference", 13
    d.Add "pageNumber", 13
    Set LoadRawFormatSpec_FontSizeSets_Set2 = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeSets_Set2", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeSets_Set3() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "label", "C" & ChrW(&H1EE1) & " ch" & ChrW(&H1EEF) & " 15"
    d.Add "nationalTitle", 15
    d.Add "nationalMotto", 15
    d.Add "partyHeader", 15
    d.Add "starSeparator", 14
    d.Add "superiorOrganName", 14
    d.Add "organName", 14
    d.Add "codeNumberNotation", 14
    d.Add "placeAndIssuedDate", 14
    d.Add "typeName", 16
    d.Add "subject", 15
    d.Add "subjectOfficialLetter", 12
    d.Add "recipientSalutation", 14
    d.Add "recipientSalutationList", 15
    d.Add "recipientSalutationInline", 14
    d.Add "recipientSalutationInlineContent", 15
    d.Add "legalBasis", 15
    d.Add "bodyText", 15
    d.Add "signerAuthority", 14
    d.Add "signerAuthorityTitle", 14
    d.Add "recipientLabel", 14
    d.Add "recipientList", 12
    d.Add "appendixLabel", 15
    d.Add "appendixTitle", 15
    d.Add "appendixReference", 15
    d.Add "pageNumber", 14
    Set LoadRawFormatSpec_FontSizeSets_Set3 = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeSets_Set3", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeSets() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "set1", LoadRawFormatSpec_FontSizeSets_Set1()
    d.Add "set2", LoadRawFormatSpec_FontSizeSets_Set2()
    d.Add "set3", LoadRawFormatSpec_FontSizeSets_Set3()
    Set LoadRawFormatSpec_FontSizeSets = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeSets", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_NationalTitle() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 12
    d.Add "max", 13
    Set LoadRawFormatSpec_FontSizeRanges_NationalTitle = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_NationalTitle", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_NationalMotto() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 13
    d.Add "max", 14
    Set LoadRawFormatSpec_FontSizeRanges_NationalMotto = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_NationalMotto", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_PartyHeader() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 15
    d.Add "max", 15
    Set LoadRawFormatSpec_FontSizeRanges_PartyHeader = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_PartyHeader", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_StarSeparator() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 13
    d.Add "max", 14
    Set LoadRawFormatSpec_FontSizeRanges_StarSeparator = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_StarSeparator", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_SuperiorOrganName() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 12
    d.Add "max", 14
    Set LoadRawFormatSpec_FontSizeRanges_SuperiorOrganName = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_SuperiorOrganName", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_OrganName() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 12
    d.Add "max", 14
    Set LoadRawFormatSpec_FontSizeRanges_OrganName = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_OrganName", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_CodeNumberNotation() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 13
    d.Add "max", 14
    Set LoadRawFormatSpec_FontSizeRanges_CodeNumberNotation = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_CodeNumberNotation", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_PlaceAndIssuedDate() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 13
    d.Add "max", 14
    Set LoadRawFormatSpec_FontSizeRanges_PlaceAndIssuedDate = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_PlaceAndIssuedDate", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_TypeName() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 13
    d.Add "max", 16
    Set LoadRawFormatSpec_FontSizeRanges_TypeName = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_TypeName", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_Subject() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 13
    d.Add "max", 15
    Set LoadRawFormatSpec_FontSizeRanges_Subject = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_Subject", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_SubjectOfficialLetter() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 12
    d.Add "max", 13
    Set LoadRawFormatSpec_FontSizeRanges_SubjectOfficialLetter = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_SubjectOfficialLetter", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_RecipientSalutation() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 13
    d.Add "max", 14
    Set LoadRawFormatSpec_FontSizeRanges_RecipientSalutation = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_RecipientSalutation", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_RecipientSalutationList() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 13
    d.Add "max", 15
    Set LoadRawFormatSpec_FontSizeRanges_RecipientSalutationList = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_RecipientSalutationList", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_RecipientSalutationInline() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 13
    d.Add "max", 14
    Set LoadRawFormatSpec_FontSizeRanges_RecipientSalutationInline = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_RecipientSalutationInline", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_RecipientSalutationInlineContent() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 13
    d.Add "max", 15
    Set LoadRawFormatSpec_FontSizeRanges_RecipientSalutationInlineContent = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_RecipientSalutationInlineContent", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_LegalBasis() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 13
    d.Add "max", 15
    Set LoadRawFormatSpec_FontSizeRanges_LegalBasis = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_LegalBasis", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_BodyText() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 13
    d.Add "max", 15
    Set LoadRawFormatSpec_FontSizeRanges_BodyText = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_BodyText", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_SignerAuthority() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 13
    d.Add "max", 14
    Set LoadRawFormatSpec_FontSizeRanges_SignerAuthority = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_SignerAuthority", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_SignerAuthorityTitle() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 13
    d.Add "max", 14
    Set LoadRawFormatSpec_FontSizeRanges_SignerAuthorityTitle = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_SignerAuthorityTitle", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_RecipientLabel() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 12
    d.Add "max", 14
    Set LoadRawFormatSpec_FontSizeRanges_RecipientLabel = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_RecipientLabel", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_RecipientList() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 11
    d.Add "max", 12
    Set LoadRawFormatSpec_FontSizeRanges_RecipientList = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_RecipientList", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_AppendixLabel() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 14
    d.Add "max", 15
    Set LoadRawFormatSpec_FontSizeRanges_AppendixLabel = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_AppendixLabel", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_AppendixTitle() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 13
    d.Add "max", 15
    Set LoadRawFormatSpec_FontSizeRanges_AppendixTitle = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_AppendixTitle", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_AppendixReference() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 13
    d.Add "max", 15
    Set LoadRawFormatSpec_FontSizeRanges_AppendixReference = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_AppendixReference", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges_PageNumber() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "min", 13
    d.Add "max", 14
    Set LoadRawFormatSpec_FontSizeRanges_PageNumber = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges_PageNumber", Err.description
End Function

Private Function LoadRawFormatSpec_FontSizeRanges() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "nationalTitle", LoadRawFormatSpec_FontSizeRanges_NationalTitle()
    d.Add "nationalMotto", LoadRawFormatSpec_FontSizeRanges_NationalMotto()
    d.Add "partyHeader", LoadRawFormatSpec_FontSizeRanges_PartyHeader()
    d.Add "starSeparator", LoadRawFormatSpec_FontSizeRanges_StarSeparator()
    d.Add "superiorOrganName", LoadRawFormatSpec_FontSizeRanges_SuperiorOrganName()
    d.Add "organName", LoadRawFormatSpec_FontSizeRanges_OrganName()
    d.Add "codeNumberNotation", LoadRawFormatSpec_FontSizeRanges_CodeNumberNotation()
    d.Add "placeAndIssuedDate", LoadRawFormatSpec_FontSizeRanges_PlaceAndIssuedDate()
    d.Add "typeName", LoadRawFormatSpec_FontSizeRanges_TypeName()
    d.Add "subject", LoadRawFormatSpec_FontSizeRanges_Subject()
    d.Add "subjectOfficialLetter", LoadRawFormatSpec_FontSizeRanges_SubjectOfficialLetter()
    d.Add "recipientSalutation", LoadRawFormatSpec_FontSizeRanges_RecipientSalutation()
    d.Add "recipientSalutationList", LoadRawFormatSpec_FontSizeRanges_RecipientSalutationList()
    d.Add "recipientSalutationInline", LoadRawFormatSpec_FontSizeRanges_RecipientSalutationInline()
    d.Add "recipientSalutationInlineContent", LoadRawFormatSpec_FontSizeRanges_RecipientSalutationInlineContent()
    d.Add "legalBasis", LoadRawFormatSpec_FontSizeRanges_LegalBasis()
    d.Add "bodyText", LoadRawFormatSpec_FontSizeRanges_BodyText()
    d.Add "signerAuthority", LoadRawFormatSpec_FontSizeRanges_SignerAuthority()
    d.Add "signerAuthorityTitle", LoadRawFormatSpec_FontSizeRanges_SignerAuthorityTitle()
    d.Add "recipientLabel", LoadRawFormatSpec_FontSizeRanges_RecipientLabel()
    d.Add "recipientList", LoadRawFormatSpec_FontSizeRanges_RecipientList()
    d.Add "appendixLabel", LoadRawFormatSpec_FontSizeRanges_AppendixLabel()
    d.Add "appendixTitle", LoadRawFormatSpec_FontSizeRanges_AppendixTitle()
    d.Add "appendixReference", LoadRawFormatSpec_FontSizeRanges_AppendixReference()
    d.Add "pageNumber", LoadRawFormatSpec_FontSizeRanges_PageNumber()
    Set LoadRawFormatSpec_FontSizeRanges = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_FontSizeRanges", Err.description
End Function

Private Function LoadRawFormatSpec_Components_NationalTitle_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", True
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_NationalTitle_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_NationalTitle_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_NationalTitle() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "cellNumber", "1"
    d.Add "text", "C" & ChrW(&H1ED8) & "NG H" & ChrW(&HD2) & "A X" & ChrW(&HC3) & " H" & ChrW(&H1ED8) & "I CH" & ChrW(&H1EE6) & " NGH" & ChrW(&H128) & "A VI" & ChrW(&H1EC6) & "T NAM"
    d.Add "letterCase", "upper"
    d.Add "style", LoadRawFormatSpec_Components_NationalTitle_Style()
    d.Add "alignment", "center"
    d.Add "zone", "topRight"
    d.Add "firstLineIndentCm", 0
    d.Add "spaceBeforePt", 0
    d.Add "lineSpacingZone", "fixed"
    Set LoadRawFormatSpec_Components_NationalTitle = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_NationalTitle", Err.description
End Function

Private Function LoadRawFormatSpec_Components_NationalMotto_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", True
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_NationalMotto_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_NationalMotto_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_NationalMotto_Underline() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "type", "solidLine"
    d.Add "widthRelativeTo", "textLength"
    d.Add "ratioMin", 1
    d.Add "ratioMax", 1
    Set LoadRawFormatSpec_Components_NationalMotto_Underline = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_NationalMotto_Underline", Err.description
End Function

Private Function LoadRawFormatSpec_Components_NationalMotto() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "cellNumber", "1"
    d.Add "text", ChrW(&H110) & ChrW(&H1ED9) & "c l" & ChrW(&H1EAD) & "p - T" & ChrW(&H1EF1) & " do - H" & ChrW(&H1EA1) & "nh ph" & ChrW(&HFA) & "c"
    d.Add "letterCase", "normal"
    d.Add "style", LoadRawFormatSpec_Components_NationalMotto_Style()
    d.Add "alignment", "center"
    d.Add "separator", " - "
    d.Add "underline", LoadRawFormatSpec_Components_NationalMotto_Underline()
    d.Add "firstLineIndentCm", 0
    d.Add "spaceBeforePt", 0
    d.Add "lineSpacingZone", "fixed"
    Set LoadRawFormatSpec_Components_NationalMotto = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_NationalMotto", Err.description
End Function

Private Function LoadRawFormatSpec_Components_PartyHeader_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", True
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_PartyHeader_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_PartyHeader_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_PartyHeader() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "cellNumber", "1"
    d.Add "text", ChrW(&H110) & ChrW(&H1EA2) & "NG C" & ChrW(&H1ED8) & "NG S" & ChrW(&H1EA2) & "N VI" & ChrW(&H1EC6) & "T NAM"
    d.Add "letterCase", "upper"
    d.Add "style", LoadRawFormatSpec_Components_PartyHeader_Style()
    d.Add "alignment", "center"
    d.Add "zone", "topRight"
    d.Add "firstLineIndentCm", 0
    d.Add "spaceBeforePt", 0
    d.Add "lineSpacingZone", "fixed"
    Set LoadRawFormatSpec_Components_PartyHeader = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_PartyHeader", Err.description
End Function

Private Function LoadRawFormatSpec_Components_SuperiorOrganName_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_SuperiorOrganName_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_SuperiorOrganName_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_SuperiorOrganName() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "cellNumber", "2"
    d.Add "letterCase", "upper"
    d.Add "style", LoadRawFormatSpec_Components_SuperiorOrganName_Style()
    d.Add "alignment", "center"
    d.Add "zone", "topLeft"
    d.Add "firstLineIndentCm", 0
    d.Add "spaceBeforePt", 0
    d.Add "lineSpacingZone", "fixed"
    Set LoadRawFormatSpec_Components_SuperiorOrganName = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_SuperiorOrganName", Err.description
End Function

Private Function LoadRawFormatSpec_Components_OrganName_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", True
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_OrganName_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_OrganName_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_OrganName_Underline() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "type", "solidLine"
    d.Add "widthRelativeTo", "textLength"
    d.Add "ratioMin", 0.333
    d.Add "ratioMax", 0.5
    Set LoadRawFormatSpec_Components_OrganName_Underline = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_OrganName_Underline", Err.description
End Function

Private Function LoadRawFormatSpec_Components_OrganName() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "cellNumber", "2"
    d.Add "letterCase", "upper"
    d.Add "style", LoadRawFormatSpec_Components_OrganName_Style()
    d.Add "alignment", "center"
    d.Add "underline", LoadRawFormatSpec_Components_OrganName_Underline()
    d.Add "firstLineIndentCm", 0
    d.Add "spaceBeforePt", 0
    d.Add "lineSpacingZone", "fixed"
    Set LoadRawFormatSpec_Components_OrganName = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_OrganName", Err.description
End Function

Private Function LoadRawFormatSpec_Components_StarSeparator_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_StarSeparator_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_StarSeparator_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_StarSeparator() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "cellNumber", "2"
    d.Add "text", "*"
    d.Add "letterCase", "normal"
    d.Add "style", LoadRawFormatSpec_Components_StarSeparator_Style()
    d.Add "alignment", "center"
    d.Add "firstLineIndentCm", 0
    d.Add "spaceBeforePt", 0
    d.Add "lineSpacingZone", "fixed"
    Set LoadRawFormatSpec_Components_StarSeparator = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_StarSeparator", Err.description
End Function

Private Function LoadRawFormatSpec_Components_CodeNumberNotation_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_CodeNumberNotation_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_CodeNumberNotation_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_CodeNumberNotation() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "cellNumber", "3"
    d.Add "prefix", "S" & ChrW(&H1ED1) & ":"
    d.Add "prefixLetterCase", "normal"
    d.Add "notationLetterCase", "upper"
    d.Add "style", LoadRawFormatSpec_Components_CodeNumberNotation_Style()
    d.Add "alignment", "center"
    d.Add "padNumberBelow", 10
    d.Add "padChar", "0"
    d.Add "separatorNumberNotation", "/"
    d.Add "separatorNotationGroups", "-"
    d.Add "noSpaceInside", True
    d.Add "firstLineIndentCm", 0
    d.Add "spaceBeforePt", 0
    d.Add "lineSpacingZone", "fixed"
    d.Add "blankNotationGapSpaces", 10
    Set LoadRawFormatSpec_Components_CodeNumberNotation = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_CodeNumberNotation", Err.description
End Function

Private Function LoadRawFormatSpec_Components_PlaceAndIssuedDate_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", True
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_PlaceAndIssuedDate_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_PlaceAndIssuedDate_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_PlaceAndIssuedDate_PadMonthsList() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add 1
    c.Add 2
    Set LoadRawFormatSpec_Components_PlaceAndIssuedDate_PadMonthsList = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_PlaceAndIssuedDate_PadMonthsList", Err.description
End Function

Private Function LoadRawFormatSpec_Components_PlaceAndIssuedDate() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "cellNumber", "4"
    d.Add "letterCase", "normal"
    d.Add "style", LoadRawFormatSpec_Components_PlaceAndIssuedDate_Style()
    d.Add "alignment", "center"
    d.Add "pattern", "{placeName}, ng" & ChrW(&HE0) & "y {dd} th" & ChrW(&HE1) & "ng {mm} n" & ChrW(&H103) & "m {yyyy}"
    d.Add "padDayBelow", 10
    d.Add "padMonthsList", LoadRawFormatSpec_Components_PlaceAndIssuedDate_PadMonthsList()
    d.Add "firstLineIndentCm", 0
    d.Add "spaceBeforePt", 0
    d.Add "lineSpacingZone", "fixed"
    d.Add "blankDayGapSpaces", 6
    d.Add "blankMonthGapSpaces", 6
    Set LoadRawFormatSpec_Components_PlaceAndIssuedDate = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_PlaceAndIssuedDate", Err.description
End Function

Private Function LoadRawFormatSpec_Components_SubjectOfficialLetter_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_SubjectOfficialLetter_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_SubjectOfficialLetter_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_SubjectOfficialLetter() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "cellNumber", "5b"
    d.Add "prefix", "V/v"
    d.Add "letterCase", "normal"
    d.Add "style", LoadRawFormatSpec_Components_SubjectOfficialLetter_Style()
    d.Add "alignment", "center"
    d.Add "firstLineIndentCm", 0
    d.Add "spaceBeforePt", 6
    d.Add "lineSpacingZone", "fixed"
    Set LoadRawFormatSpec_Components_SubjectOfficialLetter = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_SubjectOfficialLetter", Err.description
End Function

Private Function LoadRawFormatSpec_Components_TypeName_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", True
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_TypeName_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_TypeName_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_TypeName() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "cellNumber", "5a"
    d.Add "letterCase", "upper"
    d.Add "style", LoadRawFormatSpec_Components_TypeName_Style()
    d.Add "alignment", "center"
    d.Add "firstLineIndentCm", 0
    d.Add "spaceBeforePt", 6
    d.Add "lineSpacingZone", "body"
    Set LoadRawFormatSpec_Components_TypeName = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_TypeName", Err.description
End Function

Private Function LoadRawFormatSpec_Components_Subject_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", True
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_Subject_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_Subject_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_Subject_Underline() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "type", "solidLine"
    d.Add "widthRelativeTo", "textLength"
    d.Add "ratioMin", 0.333
    d.Add "ratioMax", 0.5
    Set LoadRawFormatSpec_Components_Subject_Underline = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_Subject_Underline", Err.description
End Function

Private Function LoadRawFormatSpec_Components_Subject() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "cellNumber", "5a"
    d.Add "letterCase", "normal"
    d.Add "style", LoadRawFormatSpec_Components_Subject_Style()
    d.Add "alignment", "center"
    d.Add "underline", LoadRawFormatSpec_Components_Subject_Underline()
    d.Add "firstLineIndentCm", 0
    d.Add "spaceBeforePt", 6
    d.Add "lineSpacingZone", "body"
    Set LoadRawFormatSpec_Components_Subject = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_Subject", Err.description
End Function

Private Function LoadRawFormatSpec_Components_RecipientSalutation_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_RecipientSalutation_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_RecipientSalutation_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_RecipientSalutation() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "cellNumber", "9a"
    d.Add "label", "K" & ChrW(&HED) & "nh g" & ChrW(&H1EED) & "i:"
    d.Add "letterCase", "normal"
    d.Add "style", LoadRawFormatSpec_Components_RecipientSalutation_Style()
    d.Add "alignment", "justify"
    d.Add "firstLineIndentCm", 1
    d.Add "spaceBeforePt", 6
    d.Add "lineSpacingZone", "body"
    Set LoadRawFormatSpec_Components_RecipientSalutation = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_RecipientSalutation", Err.description
End Function

Private Function LoadRawFormatSpec_Components_RecipientSalutationList_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_RecipientSalutationList_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_RecipientSalutationList_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_RecipientSalutationList() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "cellNumber", "9a"
    d.Add "letterCase", "normal"
    d.Add "style", LoadRawFormatSpec_Components_RecipientSalutationList_Style()
    d.Add "alignment", "justify"
    d.Add "bulletChar", "-"
    d.Add "itemEndChar", ";"
    d.Add "lastItemEndChar", "."
    d.Add "firstLineIndentCm", 3.5
    d.Add "spaceBeforePt", 6
    d.Add "lineSpacingZone", "body"
    Set LoadRawFormatSpec_Components_RecipientSalutationList = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_RecipientSalutationList", Err.description
End Function

Private Function LoadRawFormatSpec_Components_RecipientSalutationInline_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_RecipientSalutationInline_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_RecipientSalutationInline_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_RecipientSalutationInline() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "cellNumber", "9a"
    d.Add "label", "K" & ChrW(&HED) & "nh g" & ChrW(&H1EED) & "i:"
    d.Add "letterCase", "normal"
    d.Add "style", LoadRawFormatSpec_Components_RecipientSalutationInline_Style()
    d.Add "alignment", "center"
    d.Add "firstLineIndentCm", 0
    d.Add "spaceBeforePt", 6
    d.Add "lineSpacingZone", "body"
    Set LoadRawFormatSpec_Components_RecipientSalutationInline = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_RecipientSalutationInline", Err.description
End Function

Private Function LoadRawFormatSpec_Components_RecipientSalutationInlineContent_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_RecipientSalutationInlineContent_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_RecipientSalutationInlineContent_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_RecipientSalutationInlineContent() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "cellNumber", "9a"
    d.Add "letterCase", "normal"
    d.Add "style", LoadRawFormatSpec_Components_RecipientSalutationInlineContent_Style()
    d.Add "alignment", "center"
    d.Add "lastItemEndChar", "."
    d.Add "firstLineIndentCm", 0
    d.Add "spaceBeforePt", 6
    d.Add "lineSpacingZone", "body"
    Set LoadRawFormatSpec_Components_RecipientSalutationInlineContent = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_RecipientSalutationInlineContent", Err.description
End Function

Private Function LoadRawFormatSpec_Components_LegalBasis_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", True
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_LegalBasis_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_LegalBasis_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_LegalBasis() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "cellNumber", "6"
    d.Add "letterCase", "normal"
    d.Add "style", LoadRawFormatSpec_Components_LegalBasis_Style()
    d.Add "alignment", "justify"
    d.Add "bulletChar", ""
    d.Add "lineEndChar", ";"
    d.Add "lastLineEndChar", "."
    d.Add "firstLineIndentCm", 1
    d.Add "spaceBeforePt", 6
    d.Add "lineSpacingZone", "body"
    Set LoadRawFormatSpec_Components_LegalBasis = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_LegalBasis", Err.description
End Function

Private Function LoadRawFormatSpec_Components_BodyText_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_BodyText_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_BodyText_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_BodyText() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "cellNumber", "6"
    d.Add "letterCase", "normal"
    d.Add "style", LoadRawFormatSpec_Components_BodyText_Style()
    d.Add "alignment", "justify"
    d.Add "firstLineIndentCm", 1
    d.Add "spaceBeforePt", 6
    d.Add "lineSpacingZone", "body"
    d.Add "contentEndMark", "./."
    Set LoadRawFormatSpec_Components_BodyText = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_BodyText", Err.description
End Function

Private Function LoadRawFormatSpec_Components_SignerAuthority_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", True
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_SignerAuthority_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_SignerAuthority_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_SignerAuthority() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "cellNumber", "7a"
    d.Add "letterCase", "upper"
    d.Add "style", LoadRawFormatSpec_Components_SignerAuthority_Style()
    d.Add "alignment", "center"
    d.Add "zone", "bottomRight"
    d.Add "firstLineIndentCm", 0
    d.Add "spaceBeforePt", 0
    d.Add "lineSpacingZone", "fixed"
    Set LoadRawFormatSpec_Components_SignerAuthority = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_SignerAuthority", Err.description
End Function

Private Function LoadRawFormatSpec_Components_SignerAuthorityTitle_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", True
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_SignerAuthorityTitle_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_SignerAuthorityTitle_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_SignerAuthorityTitle() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "cellNumber", "7a"
    d.Add "letterCase", "upper"
    d.Add "style", LoadRawFormatSpec_Components_SignerAuthorityTitle_Style()
    d.Add "alignment", "center"
    d.Add "zone", "bottomRight"
    d.Add "firstLineIndentCm", 0
    d.Add "spaceBeforePt", 0
    d.Add "lineSpacingZone", "fixed"
    Set LoadRawFormatSpec_Components_SignerAuthorityTitle = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_SignerAuthorityTitle", Err.description
End Function

Private Function LoadRawFormatSpec_Components_RecipientLabel_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", True
    d.Add "italic", True
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_RecipientLabel_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_RecipientLabel_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_RecipientLabel() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "cellNumber", "9b"
    d.Add "label", "N" & ChrW(&H1A1) & "i nh" & ChrW(&H1EAD) & "n:"
    d.Add "letterCase", "normal"
    d.Add "style", LoadRawFormatSpec_Components_RecipientLabel_Style()
    d.Add "alignment", "justify"
    d.Add "zone", "bottomLeft"
    d.Add "firstLineIndentCm", 0
    d.Add "spaceBeforePt", 0
    d.Add "lineSpacingZone", "fixed"
    Set LoadRawFormatSpec_Components_RecipientLabel = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_RecipientLabel", Err.description
End Function

Private Function LoadRawFormatSpec_Components_RecipientList_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_RecipientList_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_RecipientList_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_RecipientList() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "cellNumber", "9b"
    d.Add "letterCase", "normal"
    d.Add "style", LoadRawFormatSpec_Components_RecipientList_Style()
    d.Add "alignment", "justify"
    d.Add "bulletChar", "-"
    d.Add "itemEndChar", ";"
    d.Add "archiveNotation", "VT"
    d.Add "archiveLinePattern", "^-?\s*L" & ChrW(&H1B0) & "u\s*:\s*VT\s*,\s*\S+\.\S+\.\(\d+\)\.?$"
    d.Add "firstLineIndentCm", 0
    d.Add "spaceBeforePt", 0
    d.Add "lineSpacingZone", "fixed"
    Set LoadRawFormatSpec_Components_RecipientList = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_RecipientList", Err.description
End Function

Private Function LoadRawFormatSpec_Components_AppendixLabel_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", True
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_AppendixLabel_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_AppendixLabel_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_AppendixLabel() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "letterCase", "normal"
    d.Add "style", LoadRawFormatSpec_Components_AppendixLabel_Style()
    d.Add "alignment", "center"
    d.Add "numeralType", "roman"
    d.Add "numberOnlyWhenMultiple", True
    d.Add "firstLineIndentCm", 0
    d.Add "spaceBeforePt", 6
    d.Add "lineSpacingZone", "body"
    Set LoadRawFormatSpec_Components_AppendixLabel = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_AppendixLabel", Err.description
End Function

Private Function LoadRawFormatSpec_Components_AppendixTitle_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", True
    d.Add "italic", False
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_AppendixTitle_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_AppendixTitle_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_AppendixTitle() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "letterCase", "upper"
    d.Add "style", LoadRawFormatSpec_Components_AppendixTitle_Style()
    d.Add "alignment", "center"
    d.Add "firstLineIndentCm", 0
    d.Add "spaceBeforePt", 6
    d.Add "lineSpacingZone", "body"
    Set LoadRawFormatSpec_Components_AppendixTitle = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_AppendixTitle", Err.description
End Function

Private Function LoadRawFormatSpec_Components_AppendixReference_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", True
    d.Add "underline", False
    Set LoadRawFormatSpec_Components_AppendixReference_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_AppendixReference_Style", Err.description
End Function

Private Function LoadRawFormatSpec_Components_AppendixReference() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "letterCase", "normal"
    d.Add "style", LoadRawFormatSpec_Components_AppendixReference_Style()
    d.Add "alignment", "center"
    d.Add "pattern", "(K" & ChrW(&HE8) & "m theo V" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n s" & ChrW(&H1ED1) & " {codeNumber}/{codeNotation} ng" & ChrW(&HE0) & "y {dd} th" & ChrW(&HE1) & "ng {mm} n" & ChrW(&H103) & "m {yyyy} c" & ChrW(&H1EE7) & "a {organName})"
    d.Add "skipForElectronicDocument", True
    d.Add "firstLineIndentCm", 0
    d.Add "spaceBeforePt", 6
    d.Add "lineSpacingZone", "body"
    Set LoadRawFormatSpec_Components_AppendixReference = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components_AppendixReference", Err.description
End Function

Private Function LoadRawFormatSpec_Components() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "nationalTitle", LoadRawFormatSpec_Components_NationalTitle()
    d.Add "nationalMotto", LoadRawFormatSpec_Components_NationalMotto()
    d.Add "partyHeader", LoadRawFormatSpec_Components_PartyHeader()
    d.Add "superiorOrganName", LoadRawFormatSpec_Components_SuperiorOrganName()
    d.Add "organName", LoadRawFormatSpec_Components_OrganName()
    d.Add "starSeparator", LoadRawFormatSpec_Components_StarSeparator()
    d.Add "codeNumberNotation", LoadRawFormatSpec_Components_CodeNumberNotation()
    d.Add "placeAndIssuedDate", LoadRawFormatSpec_Components_PlaceAndIssuedDate()
    d.Add "subjectOfficialLetter", LoadRawFormatSpec_Components_SubjectOfficialLetter()
    d.Add "typeName", LoadRawFormatSpec_Components_TypeName()
    d.Add "subject", LoadRawFormatSpec_Components_Subject()
    d.Add "recipientSalutation", LoadRawFormatSpec_Components_RecipientSalutation()
    d.Add "recipientSalutationList", LoadRawFormatSpec_Components_RecipientSalutationList()
    d.Add "recipientSalutationInline", LoadRawFormatSpec_Components_RecipientSalutationInline()
    d.Add "recipientSalutationInlineContent", LoadRawFormatSpec_Components_RecipientSalutationInlineContent()
    d.Add "legalBasis", LoadRawFormatSpec_Components_LegalBasis()
    d.Add "bodyText", LoadRawFormatSpec_Components_BodyText()
    d.Add "signerAuthority", LoadRawFormatSpec_Components_SignerAuthority()
    d.Add "signerAuthorityTitle", LoadRawFormatSpec_Components_SignerAuthorityTitle()
    d.Add "recipientLabel", LoadRawFormatSpec_Components_RecipientLabel()
    d.Add "recipientList", LoadRawFormatSpec_Components_RecipientList()
    d.Add "appendixLabel", LoadRawFormatSpec_Components_AppendixLabel()
    d.Add "appendixTitle", LoadRawFormatSpec_Components_AppendixTitle()
    d.Add "appendixReference", LoadRawFormatSpec_Components_AppendixReference()
    Set LoadRawFormatSpec_Components = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_Components", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Part_KeywordLine_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", True
    d.Add "italic", False
    Set LoadRawFormatSpec_StructureLevels_Part_KeywordLine_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Part_KeywordLine_Style", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Part_KeywordLine() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "alignment", "center"
    d.Add "letterCase", "normal"
    d.Add "style", LoadRawFormatSpec_StructureLevels_Part_KeywordLine_Style()
    Set LoadRawFormatSpec_StructureLevels_Part_KeywordLine = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Part_KeywordLine", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Part_TitleLine_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", True
    d.Add "italic", False
    Set LoadRawFormatSpec_StructureLevels_Part_TitleLine_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Part_TitleLine_Style", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Part_TitleLine() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "alignment", "center"
    d.Add "letterCase", "upper"
    d.Add "style", LoadRawFormatSpec_StructureLevels_Part_TitleLine_Style()
    Set LoadRawFormatSpec_StructureLevels_Part_TitleLine = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Part_TitleLine", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Part() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "keyword", "Ph" & ChrW(&H1EA7) & "n"
    d.Add "numeralType", "roman"
    d.Add "keywordLine", LoadRawFormatSpec_StructureLevels_Part_KeywordLine()
    d.Add "titleLine", LoadRawFormatSpec_StructureLevels_Part_TitleLine()
    Set LoadRawFormatSpec_StructureLevels_Part = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Part", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Chapter_KeywordLine_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", True
    d.Add "italic", False
    Set LoadRawFormatSpec_StructureLevels_Chapter_KeywordLine_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Chapter_KeywordLine_Style", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Chapter_KeywordLine() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "alignment", "center"
    d.Add "letterCase", "normal"
    d.Add "style", LoadRawFormatSpec_StructureLevels_Chapter_KeywordLine_Style()
    Set LoadRawFormatSpec_StructureLevels_Chapter_KeywordLine = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Chapter_KeywordLine", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Chapter_TitleLine_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", True
    d.Add "italic", False
    Set LoadRawFormatSpec_StructureLevels_Chapter_TitleLine_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Chapter_TitleLine_Style", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Chapter_TitleLine() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "alignment", "center"
    d.Add "letterCase", "upper"
    d.Add "style", LoadRawFormatSpec_StructureLevels_Chapter_TitleLine_Style()
    Set LoadRawFormatSpec_StructureLevels_Chapter_TitleLine = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Chapter_TitleLine", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Chapter() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "keyword", "Ch" & ChrW(&H1B0) & ChrW(&H1A1) & "ng"
    d.Add "numeralType", "roman"
    d.Add "keywordLine", LoadRawFormatSpec_StructureLevels_Chapter_KeywordLine()
    d.Add "titleLine", LoadRawFormatSpec_StructureLevels_Chapter_TitleLine()
    Set LoadRawFormatSpec_StructureLevels_Chapter = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Chapter", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Section_KeywordLine_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", True
    d.Add "italic", False
    Set LoadRawFormatSpec_StructureLevels_Section_KeywordLine_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Section_KeywordLine_Style", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Section_KeywordLine() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "alignment", "center"
    d.Add "letterCase", "normal"
    d.Add "style", LoadRawFormatSpec_StructureLevels_Section_KeywordLine_Style()
    Set LoadRawFormatSpec_StructureLevels_Section_KeywordLine = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Section_KeywordLine", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Section_TitleLine_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", True
    d.Add "italic", False
    Set LoadRawFormatSpec_StructureLevels_Section_TitleLine_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Section_TitleLine_Style", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Section_TitleLine() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "alignment", "center"
    d.Add "letterCase", "upper"
    d.Add "style", LoadRawFormatSpec_StructureLevels_Section_TitleLine_Style()
    Set LoadRawFormatSpec_StructureLevels_Section_TitleLine = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Section_TitleLine", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Section() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "keyword", "M" & ChrW(&H1EE5) & "c"
    d.Add "numeralType", "arabic"
    d.Add "keywordLine", LoadRawFormatSpec_StructureLevels_Section_KeywordLine()
    d.Add "titleLine", LoadRawFormatSpec_StructureLevels_Section_TitleLine()
    Set LoadRawFormatSpec_StructureLevels_Section = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Section", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_SubSection_KeywordLine_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", True
    d.Add "italic", False
    Set LoadRawFormatSpec_StructureLevels_SubSection_KeywordLine_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_SubSection_KeywordLine_Style", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_SubSection_KeywordLine() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "alignment", "center"
    d.Add "letterCase", "normal"
    d.Add "style", LoadRawFormatSpec_StructureLevels_SubSection_KeywordLine_Style()
    Set LoadRawFormatSpec_StructureLevels_SubSection_KeywordLine = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_SubSection_KeywordLine", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_SubSection_TitleLine_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", True
    d.Add "italic", False
    Set LoadRawFormatSpec_StructureLevels_SubSection_TitleLine_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_SubSection_TitleLine_Style", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_SubSection_TitleLine() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "alignment", "center"
    d.Add "letterCase", "upper"
    d.Add "style", LoadRawFormatSpec_StructureLevels_SubSection_TitleLine_Style()
    Set LoadRawFormatSpec_StructureLevels_SubSection_TitleLine = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_SubSection_TitleLine", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_SubSection() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "keyword", "Ti" & ChrW(&H1EC3) & "u m" & ChrW(&H1EE5) & "c"
    d.Add "numeralType", "arabic"
    d.Add "keywordLine", LoadRawFormatSpec_StructureLevels_SubSection_KeywordLine()
    d.Add "titleLine", LoadRawFormatSpec_StructureLevels_SubSection_TitleLine()
    Set LoadRawFormatSpec_StructureLevels_SubSection = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_SubSection", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Article_IndentCm_Allowed() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add 1
    c.Add 1.27
    Set LoadRawFormatSpec_StructureLevels_Article_IndentCm_Allowed = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Article_IndentCm_Allowed", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Article_IndentCm() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "allowed", LoadRawFormatSpec_StructureLevels_Article_IndentCm_Allowed()
    d.Add "default", 1
    Set LoadRawFormatSpec_StructureLevels_Article_IndentCm = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Article_IndentCm", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Article_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", True
    d.Add "italic", False
    Set LoadRawFormatSpec_StructureLevels_Article_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Article_Style", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Article() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "keyword", ChrW(&H110) & "i" & ChrW(&H1EC1) & "u"
    d.Add "numeralType", "arabic"
    d.Add "suffixAfterNumber", "."
    d.Add "indentCm", LoadRawFormatSpec_StructureLevels_Article_IndentCm()
    d.Add "style", LoadRawFormatSpec_StructureLevels_Article_Style()
    d.Add "fontSizeFollowsBody", True
    Set LoadRawFormatSpec_StructureLevels_Article = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Article", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Clause_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", False
    Set LoadRawFormatSpec_StructureLevels_Clause_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Clause_Style", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Clause_StyleWhenHasTitle() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", True
    d.Add "italic", False
    Set LoadRawFormatSpec_StructureLevels_Clause_StyleWhenHasTitle = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Clause_StyleWhenHasTitle", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Clause() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "numeralType", "arabic"
    d.Add "suffixAfterNumber", "."
    d.Add "style", LoadRawFormatSpec_StructureLevels_Clause_Style()
    d.Add "styleWhenHasTitle", LoadRawFormatSpec_StructureLevels_Clause_StyleWhenHasTitle()
    d.Add "fontSizeFollowsBody", True
    Set LoadRawFormatSpec_StructureLevels_Clause = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Clause", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Point_Style() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "bold", False
    d.Add "italic", False
    Set LoadRawFormatSpec_StructureLevels_Point_Style = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Point_Style", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels_Point() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "numeralType", "vietnameseAlphabet"
    d.Add "suffixAfterNumber", ")"
    d.Add "style", LoadRawFormatSpec_StructureLevels_Point_Style()
    d.Add "fontSizeFollowsBody", True
    Set LoadRawFormatSpec_StructureLevels_Point = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels_Point", Err.description
End Function

Private Function LoadRawFormatSpec_StructureLevels() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "part", LoadRawFormatSpec_StructureLevels_Part()
    d.Add "chapter", LoadRawFormatSpec_StructureLevels_Chapter()
    d.Add "section", LoadRawFormatSpec_StructureLevels_Section()
    d.Add "subSection", LoadRawFormatSpec_StructureLevels_SubSection()
    d.Add "article", LoadRawFormatSpec_StructureLevels_Article()
    d.Add "clause", LoadRawFormatSpec_StructureLevels_Clause()
    d.Add "point", LoadRawFormatSpec_StructureLevels_Point()
    Set LoadRawFormatSpec_StructureLevels = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_StructureLevels", Err.description
End Function

Private Sub LoadRawFormatSpec_VietnameseAlphabetPart1(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "a"
    c.Add ChrW(&H103)
    c.Add ChrW(&HE2)
    c.Add "b"
    c.Add "c"
    c.Add "d"
    c.Add ChrW(&H111)
    c.Add "e"
    c.Add ChrW(&HEA)
    c.Add "g"
    c.Add "h"
    c.Add "i"
    c.Add "k"
    c.Add "l"
    c.Add "m"
    c.Add "n"
    c.Add "o"
    c.Add ChrW(&HF4)
    c.Add ChrW(&H1A1)
    c.Add "p"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_VietnameseAlphabetPart1", Err.description
End Sub

Private Sub LoadRawFormatSpec_VietnameseAlphabetPart2(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "q"
    c.Add "r"
    c.Add "s"
    c.Add "t"
    c.Add "u"
    c.Add ChrW(&H1B0)
    c.Add "v"
    c.Add "x"
    c.Add "y"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_VietnameseAlphabetPart2", Err.description
End Sub

Private Function LoadRawFormatSpec_VietnameseAlphabet() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    LoadRawFormatSpec_VietnameseAlphabetPart1 c
    LoadRawFormatSpec_VietnameseAlphabetPart2 c
    Set LoadRawFormatSpec_VietnameseAlphabet = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_VietnameseAlphabet", Err.description
End Function

Private Sub LoadRawFormatSpec_PointOrderAlphabetPart1(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "a"
    c.Add "b"
    c.Add "c"
    c.Add "d"
    c.Add ChrW(&H111)
    c.Add "e"
    c.Add "g"
    c.Add "h"
    c.Add "i"
    c.Add "k"
    c.Add "l"
    c.Add "m"
    c.Add "n"
    c.Add "o"
    c.Add ChrW(&HF4)
    c.Add ChrW(&H1A1)
    c.Add "p"
    c.Add "q"
    c.Add "r"
    c.Add "s"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_PointOrderAlphabetPart1", Err.description
End Sub

Private Sub LoadRawFormatSpec_PointOrderAlphabetPart2(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "t"
    c.Add "u"
    c.Add ChrW(&H1B0)
    c.Add "v"
    c.Add "x"
    c.Add "y"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_PointOrderAlphabetPart2", Err.description
End Sub

Private Function LoadRawFormatSpec_PointOrderAlphabet() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    LoadRawFormatSpec_PointOrderAlphabetPart1 c
    LoadRawFormatSpec_PointOrderAlphabetPart2 c
    Set LoadRawFormatSpec_PointOrderAlphabet = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_PointOrderAlphabet", Err.description
End Function

Private Function LoadRawFormatSpec_SignerAuthorityPrefixes() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "TM.", "K" & ChrW(&HFD) & " thay m" & ChrW(&H1EB7) & "t t" & ChrW(&H1EAD) & "p th" & ChrW(&H1EC3)
    d.Add "Q.", ChrW(&H110) & ChrW(&H1B0) & ChrW(&H1EE3) & "c giao quy" & ChrW(&H1EC1) & "n c" & ChrW(&H1EA5) & "p tr" & ChrW(&H1B0) & ChrW(&H1EDF) & "ng"
    d.Add "KT.", "K" & ChrW(&HFD) & " thay ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i " & ChrW(&H111) & ChrW(&H1EE9) & "ng " & ChrW(&H111) & ChrW(&H1EA7) & "u"
    d.Add "TL.", "K" & ChrW(&HFD) & " th" & ChrW(&H1EEB) & "a l" & ChrW(&H1EC7) & "nh"
    d.Add "TUQ.", "K" & ChrW(&HFD) & " th" & ChrW(&H1EEB) & "a " & ChrW(&H1EE7) & "y quy" & ChrW(&H1EC1) & "n"
    Set LoadRawFormatSpec_SignerAuthorityPrefixes = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_SignerAuthorityPrefixes", Err.description
End Function

Private Function LoadRawFormatSpec_ConfidentialityLevels() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "TUY" & ChrW(&H1EC6) & "T M" & ChrW(&H1EAC) & "T"
    c.Add "T" & ChrW(&H1ED0) & "I M" & ChrW(&H1EAC) & "T"
    c.Add "M" & ChrW(&H1EAC) & "T"
    Set LoadRawFormatSpec_ConfidentialityLevels = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_ConfidentialityLevels", Err.description
End Function

Private Function LoadRawFormatSpec_CopyForms() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "SAO Y"
    c.Add "SAO L" & ChrW(&H1EE4) & "C"
    c.Add "TR" & ChrW(&HCD) & "CH SAO"
    Set LoadRawFormatSpec_CopyForms = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec_CopyForms", Err.description
End Function

Public Function LoadRawFormatSpec() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "2.0.0"
    d.Add "legalBasis", LoadRawFormatSpec_LegalBasis()
    d.Add "units", LoadRawFormatSpec_Units()
    d.Add "pageSetup", LoadRawFormatSpec_PageSetup()
    d.Add "font", LoadRawFormatSpec_Font()
    d.Add "pageNumber", LoadRawFormatSpec_PageNumber()
    d.Add "bodyText", LoadRawFormatSpec_BodyText()
    d.Add "charSpacing", LoadRawFormatSpec_CharSpacing()
    d.Add "fontSizeSets", LoadRawFormatSpec_FontSizeSets()
    d.Add "fontSizeRanges", LoadRawFormatSpec_FontSizeRanges()
    d.Add "components", LoadRawFormatSpec_Components()
    d.Add "structureLevels", LoadRawFormatSpec_StructureLevels()
    d.Add "vietnameseAlphabet", LoadRawFormatSpec_VietnameseAlphabet()
    d.Add "pointOrderAlphabet", LoadRawFormatSpec_PointOrderAlphabet()
    d.Add "signerAuthorityPrefixes", LoadRawFormatSpec_SignerAuthorityPrefixes()
    d.Add "confidentialityLevels", LoadRawFormatSpec_ConfidentialityLevels()
    d.Add "copyForms", LoadRawFormatSpec_CopyForms()
    Set LoadRawFormatSpec = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawFormatSpec", Err.description
End Function

Private Function LoadRawTypoDictionary_Corrections() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "s" & ChrW(&H1EED) & " l" & ChrW(&HFD), "x" & ChrW(&H1EED) & " l" & ChrW(&HFD)
    d.Add "trau gi" & ChrW(&H1ED3) & "i", "trau d" & ChrW(&H1ED3) & "i"
    d.Add "chau d" & ChrW(&H1ED3) & "i", "trau d" & ChrW(&H1ED3) & "i"
    d.Add "tr" & ChrW(&H1EA3) & "i tru" & ChrW(&H1ED1) & "t", "trau chu" & ChrW(&H1ED1) & "t"
    d.Add "tr" & ChrW(&H1EC9) & " tr" & ChrW(&HED) & "ch", "ch" & ChrW(&H1EC9) & " tr" & ChrW(&HED) & "ch"
    d.Add "tr" & ChrW(&H1EC9) & " huy", "ch" & ChrW(&H1EC9) & " huy"
    d.Add "xu" & ChrW(&H1EA5) & "t x" & ChrW(&H1EAF) & "c", "xu" & ChrW(&H1EA5) & "t s" & ChrW(&H1EAF) & "c"
    d.Add "s" & ChrW(&HE1) & "ng l" & ChrW(&H1EA1) & "ng", "x" & ChrW(&HE1) & "n l" & ChrW(&H1EA1) & "n"
    d.Add "th" & ChrW(&H103) & "m quan", "tham quan"
    d.Add "tr" & ChrW(&HED) & " m" & ChrW(&H1EA1) & "ng", "ch" & ChrW(&HED) & " m" & ChrW(&H1EA1) & "ng"
    d.Add "kh" & ChrW(&HFA) & "c tri" & ChrW(&H1EBF) & "t", "kh" & ChrW(&HFA) & "c chi" & ChrW(&H1EBF) & "t"
    d.Add "c" & ChrW(&HE2) & "u truy" & ChrW(&H1EC7) & "n", "c" & ChrW(&HE2) & "u chuy" & ChrW(&H1EC7) & "n"
    d.Add "tr" & ChrW(&HE2) & "n th" & ChrW(&HE0) & "nh", "ch" & ChrW(&HE2) & "n th" & ChrW(&HE0) & "nh"
    d.Add "ch" & ChrW(&HE2) & "n qu" & ChrW(&HFD), "tr" & ChrW(&HE2) & "n qu" & ChrW(&HFD)
    d.Add "t" & ChrW(&H1EF1) & "u chung", "t" & ChrW(&H1EF1) & "u trung"
    d.Add "giao " & ChrW(&H111) & ChrW(&H1ED9) & "ng", "dao " & ChrW(&H111) & ChrW(&H1ED9) & "ng"
    d.Add "r" & ChrW(&HF4) & "ng r" & ChrW(&HE3) & "i", "r" & ChrW(&H1ED9) & "ng r" & ChrW(&HE3) & "i"
    d.Add "chia s" & ChrW(&H1EBD), "chia s" & ChrW(&H1EBB)
    d.Add "san s" & ChrW(&H1EBD), "san s" & ChrW(&H1EBB)
    d.Add ChrW(&H111) & ChrW(&H1B0) & ChrW(&H1EDD) & "ng x" & ChrW(&HE1), ChrW(&H111) & ChrW(&H1B0) & ChrW(&H1EDD) & "ng s" & ChrW(&HE1)
    d.Add "l" & ChrW(&HE3) & "ng m" & ChrW(&H1EA1) & "ng", "l" & ChrW(&HE3) & "ng m" & ChrW(&H1EA1) & "n"
    d.Add "c" & ChrW(&H1ECD) & " s" & ChrW(&HE1) & "t", "c" & ChrW(&H1ECD) & " x" & ChrW(&HE1) & "t"
    d.Add "d" & ChrW(&H1EBD) & " tr" & ChrW(&HE1) & "i", "r" & ChrW(&H1EBD) & " tr" & ChrW(&HE1) & "i"
    d.Add "v" & ChrW(&HF4) & " h" & ChrW(&HEC) & "nh chung", "v" & ChrW(&HF4) & " h" & ChrW(&HEC) & "nh trung"
    d.Add "v" & ChrW(&HF4) & " h" & ChrW(&HEC) & "nh dung", "v" & ChrW(&HF4) & " h" & ChrW(&HEC) & "nh trung"
    d.Add "kh" & ChrW(&H1EB3) & "ng kh" & ChrW(&HE1) & "i", "kh" & ChrW(&H1EA3) & "ng kh" & ChrW(&HE1) & "i"
    d.Add "s" & ChrW(&H1EBB) & "o nh" & ChrW(&H1ECF), "x" & ChrW(&H1EBB) & "o nh" & ChrW(&H1ECF)
    d.Add "s" & ChrW(&HE1) & "t nh" & ChrW(&H1EAD) & "p", "s" & ChrW(&HE1) & "p nh" & ChrW(&H1EAD) & "p"
    Set LoadRawTypoDictionary_Corrections = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTypoDictionary_Corrections", Err.description
End Function

Private Sub LoadRawTypoDictionary_ProtectedForms_FormsPart1(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "s" & ChrW(&H1EED) & " d" & ChrW(&H1EE5) & "ng"
    c.Add "s" & ChrW(&HFA) & "c t" & ChrW(&HED) & "ch"
    c.Add "tham quan"
    c.Add "ki" & ChrW(&H1EC1) & "m ch" & ChrW(&H1EBF)
    c.Add "tr" & ChrW(&HE2) & "n tr" & ChrW(&H1ECD) & "ng"
    c.Add "tr" & ChrW(&HEC) & "u m" & ChrW(&H1EBF) & "n"
    c.Add "dao " & ChrW(&H111) & ChrW(&H1ED9) & "ng"
    c.Add "s" & ChrW(&H1EBB) & " chia"
    c.Add "t" & ChrW(&H1EF1) & "u tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng"
    c.Add "s" & ChrW(&H1EC9) & " nh" & ChrW(&H1EE5) & "c"
    c.Add ChrW(&H111) & ChrW(&H1EA3) & "m b" & ChrW(&H1EA3) & "o"
    c.Add "b" & ChrW(&H1EA3) & "o " & ChrW(&H111) & ChrW(&H1EA3) & "m"
    c.Add "kh" & ChrW(&HFA) & "c kh" & ChrW(&HED) & "ch"
    c.Add "s" & ChrW(&HE1) & "ng ki" & ChrW(&H1EBF) & "n"
    c.Add "tu" & ChrW(&H1EC7) & "ch to" & ChrW(&H1EA1) & "c"
    c.Add "gi" & ChrW(&H1EA3) & " thuy" & ChrW(&H1EBF) & "t"
    c.Add "gi" & ChrW(&H1EA3) & " thi" & ChrW(&H1EBF) & "t"
    c.Add "tranh ch" & ChrW(&H1EA5) & "p"
    c.Add "th" & ChrW(&H1EEB) & "a h" & ChrW(&HE0) & "nh"
    c.Add "th" & ChrW(&H1EEB) & "a k" & ChrW(&H1EBF)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTypoDictionary_ProtectedForms_FormsPart1", Err.description
End Sub

Private Sub LoadRawTypoDictionary_ProtectedForms_FormsPart2(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "k" & ChrW(&H1EBF) & " th" & ChrW(&H1EEB) & "a"
    c.Add "nh" & ChrW(&HE2) & "n vi" & ChrW(&HEA) & "n"
    c.Add "nh" & ChrW(&HE2) & "n v" & ChrW(&H1EAD) & "t"
    c.Add "phong ph" & ChrW(&HFA)
    c.Add ChrW(&H111) & ChrW(&H1ED9) & "t xu" & ChrW(&H1EA5) & "t"
    c.Add "ng" & ChrW(&H1EA3) & " nghi" & ChrW(&HEA) & "ng"
    c.Add "ng" & ChrW(&H1EA3) & " ng" & ChrW(&H1EDB) & "n"
    c.Add "r" & ChrW(&H1EBD) & " tr" & ChrW(&HE1) & "i"
    c.Add "trau chu" & ChrW(&H1ED1) & "t"
    c.Add "x" & ChrW(&HE1) & "n l" & ChrW(&H1EA1) & "n"
    c.Add "v" & ChrW(&HF4) & " h" & ChrW(&HEC) & "nh trung"
    c.Add "trau d" & ChrW(&H1ED3) & "i"
    c.Add "ch" & ChrW(&H1EC9) & " tr" & ChrW(&HED) & "ch"
    c.Add "ch" & ChrW(&H1EC9) & " huy"
    c.Add "xu" & ChrW(&H1EA5) & "t s" & ChrW(&H1EAF) & "c"
    c.Add "ch" & ChrW(&HED) & " m" & ChrW(&H1EA1) & "ng"
    c.Add "kh" & ChrW(&HFA) & "c chi" & ChrW(&H1EBF) & "t"
    c.Add "c" & ChrW(&HE2) & "u chuy" & ChrW(&H1EC7) & "n"
    c.Add "ch" & ChrW(&HE2) & "n th" & ChrW(&HE0) & "nh"
    c.Add "tr" & ChrW(&HE2) & "n qu" & ChrW(&HFD)
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTypoDictionary_ProtectedForms_FormsPart2", Err.description
End Sub

Private Sub LoadRawTypoDictionary_ProtectedForms_FormsPart3(ByRef c As Collection)
    On Error GoTo ErrHandler
    c.Add "t" & ChrW(&H1EF1) & "u trung"
    c.Add "r" & ChrW(&H1ED9) & "ng r" & ChrW(&HE3) & "i"
    c.Add "chia s" & ChrW(&H1EBB)
    c.Add "san s" & ChrW(&H1EBB)
    c.Add ChrW(&H111) & ChrW(&H1B0) & ChrW(&H1EDD) & "ng s" & ChrW(&HE1)
    c.Add "l" & ChrW(&HE3) & "ng m" & ChrW(&H1EA1) & "n"
    c.Add "c" & ChrW(&H1ECD) & " x" & ChrW(&HE1) & "t"
    c.Add "x" & ChrW(&H1EED) & " l" & ChrW(&HFD)
    c.Add "s" & ChrW(&HE1) & "p nh" & ChrW(&H1EAD) & "p"
    c.Add "tr" & ChrW(&H1EC9) & "a"
    c.Add "r" & ChrW(&H1B0) & ChrW(&H1EE3) & "i"
    Exit Sub
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTypoDictionary_ProtectedForms_FormsPart3", Err.description
End Sub

Private Function LoadRawTypoDictionary_ProtectedForms_Forms() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    LoadRawTypoDictionary_ProtectedForms_FormsPart1 c
    LoadRawTypoDictionary_ProtectedForms_FormsPart2 c
    LoadRawTypoDictionary_ProtectedForms_FormsPart3 c
    Set LoadRawTypoDictionary_ProtectedForms_Forms = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTypoDictionary_ProtectedForms_Forms", Err.description
End Function

Private Function LoadRawTypoDictionary_ProtectedForms() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "forms", LoadRawTypoDictionary_ProtectedForms_Forms()
    Set LoadRawTypoDictionary_ProtectedForms = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTypoDictionary_ProtectedForms", Err.description
End Function

Private Function LoadRawTypoDictionary_PendingReview_Items() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim d231 As Object
    Set d231 = CreateObject("Scripting.Dictionary")
    d231.Add "from", "chia x" & ChrW(&H1EBB)
    d231.Add "to", "chia s" & ChrW(&H1EBB)
    Dim tmp16 As String
    tmp16 = "Hai t" & ChrW(&H1EEB) & " kh" & ChrW(&HE1) & "c ngh" & ChrW(&H129) & "a th" & ChrW(&H1EAD) & "t, kh" & ChrW(&HF4) & "ng ph" & ChrW(&H1EA3) & "i l" & ChrW(&H1ED7) & "i ch" & ChrW(&HED) & "nh t" & ChrW(&H1EA3) & ": 'chia s" & ChrW(&H1EBB) & "' = san s" & ChrW(&H1EDB) & "t t" & ChrW(&HEC) & "nh c" & ChrW(&H1EA3) & "m; 'chia x" & ChrW(&H1EBB) & "' = c" & ChrW(&H1EAF) & "t chia ra (chia x" & ChrW(&H1EBB) & " " & ChrW(&H111) & ChrW(&H1EA5) & "t " & ChrW(&H111) & "ai). " & ChrW(&H110) & ChrW(&HFA) & "ng d" & ChrW(&H1EA1) & "ng x-s theo ch" & ChrW(&HED) & "nh s" & ChrW(&HE1) & "ch ch" & ChrW(&H1EE7) & " d" & ChrW(&H1EF1) & " " & ChrW(&HE1) & "n duy" & ChrW(&H1EC7)
    tmp16 = tmp16 & "t 2026-08-11, nh" & ChrW(&H1B0) & "ng KH" & ChrW(&HD4) & "NG k" & ChrW(&HED) & "ch ho" & ChrW(&H1EA1) & "t v" & ChrW(&HE0) & "o corrections v" & ChrW(&HEC) & " s" & ChrW(&H1EBD) & " s" & ChrW(&H1EED) & "a sai trong nhi" & ChrW(&H1EC1) & "u tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng h" & ChrW(&H1EE3) & "p."
    d231.Add "reason", tmp16
    Dim tmp17 As String
    tmp17 = "C" & ChrW(&H1EA6) & "N C" & ChrW(&H1A0) & " CH" & ChrW(&H1EBE) & " M" & ChrW(&H1EDA) & "I: c" & ChrW(&H1EA3) & "nh b" & ChrW(&HE1) & "o 'v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n d" & ChrW(&HF9) & "ng chia x" & ChrW(&H1EBB) & ", ki" & ChrW(&H1EC3) & "m tra l" & ChrW(&H1EA1) & "i ngh" & ChrW(&H129) & "a c" & ChrW(&HF3) & " " & ChrW(&H111) & ChrW(&HFA) & "ng " & ChrW(&HFD) & " kh" & ChrW(&HF4) & "ng' " & ChrW(&H2014) & " kh" & ChrW(&HF4) & "ng k" & ChrW(&HE8) & "m " & ChrW(&H111) & ChrW(&H1EC1) & " ngh" & ChrW(&H1ECB) & " s" & ChrW(&H1EED) & "a c" & ChrW(&H1EE5) & " th" & ChrW(&H1EC3) & ". Schema hi" & ChrW(&H1EC7) & "n t" & ChrW(&H1EA1)
    tmp17 = tmp17 & "i (corrections/protectedForms/pendingReview) ch" & ChrW(&H1B0) & "a c" & ChrW(&HF3) & " ki" & ChrW(&H1EC3) & "u 'c" & ChrW(&H1EA3) & "nh b" & ChrW(&HE1) & "o kh" & ChrW(&HF4) & "ng " & ChrW(&H111) & ChrW(&H1EC1) & " ngh" & ChrW(&H1ECB) & " s" & ChrW(&H1EED) & "a' " & ChrW(&H2014) & " c" & ChrW(&H1EA7) & "n th" & ChrW(&HEA) & "m tr" & ChrW(&H1B0) & ChrW(&H1EDB) & "c khi hi" & ChrW(&H1EC7) & "n th" & ChrW(&H1EF1) & "c h" & ChrW(&HF3) & "a TypoDictionary."
    d231.Add "recommendation", tmp17
    c.Add d231
    Set LoadRawTypoDictionary_PendingReview_Items = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTypoDictionary_PendingReview_Items", Err.description
End Function

Private Function LoadRawTypoDictionary_PendingReview() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "items", LoadRawTypoDictionary_PendingReview_Items()
    Set LoadRawTypoDictionary_PendingReview = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTypoDictionary_PendingReview", Err.description
End Function

Private Function LoadRawTypoDictionary_SkipContexts_ComponentRoles() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "nationalTitle"
    c.Add "nationalMotto"
    c.Add "superiorOrganName"
    c.Add "organName"
    c.Add "signerName"
    c.Add "placeName"
    Set LoadRawTypoDictionary_SkipContexts_ComponentRoles = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTypoDictionary_SkipContexts_ComponentRoles", Err.description
End Function

Private Function LoadRawTypoDictionary_SkipContexts() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "insideQuotes", True
    d.Add "insideUrlOrEmail", True
    d.Add "insideDocumentCode", True
    d.Add "componentRoles", LoadRawTypoDictionary_SkipContexts_ComponentRoles()
    Set LoadRawTypoDictionary_SkipContexts = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTypoDictionary_SkipContexts", Err.description
End Function

Public Function LoadRawTypoDictionary() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "1.0.0"
    d.Add "sourceLabel", "TH" & ChrW(&HD4) & "NG L" & ChrW(&H1EC6)
    d.Add "ruleCode", "LOCAL-TYPO-DICT"
    d.Add "matchWholeWord", True
    d.Add "preserveCase", True
    d.Add "requiresConfirmation", True
    d.Add "corrections", LoadRawTypoDictionary_Corrections()
    d.Add "proposedCorrections", CreateObject("Scripting.Dictionary")
    d.Add "protectedForms", LoadRawTypoDictionary_ProtectedForms()
    d.Add "pendingReview", LoadRawTypoDictionary_PendingReview()
    d.Add "skipContexts", LoadRawTypoDictionary_SkipContexts()
    Set LoadRawTypoDictionary = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawTypoDictionary", Err.description
End Function

Private Function LoadRawUnicodeToNfc_CombiningToPrecomposed() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "a" & ChrW(&H300), ChrW(&HE0)
    d.Add "A" & ChrW(&H300), ChrW(&HC0)
    d.Add "a" & ChrW(&H301), ChrW(&HE1)
    d.Add "A" & ChrW(&H301), ChrW(&HC1)
    d.Add "a" & ChrW(&H309), ChrW(&H1EA3)
    d.Add "A" & ChrW(&H309), ChrW(&H1EA2)
    d.Add "a" & ChrW(&H303), ChrW(&HE3)
    d.Add "A" & ChrW(&H303), ChrW(&HC3)
    d.Add "a" & ChrW(&H323), ChrW(&H1EA1)
    d.Add "A" & ChrW(&H323), ChrW(&H1EA0)
    d.Add "e" & ChrW(&H300), ChrW(&HE8)
    d.Add "E" & ChrW(&H300), ChrW(&HC8)
    d.Add "e" & ChrW(&H301), ChrW(&HE9)
    d.Add "E" & ChrW(&H301), ChrW(&HC9)
    d.Add "e" & ChrW(&H309), ChrW(&H1EBB)
    d.Add "E" & ChrW(&H309), ChrW(&H1EBA)
    d.Add "e" & ChrW(&H303), ChrW(&H1EBD)
    d.Add "E" & ChrW(&H303), ChrW(&H1EBC)
    d.Add "e" & ChrW(&H323), ChrW(&H1EB9)
    d.Add "E" & ChrW(&H323), ChrW(&H1EB8)
    d.Add "i" & ChrW(&H300), ChrW(&HEC)
    d.Add "I" & ChrW(&H300), ChrW(&HCC)
    d.Add "i" & ChrW(&H301), ChrW(&HED)
    d.Add "I" & ChrW(&H301), ChrW(&HCD)
    d.Add "i" & ChrW(&H309), ChrW(&H1EC9)
    d.Add "I" & ChrW(&H309), ChrW(&H1EC8)
    d.Add "i" & ChrW(&H303), ChrW(&H129)
    d.Add "I" & ChrW(&H303), ChrW(&H128)
    d.Add "i" & ChrW(&H323), ChrW(&H1ECB)
    d.Add "I" & ChrW(&H323), ChrW(&H1ECA)
    d.Add "o" & ChrW(&H300), ChrW(&HF2)
    d.Add "O" & ChrW(&H300), ChrW(&HD2)
    d.Add "o" & ChrW(&H301), ChrW(&HF3)
    d.Add "O" & ChrW(&H301), ChrW(&HD3)
    d.Add "o" & ChrW(&H309), ChrW(&H1ECF)
    d.Add "O" & ChrW(&H309), ChrW(&H1ECE)
    d.Add "o" & ChrW(&H303), ChrW(&HF5)
    d.Add "O" & ChrW(&H303), ChrW(&HD5)
    d.Add "o" & ChrW(&H323), ChrW(&H1ECD)
    d.Add "O" & ChrW(&H323), ChrW(&H1ECC)
    d.Add "u" & ChrW(&H300), ChrW(&HF9)
    d.Add "U" & ChrW(&H300), ChrW(&HD9)
    d.Add "u" & ChrW(&H301), ChrW(&HFA)
    d.Add "U" & ChrW(&H301), ChrW(&HDA)
    d.Add "u" & ChrW(&H309), ChrW(&H1EE7)
    d.Add "U" & ChrW(&H309), ChrW(&H1EE6)
    d.Add "u" & ChrW(&H303), ChrW(&H169)
    d.Add "U" & ChrW(&H303), ChrW(&H168)
    d.Add "u" & ChrW(&H323), ChrW(&H1EE5)
    d.Add "U" & ChrW(&H323), ChrW(&H1EE4)
    d.Add "y" & ChrW(&H300), ChrW(&H1EF3)
    d.Add "Y" & ChrW(&H300), ChrW(&H1EF2)
    d.Add "y" & ChrW(&H301), ChrW(&HFD)
    d.Add "Y" & ChrW(&H301), ChrW(&HDD)
    d.Add "y" & ChrW(&H309), ChrW(&H1EF7)
    d.Add "Y" & ChrW(&H309), ChrW(&H1EF6)
    d.Add "y" & ChrW(&H303), ChrW(&H1EF9)
    d.Add "Y" & ChrW(&H303), ChrW(&H1EF8)
    d.Add "y" & ChrW(&H323), ChrW(&H1EF5)
    d.Add "Y" & ChrW(&H323), ChrW(&H1EF4)
    d.Add "a" & ChrW(&H306) & ChrW(&H300), ChrW(&H1EB1)
    d.Add "A" & ChrW(&H306) & ChrW(&H300), ChrW(&H1EB0)
    d.Add "a" & ChrW(&H306) & ChrW(&H301), ChrW(&H1EAF)
    d.Add "A" & ChrW(&H306) & ChrW(&H301), ChrW(&H1EAE)
    d.Add "a" & ChrW(&H306) & ChrW(&H309), ChrW(&H1EB3)
    d.Add "A" & ChrW(&H306) & ChrW(&H309), ChrW(&H1EB2)
    d.Add "a" & ChrW(&H306) & ChrW(&H303), ChrW(&H1EB5)
    d.Add "A" & ChrW(&H306) & ChrW(&H303), ChrW(&H1EB4)
    d.Add "a" & ChrW(&H323) & ChrW(&H306), ChrW(&H1EB7)
    d.Add "A" & ChrW(&H323) & ChrW(&H306), ChrW(&H1EB6)
    d.Add "a" & ChrW(&H302) & ChrW(&H300), ChrW(&H1EA7)
    d.Add "A" & ChrW(&H302) & ChrW(&H300), ChrW(&H1EA6)
    d.Add "a" & ChrW(&H302) & ChrW(&H301), ChrW(&H1EA5)
    d.Add "A" & ChrW(&H302) & ChrW(&H301), ChrW(&H1EA4)
    d.Add "a" & ChrW(&H302) & ChrW(&H309), ChrW(&H1EA9)
    d.Add "A" & ChrW(&H302) & ChrW(&H309), ChrW(&H1EA8)
    d.Add "a" & ChrW(&H302) & ChrW(&H303), ChrW(&H1EAB)
    d.Add "A" & ChrW(&H302) & ChrW(&H303), ChrW(&H1EAA)
    d.Add "a" & ChrW(&H323) & ChrW(&H302), ChrW(&H1EAD)
    d.Add "A" & ChrW(&H323) & ChrW(&H302), ChrW(&H1EAC)
    d.Add "e" & ChrW(&H302) & ChrW(&H300), ChrW(&H1EC1)
    d.Add "E" & ChrW(&H302) & ChrW(&H300), ChrW(&H1EC0)
    d.Add "e" & ChrW(&H302) & ChrW(&H301), ChrW(&H1EBF)
    d.Add "E" & ChrW(&H302) & ChrW(&H301), ChrW(&H1EBE)
    d.Add "e" & ChrW(&H302) & ChrW(&H309), ChrW(&H1EC3)
    d.Add "E" & ChrW(&H302) & ChrW(&H309), ChrW(&H1EC2)
    d.Add "e" & ChrW(&H302) & ChrW(&H303), ChrW(&H1EC5)
    d.Add "E" & ChrW(&H302) & ChrW(&H303), ChrW(&H1EC4)
    d.Add "e" & ChrW(&H323) & ChrW(&H302), ChrW(&H1EC7)
    d.Add "E" & ChrW(&H323) & ChrW(&H302), ChrW(&H1EC6)
    d.Add "o" & ChrW(&H302) & ChrW(&H300), ChrW(&H1ED3)
    d.Add "O" & ChrW(&H302) & ChrW(&H300), ChrW(&H1ED2)
    d.Add "o" & ChrW(&H302) & ChrW(&H301), ChrW(&H1ED1)
    d.Add "O" & ChrW(&H302) & ChrW(&H301), ChrW(&H1ED0)
    d.Add "o" & ChrW(&H302) & ChrW(&H309), ChrW(&H1ED5)
    d.Add "O" & ChrW(&H302) & ChrW(&H309), ChrW(&H1ED4)
    d.Add "o" & ChrW(&H302) & ChrW(&H303), ChrW(&H1ED7)
    d.Add "O" & ChrW(&H302) & ChrW(&H303), ChrW(&H1ED6)
    d.Add "o" & ChrW(&H323) & ChrW(&H302), ChrW(&H1ED9)
    d.Add "O" & ChrW(&H323) & ChrW(&H302), ChrW(&H1ED8)
    d.Add "o" & ChrW(&H31B) & ChrW(&H300), ChrW(&H1EDD)
    d.Add "O" & ChrW(&H31B) & ChrW(&H300), ChrW(&H1EDC)
    d.Add "o" & ChrW(&H31B) & ChrW(&H301), ChrW(&H1EDB)
    d.Add "O" & ChrW(&H31B) & ChrW(&H301), ChrW(&H1EDA)
    d.Add "o" & ChrW(&H31B) & ChrW(&H309), ChrW(&H1EDF)
    d.Add "O" & ChrW(&H31B) & ChrW(&H309), ChrW(&H1EDE)
    d.Add "o" & ChrW(&H31B) & ChrW(&H303), ChrW(&H1EE1)
    d.Add "O" & ChrW(&H31B) & ChrW(&H303), ChrW(&H1EE0)
    d.Add "o" & ChrW(&H31B) & ChrW(&H323), ChrW(&H1EE3)
    d.Add "O" & ChrW(&H31B) & ChrW(&H323), ChrW(&H1EE2)
    d.Add "u" & ChrW(&H31B) & ChrW(&H300), ChrW(&H1EEB)
    d.Add "U" & ChrW(&H31B) & ChrW(&H300), ChrW(&H1EEA)
    d.Add "u" & ChrW(&H31B) & ChrW(&H301), ChrW(&H1EE9)
    d.Add "U" & ChrW(&H31B) & ChrW(&H301), ChrW(&H1EE8)
    d.Add "u" & ChrW(&H31B) & ChrW(&H309), ChrW(&H1EED)
    d.Add "U" & ChrW(&H31B) & ChrW(&H309), ChrW(&H1EEC)
    d.Add "u" & ChrW(&H31B) & ChrW(&H303), ChrW(&H1EEF)
    d.Add "U" & ChrW(&H31B) & ChrW(&H303), ChrW(&H1EEE)
    d.Add "u" & ChrW(&H31B) & ChrW(&H323), ChrW(&H1EF1)
    d.Add "U" & ChrW(&H31B) & ChrW(&H323), ChrW(&H1EF0)
    d.Add "a" & ChrW(&H306), ChrW(&H103)
    d.Add "A" & ChrW(&H306), ChrW(&H102)
    d.Add "a" & ChrW(&H302), ChrW(&HE2)
    d.Add "A" & ChrW(&H302), ChrW(&HC2)
    d.Add "e" & ChrW(&H302), ChrW(&HEA)
    d.Add "E" & ChrW(&H302), ChrW(&HCA)
    d.Add "o" & ChrW(&H302), ChrW(&HF4)
    d.Add "O" & ChrW(&H302), ChrW(&HD4)
    d.Add "o" & ChrW(&H31B), ChrW(&H1A1)
    d.Add "O" & ChrW(&H31B), ChrW(&H1A0)
    d.Add "u" & ChrW(&H31B), ChrW(&H1B0)
    d.Add "U" & ChrW(&H31B), ChrW(&H1AF)
    Set LoadRawUnicodeToNfc_CombiningToPrecomposed = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawUnicodeToNfc_CombiningToPrecomposed", Err.description
End Function

Private Function LoadRawUnicodeToNfc_HiddenChars_Remove() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add ChrW(&H200B)
    c.Add ChrW(&H200C)
    c.Add ChrW(&H200D)
    c.Add ChrW(&HFEFF)
    Set LoadRawUnicodeToNfc_HiddenChars_Remove = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawUnicodeToNfc_HiddenChars_Remove", Err.description
End Function

Private Function LoadRawUnicodeToNfc_HiddenChars_ReplaceWithSpace() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add ChrW(&H2028)
    c.Add ChrW(&HA0)
    Set LoadRawUnicodeToNfc_HiddenChars_ReplaceWithSpace = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawUnicodeToNfc_HiddenChars_ReplaceWithSpace", Err.description
End Function

Private Function LoadRawUnicodeToNfc_HiddenChars() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "remove", LoadRawUnicodeToNfc_HiddenChars_Remove()
    d.Add "replaceWithSpace", LoadRawUnicodeToNfc_HiddenChars_ReplaceWithSpace()
    Set LoadRawUnicodeToNfc_HiddenChars = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawUnicodeToNfc_HiddenChars", Err.description
End Function

Private Function LoadRawUnicodeToNfc_TestCases_MustNormalize() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim c232 As New Collection
    c232.Add "Vie" & ChrW(&H323) & ChrW(&H302) & "t Nam"
    c232.Add "Vi" & ChrW(&H1EC7) & "t Nam"
    c.Add c232
    Dim c233 As New Collection
    c233.Add "ho" & ChrW(&H300) & "a bi" & ChrW(&H300) & "nh"
    c233.Add "h" & ChrW(&HF2) & "a b" & ChrW(&HEC) & "nh"
    c.Add c233
    Dim c234 As New Collection
    c234.Add "Co" & ChrW(&H323) & ChrW(&H302) & "ng ho" & ChrW(&H300) & "a xa" & ChrW(&H303) & " ho" & ChrW(&H323) & ChrW(&H302) & "i chu" & ChrW(&H309) & " nghi" & ChrW(&H303) & "a Vie" & ChrW(&H323) & ChrW(&H302) & "t Nam"
    c234.Add "C" & ChrW(&H1ED9) & "ng h" & ChrW(&HF2) & "a x" & ChrW(&HE3) & " h" & ChrW(&H1ED9) & "i ch" & ChrW(&H1EE7) & " ngh" & ChrW(&H129) & "a Vi" & ChrW(&H1EC7) & "t Nam"
    c.Add c234
    Set LoadRawUnicodeToNfc_TestCases_MustNormalize = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawUnicodeToNfc_TestCases_MustNormalize", Err.description
End Function

Private Function LoadRawUnicodeToNfc_TestCases_MustKeepUnchanged() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "Hello World 123"
    c.Add ChrW(&H3C0) & " " & ChrW(&H2248) & " 3.14"
    c.Add ChrW(&HD83D) & ChrW(&HDE00) & " emoji khong doi"
    c.Add "Vi" & ChrW(&H1EC7) & "t Nam"
    Set LoadRawUnicodeToNfc_TestCases_MustKeepUnchanged = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawUnicodeToNfc_TestCases_MustKeepUnchanged", Err.description
End Function

Private Function LoadRawUnicodeToNfc_TestCases() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "mustNormalize", LoadRawUnicodeToNfc_TestCases_MustNormalize()
    d.Add "mustKeepUnchanged", LoadRawUnicodeToNfc_TestCases_MustKeepUnchanged()
    Set LoadRawUnicodeToNfc_TestCases = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawUnicodeToNfc_TestCases", Err.description
End Function

Public Function LoadRawUnicodeToNfc() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "1.0.0"
    d.Add "sourceLabel", "ND30"
    d.Add "ruleCode", "ND30-PL1-M1-K4-NFC"
    d.Add "actionType", "B"
    d.Add "combiningToPrecomposed", LoadRawUnicodeToNfc_CombiningToPrecomposed()
    d.Add "hiddenChars", LoadRawUnicodeToNfc_HiddenChars()
    d.Add "testCases", LoadRawUnicodeToNfc_TestCases()
    Set LoadRawUnicodeToNfc = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawUnicodeToNfc", Err.description
End Function

Private Function LoadRawCitationRules_DocumentTypes_Legislative() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "Hi" & ChrW(&H1EBF) & "n ph" & ChrW(&HE1) & "p"
    c.Add "B" & ChrW(&H1ED9) & " lu" & ChrW(&H1EAD) & "t"
    c.Add "Lu" & ChrW(&H1EAD) & "t"
    c.Add "Ph" & ChrW(&HE1) & "p l" & ChrW(&H1EC7) & "nh"
    c.Add "L" & ChrW(&H1EC7) & "nh"
    c.Add "Ngh" & ChrW(&H1ECB) & " quy" & ChrW(&H1EBF) & "t"
    c.Add "Ngh" & ChrW(&H1ECB) & " quy" & ChrW(&H1EBF) & "t li" & ChrW(&HEA) & "n t" & ChrW(&H1ECB) & "ch"
    c.Add "Ngh" & ChrW(&H1ECB) & " " & ChrW(&H111) & ChrW(&H1ECB) & "nh"
    c.Add "Quy" & ChrW(&H1EBF) & "t " & ChrW(&H111) & ChrW(&H1ECB) & "nh"
    c.Add "Th" & ChrW(&HF4) & "ng t" & ChrW(&H1B0)
    c.Add "Th" & ChrW(&HF4) & "ng t" & ChrW(&H1B0) & " li" & ChrW(&HEA) & "n t" & ChrW(&H1ECB) & "ch"
    c.Add "Ch" & ChrW(&H1EC9) & " th" & ChrW(&H1ECB)
    Set LoadRawCitationRules_DocumentTypes_Legislative = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCitationRules_DocumentTypes_Legislative", Err.description
End Function

Private Function LoadRawCitationRules_DocumentTypes_Administrative() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "Quy ch" & ChrW(&H1EBF)
    c.Add "Quy " & ChrW(&H111) & ChrW(&H1ECB) & "nh"
    c.Add "Th" & ChrW(&HF4) & "ng c" & ChrW(&HE1) & "o"
    c.Add "Th" & ChrW(&HF4) & "ng b" & ChrW(&HE1) & "o"
    c.Add "H" & ChrW(&H1B0) & ChrW(&H1EDB) & "ng d" & ChrW(&H1EAB) & "n"
    c.Add "Ch" & ChrW(&H1B0) & ChrW(&H1A1) & "ng tr" & ChrW(&HEC) & "nh"
    c.Add "K" & ChrW(&H1EBF) & " ho" & ChrW(&H1EA1) & "ch"
    c.Add "Ph" & ChrW(&H1B0) & ChrW(&H1A1) & "ng " & ChrW(&HE1) & "n"
    c.Add ChrW(&H110) & ChrW(&H1EC1) & " " & ChrW(&HE1) & "n"
    c.Add "D" & ChrW(&H1EF1) & " " & ChrW(&HE1) & "n"
    c.Add "B" & ChrW(&HE1) & "o c" & ChrW(&HE1) & "o"
    c.Add "Bi" & ChrW(&HEA) & "n b" & ChrW(&H1EA3) & "n"
    c.Add "T" & ChrW(&H1EDD) & " tr" & ChrW(&HEC) & "nh"
    c.Add "H" & ChrW(&H1EE3) & "p " & ChrW(&H111) & ChrW(&H1ED3) & "ng"
    c.Add "C" & ChrW(&HF4) & "ng v" & ChrW(&H103) & "n"
    c.Add "C" & ChrW(&HF4) & "ng " & ChrW(&H111) & "i" & ChrW(&H1EC7) & "n"
    c.Add "B" & ChrW(&H1EA3) & "n ghi nh" & ChrW(&H1EDB)
    c.Add "B" & ChrW(&H1EA3) & "n th" & ChrW(&H1ECF) & "a thu" & ChrW(&H1EAD) & "n"
    c.Add "Gi" & ChrW(&H1EA5) & "y " & ChrW(&H1EE7) & "y quy" & ChrW(&H1EC1) & "n"
    Set LoadRawCitationRules_DocumentTypes_Administrative = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCitationRules_DocumentTypes_Administrative", Err.description
End Function

Private Function LoadRawCitationRules_DocumentTypes() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "legislative", LoadRawCitationRules_DocumentTypes_Legislative()
    d.Add "administrative", LoadRawCitationRules_DocumentTypes_Administrative()
    Set LoadRawCitationRules_DocumentTypes = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCitationRules_DocumentTypes", Err.description
End Function

Private Function LoadRawCitationRules_Rules_0_Examples() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim d236 As Object
    Set d236 = CreateObject("Scripting.Dictionary")
    d236.Add "wrong", "Ngh" & ChrW(&H1ECB) & " " & ChrW(&H111) & ChrW(&H1ECB) & "nh 30/2020/N" & ChrW(&H110) & "-CP"
    d236.Add "right", "Ngh" & ChrW(&H1ECB) & " " & ChrW(&H111) & ChrW(&H1ECB) & "nh s" & ChrW(&H1ED1) & " 30/2020/N" & ChrW(&H110) & "-CP"
    c.Add d236
    Dim d237 As Object
    Set d237 = CreateObject("Scripting.Dictionary")
    d237.Add "wrong", "Th" & ChrW(&HF4) & "ng t" & ChrW(&H1B0) & " 01/2019/TT-BNV"
    d237.Add "right", "Th" & ChrW(&HF4) & "ng t" & ChrW(&H1B0) & " s" & ChrW(&H1ED1) & " 01/2019/TT-BNV"
    c.Add d237
    Dim d238 As Object
    Set d238 = CreateObject("Scripting.Dictionary")
    d238.Add "wrong", "Quy" & ChrW(&H1EBF) & "t " & ChrW(&H111) & ChrW(&H1ECB) & "nh 1989/Q" & ChrW(&H110) & "-BGD" & ChrW(&H110) & "T"
    d238.Add "right", "Quy" & ChrW(&H1EBF) & "t " & ChrW(&H111) & ChrW(&H1ECB) & "nh s" & ChrW(&H1ED1) & " 1989/Q" & ChrW(&H110) & "-BGD" & ChrW(&H110) & "T"
    c.Add d238
    Dim d239 As Object
    Set d239 = CreateObject("Scripting.Dictionary")
    d239.Add "wrong", "C" & ChrW(&HF4) & "ng v" & ChrW(&H103) & "n 123/BNV-VP"
    d239.Add "right", "C" & ChrW(&HF4) & "ng v" & ChrW(&H103) & "n s" & ChrW(&H1ED1) & " 123/BNV-VP"
    c.Add d239
    Set LoadRawCitationRules_Rules_0_Examples = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCitationRules_Rules_0_Examples", Err.description
End Function

Private Function LoadRawCitationRules_Rules_0_Exceptions_0_Examples() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "Lu" & ChrW(&H1EAD) & "t T" & ChrW(&H1ED5) & " ch" & ChrW(&H1EE9) & "c Qu" & ChrW(&H1ED1) & "c h" & ChrW(&H1ED9) & "i"
    c.Add "B" & ChrW(&H1ED9) & " lu" & ChrW(&H1EAD) & "t H" & ChrW(&HEC) & "nh s" & ChrW(&H1EF1)
    c.Add "Ph" & ChrW(&HE1) & "p l" & ChrW(&H1EC7) & "nh " & ChrW(&H1AF) & "u " & ChrW(&H111) & ChrW(&HE3) & "i ng" & ChrW(&H1B0) & ChrW(&H1EDD) & "i c" & ChrW(&HF3) & " c" & ChrW(&HF4) & "ng v" & ChrW(&H1EDB) & "i c" & ChrW(&HE1) & "ch m" & ChrW(&H1EA1) & "ng"
    Set LoadRawCitationRules_Rules_0_Exceptions_0_Examples = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCitationRules_Rules_0_Exceptions_0_Examples", Err.description
End Function

Private Function LoadRawCitationRules_Rules_0_Exceptions() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim d240 As Object
    Set d240 = CreateObject("Scripting.Dictionary")
    d240.Add "case", "Lu" & ChrW(&H1EAD) & "t v" & ChrW(&HE0) & " Ph" & ChrW(&HE1) & "p l" & ChrW(&H1EC7) & "nh"
    d240.Add "rule", "Kh" & ChrW(&HF4) & "ng ghi s" & ChrW(&H1ED1) & ", k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u, c" & ChrW(&H1A1) & " quan ban h" & ChrW(&HE0) & "nh."
    d240.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 6, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m a v" & ChrW(&HE0) & " " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    d240.Add "examples", LoadRawCitationRules_Rules_0_Exceptions_0_Examples()
    c.Add d240
    Dim d241 As Object
    Set d241 = CreateObject("Scripting.Dictionary")
    d241.Add "case", "Trong k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u c" & ChrW(&H1EE7) & "a ch" & ChrW(&HED) & "nh v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n " & ChrW(&H111) & "ang so" & ChrW(&H1EA1) & "n"
    d241.Add "rule", ChrW(&HD4) & " s" & ChrW(&H1ED1) & " 3 d" & ChrW(&HF9) & "ng d" & ChrW(&H1EA1) & "ng 'S" & ChrW(&H1ED1) & ": 15/Q" & ChrW(&H110) & "-BNV', kh" & ChrW(&HF4) & "ng ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EC7) & "n d" & ChrW(&H1EAB) & "n."
    c.Add d241
    Dim d242 As Object
    Set d242 = CreateObject("Scripting.Dictionary")
    d242.Add "case", "Trong " & ChrW(&H111) & "o" & ChrW(&H1EA1) & "n tr" & ChrW(&HED) & "ch d" & ChrW(&H1EAB) & "n nguy" & ChrW(&HEA) & "n v" & ChrW(&H103) & "n"
    d242.Add "rule", "Kh" & ChrW(&HF4) & "ng s" & ChrW(&H1EED) & "a l" & ChrW(&H1EDD) & "i v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n kh" & ChrW(&HE1) & "c."
    c.Add d242
    Set LoadRawCitationRules_Rules_0_Exceptions = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCitationRules_Rules_0_Exceptions", Err.description
End Function

Private Function LoadRawCitationRules_Rules_1_DetectPatterns() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "ng" & ChrW(&HE0) & "y\s+(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})"
    c.Add "(?<!ng" & ChrW(&HE0) & "y\s)(\d{1,2})[/.](\d{1,2})[/.](\d{4})"
    Set LoadRawCitationRules_Rules_1_DetectPatterns = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCitationRules_Rules_1_DetectPatterns", Err.description
End Function

Private Function LoadRawCitationRules_Rules_1_PadMonthsList() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add 1
    c.Add 2
    Set LoadRawCitationRules_Rules_1_PadMonthsList = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCitationRules_Rules_1_PadMonthsList", Err.description
End Function

Private Function LoadRawCitationRules_Rules_1_Examples() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim d244 As Object
    Set d244 = CreateObject("Scripting.Dictionary")
    d244.Add "wrong", "ng" & ChrW(&HE0) & "y 05/3/2020"
    d244.Add "right", "ng" & ChrW(&HE0) & "y 05 th" & ChrW(&HE1) & "ng 3 n" & ChrW(&H103) & "m 2020"
    c.Add d244
    Dim d245 As Object
    Set d245 = CreateObject("Scripting.Dictionary")
    d245.Add "wrong", "ng" & ChrW(&HE0) & "y 25/5/2018"
    d245.Add "right", "ng" & ChrW(&HE0) & "y 25 th" & ChrW(&HE1) & "ng 5 n" & ChrW(&H103) & "m 2018"
    c.Add d245
    Dim d246 As Object
    Set d246 = CreateObject("Scripting.Dictionary")
    d246.Add "wrong", "ng" & ChrW(&HE0) & "y 1/1/2020"
    d246.Add "right", "ng" & ChrW(&HE0) & "y 01 th" & ChrW(&HE1) & "ng 01 n" & ChrW(&H103) & "m 2020"
    c.Add d246
    Dim d247 As Object
    Set d247 = CreateObject("Scripting.Dictionary")
    d247.Add "wrong", "ng" & ChrW(&HE0) & "y 15-11-2021"
    d247.Add "right", "ng" & ChrW(&HE0) & "y 15 th" & ChrW(&HE1) & "ng 11 n" & ChrW(&H103) & "m 2021"
    c.Add d247
    Set LoadRawCitationRules_Rules_1_Examples = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCitationRules_Rules_1_Examples", Err.description
End Function

Private Function LoadRawCitationRules_Rules_1_Exceptions() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim d248 As Object
    Set d248 = CreateObject("Scripting.Dictionary")
    d248.Add "case", "Trong b" & ChrW(&H1EA3) & "ng bi" & ChrW(&H1EC3) & "u"
    d248.Add "rule", "B" & ChrW(&H1EA3) & "ng bi" & ChrW(&H1EC3) & "u th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng d" & ChrW(&HF9) & "ng d" & ChrW(&H1EA1) & "ng r" & ChrW(&HFA) & "t g" & ChrW(&H1ECD) & "n " & ChrW(&H111) & ChrW(&H1EC3) & " ti" & ChrW(&H1EBF) & "t ki" & ChrW(&H1EC7) & "m ch" & ChrW(&H1ED7) & "."
    d248.Add "recommendation", "c" & ChrW(&H1EA3) & "nh b" & ChrW(&HE1) & "o, kh" & ChrW(&HF4) & "ng t" & ChrW(&H1EF1) & " s" & ChrW(&H1EED) & "a"
    c.Add d248
    Dim d249 As Object
    Set d249 = CreateObject("Scripting.Dictionary")
    d249.Add "case", "Trong " & ChrW(&H111) & "o" & ChrW(&H1EA1) & "n tr" & ChrW(&HED) & "ch d" & ChrW(&H1EAB) & "n nguy" & ChrW(&HEA) & "n v" & ChrW(&H103) & "n"
    d249.Add "rule", "Kh" & ChrW(&HF4) & "ng s" & ChrW(&H1EED) & "a."
    c.Add d249
    Set LoadRawCitationRules_Rules_1_Exceptions = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCitationRules_Rules_1_Exceptions", Err.description
End Function

Private Function LoadRawCitationRules_Rules_2_SubsequentCitation() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "rule", "Trong c" & ChrW(&HE1) & "c l" & ChrW(&H1EA7) & "n vi" & ChrW(&H1EC7) & "n d" & ChrW(&H1EAB) & "n ti" & ChrW(&H1EBF) & "p theo, ch" & ChrW(&H1EC9) & " ghi t" & ChrW(&HEA) & "n lo" & ChrW(&H1EA1) & "i v" & ChrW(&HE0) & " s" & ChrW(&H1ED1) & ", k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u c" & ChrW(&H1EE7) & "a v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n " & ChrW(&H111) & ChrW(&HF3) & "."
    d.Add "example", "Ngh" & ChrW(&H1ECB) & " " & ChrW(&H111) & ChrW(&H1ECB) & "nh s" & ChrW(&H1ED1) & " 30/2020/N" & ChrW(&H110) & "-CP"
    d.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, Ph" & ChrW(&H1EA7) & "n I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 6, " & ChrW(&H111) & "i" & ChrW(&H1EC3) & "m b"
    Set LoadRawCitationRules_Rules_2_SubsequentCitation = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCitationRules_Rules_2_SubsequentCitation", Err.description
End Function

Private Function LoadRawCitationRules_Rules() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim d235 As Object
    Set d235 = CreateObject("Scripting.Dictionary")
    d235.Add "ruleCode", "ND30-PL1-M2-K6B-SO"
    d235.Add "name", "requireSoKeyword"
    d235.Add "description", "Sau t" & ChrW(&HEA) & "n lo" & ChrW(&H1EA1) & "i v" & ChrW(&H103) & "n b" & ChrW(&H1EA3) & "n ph" & ChrW(&H1EA3) & "i c" & ChrW(&HF3) & " ch" & ChrW(&H1EEF) & " 's" & ChrW(&H1ED1) & "' tr" & ChrW(&H1B0) & ChrW(&H1EDB) & "c s" & ChrW(&H1ED1) & " hi" & ChrW(&H1EC7) & "u."
    d235.Add "detectPattern", "(?<typeName>{documentTypes})\s+(?<code>\d+[/-][\p{L}\d/-]+)"
    d235.Add "fixTemplate", "{typeName} s" & ChrW(&H1ED1) & " {code}"
    d235.Add "examples", LoadRawCitationRules_Rules_0_Examples()
    d235.Add "exceptions", LoadRawCitationRules_Rules_0_Exceptions()
    c.Add d235
    Dim d243 As Object
    Set d243 = CreateObject("Scripting.Dictionary")
    d243.Add "ruleCode", "ND30-PL1-M2-K6B-DATE"
    d243.Add "name", "requireFullDateForm"
    d243.Add "description", "Th" & ChrW(&H1EDD) & "i gian ban h" & ChrW(&HE0) & "nh khi vi" & ChrW(&H1EC7) & "n d" & ChrW(&H1EAB) & "n ph" & ChrW(&H1EA3) & "i vi" & ChrW(&H1EBF) & "t " & ChrW(&H111) & ChrW(&H1EA7) & "y " & ChrW(&H111) & ChrW(&H1EE7) & " d" & ChrW(&H1EA1) & "ng 'ng" & ChrW(&HE0) & "y ... th" & ChrW(&HE1) & "ng ... n" & ChrW(&H103) & "m ...'."
    d243.Add "detectPatterns", LoadRawCitationRules_Rules_1_DetectPatterns()
    d243.Add "fixTemplate", "ng" & ChrW(&HE0) & "y {dd} th" & ChrW(&HE1) & "ng {m} n" & ChrW(&H103) & "m {yyyy}"
    d243.Add "padDayBelow", 10
    d243.Add "padMonthsList", LoadRawCitationRules_Rules_1_PadMonthsList()
    d243.Add "examples", LoadRawCitationRules_Rules_1_Examples()
    d243.Add "exceptions", LoadRawCitationRules_Rules_1_Exceptions()
    c.Add d243
    Dim d250 As Object
    Set d250 = CreateObject("Scripting.Dictionary")
    d250.Add "ruleCode", "ND30-PL1-M2-K6B-CITE"
    d250.Add "name", "requireFullFirstCitation"
    d250.Add "description", "Vi" & ChrW(&H1EC7) & "n d" & ChrW(&H1EAB) & "n l" & ChrW(&H1EA7) & "n " & ChrW(&H111) & ChrW(&H1EA7) & "u ph" & ChrW(&H1EA3) & "i ghi " & ChrW(&H111) & ChrW(&H1EA7) & "y " & ChrW(&H111) & ChrW(&H1EE7) & ": t" & ChrW(&HEA) & "n lo" & ChrW(&H1EA1) & "i, s" & ChrW(&H1ED1) & ", k" & ChrW(&HFD) & " hi" & ChrW(&H1EC7) & "u, th" & ChrW(&H1EDD) & "i gian ban h" & ChrW(&HE0) & "nh, t" & ChrW(&HEA) & "n c" & ChrW(&H1A1) & " quan ban h" & ChrW(&HE0) & "nh, tr" & ChrW(&HED) & "ch y" & ChrW(&H1EBF) & "u n" & ChrW(&H1ED9) & "i dung."
    d250.Add "checkability", "partial"
    d250.Add "fullForm", "{typeName} s" & ChrW(&H1ED1) & " {codeNumber}/{codeNotation} ng" & ChrW(&HE0) & "y {dd} th" & ChrW(&HE1) & "ng {m} n" & ChrW(&H103) & "m {yyyy} c" & ChrW(&H1EE7) & "a {organName} v" & ChrW(&H1EC1) & " {subject}"
    d250.Add "example", "Ngh" & ChrW(&H1ECB) & " " & ChrW(&H111) & ChrW(&H1ECB) & "nh s" & ChrW(&H1ED1) & " 30/2020/N" & ChrW(&H110) & "-CP ng" & ChrW(&HE0) & "y 05 th" & ChrW(&HE1) & "ng 3 n" & ChrW(&H103) & "m 2020 c" & ChrW(&H1EE7) & "a Ch" & ChrW(&HED) & "nh ph" & ChrW(&H1EE7) & " v" & ChrW(&H1EC1) & " c" & ChrW(&HF4) & "ng t" & ChrW(&HE1) & "c v" & ChrW(&H103) & "n th" & ChrW(&H1B0)
    d250.Add "subsequentCitation", LoadRawCitationRules_Rules_2_SubsequentCitation()
    c.Add d250
    Set LoadRawCitationRules_Rules = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCitationRules_Rules", Err.description
End Function

Private Function LoadRawCitationRules_TestCases_MustFlag() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "Ngh" & ChrW(&H1ECB) & " " & ChrW(&H111) & ChrW(&H1ECB) & "nh 30/2020/N" & ChrW(&H110) & "-CP ng" & ChrW(&HE0) & "y 05/3/2020"
    c.Add "C" & ChrW(&H103) & "n c" & ChrW(&H1EE9) & " Th" & ChrW(&HF4) & "ng t" & ChrW(&H1B0) & " 01/2019/TT-BNV"
    c.Add "theo Quy" & ChrW(&H1EBF) & "t " & ChrW(&H111) & ChrW(&H1ECB) & "nh 1989/Q" & ChrW(&H110) & "-BGD" & ChrW(&H110) & "T ng" & ChrW(&HE0) & "y 25/5/2018"
    Set LoadRawCitationRules_TestCases_MustFlag = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCitationRules_TestCases_MustFlag", Err.description
End Function

Private Function LoadRawCitationRules_TestCases_MustNotFlag() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "Ngh" & ChrW(&H1ECB) & " " & ChrW(&H111) & ChrW(&H1ECB) & "nh s" & ChrW(&H1ED1) & " 30/2020/N" & ChrW(&H110) & "-CP ng" & ChrW(&HE0) & "y 05 th" & ChrW(&HE1) & "ng 3 n" & ChrW(&H103) & "m 2020 c" & ChrW(&H1EE7) & "a Ch" & ChrW(&HED) & "nh ph" & ChrW(&H1EE7)
    c.Add "Lu" & ChrW(&H1EAD) & "t T" & ChrW(&H1ED5) & " ch" & ChrW(&H1EE9) & "c Qu" & ChrW(&H1ED1) & "c h" & ChrW(&H1ED9) & "i"
    c.Add "B" & ChrW(&H1ED9) & " lu" & ChrW(&H1EAD) & "t H" & ChrW(&HEC) & "nh s" & ChrW(&H1EF1)
    c.Add "S" & ChrW(&H1ED1) & ": 15/Q" & ChrW(&H110) & "-BNV"
    c.Add "Ngh" & ChrW(&H1ECB) & " " & ChrW(&H111) & ChrW(&H1ECB) & "nh s" & ChrW(&H1ED1) & " 30/2020/N" & ChrW(&H110) & "-CP"
    Set LoadRawCitationRules_TestCases_MustNotFlag = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCitationRules_TestCases_MustNotFlag", Err.description
End Function

Private Function LoadRawCitationRules_TestCases() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "mustFlag", LoadRawCitationRules_TestCases_MustFlag()
    d.Add "mustNotFlag", LoadRawCitationRules_TestCases_MustNotFlag()
    Set LoadRawCitationRules_TestCases = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCitationRules_TestCases", Err.description
End Function

Public Function LoadRawCitationRules() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "1.0.0"
    d.Add "sourceLabel", "ND30"
    d.Add "actionType", "C"
    d.Add "documentTypes", LoadRawCitationRules_DocumentTypes()
    d.Add "rules", LoadRawCitationRules_Rules()
    d.Add "testCases", LoadRawCitationRules_TestCases()
    Set LoadRawCitationRules = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawCitationRules", Err.description
End Function

Private Function LoadRawSpecialCapitalizations_Entries() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim d251 As Object
    Set d251 = CreateObject("Scripting.Dictionary")
    d251.Add "phrase", "Th" & ChrW(&H1EE7) & " " & ChrW(&H111) & ChrW(&HF4) & " H" & ChrW(&HE0) & " N" & ChrW(&H1ED9) & "i"
    d251.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c III, kho" & ChrW(&H1EA3) & "n 1 " & ChrW(&H2014) & " vi" & ChrW(&H1EBF) & "t hoa " & ChrW(&H111) & ChrW(&H1EB7) & "c bi" & ChrW(&H1EC7) & "t"
    d251.Add "note", "Ch" & ChrW(&H1EEF) & " 'Th" & ChrW(&H1EE7) & " " & ChrW(&H111) & ChrW(&HF4) & "' vi" & ChrW(&H1EBF) & "t hoa c" & ChrW(&H1EA3) & " hai " & ChrW(&HE2) & "m ti" & ChrW(&H1EBF) & "t, kh" & ChrW(&HE1) & "c quy t" & ChrW(&H1EAF) & "c chung c" & ChrW(&H1EE7) & "a danh t" & ChrW(&H1EEB) & " chung + t" & ChrW(&HEA) & "n ri" & ChrW(&HEA) & "ng"
    c.Add d251
    Dim d252 As Object
    Set d252 = CreateObject("Scripting.Dictionary")
    d252.Add "phrase", "Th" & ChrW(&HE0) & "nh ph" & ChrW(&H1ED1) & " H" & ChrW(&H1ED3) & " Ch" & ChrW(&HED) & " Minh"
    d252.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c III, kho" & ChrW(&H1EA3) & "n 1 " & ChrW(&H2014) & " vi" & ChrW(&H1EBF) & "t hoa " & ChrW(&H111) & ChrW(&H1EB7) & "c bi" & ChrW(&H1EC7) & "t"
    Dim tmp18 As String
    tmp18 = "Ch" & ChrW(&H1EEF) & " 'Th" & ChrW(&HE0) & "nh ph" & ChrW(&H1ED1) & "' vi" & ChrW(&H1EBF) & "t hoa c" & ChrW(&H1EA3) & " hai " & ChrW(&HE2) & "m ti" & ChrW(&H1EBF) & "t. Kh" & ChrW(&HE1) & "c v" & ChrW(&H1EDB) & "i 'th" & ChrW(&HE0) & "nh ph" & ChrW(&H1ED1) & " Th" & ChrW(&HE1) & "i Nguy" & ChrW(&HEA) & "n' " & ChrW(&H2014) & " danh t" & ChrW(&H1EEB) & " chung 'th" & ChrW(&HE0) & "nh ph" & ChrW(&H1ED1) & "' vi" & ChrW(&H1EBF) & "t th" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng v" & ChrW(&HEC) & " kh" & ChrW(&HF4) & "ng ph" & ChrW(&H1EA3) & "i tr" & ChrW(&H1B0) & ChrW(&H1EDD) & "ng h" & ChrW(&H1EE3) & "p " & ChrW(&H111) & ChrW(&H1EB7) & "c bi" & ChrW(&H1EC7) & "t n"
    tmp18 = tmp18 & ChrW(&HE0) & "y"
    d252.Add "note", tmp18
    c.Add d252
    Dim d253 As Object
    Set d253 = CreateObject("Scripting.Dictionary")
    d253.Add "phrase", "Ban Ch" & ChrW(&H1EA5) & "p h" & ChrW(&HE0) & "nh Trung " & ChrW(&H1B0) & ChrW(&H1A1) & "ng " & ChrW(&H110) & ChrW(&H1EA3) & "ng C" & ChrW(&H1ED9) & "ng s" & ChrW(&H1EA3) & "n Vi" & ChrW(&H1EC7) & "t Nam"
    d253.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c IV, kho" & ChrW(&H1EA3) & "n 1 " & ChrW(&H2014) & " vi" & ChrW(&H1EBF) & "t hoa " & ChrW(&H111) & ChrW(&H1EB7) & "c bi" & ChrW(&H1EC7) & "t"
    d253.Add "note", "T" & ChrW(&HEA) & "n c" & ChrW(&H1A1) & " quan l" & ChrW(&HE3) & "nh " & ChrW(&H111) & ChrW(&H1EA1) & "o " & ChrW(&H110) & ChrW(&H1EA3) & "ng, vi" & ChrW(&H1EBF) & "t hoa theo danh s" & ChrW(&HE1) & "ch c" & ChrW(&H1EE9) & "ng"
    c.Add d253
    Dim d254 As Object
    Set d254 = CreateObject("Scripting.Dictionary")
    d254.Add "phrase", "V" & ChrW(&H103) & "n ph" & ChrW(&HF2) & "ng Trung " & ChrW(&H1B0) & ChrW(&H1A1) & "ng " & ChrW(&H110) & ChrW(&H1EA3) & "ng"
    d254.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c II, M" & ChrW(&H1EE5) & "c IV, kho" & ChrW(&H1EA3) & "n 1 " & ChrW(&H2014) & " vi" & ChrW(&H1EBF) & "t hoa " & ChrW(&H111) & ChrW(&H1EB7) & "c bi" & ChrW(&H1EC7) & "t"
    d254.Add "note", "T" & ChrW(&HEA) & "n c" & ChrW(&H1A1) & " quan, vi" & ChrW(&H1EBF) & "t hoa theo danh s" & ChrW(&HE1) & "ch c" & ChrW(&H1EE9) & "ng"
    c.Add d254
    Set LoadRawSpecialCapitalizations_Entries = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawSpecialCapitalizations_Entries", Err.description
End Function

Public Function LoadRawSpecialCapitalizations() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "1.0.0"
    d.Add "status", "day du theo vi du neu trong Phu luc II"
    d.Add "sourceLabel", "ND30"
    d.Add "actionType", "TU SUA"
    d.Add "entries", LoadRawSpecialCapitalizations_Entries()
    Set LoadRawSpecialCapitalizations = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawSpecialCapitalizations", Err.description
End Function

Private Function LoadRawNonSentenceEndingAbbreviations_Entries() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    Dim d255 As Object
    Set d255 = CreateObject("Scripting.Dictionary")
    d255.Add "abbreviation", "TM."
    d255.Add "meaning", "Thay m" & ChrW(&H1EB7) & "t"
    d255.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 7"
    c.Add d255
    Dim d256 As Object
    Set d256 = CreateObject("Scripting.Dictionary")
    d256.Add "abbreviation", "KT."
    d256.Add "meaning", "K" & ChrW(&HFD) & " thay"
    d256.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 7"
    c.Add d256
    Dim d257 As Object
    Set d257 = CreateObject("Scripting.Dictionary")
    d257.Add "abbreviation", "TL."
    d257.Add "meaning", "Th" & ChrW(&H1EEB) & "a l" & ChrW(&H1EC7) & "nh"
    d257.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 7"
    c.Add d257
    Dim d258 As Object
    Set d258 = CreateObject("Scripting.Dictionary")
    d258.Add "abbreviation", "TUQ."
    d258.Add "meaning", "Th" & ChrW(&H1EEB) & "a " & ChrW(&H1EE7) & "y quy" & ChrW(&H1EC1) & "n"
    d258.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 7"
    c.Add d258
    Dim d259 As Object
    Set d259 = CreateObject("Scripting.Dictionary")
    d259.Add "abbreviation", "Q."
    d259.Add "meaning", "Quy" & ChrW(&H1EC1) & "n"
    d259.Add "citation", "Ph" & ChrW(&H1EE5) & " l" & ChrW(&H1EE5) & "c I, M" & ChrW(&H1EE5) & "c II, kho" & ChrW(&H1EA3) & "n 7"
    c.Add d259
    Dim d260 As Object
    Set d260 = CreateObject("Scripting.Dictionary")
    d260.Add "abbreviation", "P."
    d260.Add "meaning", "Ph" & ChrW(&HF3)
    d260.Add "citation", "th" & ChrW(&HF4) & "ng l" & ChrW(&H1EC7)
    c.Add d260
    Dim d261 As Object
    Set d261 = CreateObject("Scripting.Dictionary")
    d261.Add "abbreviation", "v.v."
    d261.Add "meaning", "v" & ChrW(&HE2) & "n v" & ChrW(&HE2) & "n"
    d261.Add "citation", "th" & ChrW(&HF4) & "ng l" & ChrW(&H1EC7)
    c.Add d261
    Dim d262 As Object
    Set d262 = CreateObject("Scripting.Dictionary")
    d262.Add "abbreviation", "TP."
    d262.Add "meaning", "Th" & ChrW(&HE0) & "nh ph" & ChrW(&H1ED1)
    d262.Add "citation", "th" & ChrW(&HF4) & "ng l" & ChrW(&H1EC7)
    c.Add d262
    Dim d263 As Object
    Set d263 = CreateObject("Scripting.Dictionary")
    d263.Add "abbreviation", ChrW(&H110) & "T."
    d263.Add "meaning", ChrW(&H110) & "i" & ChrW(&H1EC7) & "n tho" & ChrW(&H1EA1) & "i"
    d263.Add "citation", "th" & ChrW(&HF4) & "ng l" & ChrW(&H1EC7)
    c.Add d263
    Dim d264 As Object
    Set d264 = CreateObject("Scripting.Dictionary")
    d264.Add "abbreviation", "TS."
    d264.Add "meaning", "Ti" & ChrW(&H1EBF) & "n s" & ChrW(&H129)
    d264.Add "citation", "th" & ChrW(&HF4) & "ng l" & ChrW(&H1EC7)
    c.Add d264
    Dim d265 As Object
    Set d265 = CreateObject("Scripting.Dictionary")
    d265.Add "abbreviation", "GS."
    d265.Add "meaning", "Gi" & ChrW(&HE1) & "o s" & ChrW(&H1B0)
    d265.Add "citation", "th" & ChrW(&HF4) & "ng l" & ChrW(&H1EC7)
    c.Add d265
    Dim d266 As Object
    Set d266 = CreateObject("Scripting.Dictionary")
    d266.Add "abbreviation", "PGS."
    d266.Add "meaning", "Ph" & ChrW(&HF3) & " Gi" & ChrW(&HE1) & "o s" & ChrW(&H1B0)
    d266.Add "citation", "th" & ChrW(&HF4) & "ng l" & ChrW(&H1EC7)
    c.Add d266
    Dim d267 As Object
    Set d267 = CreateObject("Scripting.Dictionary")
    d267.Add "abbreviation", "ThS."
    d267.Add "meaning", "Th" & ChrW(&H1EA1) & "c s" & ChrW(&H129)
    d267.Add "citation", "th" & ChrW(&HF4) & "ng l" & ChrW(&H1EC7)
    c.Add d267
    Set LoadRawNonSentenceEndingAbbreviations_Entries = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawNonSentenceEndingAbbreviations_Entries", Err.description
End Function

Public Function LoadRawNonSentenceEndingAbbreviations() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "1.0.0"
    d.Add "status", "hat giong - cac muc thuong gap nhat, chua day du"
    d.Add "sourceLabel", "SUY RA"
    d.Add "entries", LoadRawNonSentenceEndingAbbreviations_Entries()
    Set LoadRawNonSentenceEndingAbbreviations = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawNonSentenceEndingAbbreviations", Err.description
End Function

Private Function LoadRawRegionNames_Regions() As Collection
    On Error GoTo ErrHandler
    Dim c As New Collection
    c.Add "B" & ChrW(&H1EAF) & "c B" & ChrW(&H1ED9)
    c.Add "Trung B" & ChrW(&H1ED9)
    c.Add "Nam B" & ChrW(&H1ED9)
    c.Add "B" & ChrW(&H1EAF) & "c Trung B" & ChrW(&H1ED9)
    c.Add "Nam Trung B" & ChrW(&H1ED9)
    c.Add "Duy" & ChrW(&HEA) & "n h" & ChrW(&H1EA3) & "i Nam Trung B" & ChrW(&H1ED9)
    c.Add "Duy" & ChrW(&HEA) & "n h" & ChrW(&H1EA3) & "i mi" & ChrW(&H1EC1) & "n Trung"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng B" & ChrW(&H1EAF) & "c"
    c.Add "T" & ChrW(&HE2) & "y B" & ChrW(&H1EAF) & "c"
    c.Add "T" & ChrW(&HE2) & "y Nguy" & ChrW(&HEA) & "n"
    c.Add ChrW(&H110) & ChrW(&HF4) & "ng Nam B" & ChrW(&H1ED9)
    c.Add "T" & ChrW(&HE2) & "y Nam B" & ChrW(&H1ED9)
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng b" & ChrW(&H1EB1) & "ng s" & ChrW(&HF4) & "ng H" & ChrW(&H1ED3) & "ng"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng b" & ChrW(&H1EB1) & "ng s" & ChrW(&HF4) & "ng C" & ChrW(&H1EED) & "u Long"
    c.Add ChrW(&H110) & ChrW(&H1ED3) & "ng b" & ChrW(&H1EB1) & "ng B" & ChrW(&H1EAF) & "c B" & ChrW(&H1ED9)
    c.Add "Mi" & ChrW(&H1EC1) & "n B" & ChrW(&H1EAF) & "c"
    c.Add "Mi" & ChrW(&H1EC1) & "n Trung"
    c.Add "Mi" & ChrW(&H1EC1) & "n Nam"
    Set LoadRawRegionNames_Regions = c
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegionNames_Regions", Err.description
End Function

Public Function LoadRawRegionNames() As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d.Add "schemaVersion", "1.0.0"
    d.Add "status", "day du theo cach phan vung dia ly pho thong"
    d.Add "sourceLabel", "ND30"
    d.Add "actionType", "TU SUA"
    d.Add "regions", LoadRawRegionNames_Regions()
    Set LoadRawRegionNames = d
    Exit Function
ErrHandler:
    Err.Raise Err.number, "RuleData.LoadRawRegionNames", Err.description
End Function
