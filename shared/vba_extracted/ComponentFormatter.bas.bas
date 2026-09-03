Attribute VB_Name = "ComponentFormatter"
'==============================================================
' ComponentFormatter -- Dinh dang 17 thanh phan the thuc DA CO SAN trong van ban ( doi
' QUYET DINH Q-02 (11 thang 8 nam 2026): module nay CHI dinh dang thanh phan DA TON TAI trong tai
' lieu, KHONG BAO GIO chen noi dung moi. Vai tro khong ton tai (khong doan nao khop trong
' LayoutMap) -> khong lam gi.
' Nhan dien sai thi KHONG SUA (CLAUDE.md muc 5): ComponentDetector.DetectComponents.LayoutMap
' CHI gom doan co do tin cay "high"/"medium" -- do tin cay "low" khong bao gio vao LayoutMap, nen
' bo loc theo LayoutMap o duoi TU DONG thoa "Rang buoc" cua.
' KIEN TRUC BOC Utils.BeginOperation/EndOperation -- HAI kieu ham khac nhau trong file nay:
' - Ham TRUC TIEP GHI DE van ban (sau ham dung tam sau "Sua*"/goi ben trong)
'   ApplyComponentStyle/FormatComponentByRole/FixNationalTitleMottoSpacing: chi DIEU PHOI, goi lai
'   cac primitive DA TU BOC rieng cua TextFormatter.bas/ParagraphFormatter.bas -- BAN THAN KHONG
'   mo BeginOperation, tranh long nhau (hai bien module-level cua Utils.bas khong tai nhap duoc)
'   -- dung nguyen tac da ghi o dau TextFormatter.bas.
' - Sau ham con lai (FixNationalMottoSeparator, FixCodeNumberNotation, FixPlaceAndIssuedDate,
'   FixLegalBasisPunctuation, FixRecipientSalutationColon) TU GHI DE Range.Text truc tiep (khong
'   qua primitive nao), nen TU boc MOT lan Utils.BeginOperation/EndOperation cho CA VONG LAP nhieu
'   doan -- dung dung idiom da co san cua ToneNormalizer.ApplyToneStyle: mot lan bam nut = mot muc
'   trong nhat ky thao tac, du sua nhieu doan.
' CMP-08, CMP-09 (cap bo cuc Phan/Chuong/Muc/Dieu/Khoan/Diem) -- KHONG thuoc pham vi, xem
' StructureFormatter.bas. Muc "Viec phai lam" #2 cua goi y dung Shapes.AddTextbox cho hai o nay
' khi lam duoc -- CHUA lam vi thieu buoc nhan dien truoc do, KHONG phai vi thieu API
' (Shapes.AddTextbox co san trong Word Object Model).
' Moi thu tuc cong khai bat dau bang On Error GoTo ErrHandler. Dung AscW/ChrW, khong dung Asc/Chr
' (phu thuoc code page).
'==============================================================
Option Explicit

' ============================================================================
' Tien ich chung -- nhan dien LayoutMap, chi so doan theo vai tro (co sap xep), Word.Range cua MOT
' chi so doan.
' ============================================================================

Public Function DetectLayoutMap(Optional ByVal regime As String = "ND30") As Object
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ComponentFormatter.DetectLayoutMap")
    On Error GoTo ErrHandler
    Dim t0 As Double: t0 = Timer
    ' CHI can noi dung doan van + DetectDocumentType/DetectComponents (ca hai cung CHI doc
    ' snapshot("Paragraphs")) - dung ban chup NHE, tranh cham anh loi khong can thiet (T-71, xem
    ' ghi chu dau DocumentSnapshot.CaptureParagraphsOnlySnapshot).
    Dim snapshot As Object: Set snapshot = DocumentSnapshot.CaptureParagraphsOnlySnapshot()
    DebugTrace.Log "ComponentFormatter.DetectLayoutMap", "CaptureParagraphsOnlySnapshot xong, " & Format$(Timer - t0, "0.00") & "s"
    Dim typeResult As Object: Set typeResult = DocumentTypeDetector.DetectDocumentType(snapshot, regime)
    DebugTrace.Log "ComponentFormatter.DetectLayoutMap", "DetectDocumentType xong, " & Format$(Timer - t0, "0.00") & "s"
    Dim detection As Object
    Set detection = ComponentDetector.DetectComponents(snapshot, CStr(typeResult("Type")), regime)
    DebugTrace.Log "ComponentFormatter.DetectLayoutMap", "DetectComponents xong, " & Format$(Timer - t0, "0.00") & "s"
    Set DetectLayoutMap = detection("LayoutMap")
    Exit Function
ErrHandler:
    DebugTrace.LogErr "ComponentFormatter.DetectLayoutMap", "loi giua chung", Err.number, Err.description
    Err.Raise Err.number, "ComponentFormatter.DetectLayoutMap", Err.description
End Function
