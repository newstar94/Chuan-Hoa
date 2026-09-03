Attribute VB_Name = "ProcessingIndicator"
'==============================================================
' "Voi moi
' chuc nang: truoc khi thuc thi, hien hop thoai nho thong bao Dang xu ly o chinh giua man hinh
' (khong dung MsgBox). Hop thoai ton tai xuyen suot trong qua trinh xu ly, sau khi xu ly xong hop
' thoai tu dong dong").
' Choke point DUY NHAT goi frmProcessing.Show/Unload - KHONG noi nao khac trong du an duoc
' Show/Unload form nay truc tiep (xem UserForm_QueryClose cua chinh form). Gan vao
' Utils.BeginOperation/EndOperation/AbortOperation (bao trum HAU HET fix routine cua ca du an,
' cung choke point voi DebugTrace/OperationLogger) - rieng DataReader.RunAndReport (thao tac CHI
' DOC, khong di qua BeginOperation, xem dau DataReader.bas) tu goi ShowProcessing/ HideProcessing
' rieng.
' THEM mot ShowProcessing o dau FindingReporter.RunCheckAndReport, nhung MOI buoc con tu dong
' (LineBreakNormalizer/EdgeWhitespaceTrimmer/FontVariantNormalizer/ MultiSpaceCollapser) van TU
' MO/DONG Utils.BeginOperation/EndOperation cua CHINH no - lam viec CO CHU DICH cua de tranh long
' UndoRecord (Application.UndoRecord KHONG cho long nhau) -- nhung day LAI la NGUYEN NHAN GOC that
' su cua "popup chop tat": HideProcessing truoc day KHONG dem so lan long nhau, nen lan
' EndOperation cua MOT BUOC CON (vi du LineBreakNormalizer) dong luon popup cua BUOC CHA
' (RunCheckAndReport) dang con chay do choc benh, dan den "khoang giua khong thay popup" trong luc
' cac buoc con khac (EdgeWhitespaceTrimmer, MultiSpaceCollapser...) van dang lam viec.
' THEM dem so lan long nhau (mDepth) - Show/Hide nay AN TOAN goi long nhau: Show chi THAT SU hien
' form o lan GOI DAU TIEN (mDepth 0 -> 1), Hide chi THAT SU an form o lan GIAM VE 0 - cac lan
' Show/Hide "long ben trong" chi tang/giam bo dem, KHONG dong form som. Quyet dinh cu "KHONG dem
' so lan long nhau" noi ve viec KHONG MO THEM mot UndoRecord long nhau - KHONG lien quan gi den
' viec popup THI GIAC co duoc phep long hay khong; hai the nay tach biet hoan toan tu dot nay tro
' di.
' ResetDepth la CHOT AN TOAN: goi o MOI diem vao cap cao nhat THAT SU cua san pham (hien tai:
' FindingReporter.RunCheckAndReport, DataReader.RunAndReport) TRUOC lan ShowProcessing dau tien -
' dam bao mDepth KHONG BAO GIO ket lai o gia tri khac 0 dai dang (vi du mot loi VBA hiem gap lam
' mat can bang Show/Hide o lan chay TRUOC) lam popup khong bao gio hien duoc o lan chay TIEP THEO.
'==============================================================
Option Explicit

Private mDepth As Long

' Goi TRUOC lan ShowProcessing dau tien o MOI diem vao cap cao nhat (xem giai thich tren).
Public Sub ResetDepth()
    mDepth = 0
End Sub

Public Sub ShowProcessing()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ProcessingIndicator.ShowProcessing")
    On Error Resume Next
    mDepth = mDepth + 1
    If mDepth > 1 Then
        DebugTrace.Log "ProcessingIndicator.ShowProcessing", "Da hien tu buoc cha (depth=" & CStr(mDepth) & ") - bo qua"
        Exit Sub
    End If

    DebugTrace.Log "ProcessingIndicator.ShowProcessing", "Bat dau"
    frmProcessing.Show vbModeless

    Dim i As Long
    For i = 1 To 5
        DoEvents
    Next i
    frmProcessing.Repaint
    DoEvents

    DebugTrace.Log "ProcessingIndicator.ShowProcessing", "Hoan tat, frmProcessing.Visible=" & CStr(frmProcessing.Visible)
End Sub

Public Sub HideProcessing()
    Dim dbgScope As TraceScope: Set dbgScope = DebugTrace.EnterScope("ProcessingIndicator.HideProcessing")
    On Error Resume Next
    If mDepth > 0 Then mDepth = mDepth - 1
    If mDepth > 0 Then
        DebugTrace.Log "ProcessingIndicator.HideProcessing", "Con buoc cha dang chay (depth=" & CStr(mDepth) & ") - chua dong"
        Exit Sub
    End If

    Unload frmProcessing
    DebugTrace.Log "ProcessingIndicator.HideProcessing", "Da Unload"
End Sub
