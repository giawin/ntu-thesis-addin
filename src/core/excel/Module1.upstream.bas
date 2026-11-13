Attribute VB_Name = "Module1"
Private mHT_Co As Boolean
Private mHT_Ws As Object
Private mHT_R1 As Long, mHT_R2 As Long, mHT_C1 As Long, mHT_C2 As Long
Private mHT_ColW() As Double
Private mHT_RowH() As Double
Private mHT_FontAll As Variant, mHT_VAll As Variant, mHT_ShrinkAll As Variant, mHT_WrapAll As Variant
Private mHT_Font() As String
Private mHT_VAl() As Long
Private mHT_Shrink() As Boolean
Private mHT_Wrap() As Boolean
Private mHT_PerCell As Boolean
Private mHT_HBreaks As String
Private mHT_VBreaks As String
Private mHT_PS(1 To 20) As Variant
Private Const LIC_CAN_NHAP_MA As Boolean = False
Public Sub RibbonChuanHoaExcel(control As Object)
Call ChuanHoa_TrangIn_Excel(False)
End Sub
Public Sub ChuanHoa_TrangIn_Excel_Silent()
Call ChuanHoa_TrangIn_Excel(True)
End Sub
Private Sub HT_InAn(ByVal bat As Boolean)
On Error Resume Next
Application.PrintCommunication = bat
On Error GoTo 0
End Sub
Private Sub HT_ChupTrangThai(ByVal ws As Worksheet, ByVal r1 As Long, ByVal r2 As Long, ByVal c1 As Long, ByVal c2 As Long)
mHT_Co = False
On Error Resume Next
Set mHT_Ws = ws
mHT_R1 = r1: mHT_R2 = r2: mHT_C1 = c1: mHT_C2 = c2
Dim r As Long, c As Long
ReDim mHT_ColW(c1 To c2)
For c = c1 To c2
mHT_ColW(c) = ws.Columns(c).ColumnWidth
Next c
ReDim mHT_RowH(r1 To r2)
For r = r1 To r2
mHT_RowH(r) = ws.Rows(r).RowHeight
Next r
Dim rg As Range
Set rg = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
mHT_FontAll = rg.Font.Name
mHT_VAll = rg.VerticalAlignment
mHT_ShrinkAll = rg.ShrinkToFit
mHT_WrapAll = rg.WrapText
mHT_PerCell = (IsNull(mHT_FontAll) Or IsNull(mHT_VAll) Or IsNull(mHT_ShrinkAll) Or IsNull(mHT_WrapAll)) And _
(CDbl(r2 - r1 + 1) * CDbl(c2 - c1 + 1) <= 200000#)
If mHT_PerCell Then
ReDim mHT_Font(r1 To r2, c1 To c2)
ReDim mHT_VAl(r1 To r2, c1 To c2)
ReDim mHT_Shrink(r1 To r2, c1 To c2)
ReDim mHT_Wrap(r1 To r2, c1 To c2)
For r = r1 To r2
For c = c1 To c2
With ws.Cells(r, c)
mHT_Font(r, c) = .Font.Name
mHT_VAl(r, c) = .VerticalAlignment
mHT_Shrink(r, c) = .ShrinkToFit
mHT_Wrap(r, c) = .WrapText
End With
Next c
Next r
End If
mHT_HBreaks = "": mHT_VBreaks = ""
Dim hb As HPageBreak, vb As VPageBreak
For Each hb In ws.HPageBreaks
If hb.Type = xlPageBreakManual Then mHT_HBreaks = mHT_HBreaks & hb.Location.Row & ","
Next hb
For Each vb In ws.VPageBreaks
If vb.Type = xlPageBreakManual Then mHT_VBreaks = mHT_VBreaks & vb.Location.Column & ","
Next vb
With ws.PageSetup
mHT_PS(1) = .PaperSize
mHT_PS(2) = .Orientation
mHT_PS(3) = .LeftMargin
mHT_PS(4) = .RightMargin
mHT_PS(5) = .TopMargin
mHT_PS(6) = .BottomMargin
mHT_PS(7) = .HeaderMargin
mHT_PS(8) = .FooterMargin
mHT_PS(9) = .CenterHorizontally
mHT_PS(10) = .CenterVertically
mHT_PS(11) = .PrintGridlines
mHT_PS(12) = .PrintHeadings
mHT_PS(13) = .PrintArea
mHT_PS(14) = .PrintTitleRows
mHT_PS(15) = .Zoom
mHT_PS(16) = .FitToPagesWide
mHT_PS(17) = .FitToPagesTall
mHT_PS(18) = .CenterFooter
End With
On Error GoTo 0
mHT_Co = True
End Sub
Private Function HT_KhoiPhuc() As Boolean
HT_KhoiPhuc = False
If Not mHT_Co Then Exit Function
Dim ws As Worksheet
Set ws = Nothing
On Error Resume Next
Set ws = mHT_Ws
Dim chk As String
chk = ws.Name
If Err.Number <> 0 Or ws Is Nothing Then
mHT_Co = False
Exit Function
End If
Application.ScreenUpdating = False
Dim r As Long, c As Long
Dim rg As Range
Set rg = ws.Range(ws.Cells(mHT_R1, mHT_C1), ws.Cells(mHT_R2, mHT_C2))
If Not IsNull(mHT_FontAll) Then rg.Font.Name = mHT_FontAll
If Not IsNull(mHT_VAll) Then rg.VerticalAlignment = mHT_VAll
If Not IsNull(mHT_ShrinkAll) Then rg.ShrinkToFit = mHT_ShrinkAll
If Not IsNull(mHT_WrapAll) Then rg.WrapText = mHT_WrapAll
If mHT_PerCell Then
For r = mHT_R1 To mHT_R2
For c = mHT_C1 To mHT_C2
With ws.Cells(r, c)
If IsNull(mHT_FontAll) Then .Font.Name = mHT_Font(r, c)
If IsNull(mHT_VAll) Then .VerticalAlignment = mHT_VAl(r, c)
If IsNull(mHT_ShrinkAll) Then .ShrinkToFit = mHT_Shrink(r, c)
If IsNull(mHT_WrapAll) Then .WrapText = mHT_Wrap(r, c)
End With
Next c
Next r
End If
ws.ResetAllPageBreaks
Dim parts() As String, k As Long
parts = Split(mHT_HBreaks, ",")
For k = LBound(parts) To UBound(parts)
If Len(parts(k)) > 0 Then ws.HPageBreaks.Add Before:=ws.Rows(CLng(parts(k)))
Next k
parts = Split(mHT_VBreaks, ",")
For k = LBound(parts) To UBound(parts)
If Len(parts(k)) > 0 Then ws.VPageBreaks.Add Before:=ws.Columns(CLng(parts(k)))
Next k
Call HT_InAn(False)
With ws.PageSetup
.PaperSize = mHT_PS(1)
.Orientation = mHT_PS(2)
.LeftMargin = mHT_PS(3)
.RightMargin = mHT_PS(4)
.TopMargin = mHT_PS(5)
.BottomMargin = mHT_PS(6)
.HeaderMargin = mHT_PS(7)
.FooterMargin = mHT_PS(8)
.CenterHorizontally = mHT_PS(9)
.CenterVertically = mHT_PS(10)
.PrintGridlines = mHT_PS(11)
.PrintHeadings = mHT_PS(12)
.PrintArea = mHT_PS(13)
.PrintTitleRows = mHT_PS(14)
If VarType(mHT_PS(15)) = vbBoolean Then
.Zoom = False
.FitToPagesWide = mHT_PS(16)
.FitToPagesTall = mHT_PS(17)
Else
.Zoom = mHT_PS(15)
End If
.CenterFooter = mHT_PS(18)
End With
Call HT_InAn(True)
For c = mHT_C1 To mHT_C2
ws.Columns(c).ColumnWidth = mHT_ColW(c)
Next c
For r = mHT_R1 To mHT_R2
ws.Rows(r).RowHeight = mHT_RowH(r)
Next r
Application.ScreenUpdating = True
On Error GoTo 0
mHT_Co = False
HT_KhoiPhuc = True
End Function
Public Sub HoanTac_ChuanHoa_Excel()
If Not HT_KhoiPhuc() Then
MsgBox "Khong co lan chuan hoa nao de hoan tac (hoac Excel da dong / mo lai sau lan chuan hoa).", vbInformation, "Hoan tac"
End If
End Sub
Public Sub RibbonHoanTacExcel(control As Object)
Call HoanTac_ChuanHoa_Excel
End Sub
Sub ChuanHoa_TrangIn_Excel(Optional ByVal silentMode As Boolean = False)
If Not silentMode Then
Call CN_TuKiemTraHangTuan(True)
If Not CheckLicenseActivation() Then Exit Sub
End If
Dim ws As Worksheet
Set ws = Nothing
On Error Resume Next
Set ws = ActiveSheet
On Error GoTo 0
If ws Is Nothing Then Exit Sub
If ws.ProtectContents Then
If Not silentMode Then MsgBox "Sheet dang bi khoa (Protect Sheet). Hay bo khoa sheet truoc khi chuan hoa trang in.", vbExclamation, "Chuan Hoa Excel"
Exit Sub
End If
On Error GoTo LoiChuanHoa
Application.ScreenUpdating = False
Application.DisplayAlerts = False
Dim firstRow As Long, lastRow As Long, firstCol As Long, lastCol As Long
If Not GetTrueDataBounds(ws, firstRow, lastRow, firstCol, lastCol) Then
Application.ScreenUpdating = True
Application.DisplayAlerts = True
If Not silentMode Then
MsgBox "Bang tinh hien tai khong co du lieu!", vbExclamation, "Chuan Hoa Excel"
End If
Exit Sub
End If
Dim rng As Range
Set rng = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol))
Call HT_ChupTrangThai(ws, firstRow, lastRow, firstCol, lastCol)
Dim origOrientation As Long
origOrientation = xlPortrait
On Error Resume Next
origOrientation = ws.PageSetup.Orientation
On Error GoTo LoiChuanHoa
ws.ResetAllPageBreaks
rng.Font.Name = "Times New Roman"
rng.VerticalAlignment = xlCenter
rng.ShrinkToFit = False
Dim rIdx As Long, cIdx As Long
Dim hasHeader As Boolean
Dim hStart As Long, hEnd As Long
hasHeader = FindTableHeaderRows(ws, firstRow, lastRow, firstCol, lastCol, hStart, hEnd)
If hasHeader Then
Call ClampLastColToTableHeader(ws, firstRow, lastRow, firstCol, lastCol, hStart, hEnd)
Set rng = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol))
End If
Dim tableEndRow As Long
tableEndRow = FindTableDataEndRow(ws, firstRow, lastRow, firstCol, lastCol, hEnd, hasHeader)
Call ApplySmartWrapText(ws, firstRow, lastRow, firstCol, lastCol, hStart, tableEndRow, hasHeader)
Call FitColumnsSmart(ws, firstRow, lastRow, firstCol, lastCol, hStart, tableEndRow, hasHeader)
Dim totalWidthPt As Double
Dim visibleCols As Long
totalWidthPt = GetVisibleColumnsWidthPt(ws, firstCol, lastCol, visibleCols)
Dim logicalCols As Long
logicalCols = visibleCols
If hasHeader Then
Dim lCnt As Long: lCnt = 0
For cIdx = firstCol To lastCol
If Not ws.Columns(cIdx).Hidden Then
If ws.Cells(hStart, cIdx).Address = ws.Cells(hStart, cIdx).MergeArea.Cells(1, 1).Address Then
lCnt = lCnt + 1
End If
End If
Next cIdx
If lCnt >= 2 Then logicalCols = lCnt
End If
Dim chosenOrientation As Long
Dim targetPageWidthPt As Double
Dim targetPageHeightPt As Double
If origOrientation = xlLandscape Then
chosenOrientation = xlLandscape
targetPageWidthPt = 762#
targetPageHeightPt = 500#
ElseIf visibleCols <= 6 And logicalCols <= 6 And totalWidthPt <= 545 Then
chosenOrientation = xlPortrait
targetPageWidthPt = 518#
targetPageHeightPt = 745#
Else
chosenOrientation = xlLandscape
targetPageWidthPt = 762#
targetPageHeightPt = 500#
End If
If totalWidthPt > 40 And totalWidthPt < targetPageWidthPt * 0.78 Then
Dim expandRatio As Double
expandRatio = (targetPageWidthPt * 0.92) / totalWidthPt
If expandRatio > 1.35 Then expandRatio = 1.35
For cIdx = firstCol To lastCol
If Not ws.Columns(cIdx).Hidden Then
Dim newColW As Double
newColW = ws.Columns(cIdx).ColumnWidth * expandRatio
If newColW <= 75 Then ws.Columns(cIdx).ColumnWidth = newColW
End If
Next cIdx
totalWidthPt = GetVisibleColumnsWidthPt(ws, firstCol, lastCol, visibleCols)
End If
Dim printScale As Double
printScale = 1#
If totalWidthPt > targetPageWidthPt And totalWidthPt > 0 Then
printScale = targetPageWidthPt / totalWidthPt
If printScale < 0.45 Then printScale = 0.45
End If
Call AutoFitRowsAccurate(ws, firstRow, lastRow, firstCol, lastCol, hStart, hEnd, tableEndRow, hasHeader, printScale)
Dim totalHeightPt As Double
totalHeightPt = GetVisibleRowsHeightPt(ws, firstRow, lastRow)
Dim tableHeightPt As Double
tableHeightPt = GetVisibleRowsHeightPt(ws, firstRow, tableEndRow)
Dim scaledHeightPt As Double
scaledHeightPt = totalHeightPt * printScale
Dim scaledTableHeightPt As Double
scaledTableHeightPt = tableHeightPt * printScale
Dim targetTallPages As Long
targetTallPages = 0
If scaledHeightPt <= targetPageHeightPt * 1.50 Then
targetTallPages = 1
ElseIf tableEndRow < lastRow And scaledTableHeightPt <= targetPageHeightPt * 1.35 And scaledHeightPt <= targetPageHeightPt * 1.75 Then
targetTallPages = 1
Else
Dim fullPages As Long
fullPages = Int(scaledHeightPt / targetPageHeightPt)
Dim remPages As Double
remPages = (scaledHeightPt / targetPageHeightPt) - fullPages
If fullPages >= 1 And fullPages <= 12 Then
If remPages > 0 And remPages <= 0.45 Then
targetTallPages = fullPages
ElseIf tableEndRow < lastRow And Int(scaledTableHeightPt / targetPageHeightPt) <= fullPages And remPages <= 0.65 Then
targetTallPages = fullPages
End If
End If
End If
Dim zoomW As Double, zoomH As Double
zoomW = 100#
If totalWidthPt > targetPageWidthPt And totalWidthPt > 0 Then
zoomW = (targetPageWidthPt / totalWidthPt) * 98#
End If
zoomH = 100#
If targetTallPages > 0 And totalHeightPt > 0 Then
Dim availTallPt As Double
availTallPt = targetPageHeightPt * targetTallPages
If totalHeightPt > availTallPt Then
zoomH = (availTallPt / totalHeightPt) * 97#
End If
End If
Dim calcZoom As Long
If zoomH < zoomW Then
calcZoom = Int(zoomH)
Else
calcZoom = Int(zoomW)
End If
If calcZoom > 100 Then calcZoom = 100
If calcZoom < 45 Then calcZoom = 45
Call HT_InAn(False)
On Error Resume Next
With ws.PageSetup
.PaperSize = xlPaperA4
.Orientation = chosenOrientation
.LeftMargin = Application.CentimetersToPoints(1#)
.RightMargin = Application.CentimetersToPoints(1#)
.TopMargin = Application.CentimetersToPoints(1.2)
.BottomMargin = Application.CentimetersToPoints(1.2)
.HeaderMargin = Application.CentimetersToPoints(0.5)
.FooterMargin = Application.CentimetersToPoints(0.5)
.CenterHorizontally = True
.CenterVertically = False
.PrintGridlines = False
.PrintHeadings = False
.PrintArea = rng.Address
If hasHeader And targetTallPages <> 1 And hStart <= 10 Then
.PrintTitleRows = "$" & hStart & ":$" & hEnd
Else
.PrintTitleRows = ""
End If
.FitToPagesWide = False
.FitToPagesTall = False
.Zoom = calcZoom
End With
On Error GoTo LoiChuanHoa
Call HT_InAn(True)
Dim stepGuard As Long
stepGuard = 0
Do While stepGuard < 5 And calcZoom > 45
Dim needShrink As Boolean
needShrink = False
If ws.VPageBreaks.Count > 0 Then needShrink = True
If targetTallPages > 0 Then
If ws.HPageBreaks.Count >= targetTallPages Then needShrink = True
End If
If Not needShrink Then Exit Do
calcZoom = calcZoom - 3
If calcZoom < 45 Then calcZoom = 45
ws.PageSetup.Zoom = calcZoom
stepGuard = stepGuard + 1
Loop
Dim soTrang As Long
soTrang = 1
On Error Resume Next
soTrang = (ws.HPageBreaks.Count + 1) * (ws.VPageBreaks.Count + 1)
Dim uocTinh As Long
uocTinh = Int((GetVisibleRowsHeightPt(ws, firstRow, lastRow) * calcZoom / 100) / targetPageHeightPt) + 1
If uocTinh > soTrang Then soTrang = uocTinh
If soTrang > 1 And Len(CStr(mHT_PS(18))) = 0 Then
ws.PageSetup.CenterFooter = "&""Times New Roman""&11Trang &P/&N"
End If
On Error GoTo LoiChuanHoa
Dim minFont As Double
minFont = 0
Dim fc As Range
Dim scanned As Long
On Error Resume Next
For Each fc In rng.Cells
If Len(CStr(fc.Value)) > 0 Then
Dim fsz As Double
fsz = Val(CStr(fc.Font.Size))
If fsz > 0 And (minFont = 0 Or fsz < minFont) Then minFont = fsz
End If
scanned = scanned + 1
If scanned > 20000 Then Exit For
Next fc
On Error GoTo LoiChuanHoa
Dim chuNho As String
chuNho = ""
If minFont > 0 Then
If minFont * calcZoom / 100 < 8.5 Then
chuNho = vbCrLf & "- CANH BAO: chu in ra chi con khoang " & Format(minFont * calcZoom / 100, "0.0") & " pt (thu nho " & calcZoom & "%), co the kho doc." & _
vbCrLf & "  Nen chuyen kho ngang, bo bot cot hoac tach bang."
End If
End If
On Error Resume Next
Application.OnUndo "Hoan tac Chuan hoa trang in", "HoanTac_ChuanHoa_Excel"
On Error GoTo LoiChuanHoa
Application.ScreenUpdating = True
Application.DisplayAlerts = True
If Not silentMode Then
Dim orientName As String
If chosenOrientation = xlPortrait Then
orientName = "A4 Doc (Portrait)"
Else
orientName = "A4 Ngang (Landscape)"
End If
MsgBox "Da chuan hoa bang bieu & trang in hoan hao!" & vbCrLf & _
"- Kho giay tu dong: " & orientName & vbCrLf & _
"- Da can chinh Fit ngang doc & chong che khuat chu 100%." & chuNho & vbCrLf & vbCrLf & _
"Bam Ctrl+Z (hoac nut Hoan tac) de tra lai nhu truoc khi chuan hoa." & vbCrLf & vbCrLf & _
"Created by VPDU xa Thanh Phu", vbInformation, "Chuan Hoa Trang In Excel"
End If
On Error GoTo 0
Exit Sub
LoiChuanHoa:
Dim errMsg As String
errMsg = Err.Number & " - " & Err.Description
On Error Resume Next
Call HT_InAn(True)
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Dim daTra As Boolean
daTra = HT_KhoiPhuc()
On Error GoTo 0
If Not silentMode Then
If daTra Then
MsgBox "Co loi khi chuan hoa: " & errMsg & vbCrLf & "Da tra bang tinh ve nhu truoc khi chuan hoa.", vbExclamation, "Chuan Hoa Excel"
Else
MsgBox "Co loi khi chuan hoa: " & errMsg, vbExclamation, "Chuan Hoa Excel"
End If
End If
End Sub
Private Function GetTrueDataBounds(ByVal ws As Worksheet, ByRef firstRow As Long, ByRef lastRow As Long, ByRef firstCol As Long, ByRef lastCol As Long) As Boolean
On Error Resume Next
GetTrueDataBounds = False
Dim lastRCell As Range, lastCCell As Range
Dim firstRCell As Range, firstCCell As Range
Set lastRCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
If lastRCell Is Nothing Then Exit Function
Set lastCCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
Set firstRCell = ws.Cells.Find(What:="*", After:=ws.Cells(lastRCell.Row, lastCCell.Column), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
Set firstCCell = ws.Cells.Find(What:="*", After:=ws.Cells(lastRCell.Row, lastCCell.Column), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
firstRow = 1
firstCol = 1
If Not firstRCell Is Nothing Then firstRow = firstRCell.Row
If Not firstCCell Is Nothing Then firstCol = firstCCell.Column
If firstRow > 3 Then firstRow = 1
If firstCol > 2 Then firstCol = 1
lastRow = lastRCell.Row
lastCol = lastCCell.Column
Dim r As Long, c As Long
For r = firstRow To lastRow
Dim chkCell As Range
Set chkCell = ws.Cells(r, lastCol)
If chkCell.MergeCells Then
Dim mEndCol As Long
mEndCol = chkCell.MergeArea.Column + chkCell.MergeArea.Columns.Count - 1
If mEndCol > lastCol And mEndCol <= lastCol + 15 Then
lastCol = mEndCol
End If
End If
Next r
For c = firstCol To lastCol
Dim chkBotCell As Range
Set chkBotCell = ws.Cells(lastRow, c)
If chkBotCell.MergeCells Then
Dim mEndRow As Long
mEndRow = chkBotCell.MergeArea.Row + chkBotCell.MergeArea.Rows.Count - 1
If mEndRow > lastRow And mEndRow <= lastRow + 15 Then
lastRow = mEndRow
End If
End If
Next c
If lastRow > firstRow + 6 Then
Dim botNonEmpty As Long: botNonEmpty = 0
Dim botIsNum As Boolean: botIsNum = False
For c = firstCol To lastCol
If Not IsEmpty(ws.Cells(lastRow, c).Value) Then
botNonEmpty = botNonEmpty + 1
If IsNumeric(ws.Cells(lastRow, c).Value) And ws.Cells(lastRow, c).Borders(xlEdgeTop).LineStyle = xlNone Then
botIsNum = True
End If
End If
Next c
If botNonEmpty = 1 And botIsNum Then
Dim prevHasData As Boolean: prevHasData = False
For r = lastRow - 2 To lastRow - 1
For c = firstCol To lastCol
If Not IsEmpty(ws.Cells(r, c).Value) Then prevHasData = True
Next c
Next r
If Not prevHasData Then
For r = lastRow - 3 To firstRow Step -1
For c = firstCol To lastCol
If Not IsEmpty(ws.Cells(r, c).Value) Then
lastRow = r
GoTo DoneTrimBot
End If
Next c
Next r
End If
End If
End If
DoneTrimBot:
GetTrueDataBounds = (lastRow >= firstRow And lastCol >= firstCol)
On Error GoTo 0
End Function
Private Sub ApplySmartWrapText(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstCol As Long, ByVal lastCol As Long, ByVal hStart As Long, ByVal tableEndRow As Long, ByVal hasHeader As Boolean)
On Error Resume Next
Dim fullRng As Range
Set fullRng = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol))
fullRng.WrapText = True
If lastCol <= firstCol Then Exit Sub
Dim r As Long, c As Long
For r = firstRow To lastRow
Dim isOutsideTable As Boolean
isOutsideTable = False
If hasHeader And r < hStart Then isOutsideTable = True
If r > tableEndRow Then isOutsideTable = True
If isOutsideTable And Not ws.Rows(r).Hidden Then
For c = firstCol To lastCol - 1
Dim cell As Range
Set cell = ws.Cells(r, c)
If Not cell.MergeCells Then
Dim txt As String
txt = Trim(CStr(cell.Value))
If Len(txt) > 0 And Len(txt) > ws.Columns(c).ColumnWidth Then
Dim nextCell As Range
Set nextCell = ws.Cells(r, c + 1)
If Not nextCell.MergeCells And Len(Trim(CStr(nextCell.Value))) = 0 Then
cell.WrapText = False
End If
End If
End If
Next c
End If
Next r
On Error GoTo 0
End Sub
Private Sub FitColumnsSmart(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstCol As Long, ByVal lastCol As Long, ByVal hStart As Long, ByVal tableEndRow As Long, ByVal hasHeader As Boolean)
On Error Resume Next
Dim c As Long, r As Long
Dim scanStart As Long
scanStart = IIf(hasHeader, hStart, firstRow)
For c = firstCol To lastCol
Dim colObj As Range
Set colObj = ws.Columns(c)
If Not colObj.Hidden Then
Dim curW As Double
curW = colObj.ColumnWidth
Dim maxHashNumLen As Long: maxHashNumLen = 0
Dim hasAnyText As Boolean: hasAnyText = False
For r = scanStart To tableEndRow
If Not ws.Rows(r).Hidden Then
Dim cell As Range
Set cell = ws.Cells(r, c)
If cell.MergeArea.Columns.Count = 1 Then
If Not IsEmpty(cell.Value) Then
hasAnyText = True
Dim dispTxt As String
dispTxt = Trim(CStr(cell.Text))
If Left(dispTxt, 1) = "#" Then
Dim cleanNumStr As String
If IsNumeric(cell.Value) Then
cleanNumStr = CStr(Round(CDbl(cell.Value), 2))
Else
cleanNumStr = Trim(CStr(cell.Value))
End If
If Len(cleanNumStr) > maxHashNumLen Then maxHashNumLen = Len(cleanNumStr)
End If
End If
End If
End If
Next r
If maxHashNumLen > 0 Then
Dim reqNumW As Double
reqNumW = (maxHashNumLen * 1.15) + 2.2
If reqNumW > 26 Then reqNumW = 26
If curW < reqNumW Then colObj.ColumnWidth = reqNumW
ElseIf hasAnyText And curW < 4.2 Then
colObj.ColumnWidth = 4.5
End If
End If
Next c
On Error GoTo 0
End Sub
Private Sub AutoFitRowsAccurate(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstCol As Long, ByVal lastCol As Long, ByVal hStart As Long, ByVal hEnd As Long, ByVal tableEndRow As Long, ByVal hasHeader As Boolean, ByVal printScale As Double)
On Error Resume Next
Dim fullRng As Range
Set fullRng = ws.Range(ws.Cells(firstRow, firstCol), ws.Cells(lastRow, lastCol))
fullRng.Rows.AutoFit
Dim scaleSafety As Double
scaleSafety = 1.04
If printScale < 0.9 Then
scaleSafety = 1.04 + ((0.9 - printScale) * 0.1)
If scaleSafety > 1.08 Then scaleSafety = 1.08
End If
Dim r As Long, c As Long
For r = firstRow To lastRow
Dim rowObj As Range
Set rowObj = ws.Rows(r)
If Not rowObj.Hidden Then
Dim rowHasText As Boolean
rowHasText = False
Dim maxRowFontSize As Double
maxRowFontSize = 0#
For c = firstCol To lastCol
Dim scanCell As Range
Set scanCell = ws.Cells(r, c)
If Not IsEmpty(scanCell.Value) Then
If Len(Trim(CStr(scanCell.Value))) > 0 Then
rowHasText = True
Dim cfs As Double
cfs = Val(CStr(scanCell.Font.Size))
If cfs >= 6 And cfs <= 36 And cfs > maxRowFontSize Then
maxRowFontSize = cfs
End If
End If
End If
Next c
If Not rowHasText Then
If r > tableEndRow And rowObj.RowHeight > 15.5 Then
rowObj.RowHeight = 14.5
ElseIf rowObj.RowHeight > 20# Then
rowObj.RowHeight = 16#
End If
Else
If maxRowFontSize < 6 Then maxRowFontSize = 11#
Dim minRowH As Double
minRowH = (maxRowFontSize * 1.35) + 4#
If hasHeader And r >= hStart And r <= hEnd Then
If minRowH < (maxRowFontSize * 1.45) + 5# Then
minRowH = (maxRowFontSize * 1.45) + 5#
End If
End If
Dim maxReqHeight As Double
maxReqHeight = minRowH
Dim hasVertMergeStart As Boolean
hasVertMergeStart = False
For c = firstCol To lastCol
Dim chkV As Range
Set chkV = ws.Cells(r, c)
If chkV.MergeCells Then
If chkV.MergeArea.Row = r And chkV.MergeArea.Rows.Count > 1 Then
If Len(Trim(CStr(chkV.MergeArea.Cells(1, 1).Value))) > 12 Then
hasVertMergeStart = True
Exit For
End If
End If
End If
Next c
If Not hasVertMergeStart Then
If rowObj.RowHeight > maxReqHeight And rowObj.RowHeight <= 400 Then
maxReqHeight = rowObj.RowHeight + 3#
End If
End If
For c = firstCol To lastCol
Dim cell As Range
Set cell = ws.Cells(r, c)
If cell.Address = cell.MergeArea.Cells(1, 1).Address Then
If cell.MergeArea.Rows.Count = 1 Then
Dim cellReqH As Double
cellReqH = CalcCellRequiredHeightPt(cell, scaleSafety)
If cellReqH > maxReqHeight Then
maxReqHeight = cellReqH
End If
End If
End If
Next c
If maxReqHeight > 409# Then maxReqHeight = 409#
rowObj.RowHeight = maxReqHeight
End If
End If
Next r
Dim targetSpan As Long
For targetSpan = 2 To 12
For r = firstRow To lastRow
If Not ws.Rows(r).Hidden Then
For c = firstCol To lastCol
Dim mCell As Range
Set mCell = ws.Cells(r, c)
If mCell.MergeCells Then
Dim mRowSpan As Long
mRowSpan = mCell.MergeArea.Rows.Count
If mCell.Address = mCell.MergeArea.Cells(1, 1).Address And _
((targetSpan < 12 And mRowSpan = targetSpan) Or (targetSpan = 12 And mRowSpan >= 12)) Then
Dim totalReqH As Double
totalReqH = CalcCellRequiredHeightPt(mCell, scaleSafety)
If totalReqH > 0 Then
Dim mStartR As Long, mEndR As Long, mr As Long
mStartR = mCell.MergeArea.Row
mEndR = mStartR + mRowSpan - 1
Dim curTotalH As Double: curTotalH = 0
Dim visRowsInMerge As Long: visRowsInMerge = 0
For mr = mStartR To mEndR
If Not ws.Rows(mr).Hidden Then
curTotalH = curTotalH + ws.Rows(mr).RowHeight
visRowsInMerge = visRowsInMerge + 1
End If
Next mr
If visRowsInMerge > 0 And curTotalH < totalReqH Then
Dim addPerRow As Double
addPerRow = (totalReqH - curTotalH) / visRowsInMerge
For mr = mStartR To mEndR
If Not ws.Rows(mr).Hidden Then
Dim newH As Double
newH = ws.Rows(mr).RowHeight + addPerRow
If newH > 409# Then newH = 409#
ws.Rows(mr).RowHeight = newH
End If
Next mr
End If
End If
End If
End If
Next c
End If
Next r
Next targetSpan
On Error GoTo 0
End Sub
Private Function CalcCellRequiredHeightPt(ByVal cell As Range, ByVal scaleSafety As Double) As Double
On Error Resume Next
CalcCellRequiredHeightPt = 0
If IsEmpty(cell.Value) Then Exit Function
Dim valStr As String
If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
Dim dispNum As String
dispNum = Trim(CStr(cell.Text))
If Len(dispNum) > 0 And Left(dispNum, 1) <> "#" Then
valStr = dispNum
Else
valStr = CStr(Round(CDbl(cell.Value), 2))
End If
Else
valStr = Trim(CStr(cell.Value))
End If
If Len(valStr) = 0 Then Exit Function
valStr = Replace(valStr, ChrW(8230), "...")
Dim usableWidthPt As Double
usableWidthPt = GetVisibleMergeAreaWidthPt(cell.MergeArea) - 10#
If usableWidthPt < 24# Then usableWidthPt = 24#
Dim fSize As Double
fSize = Val(CStr(cell.Font.Size))
If fSize < 6 Or fSize > 36 Then fSize = 12#
Dim isBold As Boolean
isBold = False
If IsNull(cell.Font.Bold) Then
isBold = True
ElseIf CBool(cell.Font.Bold) = True Then
isBold = True
End If
Dim charWidthPt As Double
charWidthPt = fSize * 0.47
If isBold Then charWidthPt = charWidthPt * 1.1
If valStr = UCase(valStr) And Len(valStr) > 6 Then charWidthPt = charWidthPt * 1.12
charWidthPt = charWidthPt * scaleSafety
Dim maxCharsPerLine As Long
maxCharsPerLine = Int(usableWidthPt / charWidthPt)
If maxCharsPerLine < 4 Then maxCharsPerLine = 4
Dim normalizedTxt As String
normalizedTxt = Replace(valStr, vbCrLf, vbLf)
normalizedTxt = Replace(normalizedTxt, vbCr, vbLf)
Dim hardLines() As String
hardLines = Split(normalizedTxt, vbLf)
Dim totalLines As Long
totalLines = 0
If Not cell.MergeCells And CBool(cell.WrapText) = False Then
totalLines = UBound(hardLines) - LBound(hardLines) + 1
Else
Dim k As Long
For k = LBound(hardLines) To UBound(hardLines)
Dim pLine As String
pLine = Trim(hardLines(k))
If Len(pLine) = 0 Then
totalLines = totalLines + 1
Else
totalLines = totalLines + CountWordWrappedLines(pLine, maxCharsPerLine)
End If
Next k
End If
If totalLines < 1 Then totalLines = 1
Dim lineHeightPt As Double
lineHeightPt = fSize * 1.35
CalcCellRequiredHeightPt = (totalLines * lineHeightPt) + 5#
On Error GoTo 0
End Function
Private Function CountWordWrappedLines(ByVal paragraphText As String, ByVal maxCharsPerLine As Long) As Long
Dim words() As String
words = Split(paragraphText, " ")
Dim linesCount As Long: linesCount = 1
Dim curLineLen As Long: curLineLen = 0
Dim i As Long
For i = LBound(words) To UBound(words)
Dim w As String
w = words(i)
Dim wLen As Long
wLen = Len(w)
If wLen > 0 Then
If curLineLen = 0 Then
If wLen <= maxCharsPerLine Then
curLineLen = wLen
Else
Dim extraLines As Long
extraLines = Int((wLen - 1) / maxCharsPerLine)
linesCount = linesCount + extraLines
curLineLen = wLen - (extraLines * maxCharsPerLine)
End If
ElseIf curLineLen + 1 + wLen <= maxCharsPerLine Then
curLineLen = curLineLen + 1 + wLen
Else
linesCount = linesCount + 1
If wLen <= maxCharsPerLine Then
curLineLen = wLen
Else
Dim extraWLines As Long
extraWLines = Int((wLen - 1) / maxCharsPerLine)
linesCount = linesCount + extraWLines
curLineLen = wLen - (extraWLines * maxCharsPerLine)
End If
End If
End If
Next i
CountWordWrappedLines = linesCount
End Function
Private Function GetVisibleMergeAreaWidthPt(ByVal mergeRng As Range) As Double
On Error Resume Next
Dim totalPt As Double: totalPt = 0
Dim mCol As Range
For Each mCol In mergeRng.Columns
If Not mCol.Hidden Then
totalPt = totalPt + mCol.Width
End If
Next mCol
GetVisibleMergeAreaWidthPt = totalPt
On Error GoTo 0
End Function
Private Function GetVisibleColumnsWidthPt(ByVal ws As Worksheet, ByVal firstCol As Long, ByVal lastCol As Long, ByRef visibleColsCount As Long) As Double
On Error Resume Next
Dim totalPt As Double: totalPt = 0
visibleColsCount = 0
Dim c As Long
For c = firstCol To lastCol
If Not ws.Columns(c).Hidden Then
totalPt = totalPt + ws.Columns(c).Width
visibleColsCount = visibleColsCount + 1
End If
Next c
GetVisibleColumnsWidthPt = totalPt
On Error GoTo 0
End Function
Private Function GetVisibleRowsHeightPt(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long) As Double
On Error Resume Next
Dim totalPt As Double: totalPt = 0
Dim r As Long
For r = firstRow To lastRow
If Not ws.Rows(r).Hidden Then
totalPt = totalPt + ws.Rows(r).RowHeight
End If
Next r
GetVisibleRowsHeightPt = totalPt
On Error GoTo 0
End Function
Private Function FindTableDataEndRow(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstCol As Long, ByVal lastCol As Long, ByVal hEnd As Long, ByVal hasHeader As Boolean) As Long
FindTableDataEndRow = lastRow
Dim minCheckRow As Long
minCheckRow = IIf(hasHeader, hEnd + 1, firstRow + 1)
If lastRow <= minCheckRow Then Exit Function
Dim startScanFooter As Long
startScanFooter = lastRow - 18
If startScanFooter < minCheckRow Then startScanFooter = minCheckRow
Dim r As Long, c As Long
For r = startScanFooter To lastRow
Dim rowTextAsc As String: rowTextAsc = ""
Dim nonEmptyCells As Long: nonEmptyCells = 0
For c = firstCol To lastCol
If ws.Cells(r, c).Address = ws.Cells(r, c).MergeArea.Cells(1, 1).Address Then
Dim cellVal As String
cellVal = Trim(CStr(ws.Cells(r, c).Value))
If Len(cellVal) > 0 Then
nonEmptyCells = nonEmptyCells + 1
rowTextAsc = rowTextAsc & " " & ToAscii(cellVal)
End If
End If
Next c
If nonEmptyCells > 0 And nonEmptyCells <= 4 Then
If InStr(rowTextAsc, "T/M ") > 0 Or InStr(rowTextAsc, "TM.") > 0 Or _
InStr(rowTextAsc, "K/T ") > 0 Or InStr(rowTextAsc, "KT.") > 0 Or _
InStr(rowTextAsc, "T/L ") > 0 Or InStr(rowTextAsc, "TL.") > 0 Or _
InStr(rowTextAsc, "NGUOI LAP") > 0 Or InStr(rowTextAsc, "LAP BIEU") > 0 Or _
InStr(rowTextAsc, "BI THU") > 0 Or InStr(rowTextAsc, "PHO BI THU") > 0 Or _
InStr(rowTextAsc, "CHU TICH") > 0 Or InStr(rowTextAsc, "TRUONG BAN") > 0 Or _
InStr(rowTextAsc, "CHANH VAN PHONG") > 0 Or InStr(rowTextAsc, "VAN PHONG") > 0 Or _
InStr(rowTextAsc, "THU TRUONG") > 0 Or InStr(rowTextAsc, "GIAM DOC") > 0 Or _
InStr(rowTextAsc, "XAC NHAN") > 0 Or InStr(rowTextAsc, "DUYET") > 0 Or _
InStr(rowTextAsc, "NOI NHAN") > 0 Or _
(InStr(rowTextAsc, "NGAY") > 0 And InStr(rowTextAsc, "THANG") > 0 And InStr(rowTextAsc, "NAM") > 0) Then
If r - 1 >= minCheckRow Then
FindTableDataEndRow = r - 1
Exit Function
End If
End If
End If
Next r
End Function
Private Function FindTableHeaderRows(ByVal ws As Worksheet, ByVal startRow As Long, ByVal endRow As Long, ByVal startCol As Long, ByVal endCol As Long, ByRef hStart As Long, ByRef hEnd As Long) As Boolean
Dim r As Long, c As Long
Dim maxScanRow As Long
maxScanRow = startRow + 25
If maxScanRow > endRow Then maxScanRow = endRow
Dim bestRow As Long: bestRow = 0
Dim maxScore As Long: maxScore = 0
For r = startRow To maxScanRow
Dim colCountWithText As Long: colCountWithText = 0
Dim coveredCols As Long: coveredCols = 0
Dim rowAscText As String: rowAscText = ""
For c = startCol To endCol
If ws.Cells(r, c).Address = ws.Cells(r, c).MergeArea.Cells(1, 1).Address Then
Dim cellVal As String
cellVal = Trim(CStr(ws.Cells(r, c).Value))
If Len(cellVal) > 0 Then
colCountWithText = colCountWithText + 1
coveredCols = coveredCols + ws.Cells(r, c).MergeArea.Columns.Count
rowAscText = rowAscText & " " & ToAscii(cellVal)
End If
End If
Next c
Dim isHeaderCandidate As Boolean: isHeaderCandidate = False
If colCountWithText >= 3 Then
isHeaderCandidate = True
ElseIf colCountWithText >= 2 And coveredCols >= 3 Then
isHeaderCandidate = True
End If
If isHeaderCandidate Then
Dim isTieuNgu As Boolean: isTieuNgu = False
If InStr(rowAscText, "DANG CONG SAN") > 0 Or _
InStr(rowAscText, "CONG HOA XA HOI") > 0 Or _
InStr(rowAscText, "DOC LAP") > 0 Then
isTieuNgu = True
ElseIf InStr(rowAscText, "DANG BO") > 0 Or InStr(rowAscText, "DANG UY") > 0 Or InStr(rowAscText, "UY BAN") > 0 Then
isTieuNgu = True
ElseIf InStr(rowAscText, "NGAY") > 0 And InStr(rowAscText, "THANG") > 0 And InStr(rowAscText, "NAM") > 0 Then
isTieuNgu = True
End If
If Not isTieuNgu Then
Dim isDocTitle As Boolean: isDocTitle = False
If colCountWithText <= 2 Then
If InStr(rowAscText, "DANH MUC") > 0 Or _
InStr(rowAscText, "PHU LUC") > 0 Or _
InStr(rowAscText, "KE HOACH") > 0 Or _
InStr(rowAscText, "BAO CAO") > 0 Or _
InStr(rowAscText, "DANH SACH") > 0 Or _
InStr(rowAscText, "BANG TONG HOP") > 0 Or _
InStr(rowAscText, "BANG PHAN CONG") > 0 Or _
InStr(rowAscText, "QUYET DINH") > 0 Or _
InStr(rowAscText, "THONG BAO") > 0 Then
isDocTitle = True
End If
End If
If Not isDocTitle Then
Dim currentScore As Long
currentScore = colCountWithText * 10
Dim kws As Variant
kws = Array("STT", "TT", "TRUC", "NHOM", "NOI DUNG", "NHIEM VU", "SAN PHAM", _
"DON VI", "CHI TIEU", "THOI GIAN", "PHAN CONG", "HO VA TEN", "HO TEN", _
"CHUC VU", "SO LUONG", "GHI CHU", "TIEN DO", "KET QUA", _
"MO TA", "CHI TIET", "DIEN GIAI", "DON GIA", "THANH TIEN", "NGUOI THUC HIEN", _
"DU TOAN", "DIEU CHINH", "TANG", "GIAM", "KINH PHI")
Dim kwIdx As Long, kwHits As Long
kwHits = 0
For kwIdx = LBound(kws) To UBound(kws)
If InStr(rowAscText, CStr(kws(kwIdx))) > 0 Then
kwHits = kwHits + 1
End If
Next kwIdx
currentScore = currentScore + (kwHits * 20)
If (kwHits >= 1 Or colCountWithText >= 4) And currentScore >= 50 Then
maxScore = currentScore
bestRow = r
Exit For
End If
End If
End If
End If
Next r
If bestRow = 0 Or maxScore < 40 Then
FindTableHeaderRows = False
Exit Function
End If
hStart = bestRow
hEnd = bestRow
Dim maxMergeRow As Long
maxMergeRow = bestRow
For c = startCol To endCol
If ws.Cells(bestRow, c).MergeCells Then
Dim mRows As Long
mRows = ws.Cells(bestRow, c).MergeArea.Row + ws.Cells(bestRow, c).MergeArea.Rows.Count - 1
If mRows > maxMergeRow Then maxMergeRow = mRows
End If
Next c
If maxMergeRow > hEnd Then hEnd = maxMergeRow
If hEnd < endRow Then
Dim nextRow As Long
nextRow = hEnd + 1
Dim numColsWithNumber As Long: numColsWithNumber = 0
For c = startCol To endCol
Dim nVal As String
nVal = Trim(CStr(ws.Cells(nextRow, c).Value))
If Len(nVal) > 0 Then
If nVal Like "(*)" Or IsNumeric(nVal) Or nVal = "a" Or nVal = "b" Or InStr(nVal, "=") > 0 Then
numColsWithNumber = numColsWithNumber + 1
End If
End If
Next c
If numColsWithNumber >= 3 Then
hEnd = nextRow
End If
End If
FindTableHeaderRows = True
End Function
Private Sub ClampLastColToTableHeader(ByVal ws As Worksheet, ByVal firstRow As Long, ByVal lastRow As Long, ByVal firstCol As Long, ByRef lastCol As Long, ByVal hStart As Long, ByVal hEnd As Long)
On Error Resume Next
If lastCol <= firstCol + 1 Then Exit Sub
Dim trueTableLastCol As Long
trueTableLastCol = 0
Dim c As Long
For c = firstCol To lastCol
Dim topHdrCell As Range
Set topHdrCell = ws.Cells(hStart, c).MergeArea.Cells(1, 1)
Dim hdrTxt As String
hdrTxt = Trim(CStr(topHdrCell.Value))
If Len(hdrTxt) = 0 Then
Exit For
End If
Dim hdrAsc As String
hdrAsc = ToAscii(hdrTxt)
If InStr(hdrAsc, "KEM THEO") > 0 Then
Exit For
End If
trueTableLastCol = c
Next c
If trueTableLastCol >= firstCol + 1 And trueTableLastCol < lastCol Then
lastCol = trueTableLastCol
End If
On Error GoTo 0
End Sub
Private Function LicApiUrl() As String
LicApiUrl = "https://script.google.com/macros/s/AKfycbxSZyf-7FeVZ6xbhBwuT5aVVGJDdqP8AJCYek_Wp3Ila5yWLGnmq_rIua3kL1dFsvzOug/exec"
End Function
Private Function CheckLicenseActivation() As Boolean
Dim appName As String, sectionName As String
Dim savedKey As String, currentMachine As String
Dim lastCheckDate As String
Dim apiUrl As String
appName = "OfficePartySuite"
sectionName = "License"
apiUrl = LicApiUrl()
On Error Resume Next
savedKey = GetSetting(appName, sectionName, "ActiveKey", "")
lastCheckDate = GetSetting(appName, sectionName, "LastCheckDate", "")
On Error GoTo 0
If Len(savedKey) > 0 Then
If Len(lastCheckDate) > 0 Then
Dim isRecent As Boolean: isRecent = False
On Error Resume Next
If IsDate(lastCheckDate) Then
If Abs(DateDiff("d", CDate(lastCheckDate), Now)) < 7 Then
isRecent = True
End If
End If
On Error GoTo 0
If isRecent Then
CheckLicenseActivation = True
Exit Function
End If
End If
currentMachine = GetMachineName()
Dim checkUrl As String
checkUrl = apiUrl & "?key=" & EncodeUrl(savedKey) & "&machine=" & EncodeUrl(currentMachine)
Dim checkResp As String
checkResp = CallHttpApi(checkUrl)
If Len(checkResp) > 0 Then
Dim checkStatus As String
checkStatus = ParseJsonField(checkResp, "status")
If checkStatus = "ERROR" Then
SaveSetting appName, sectionName, "ActiveKey", ""
SaveSetting appName, sectionName, "ActiveMachine", ""
SaveSetting appName, sectionName, "LastCheckDate", ""
Dim errReason As String
errReason = ParseJsonField(checkResp, "message")
If Len(errReason) = 0 Then errReason = "Ban quyen cua ban da bi khoa boi VPDU xa Thanh Phu!"
MsgBox "CANH BAO BAN QUYEN:" & vbCrLf & errReason & vbCrLf & vbCrLf & _
"Tien ich da bi dung hoat dong tren may nay.", vbCritical, "Ban Quyen Bi Khoa"
CheckLicenseActivation = False
Exit Function
Else
On Error Resume Next
SaveSetting appName, sectionName, "LastCheckDate", Format(Now, "yyyy-mm-dd")
On Error GoTo 0
End If
End If
CheckLicenseActivation = True
Exit Function
End If
currentMachine = GetMachineName()
If Not LIC_CAN_NHAP_MA Then
If GetSetting(appName, sectionName, "AutoTryDate", "") = Format(Now, "yyyy-mm-dd") Then
CheckLicenseActivation = True
Exit Function
End If
Dim autoResp As String
autoResp = CallHttpApi(apiUrl & "?machine=" & EncodeUrl(currentMachine))
If ParseJsonField(autoResp, "khoa") = "1" Then
MsgBox "CANH BAO BAN QUYEN:" & vbCrLf & ParseJsonField(autoResp, "message"), vbCritical, "Ban Quyen Add-in"
CheckLicenseActivation = False
Exit Function
End If
If ParseJsonField(autoResp, "status") = "SUCCESS" And Len(ParseJsonField(autoResp, "key")) > 0 Then
SaveSetting appName, sectionName, "ActiveKey", ParseJsonField(autoResp, "key")
SaveSetting appName, sectionName, "ActiveMachine", currentMachine
SaveSetting appName, sectionName, "LastCheckDate", Format(Now, "yyyy-mm-dd")
Else
SaveSetting appName, sectionName, "AutoTryDate", Format(Now, "yyyy-mm-dd")
End If
CheckLicenseActivation = True
Exit Function
End If
Dim userInput As String
userInput = InputBox("Vui long nhap Ma kich hoat ban quyen de su dung tien ich:" & vbCrLf & _
"(Moi may tinh chi can nhap kich hoat 1 lan duy nhat)" & vbCrLf & vbCrLf & _
"Ten may tinh: " & currentMachine & vbCrLf & _
"Don vi phat trien: Van phong Dang uy xa Thanh Phu", _
"Kich Hoat Ban Quyen Add-in")
userInput = Trim(UCase(userInput))
If Len(userInput) = 0 Then
MsgBox "Ban chua nhap ma kich hoat! Khong the thuc hien chuan hoa.", vbExclamation, "Ban Quyen Add-in"
CheckLicenseActivation = False
Exit Function
End If
Dim actUrl As String
actUrl = apiUrl & "?key=" & EncodeUrl(userInput) & "&machine=" & EncodeUrl(currentMachine)
Dim actResp As String
actResp = CallHttpApi(actUrl)
If Len(actResp) = 0 Then
MsgBox "Khong the ket noi den may chu xac thuc ban quyen!" & vbCrLf & _
"Vui long kiem tra ket noi Internet va thu lai.", vbExclamation, "Loi Ket Noi"
CheckLicenseActivation = False
Exit Function
End If
Dim actStatus As String, actMsg As String
actStatus = ParseJsonField(actResp, "status")
actMsg = ParseJsonField(actResp, "message")
If actStatus = "SUCCESS" Then
SaveSetting appName, sectionName, "ActiveKey", userInput
SaveSetting appName, sectionName, "ActiveMachine", currentMachine
SaveSetting appName, sectionName, "LastCheckDate", Format(Now, "yyyy-mm-dd")
MsgBox "Kich hoat ban quyen thanh cong!" & vbCrLf & _
"May tinh: " & currentMachine & vbCrLf & _
"Ban co the su dung tien ich vinh vien tren may nay.", vbInformation, "Kich Hoat Thanh Cong"
CheckLicenseActivation = True
Else
MsgBox "Kich hoat that bai!" & vbCrLf & actMsg & vbCrLf & vbCrLf & _
"Vui long lien he Van phong Dang uy xa Thanh Phu de duoc cap ma hop le.", vbCritical, "Kich Hoat That Bai"
CheckLicenseActivation = False
End If
End Function
Private Function CallHttpApi(ByVal url As String) As String
Dim res As String
res = ""
On Error Resume Next
Dim http As Object
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
If http Is Nothing Then
Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
End If
http.Open "GET", url, False
http.setRequestHeader "User-Agent", "Mozilla/5.0"
http.setTimeouts 5000, 8000, 8000, 8000
http.Send
If Err.Number = 0 Then
res = http.responseText
End If
On Error GoTo 0
CallHttpApi = Trim(res)
End Function
Private Function GetMachineName() As String
Dim mName As String
mName = Environ("COMPUTERNAME")
If Len(mName) = 0 Then mName = Environ("USERNAME")
If Len(mName) = 0 Then mName = "UNKNOWN_PC"
GetMachineName = mName
End Function
Private Function ParseJsonField(ByVal json As String, ByVal fieldName As String) As String
Dim pos As Long, startPos As Long, endPos As Long
pos = InStr(json, """" & fieldName & """")
If pos > 0 Then
startPos = InStr(pos + Len(fieldName) + 2, json, """")
If startPos > 0 Then
endPos = InStr(startPos + 1, json, """")
If endPos > startPos Then
ParseJsonField = Mid(json, startPos + 1, endPos - startPos - 1)
End If
End If
End If
End Function
Private Function EncodeUrl(ByVal s As String) As String
Dim i As Long, ch As String, code As Long
Dim res As String
res = ""
For i = 1 To Len(s)
ch = Mid(s, i, 1)
code = AscW(ch) And &HFFFF&
If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Or ch = "-" Or ch = "_" Or ch = "." Then
res = res & ch
ElseIf code < 128 Then
res = res & "%" & Right("0" & Hex(code), 2)
Else
res = res & "_"
End If
Next i
EncodeUrl = res
End Function
Private Function ToAscii(ByVal s As String) As String
Dim n As Long
n = Len(s)
If n = 0 Then
ToAscii = ""
Exit Function
End If
Dim res As String
res = Space$(n)
Dim k As Long, w As Long
For k = 1 To n
w = AscW(Mid$(s, k, 1))
If w < 0 Then w = w + 65536
Select Case w
Case 272, 273: Mid$(res, k, 1) = "D"
Case 192 To 195, 224 To 227, 258, 259, 7840 To 7863: Mid$(res, k, 1) = "A"
Case 200 To 202, 232 To 234, 278, 279, 7864 To 7879: Mid$(res, k, 1) = "E"
Case 204, 205, 236, 237, 296, 297, 7880 To 7883: Mid$(res, k, 1) = "I"
Case 210 To 214, 242 To 246, 416, 417, 7884 To 7907: Mid$(res, k, 1) = "O"
Case 217, 218, 249, 250, 360, 361, 431, 432, 7908 To 7921: Mid$(res, k, 1) = "U"
Case 221, 253, 7922 To 7929: Mid$(res, k, 1) = "Y"
Case Else
If w >= 97 And w <= 122 Then
Mid$(res, k, 1) = Chr$(w - 32)
ElseIf (w >= 65 And w <= 90) Or (w >= 48 And w <= 57) Or w = 32 Or w = 42 Or w = 45 Or w = 47 Then
Mid$(res, k, 1) = Chr$(w)
Else
Mid$(res, k, 1) = " "
End If
End Select
Next k
ToAscii = res
End Function
Public Sub Auto_Open()
Call CN_HenKiemTraKhiMo
End Sub
Public Sub RibbonOnLoad(ribbon As Object)
Call CN_HenKiemTraKhiMo
End Sub
Public Function CN_PhienBan() As String
CN_PhienBan = "2026.11.12.2"
End Function
Private Function CN_ManifestUrl() As String
CN_ManifestUrl = "https://github.com/CuongBC195/add-in-office/releases/latest/download/version.json"
End Function
Private Function CN_SoSanh(ByVal a As String, ByVal b As String) As Integer
Dim pa() As String, pb() As String
pa = Split(a, ".")
pb = Split(b, ".")
Dim i As Integer, va As Long, vb As Long
For i = 0 To 5
va = 0: vb = 0
If i <= UBound(pa) Then va = Val(pa(i))
If i <= UBound(pb) Then vb = Val(pb(i))
If va > vb Then
CN_SoSanh = 1
Exit Function
ElseIf va < vb Then
CN_SoSanh = -1
Exit Function
End If
Next i
CN_SoSanh = 0
End Function
Private Function CN_FileTuChay() As String
On Error Resume Next
CN_FileTuChay = ThisWorkbook.FullName
On Error GoTo 0
End Function
Private Function CN_DungFileAddIn() As Boolean
Dim n As String
n = LCase(Mid(CN_FileTuChay(), InStrRev(Replace(CN_FileTuChay(), "\", "/"), "/") + 1))
CN_DungFileAddIn = (n = "chuanhoa_excel.xlam")
End Function
Private Function CN_ShQ(ByVal s As String) As String
CN_ShQ = "'" & Replace(s, "'", "'\''") & "'"
End Function
Private Function CN_TaiManifest(Optional ByVal nhanh As Boolean = False) As String
Dim i As Integer, res As String, soLan As Integer, cho As Integer
soLan = IIf(nhanh, 1, 2)
cho = IIf(nhanh, 3, 15)
For i = 1 To soLan
res = ""
On Error Resume Next
Dim http As Object
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
http.Open "GET", CN_ManifestUrl(), False
http.setRequestHeader "User-Agent", "Mozilla/5.0"
http.setTimeouts 5000, 5000, cho * 1000, cho * 1000
http.Send
If Err.Number = 0 Then
If http.Status = 200 Then res = http.responseText
End If
On Error GoTo 0
res = Trim(res)
If Len(ParseJsonField(res, "version")) > 0 Then Exit For
Next i
CN_TaiManifest = res
End Function
Public Sub RibbonCapNhat(control As Object)
Call CN_KiemTraCapNhat(True)
End Sub
Public Sub CN_HenKiemTraKhiMo()
Static daHen As Boolean
If daHen Then Exit Sub
daHen = True
On Error Resume Next
Application.OnTime Now + TimeSerial(0, 0, 20), "CN_KiemTraNhanhKhiMo"
End Sub
Public Sub CN_KiemTraNhanhKhiMo()
Call CN_TuKiemTraHangTuan(True)
End Sub
Public Sub CN_TuKiemTraHangTuan(Optional ByVal nhanh As Boolean = False)
On Error Resume Next
Dim lastChk As String, daDocDuoc As Boolean
lastChk = GetSetting("OfficePartySuite", "CapNhat", "LanKiemTra_EXCEL", "")
If Len(lastChk) > 0 Then
If IsDate(lastChk) Then
If DateDiff("d", CDate(lastChk), Date) < 1 Then Exit Sub
End If
End If
Dim lanThu As String
lanThu = GetSetting("OfficePartySuite", "CapNhat", "LanThu_EXCEL", "")
If Len(lanThu) > 0 Then
If IsDate(lanThu) Then
If DateDiff("n", CDate(lanThu), Now) < 60 Then Exit Sub
End If
End If
SaveSetting "OfficePartySuite", "CapNhat", "LanThu_EXCEL", Format(Now, "yyyy-mm-dd hh:nn:ss")
On Error GoTo 0
Call CN_KiemTraCapNhat(False, nhanh, daDocDuoc)
If daDocDuoc Then SaveSetting "OfficePartySuite", "CapNhat", "LanKiemTra_EXCEL", Format(Date, "yyyy-mm-dd")
End Sub
Public Sub CN_KiemTraCapNhat(ByVal hoiKhiKhongCo As Boolean, Optional ByVal nhanh As Boolean = False, Optional ByRef daDocDuoc As Boolean = False)
Dim js As String
js = CN_TaiManifest(nhanh)
Dim ver As String, url As String, sha As String, notes As String, kt As Long
ver = ParseJsonField(js, "version")
url = ParseJsonField(js, "excel_url")
sha = LCase(ParseJsonField(js, "excel_sha256"))
notes = ParseJsonField(js, "notes_ascii")
kt = Val(ParseJsonField(js, "excel_size"))
If Len(ParseJsonField(js, "excel_url_win")) > 0 Then
url = ParseJsonField(js, "excel_url_win")
sha = LCase(ParseJsonField(js, "excel_sha256_win"))
kt = Val(ParseJsonField(js, "excel_size_win"))
End If
If Len(ver) = 0 Or Len(url) = 0 Or Len(sha) <> 64 Then
If hoiKhiKhongCo Then MsgBox "Khong ket noi duoc may chu cap nhat (GitHub). Kiem tra mang roi thu lai." & vbCrLf & "Phien ban dang dung: " & CN_PhienBan(), vbExclamation, "Cap nhat Add-in"
Exit Sub
End If
daDocDuoc = True
If CN_SoSanh(ver, CN_PhienBan()) <= 0 Then
If hoiKhiKhongCo Then MsgBox "Ban dang dung phien ban moi nhat: " & CN_PhienBan(), vbInformation, "Cap nhat Add-in"
Exit Sub
End If
If Not CN_DungFileAddIn() Then
MsgBox "Co phien ban moi " & ver & " nhung macro dang chay tu '" & CN_FileTuChay() & "', khong phai file add-in da cai." & vbCrLf & _
"Hay cai lai bang bo cai dat.", vbExclamation, "Cap nhat Add-in"
Exit Sub
End If
Dim ans As VbMsgBoxResult
ans = MsgBox("Co phien ban moi: " & ver & " (dang dung " & CN_PhienBan() & ")." & vbCrLf & vbCrLf & notes & vbCrLf & vbCrLf & "Cap nhat ngay?", vbYesNo + vbQuestion, "Cap nhat Add-in")
If ans <> vbYes Then Exit Sub
Dim msg As String
msg = CN_TaiVaCai(url, sha, kt)
If Len(msg) > 0 Then MsgBox msg, vbInformation, "Cap nhat Add-in"
End Sub
Private Function CN_TaiVaCai(ByVal url As String, ByVal sha As String, Optional ByVal kichThuoc As Long = 0) As String
Dim target As String
target = CN_FileTuChay()
On Error GoTo WinFail
Dim http As Object
Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
http.Open "GET", url, False
http.setRequestHeader "User-Agent", "Mozilla/5.0"
http.setTimeouts 10000, 15000, 60000, 180000
http.Send
If http.Status <> 200 Then GoTo WinFail
Dim b() As Byte, n As Long
b = http.responseBody
n = UBound(b) - LBound(b) + 1
If n < 10240 Then GoTo WinBad
If b(LBound(b)) <> 80 Or b(LBound(b) + 1) <> 75 Then GoTo WinBad
If kichThuoc > 0 And n <> kichThuoc Then GoTo WinBad
Dim thuMuc As String
thuMuc = Environ("APPDATA") & "\OfficePartySuite"
If Len(Dir(thuMuc, vbDirectory)) = 0 Then MkDir thuMuc
Dim moiW As String
moiW = thuMuc & "\capnhat_EXCEL.moi"
If Len(Dir(moiW)) > 0 Then Kill moiW
Dim f As Integer
f = FreeFile
Open moiW For Binary Access Write As #f
Put #f, 1, b
Close #f
If FileLen(moiW) <> n Then GoTo WinBad
Dim tgtW As String
tgtW = target
If Len(Environ("APPDATA")) > 0 And LCase(Left(tgtW, Len(Environ("APPDATA")))) = LCase(Environ("APPDATA")) Then
tgtW = "%APPDATA%" & Mid(tgtW, Len(Environ("APPDATA")) + 1)
End If
Dim cmdPath As String
cmdPath = thuMuc & "\capnhat_EXCEL.cmd"
f = FreeFile
Open cmdPath For Output As #f
Print #f, "@echo off"
Print #f, ":cho"
Print #f, "tasklist /FI ""IMAGENAME eq EXCEL.EXE"" 2>nul | find /I ""EXCEL.EXE"" >nul"
Print #f, "if not errorlevel 1 (timeout /t 2 /nobreak >nul & goto cho)"
Print #f, "copy /Y ""%APPDATA%\OfficePartySuite\capnhat_EXCEL.moi"" """ & tgtW & """ >nul"
Print #f, "del ""%APPDATA%\OfficePartySuite\capnhat_EXCEL.moi"" >nul 2>&1"
Close #f
Shell "cmd.exe /c """ & cmdPath & """", vbMinimizedNoFocus
CN_TaiVaCai = "Da tai ban moi. Hay DONG HET cua so Excel: ban moi se tu cai dat, mo lai Excel la dung duoc."
Exit Function
WinBad:
On Error Resume Next
Close #f
If Len(moiW) > 0 Then Kill moiW
CN_TaiVaCai = "File tai ve khong dung (thieu hoac hong). Chua thay doi gi, hay thu lai sau."
Exit Function
WinFail:
CN_TaiVaCai = "Tai ban cap nhat that bai (kiem tra mang). Chua thay doi gi."
End Function
Public Sub CN_KiemThuCapNhat()
Dim h As String
h = Environ("HOME")
If Len(Dir(h & "/thethuc_debug.on")) = 0 Then Exit Sub
Dim js As String
js = CN_TaiManifest()
Dim ver As String, url As String, sha As String, res As String
ver = ParseJsonField(js, "version")
url = ParseJsonField(js, "excel_url")
sha = LCase(ParseJsonField(js, "excel_sha256"))
res = "file=" & CN_FileTuChay() & " | dungFile=" & CN_DungFileAddIn() & " | dangDung=" & CN_PhienBan() & " | banMoi=" & ver
If Len(ver) > 0 And CN_SoSanh(ver, CN_PhienBan()) > 0 And CN_DungFileAddIn() Then res = res & " | " & CN_TaiVaCai(url, sha)
Dim f As Integer
f = FreeFile
Open h & "/cn_test_EXCEL.log" For Output As #f
Print #f, res
Close #f
End Sub
