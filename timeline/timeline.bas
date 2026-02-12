Option Explicit

Private Const GRID_SEC As Double = 1
Private Const COL_WIDTH As Double = 0.5
Private Const LABEL_STEP As Long = 10
Private Const HEADER_ROW As Long = 6
Private Const DATA_ROW_START As Long = 7
Private Const CSV_SUBFOLDER As String = "csv_data"

Private Function GetCfg(ws As Worksheet, addr As String, defVal As Double) As Double
    Dim v As Variant
    v = ws.Range(addr).Value
    If IsNumeric(v) Then
        GetCfg = CDbl(v)
    Else
        GetCfg = defVal
    End If
End Function

Private Sub ClearGanttShapes(ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).Name = "gantt_picture" Then
            ws.Shapes(i).Delete
        End If
    Next i
End Sub

Sub BuildGantt()
    Dim wsD As Worksheet, wsG As Worksheet
    Dim lastRow As Long
    Dim rowH As Double
    Dim nodeInfo As Object, rowMap As Object
    Dim node As String
    Dim rowIdx As Long
    Dim startT As Double, endT As Double
    Dim state As String, fillColor As Long
    Dim data As Variant
    Dim i As Long
    Dim minStart As Double, maxEnd As Double
    Dim nodes() As String, orders() As Double
    Dim nodeCount As Long
    Dim ordVal As Double
    Dim screenState As Boolean, eventsState As Boolean
    Dim calcState As XlCalculation
    Dim gridStart As Double, gridEnd As Double
    Dim numCols As Long
    Dim lastRowUsed As Long, lastColUsed As Long
    Dim clearLastRow As Long, clearLastCol As Long
    Dim colStart As Long, colEnd As Long
    Dim t As Double
    Dim row As Long
    Dim gridSec As Double
    Dim labelStep As Double
    Dim labelCols As Long, labelEnd As Long
    Dim headerRange As Range, gridRange As Range
    Dim gridColor As Long
    Dim stdH As Double

    Set wsD = ThisWorkbook.Sheets("Data")
    Set wsG = ThisWorkbook.Sheets("Gantt")

    stdH = wsG.StandardHeight
    wsG.Rows("1:3").Hidden = False
    If stdH > 0 Then wsG.Rows("1:3").RowHeight = stdH

    screenState = Application.ScreenUpdating
    eventsState = Application.EnableEvents
    calcState = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    rowH = GetCfg(wsG, "B4", 18)

    Call ClearGanttShapes(wsG)

    lastRow = wsD.Cells(wsD.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then GoTo CleanUp

    data = wsD.Range("A2:I" & lastRow).Value

    Set nodeInfo = CreateObject("Scripting.Dictionary")
    minStart = 1E+99
    maxEnd = -1E+99

    For i = 1 To UBound(data, 1)
        node = Trim$(CStr(data(i, 1)))
        If node <> "" Then
            If Not nodeInfo.Exists(node) Then
                If IsNumeric(data(i, 3)) Then
                    ordVal = CDbl(data(i, 3))
                Else
                    ordVal = 9.9E+99
                End If
                nodeInfo.Add node, ordVal
            End If

            startT = CDbl(Val(data(i, 4)))
            endT = CDbl(Val(data(i, 5)))
            If startT < minStart Then minStart = startT
            If endT > maxEnd Then maxEnd = endT
        End If
    Next i

    If nodeInfo.Count = 0 Then GoTo CleanUp
    If minStart = 1E+99 Then GoTo CleanUp

    gridSec = GRID_SEC
    If gridSec <= 0 Then gridSec = 1

    gridStart = Fix(minStart / gridSec) * gridSec
    gridEnd = Fix((maxEnd - 0.0000001) / gridSec) * gridSec
    numCols = CLng((gridEnd - gridStart) / gridSec) + 1
    If numCols < 1 Then GoTo CleanUp

    nodeCount = nodeInfo.Count
    ReDim nodes(1 To nodeCount)
    ReDim orders(1 To nodeCount)

    i = 0
    Dim k As Variant
    For Each k In nodeInfo.Keys
        i = i + 1
        nodes(i) = CStr(k)
        orders(i) = CDbl(nodeInfo(k))
    Next k

    Call SortNodesByOrder(nodes, orders, 1, nodeCount)

    lastRowUsed = wsG.Cells(wsG.Rows.Count, 1).End(xlUp).Row
    lastColUsed = wsG.Cells(HEADER_ROW, wsG.Columns.Count).End(xlToLeft).Column
    If lastRowUsed < HEADER_ROW Then lastRowUsed = HEADER_ROW
    If lastColUsed < 2 Then lastColUsed = 2
    clearLastRow = Application.WorksheetFunction.Max(lastRowUsed, HEADER_ROW + nodeCount)
    clearLastCol = Application.WorksheetFunction.Max(lastColUsed, 1 + numCols)

    wsG.Range(wsG.Cells(HEADER_ROW, 1), wsG.Cells(clearLastRow, clearLastCol)).ClearContents
    wsG.Range(wsG.Cells(HEADER_ROW, 2), wsG.Cells(clearLastRow, clearLastCol)).Borders.LineStyle = xlNone
    wsG.Range(wsG.Cells(DATA_ROW_START, 2), wsG.Cells(clearLastRow, clearLastCol)).Interior.Pattern = xlNone

    wsG.Cells(HEADER_ROW, 1).Value = "Node"
    wsG.Rows(HEADER_ROW).RowHeight = rowH

    Set headerRange = wsG.Range(wsG.Cells(HEADER_ROW, 2), wsG.Cells(HEADER_ROW, 1 + numCols))
    headerRange.UnMerge
    headerRange.HorizontalAlignment = xlCenter
    headerRange.VerticalAlignment = xlCenter
    headerRange.NumberFormat = "0"

    wsG.Columns(2).Resize(, numCols).ColumnWidth = COL_WIDTH

    labelStep = LABEL_STEP
    If labelStep < gridSec Then labelStep = gridSec
    labelCols = CLng(labelStep / gridSec)
    If labelCols < 1 Then labelCols = 1
    gridColor = RGB(200, 200, 200)

    For t = gridStart To gridEnd Step labelStep
        colStart = 2 + CLng((t - gridStart) / gridSec)
        labelEnd = colStart + labelCols - 1
        If labelEnd > 1 + numCols Then labelEnd = 1 + numCols
        With wsG.Range(wsG.Cells(HEADER_ROW, colStart), wsG.Cells(HEADER_ROW, labelEnd))
            .Merge
            .Value = t - gridStart
        End With

        Set gridRange = wsG.Range(wsG.Cells(HEADER_ROW, colStart), wsG.Cells(HEADER_ROW + nodeCount, colStart))
        With gridRange.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = gridColor
        End With
    Next t

    Set rowMap = CreateObject("Scripting.Dictionary")
    For i = 1 To nodeCount
        rowIdx = i
        rowMap.Add nodes(i), rowIdx
        row = HEADER_ROW + rowIdx
        wsG.Cells(row, 1).Value = nodes(i)
        wsG.Rows(row).RowHeight = rowH
    Next i

    For i = 1 To UBound(data, 1)
        node = Trim$(CStr(data(i, 1)))
        If node = "" Then GoTo NextRow
        If Not rowMap.Exists(node) Then GoTo NextRow

        startT = CDbl(Val(data(i, 4)))
        endT = CDbl(Val(data(i, 5)))
        If endT <= startT Then GoTo NextRow

        colStart = 2 + CLng(Fix((startT - gridStart) / gridSec))
        colEnd = 2 + CLng(Fix((endT - 0.0000001 - gridStart) / gridSec))
        If colEnd < colStart Then GoTo NextRow
        If colStart < 2 Then colStart = 2
        If colEnd > 1 + numCols Then colEnd = 1 + numCols

        state = LCase$(Trim$(CStr(data(i, 7))))
        Select Case state
            Case "process"
                fillColor = RGB(46, 204, 113)
            Case "wait"
                fillColor = RGB(243, 156, 18)
            Case "down"
                fillColor = RGB(52, 152, 219)
            Case "idle"
                fillColor = RGB(241, 196, 15)
            Case Else
                fillColor = RGB(156, 163, 175)
        End Select

        row = HEADER_ROW + rowMap(node)
        wsG.Range(wsG.Cells(row, colStart), wsG.Cells(row, colEnd)).Interior.Color = fillColor

NextRow:
        If (i Mod 1000) = 0 Then DoEvents
    Next i

    On Error Resume Next
    wsG.Activate
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    On Error GoTo 0

CleanUp:
    Application.ScreenUpdating = screenState
    Application.EnableEvents = eventsState
    Application.Calculation = calcState

End Sub

Sub ImportTimelineCsv()
    Dim folderPath As String
    Dim filePath As String

    folderPath = GetDefaultCsvFolder()
    filePath = PickCsvFile(folderPath)
    If filePath = "" Then Exit Sub

    ImportCsvToData filePath
End Sub

Sub ImportTimelineCsvAndBuild()
    Dim folderPath As String
    Dim filePath As String

    folderPath = GetDefaultCsvFolder()
    filePath = PickCsvFile(folderPath)
    If filePath = "" Then Exit Sub

    ImportCsvToData filePath
    BuildGantt
End Sub

Private Function GetDefaultCsvFolder() As String
    Dim p As String
    p = ThisWorkbook.Path & "\" & CSV_SUBFOLDER
    If Dir(p, vbDirectory) <> "" Then
        GetDefaultCsvFolder = p
    Else
        GetDefaultCsvFolder = ThisWorkbook.Path
    End If
End Function

Private Function PickCsvFile(ByVal initialFolder As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select CSV timeline file"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        If Right$(initialFolder, 1) <> "\" Then initialFolder = initialFolder & "\"
        .InitialFileName = initialFolder
        If .Show <> -1 Then
            PickCsvFile = ""
        Else
            PickCsvFile = .SelectedItems(1)
        End If
    End With
End Function

Private Sub ImportCsvToData(ByVal filePath As String)
    Dim wsD As Worksheet
    Dim wbCSV As Workbook
    Dim rng As Range
    Dim screenState As Boolean, eventsState As Boolean
    Dim calcState As XlCalculation

    Set wsD = ThisWorkbook.Sheets("Data")

    screenState = Application.ScreenUpdating
    eventsState = Application.EnableEvents
    calcState = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanUp

    wsD.Cells.ClearContents

    Workbooks.OpenText Filename:=filePath, Origin:=xlMSDOS, StartRow:=1, _
        DataType:=xlDelimited, TextQualifier:=xlTextQualifierDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False, _
        Comma:=True, Space:=False, Other:=False

    Set wbCSV = ActiveWorkbook
    Set rng = wbCSV.Sheets(1).UsedRange

    wsD.Range("A1").Resize(rng.Rows.Count, rng.Columns.Count).Value = rng.Value

CleanUp:
    On Error Resume Next
    If Not wbCSV Is Nothing Then wbCSV.Close SaveChanges:=False
    Application.ScreenUpdating = screenState
    Application.EnableEvents = eventsState
    Application.Calculation = calcState
End Sub

Private Sub SortNodesByOrder(nodes() As String, orders() As Double, ByVal first As Long, ByVal last As Long)
    Dim i As Long, j As Long
    Dim pivot As Double
    Dim tmpO As Double, tmpN As String
    i = first
    j = last
    pivot = orders((first + last) \ 2)
    Do While i <= j
        Do While orders(i) < pivot
            i = i + 1
        Loop
        Do While orders(j) > pivot
            j = j - 1
        Loop
        If i <= j Then
            tmpO = orders(i)
            orders(i) = orders(j)
            orders(j) = tmpO
            tmpN = nodes(i)
            nodes(i) = nodes(j)
            nodes(j) = tmpN
            i = i + 1
            j = j - 1
        End If
    Loop
    If first < j Then SortNodesByOrder nodes, orders, first, j
    If i < last Then SortNodesByOrder nodes, orders, i, last
End Sub
