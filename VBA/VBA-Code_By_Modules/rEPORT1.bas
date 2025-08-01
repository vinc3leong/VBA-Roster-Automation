Attribute VB_Name = "rEPORT1"
Sub GenerateMorningShiftAnalysis()
    Dim wsPersonnel As Worksheet
    Dim wsAnalysis As Worksheet
    Dim wsRosterCopy As Worksheet
    Dim tbl As ListObject
    Dim nameList As Range
    Dim dutyCounterList As Range
    Dim lastRow As Long, i As Long
    Dim dict As Object
    Dim empName As String
    Dim latestRosterName As String
    Dim sht As Worksheet
    Dim newestDate As Date
    Dim START_ROW As Long: START_ROW = 6
    Dim MOR_COL As Long: MOR_COL = 6
    Dim cell As Range, cellValue As String, currStaff As String
    Dim NextRow As Long, tableStartRow As Long

    ' Find latest ActualRoster_* sheet
    newestDate = 0
    For Each sht In ThisWorkbook.Sheets
        If sht.Name Like "ActualRoster_*" Then
            Dim dtPart As String
            dtPart = Replace(Mid(sht.Name, 14), "_", " ")
            On Error Resume Next
            Dim parsedDate As Date
            parsedDate = CDate(Left(dtPart, 4) & "/" & Mid(dtPart, 5, 2) & "/" & Mid(dtPart, 7, 2) & " " & Mid(dtPart, 10, 2) & ":" & Mid(dtPart, 12, 2))
            If Err.Number = 0 Then
                If parsedDate > newestDate Then
                    newestDate = parsedDate
                    latestRosterName = sht.Name
                End If
            End If
            On Error GoTo 0
        End If
    Next sht
    
    If latestRosterName = "" Then
        MsgBox "No ActualRoster_* sheet found.", vbExclamation
        Exit Sub
    End If

    ' Add big title at row 1
    With wsAnalysis.Range("A1:E1")
        .Merge
        .Value = "Analysis Report"
        .Interior.Color = RGB(255, 199, 206) ' Light red
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Set wsRosterCopy = Sheets(latestRosterName)
    Set wsPersonnel = Sheets("Morning PersonnelList")

    ' Create or clear analysis sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("MorningAnalysis").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsAnalysis = Sheets.Add(After:=Sheets(Sheets.count))
    wsAnalysis.Name = "MorningAnalysis"

    ' Write title
    wsAnalysis.Range("A1").Value = "Morning Slot Analysis"
    wsAnalysis.Range("A1").Font.Bold = True
    wsAnalysis.Range("A1").Font.Size = 14
    tableStartRow = 4 ' header starts here

    ' Header row
    With wsAnalysis
        .Range("A" & tableStartRow).Value = "Name"
        .Range("B" & tableStartRow).Value = "System Counter"
        .Range("C" & tableStartRow).Value = "Actual Counter"
        .Range("D" & tableStartRow).Value = "Difference"
        .Range("E" & tableStartRow).Value = "% Difference"
    End With

    ' Get personnel table
    Set tbl = wsPersonnel.ListObjects("MorningMainList")
    Set nameList = tbl.ListColumns("Name").DataBodyRange
    Set dutyCounterList = tbl.ListColumns("Duties Counter").DataBodyRange

    ' Create dictionary and fill initial system counter
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To nameList.Rows.count
        empName = UCase(Trim(nameList.Cells(i, 1).Value))
        wsAnalysis.Cells(i + tableStartRow, 1).Value = empName
        wsAnalysis.Cells(i + tableStartRow, 2).Value = dutyCounterList.Cells(i, 1).Value
        dict(empName) = 0
    Next i

    ' Count actual appearances from roster
    For i = START_ROW To 186
        Set cell = wsRosterCopy.Cells(i, MOR_COL)
        cellValue = cell.Value

        If InStr(cellValue, vbNewLine) > 0 Then
            currStaff = UCase(Trim(Replace(Split(cellValue, vbNewLine)(0), Chr(160), " ")))
        Else
            currStaff = UCase(Trim(cellValue))
        End If

        If Len(currStaff) > 0 And currStaff <> "CLOSED" Then
            If dict.Exists(currStaff) Then
                dict(currStaff) = dict(currStaff) + 1
            Else
                ' New staff found
                NextRow = wsAnalysis.Cells(wsAnalysis.Rows.count, 1).End(xlUp).row + 1
                wsAnalysis.Cells(NextRow, 1).Value = currStaff
                wsAnalysis.Cells(NextRow, 2).Value = 0
                wsAnalysis.Cells(NextRow, 3).Value = 1
                wsAnalysis.Cells(NextRow, 4).FormulaR1C1 = "=RC[-1]-RC[-2]"
                wsAnalysis.Cells(NextRow, 5).FormulaR1C1 = "=IF(RC[-3]=0,"""",RC[-1]/RC[-3]*100)"
                dict(currStaff) = 1

                ' Highlight new row
                wsAnalysis.Range(wsAnalysis.Cells(NextRow, 1), wsAnalysis.Cells(NextRow, 5)).Interior.Color = RGB(255, 255, 153)
            End If
        End If
    Next i

    ' Fill actual count and compute difference + percentage
    For i = tableStartRow + 1 To wsAnalysis.Cells(wsAnalysis.Rows.count, 1).End(xlUp).row
        empName = UCase(Trim(wsAnalysis.Cells(i, 1).Value))
        If dict.Exists(empName) Then
            wsAnalysis.Cells(i, 3).Value = dict(empName)
            wsAnalysis.Cells(i, 4).FormulaR1C1 = "=RC[-1]-RC[-2]"
            wsAnalysis.Cells(i, 5).FormulaR1C1 = "=IF(RC[-3]=0,0,RC[-1]/RC[-3]*100)"
        End If
    Next i

    ' Format as Table
    Dim tableRange As Range
    Dim analysisTable As ListObject

    lastRow = wsAnalysis.Cells(wsAnalysis.Rows.count, 1).End(xlUp).row
    Set tableRange = wsAnalysis.Range("A" & tableStartRow & ":E" & lastRow)
    Set analysisTable = wsAnalysis.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    analysisTable.Name = "MorningShiftTable"

    ' Format % column to 2 decimal places
    analysisTable.ListColumns("% Difference").DataBodyRange.NumberFormat = "0.00"


    ' Add small section title at row 3
    With wsAnalysis.Range("A3:C3")
        .Merge
        .Value = "Morning Slot Analysis"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = analysisTable.HeaderRowRange.Interior.Color
    End With

    MsgBox "Morning shift analysis generated using '" & latestRosterName & "'.", vbInformation
End Sub

