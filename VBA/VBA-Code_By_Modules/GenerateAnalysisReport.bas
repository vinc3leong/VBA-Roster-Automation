Attribute VB_Name = "GenerateAnalysisReport"
Private last_row_roster As Long
' Reusable function to generate side-by-side shift analysis blocks
Sub GenerateShiftAnalysisBlock(wsAnalysis As Worksheet, rosterSheet As Worksheet, _
                                personnelSheetName As String, tableName As String, _
                                slotTitle As String, rosterCol1 As Long, startCol As Long, _
                                Optional rosterCol2 As Variant)

    Dim wsPersonnel As Worksheet
    Dim tbl As ListObject
    Dim nameList As Range, dutyCounterList As Range
    Dim dict As Object
    Dim empName As String
    Dim cell As Range, cellValue As String, currStaff As String
    Dim rowOffset As Long: rowOffset = 4
    Dim lastRow As Long, NextRow As Long
    Dim tableRange As Range, analysisTable As ListObject
    Dim i As Long, tableWidth As Long: tableWidth = 5
    Dim enteredPassword As String

    Set wsPersonnel = Sheets(personnelSheetName)
    Set tbl = wsPersonnel.ListObjects(tableName)
    Set nameList = tbl.ListColumns("Name").DataBodyRange
    Set dutyCounterList = tbl.ListColumns("Duties Counter").DataBodyRange

    ' Add small section title
    With wsAnalysis.Cells(3, startCol).Resize(1, 3)
        .Merge
        .Value = slotTitle
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(184, 204, 228)
    End With

    ' Add header
    With wsAnalysis
        .Cells(rowOffset, startCol).Value = "Name"
        .Cells(rowOffset, startCol + 1).Value = "System Counter"
        .Cells(rowOffset, startCol + 2).Value = "Actual Counter"
        .Cells(rowOffset, startCol + 3).Value = "Difference"
        .Cells(rowOffset, startCol + 4).Value = "% Difference"
    End With

    ' Check if table exists and if it is empty
    On Error Resume Next
    Set tbl = wsPersonnel.ListObjects(tableName)
    On Error GoTo 0
    
    If tbl Is Nothing Or tbl.ListRows.count = 0 Then
        ' Create an empty analysis table with only headers
        Set tableRange = wsAnalysis.Range(wsAnalysis.Cells(rowOffset, startCol), _
                                          wsAnalysis.Cells(rowOffset, startCol + tableWidth - 1))
        
        Set analysisTable = wsAnalysis.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
        analysisTable.Name = Replace(slotTitle, " ", "") & "Table"
        
        ' No data rows, so skip formatting % Difference
        Exit Sub
    End If


    Set dict = CreateObject("Scripting.Dictionary")
    ' Load initial staff counters
    For i = 1 To nameList.Rows.count
        empName = UCase(Trim(nameList.Cells(i, 1).Value))
        wsAnalysis.Cells(rowOffset + i, startCol).Value = empName
        wsAnalysis.Cells(rowOffset + i, startCol + 1).Value = dutyCounterList.Cells(i, 1).Value
        dict(empName) = 0
    Next i

    Dim colIndex As Variant
    Dim colList As Variant
    
    If IsMissing(rosterCol2) Then
        colList = Array(rosterCol1)
    Else
        colList = Array(rosterCol1, rosterCol2)
    End If
    
    For Each colIndex In colList
        For i = 6 To last_row_roster
            Set cell = rosterSheet.Cells(i, colIndex)
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
                    NextRow = wsAnalysis.Cells(wsAnalysis.Rows.count, startCol).End(xlUp).row + 1
                    wsAnalysis.Cells(NextRow, startCol).Value = currStaff
                    wsAnalysis.Cells(NextRow, startCol + 1).Value = 0
                    wsAnalysis.Cells(NextRow, startCol + 2).Value = 1
                    wsAnalysis.Cells(NextRow, startCol + 3).FormulaR1C1 = "=RC[-1]-RC[-2]"
                    wsAnalysis.Cells(NextRow, startCol + 4).FormulaR1C1 = "=IF(RC[-3]=0,"""",RC[-1]/RC[-3]*100)"
                    dict(currStaff) = 1
                    wsAnalysis.Range(wsAnalysis.Cells(NextRow, startCol), wsAnalysis.Cells(NextRow, startCol + 4)).Interior.Color = RGB(255, 255, 153)
                End If
            End If
        Next i
    Next colIndex

    For i = rowOffset + 1 To wsAnalysis.Cells(wsAnalysis.Rows.count, startCol).End(xlUp).row
        empName = UCase(Trim(wsAnalysis.Cells(i, startCol).Value))
        If dict.Exists(empName) Then
            wsAnalysis.Cells(i, startCol + 2).Value = dict(empName)
            wsAnalysis.Cells(i, startCol + 3).FormulaR1C1 = "=RC[-1]-RC[-2]"
            wsAnalysis.Cells(i, startCol + 4).FormulaR1C1 = "=IF(RC[-3]=0,0,RC[-1]/RC[-3]*100)"
        End If
    Next i

    ' Format as Table
    lastRow = wsAnalysis.Cells(wsAnalysis.Rows.count, startCol).End(xlUp).row
    Set tableRange = wsAnalysis.Range(wsAnalysis.Cells(rowOffset, startCol), wsAnalysis.Cells(lastRow, startCol + tableWidth - 1))
    Set analysisTable = wsAnalysis.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    analysisTable.Name = Replace(slotTitle, " ", "") & "Table"
    analysisTable.ListColumns("% Difference").DataBodyRange.NumberFormat = "0.00"
End Sub
Sub GenerateTotalSummaryTable(wsAnalysis As Worksheet)
    Dim summaryDict As Object
    Set summaryDict = CreateObject("Scripting.Dictionary")
    
    Dim tableNames As Variant
    tableNames = Array("LoanMailBoxSlotAnalysisTable", "MorningSlotAnalysisTable", _
                       "AfternoonSlotAnalysisTable", "AOHSlotAnalysisTable", "SatAOHSlotAnalysisTable")

    Dim tbl As ListObject, row As ListRow
    Dim empName As String
    Dim sysCount As Long, actCount As Long
    Dim tblName As Variant

    ' Loop through each analysis table
    For Each tblName In tableNames
        On Error Resume Next
        Set tbl = wsAnalysis.ListObjects(tblName)
        On Error GoTo 0

        ' Skip if table doesn't exist or is empty
        If Not tbl Is Nothing Then
            If tbl.ListRows.count > 0 Then
                ' Loop through each row in the table
                For Each row In tbl.ListRows
                    empName = UCase(Trim(row.Range.Cells(1, 1).Value))
                    sysCount = Val(row.Range.Cells(1, 2).Value)
                    actCount = Val(row.Range.Cells(1, 3).Value)
                    
                    If summaryDict.Exists(empName) Then
                        Dim counts As Variant
                        counts = summaryDict(empName)
                        counts(0) = counts(0) + sysCount
                        counts(1) = counts(1) + actCount
                        summaryDict(empName) = counts
                    Else
                        summaryDict(empName) = Array(sysCount, actCount)
                    End If
                Next row
            End If
        End If
        Set tbl = Nothing
    Next tblName

    ' Determine where to place the summary table
    Dim startCol As Long
    startCol = wsAnalysis.Cells(4, wsAnalysis.Columns.count).End(xlToLeft).Column + 2
    Dim rowOffset As Long: rowOffset = 4

    ' Add table header
    With wsAnalysis
        .Cells(3, startCol).Resize(1, 5).Merge
        .Cells(3, startCol).Value = "Total Summary"
        .Cells(3, startCol).Font.Bold = True
        .Cells(3, startCol).Interior.Color = RGB(184, 204, 228)
        .Cells(3, startCol).HorizontalAlignment = xlCenter

        .Cells(rowOffset, startCol).Value = "Name"
        .Cells(rowOffset, startCol + 1).Value = "System Counter"
        .Cells(rowOffset, startCol + 2).Value = "Actual Counter"
        .Cells(rowOffset, startCol + 3).Value = "Difference"
        .Cells(rowOffset, startCol + 4).Value = "% Difference"
        .Range(.Cells(rowOffset, startCol), .Cells(rowOffset, startCol + 4)).Font.Bold = True
    End With

    ' Write data into sheet
    Dim i As Long: i = rowOffset + 1
    Dim diff As Long, pctDiff As Double
    Dim sysVal As Long, actVal As Long
    Dim staff As Variant

    For Each staff In summaryDict.Keys
        sysVal = summaryDict(staff)(0)
        actVal = summaryDict(staff)(1)
        diff = actVal - sysVal
        
        If sysVal <> 0 Then
            pctDiff = diff / sysVal * 100
        Else
            pctDiff = 0
        End If

        wsAnalysis.Cells(i, startCol).Value = staff
        wsAnalysis.Cells(i, startCol + 1).Value = sysVal
        wsAnalysis.Cells(i, startCol + 2).Value = actVal
        wsAnalysis.Cells(i, startCol + 3).Value = diff
        wsAnalysis.Cells(i, startCol + 4).Value = pctDiff
        i = i + 1
    Next staff

    ' Create an empty table if no data exists
    Dim lastRow As Long
    If summaryDict.count = 0 Then
        lastRow = rowOffset
    Else
        lastRow = wsAnalysis.Cells(wsAnalysis.Rows.count, startCol).End(xlUp).row
    End If

    ' Format as table
    Dim tableRange As Range
    Set tableRange = wsAnalysis.Range(wsAnalysis.Cells(rowOffset, startCol), wsAnalysis.Cells(lastRow, startCol + 4))

    Dim summaryTable As ListObject
    Set summaryTable = wsAnalysis.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    summaryTable.Name = "TotalSummaryTable"
    If summaryTable.ListRows.count > 0 Then
        summaryTable.ListColumns("% Difference").DataBodyRange.NumberFormat = "0.00"
    End If
End Sub



Sub MasterGenerateAllAnalyses()
    Dim wsAnalysis As Worksheet
    Dim rosterSheet As Worksheet
    Dim userRange As Range
    Dim selectedSheet As Worksheet
    Const password As String = "rostering2025"
    
    enteredPassword = InputBox("Please enter the password to generate analysis report:", "Password Authentication")
    If enteredPassword <> password Then
        MsgBox "Incorrect password. Unable to generate report.", vbCritical
        Exit Sub
    End If

    ' Prompt user to click on any cell in the target ActualRoster_* sheet
    On Error Resume Next
    Set userRange = Application.InputBox( _
        Prompt:="Please choose one 'ActualRoster' sheet to analyse." & vbCrLf & _
                "After that, click on any cell on the selected 'ActualRoster' sheet." & vbCrLf & _
                "The sheet name must start with 'ActualRoster_'", _
        title:="Select Actual Roster Sheet", _
        Type:=8)
    On Error GoTo 0

    If userRange Is Nothing Then Exit Sub ' User cancelled

    Set selectedSheet = userRange.Worksheet
    If selectedSheet.Name Like "ActualRoster_*" = False Then
        MsgBox "Invalid selection. Please choose a sheet that starts with 'ActualRoster_'.", vbExclamation
        Exit Sub
    End If
    Set rosterSheet = selectedSheet

    ' Create ReportAnalysis sheet
    Dim dtNamePart As String, formattedName As String
    dtNamePart = Mid(selectedSheet.Name, 14)
    formattedName = "AnalysisReport_" & dtNamePart
    
    Set wsAnalysis = Sheets.Add(After:=Sheets(Sheets.count))
    On Error Resume Next
    wsAnalysis.Name = formattedName
    If Err.Number <> 0 Then
        MsgBox "Could not name sheet as " & formattedName & ". It may already exist.", vbExclamation
        wsAnalysis.Name = "AnalysisReport_" & Format(Now, "yyyymmdd_hhnnss")
    End If
    On Error GoTo 0

    'Find last row of roster
    If selectedSheet.Cells(2, 10).Value = "Jan-Jun" And selectedSheet.Cells(2, 13).Value Mod 4 = 0 Then
        last_row_roster = 187
    ElseIf selectedSheet.Cells(2, 10).Value = "Jan-Jun" Then
        last_row_roster = 186
    Else
        last_row_roster = 189
    End If
    
    ' Big title
    With wsAnalysis.Range("A1:Z1")
        .Merge
        .Value = "Analysis Report"
        .Font.Size = 16
        .Font.Bold = True
        .Interior.Color = RGB(255, 199, 206)
        .HorizontalAlignment = xlCenter
    End With

    ' Show selected ActualRoster sheet name in row 2
    With wsAnalysis.Range("A2:Z2")
        .Merge
        .Value = "Based on: " & rosterSheet.Name
        .Font.Italic = True
        .HorizontalAlignment = xlCenter
    End With

    ' Generate all 5 analyses side by side
    GenerateShiftAnalysisBlock wsAnalysis, rosterSheet, "Loan Mail Box PersonnelList", "LoanMailBoxMainList", "Loan Mail Box Slot Analysis", LMB_COL, 1
    GenerateShiftAnalysisBlock wsAnalysis, rosterSheet, "Morning PersonnelList", "MorningMainList", "Morning Slot Analysis", MOR_COL, 7
    GenerateShiftAnalysisBlock wsAnalysis, rosterSheet, "Afternoon PersonnelList", "AfternoonMainList", "Afternoon Slot Analysis", AFT_COL, 13
    GenerateShiftAnalysisBlock wsAnalysis, rosterSheet, "AOH PersonnelList", "AOHMainList", "AOH Slot Analysis", AOH_COL, 19
    GenerateShiftAnalysisBlock wsAnalysis, rosterSheet, "Sat AOH PersonnelList", "SatAOHMainList", _
    "Sat AOH Slot Analysis", SAT_AOH_COL1, 25, SAT_AOH_COL2
    GenerateTotalSummaryTable wsAnalysis
    
    With wsAnalysis.Cells
        .Locked = True
    End With
    
    wsAnalysis.Protect password:="nuslib2017@52", _
                        AllowSorting:=True, _
                        AllowFiltering:=True, _
                        AllowFormattingCells:=True
                        
    

    MsgBox "All shift analyses completed for '" & rosterSheet.Name & "'!", vbInformation
End Sub


Sub TestGenerateSummary()
    Call GenerateTotalSummaryTable(Sheets("AnalysisReport_20250724_0306"))
End Sub


