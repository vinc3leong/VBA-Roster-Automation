Attribute VB_Name = "Report"
Sub GenerateMorningShiftAnalysis()
    Dim wsPersonnel As Worksheet, wsRoster As Worksheet, wsAnalysis As Worksheet
    Dim tbl As ListObject
    Dim nameList As Range, dutyCounterList As Range
    Dim lastRow As Long, i As Long
    Dim dict As Object
    Dim empName As String
    Dim MOR_COL As Long: MOR_COL = 6 ' Morning Shift column
    Dim START_ROW As Long: START_ROW = 6 ' Data start row

    ' Set sheets
    Set wsPersonnel = Sheets("Morning PersonnelList")
    Set wsRoster = Sheets("Roster")
    
    ' Create or clear analysis sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("MorningAnalysis").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsAnalysis = Sheets.Add(After:=Sheets(Sheets.count))
    wsAnalysis.Name = "MorningAnalysis"
    
    ' Get the table
    Set tbl = wsPersonnel.ListObjects("MorningMainList")
    Set nameList = tbl.ListColumns("Name").DataBodyRange
    Set dutyCounterList = tbl.ListColumns("Duties Counter").DataBodyRange

    ' Header row for analysis
    With wsAnalysis
        .Range("A1").Value = "Name"
        .Range("B1").Value = "System Counter"
        .Range("C1").Value = "Actual Counter"
        .Range("D1").Value = "Difference"
    End With

    ' Copy names and system counter to analysis sheet
    For i = 1 To nameList.Rows.count
        wsAnalysis.Cells(i + 1, 1).Value = nameList.Cells(i, 1).Value
        wsAnalysis.Cells(i + 1, 2).Value = dutyCounterList.Cells(i, 1).Value
    Next i

    ' Create dictionary to count actual appearances
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To nameList.Rows.count
        empName = nameList.Cells(i, 1).Value
        dict(empName) = 0
    Next i

    ' Walk through the roster and count appearances
    lastRow = wsRoster.Cells(wsRoster.Rows.count, MOR_COL).End(xlUp).row
    For i = START_ROW To lastRow
        empName = Trim(wsRoster.Cells(i, MOR_COL).Value)
        If dict.Exists(empName) Then
            dict(empName) = dict(empName) + 1
        End If
    Next i

    ' Write actual counter and difference to sheet
    For i = 2 To nameList.Rows.count + 1
        empName = wsAnalysis.Cells(i, 1).Value
        wsAnalysis.Cells(i, 3).Value = dict(empName) ' Actual Counter
        wsAnalysis.Cells(i, 4).FormulaR1C1 = "=RC[-2]-RC[-1]" ' System - Actual
    Next i

    MsgBox "Morning shift analysis generated in 'MorningAnalysis' sheet.", vbInformation
End Sub
