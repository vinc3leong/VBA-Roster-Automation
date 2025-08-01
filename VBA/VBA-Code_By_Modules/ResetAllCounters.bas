Attribute VB_Name = "ResetAllCounters"
Sub ResetAllCounters()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim counterColumn As ListColumn
    Dim tblNames As Variant
    Dim sheetNames As Variant
    Dim i As Integer
    Dim wsRoster As Worksheet
    
    ' Define your table and corresponding sheet names
    sheetNames = Array("Loan Mail Box PersonnelList", "Morning PersonnelList", _
                       "Afternoon PersonnelList", "AOH PersonnelList", "Sat AOH PersonnelList")
    tblNames = Array("LoanMailBoxMainList", "MorningMainList", "AfternoonMainList", _
                     "AOHMainList", "SatAOHMainList")
                     
    'if table is empty exit sub
    
    ' Loop through all specified tables and reset Duties Counter column to 0
    For i = 0 To UBound(sheetNames)
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        'ws.Unprotect
        Set tbl = ws.ListObjects(tblNames(i))
        
        ' Check if the table is empty
        If tbl.ListRows.count = 0 Then
            Debug.Print "Table " & tblNames(i) & " on sheet " & sheetNames(i) & " is empty. Skipping reset."
            GoTo NextSheet
        End If
        
        Set counterColumn = tbl.ListColumns("Duties Counter")
        
        Dim j As Long
        For j = 1 To counterColumn.DataBodyRange.Rows.count
            counterColumn.DataBodyRange.Cells(j, 1).Value = 0
        Next j
NextSheet:
    Next i
        ' Reprotect the worksheet and lock table ranges
        'With ws
        '    If Not tbl Is Nothing Then
        '        .ListObjects(tbl.Name).Range.Locked = True
        '    End If
        '    .Range("D5:D9").Locked = False ' keep data entry remains unlocked
        '    .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                     AllowFiltering:=True, AllowSorting:=True, AllowUsingPivotTables:=True
        'End With
    
    ' Activate Roster sheet
    Set wsRoster = Sheets("Roster")
    wsRoster.Activate


End Sub




