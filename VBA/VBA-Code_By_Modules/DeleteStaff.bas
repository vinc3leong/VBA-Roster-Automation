Attribute VB_Name = "DeleteStaff"
Sub DeleteStaff()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim selectedCell As Range
    Dim rowToDelete As ListRow
    Dim dutyType As String
    Dim specificDaysTbl As ListObject
    Dim availIndex As Long
    Const password As String = "rostering2025" ' Password just for deletion authorization
    Dim enteredPassword As String
    Dim sdRow As ListRow
    Dim nameIndex As Long, rowIndex As Long, sdNameIndex As Long
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then
        MsgBox "No active worksheet found.", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    ws.Unprotect
    On Error GoTo ErrHandler
    
    ' Determine the duty type based on the active sheet with trimming
    Select Case Trim(UCase(ws.Name))
        Case "LOAN MAIL BOX PERSONNELLIST"
            dutyType = "LOANMAILBOX"
            Set tbl = ws.ListObjects("LoanMailBoxMainList")
            Set specificDaysTbl = ws.ListObjects("LoanMailBoxSpecificDaysWorkingStaff")
        Case "MORNING PERSONNELLIST"
            dutyType = "MORNING"
            Set tbl = ws.ListObjects("MorningMainList")
            Set specificDaysTbl = ws.ListObjects("MorningSpecificDaysWorkingStaff")
        Case "AFTERNOON PERSONNELLIST"
            dutyType = "AFTERNOON"
            Set tbl = ws.ListObjects("AfternoonMainList")
            Set specificDaysTbl = ws.ListObjects("AfternoonSpecificDaysWorkingStaff")
        Case "AOH PERSONNELLIST"
            dutyType = "AOH"
            Set tbl = ws.ListObjects("AOHMainList")
            Set specificDaysTbl = ws.ListObjects("AOHSpecificDaysWorkingStaff")
        Case "SAT AOH PERSONNELLIST"
            dutyType = "SAT_AOH"
            Set tbl = ws.ListObjects("SatAOHMainList")
            Set specificDaysTbl = Nothing
        Case Else
            MsgBox "This sheet is not a recognized personnel list: '" & ws.Name & "'. Please select a valid personnel list sheet.", vbExclamation
            GoTo ReprotectAndExit
    End Select
    
    ' Validate main table exists
    If tbl Is Nothing Then
        MsgBox "Main table not found on '" & ws.Name & "'.", vbExclamation
        GoTo ReprotectAndExit
    End If
    
    ' Get the selected cell
    Set selectedCell = ActiveCell
    If Not Intersect(selectedCell, tbl.Range) Is Nothing Then
        ' Check if the selected cell is in the "Name" column
        nameIndex = GetColumnIndex(tbl, "Name")
        If nameIndex = -1 Then
            MsgBox "Column 'Name' not found in main table.", vbExclamation
            GoTo ReprotectAndExit
        End If
        If selectedCell.Column <> tbl.Range.Cells(1, nameIndex).Column Then
            MsgBox "Please select a cell in the 'Name' column to delete the staff.", vbExclamation
            GoTo ReprotectAndExit
        End If
        
        ' Find the row in the table
        rowIndex = selectedCell.row - tbl.Range.row
        If rowIndex <= 0 Or rowIndex > tbl.ListRows.count Then
            MsgBox "Selected cell is not within a valid table row.", vbExclamation
            GoTo ReprotectAndExit
        End If
        Set rowToDelete = tbl.ListRows(rowIndex)
        
        ' Only check Availability Type if this isn't Sat AOH sheet
        If dutyType <> "SAT_AOH" Then
            availIndex = GetColumnIndex(tbl, "Availability Type")
            If availIndex = -1 Then
                MsgBox "Column 'Availability Type' not found in main table.", vbExclamation
                GoTo ReprotectAndExit
            End If
            
            ' Clear filters if any (for validation purposes)
            If tbl.ShowAutoFilter Then
                tbl.AutoFilter.ShowAllData
            End If
            If Not specificDaysTbl Is Nothing And specificDaysTbl.ShowAutoFilter Then
                specificDaysTbl.AutoFilter.ShowAllData
            End If
        End If
    Else
        MsgBox "Please select a cell within the main table to delete a staff.", vbExclamation
        GoTo ReprotectAndExit
    End If
    
    ' Only ask for password AFTER all validations pass
    enteredPassword = InputBox("Please enter the password for deletion:", "Password Authentication")
    If enteredPassword <> password Then
        MsgBox "Incorrect password. Delete operation declined.", vbCritical
        GoTo ReprotectAndExit
    End If
    
    ' Perform deletion within the current sheets contextt
    If dutyType = "SAT_AOH" Then
        rowToDelete.Delete
    Else
        ' For other sheets, handle specific days workers if needed
        If UCase(Trim(rowToDelete.Range.Cells(1, availIndex).Value)) = "SPECIFIC DAYS" And Not specificDaysTbl Is Nothing Then
            sdNameIndex = GetColumnIndex(specificDaysTbl, "Name")
            If sdNameIndex <> -1 Then
                ' Clear filters on specific days table
                If specificDaysTbl.ShowAutoFilter Then
                    specificDaysTbl.AutoFilter.ShowAllData
                End If
                ' Find and delete the corresponding row in the same sheets specific days tablee
                For Each sdRow In specificDaysTbl.ListRows
                    If UCase(Trim(sdRow.Range.Cells(1, sdNameIndex).Value)) = UCase(Trim(selectedCell.Value)) Then
                        sdRow.Delete
                        Exit For
                    End If
                Next sdRow
            End If
        End If
        rowToDelete.Delete
    End If
    
    CalculateMaxDuties.CalculateMaxDuties dutyType
    MsgBox "Staff deleted and Max Duties recalculated successfully for " & dutyType & " on " & ws.Name & ".", vbInformation

ReprotectAndExit:
    ' Reprotect the worksheet WITHOUT password but keep D5:D9 unlocked
    If Not ws Is Nothing Then
        On Error Resume Next ' In case sheet is already protected
        With ws
            .Unprotect ' Ensure sheet is unprotected before setting locked ranges
            If Not tbl Is Nothing Then
                .ListObjects(tbl.Name).Range.Locked = True
            End If
            If Not specificDaysTbl Is Nothing Then
                .ListObjects(specificDaysTbl.Name).Range.Locked = True
            End If
            .Range("D5:D9").Locked = False ' Keep data entry unlocked
            .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                     AllowFiltering:=True, AllowSorting:=True, AllowUsingPivotTables:=True
        End With
        On Error GoTo 0
    End If
    ws.Activate
    Exit Sub

ErrHandler:
    MsgBox "An error occurred: " & Err.Description & vbCrLf & _
           "Line: " & Erl & vbCrLf & _
           "Duty Type: " & dutyType & vbCrLf & _
           "Worksheet: " & IIf(ws Is Nothing, "Not Set", ws.Name), vbCritical
    GoTo ReprotectAndExit
End Sub

' Helper function to get column index
Private Function GetColumnIndex(tbl As ListObject, columnName As String) As Long
    On Error Resume Next
    GetColumnIndex = tbl.ListColumns(columnName).Index
    If Err.Number <> 0 Then GetColumnIndex = -1
    On Error GoTo 0
End Function

