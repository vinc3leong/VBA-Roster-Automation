Attribute VB_Name = "InsertStaff"
Sub InsertStaff(dutyType As String)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim staffName As String, dept As String
    Dim availType As String, workDays As String, percentage As String
    Dim checkRow As Long
    Dim specificDaysTbl As ListObject
    Dim specificRow As ListRow
    
    ' Set worksheet and tables based on dutyType
    Select Case UCase(dutyType)
        Case "LOANMAILBOX"
            Set ws = ThisWorkbook.Sheets("Loan Mail Box PersonnelList")
            Set tbl = ws.ListObjects("LoanMailBoxMainList")
            Set specificDaysTbl = ws.ListObjects("LoanMailBoxSpecificDaysWorkingStaff")
        Case "MORNING"
            Set ws = ThisWorkbook.Sheets("Morning PersonnelList")
            Set tbl = ws.ListObjects("MorningMainList")
            Set specificDaysTbl = ws.ListObjects("MorningSpecificDaysWorkingStaff")
        Case "AFTERNOON"
            Set ws = ThisWorkbook.Sheets("Afternoon PersonnelList")
            Set tbl = ws.ListObjects("AfternoonMainList")
            Set specificDaysTbl = ws.ListObjects("AfternoonSpecificDaysWorkingStaff")
        Case "AOH"
            Set ws = ThisWorkbook.Sheets("AOH PersonnelList")
            Set tbl = ws.ListObjects("AOHMainList")
            Set specificDaysTbl = ws.ListObjects("AOHSpecificDaysWorkingStaff")
        Case "SAT_AOH"
            Set ws = ThisWorkbook.Sheets("Sat AOH PersonnelList")
            Set tbl = ws.ListObjects("SatAOHMainList")
            ' No specificDaysTbl for Sat AOH
        Case Else
            MsgBox "Invalid duty type. Use 'LoanMailBox', 'Morning', 'Afternoon', 'AOH', or 'Sat_AOH'.", vbExclamation
            Exit Sub
    End Select
    
    ' Validate worksheet and table
    If ws Is Nothing Then
        MsgBox "Worksheet for " & dutyType & " not found.", vbExclamation
        Exit Sub
    End If
    If tbl Is Nothing Then
        MsgBox "Table 'MainList' not found on '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If
    If specificDaysTbl Is Nothing And UCase(Trim(ws.Range("D7").Value)) = "SPECIFIC DAYS" And UCase(dutyType) <> "SAT_AOH" Then
        MsgBox "Table 'SpecificDaysWorkingStaff' not found on '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If

    ' Unprotect the worksheet (unlock)
    ws.Unprotect
    
    ' to remove the line breaks before checking for name duplication
    Dim cleanInput As String
    cleanInput = RemoveLineBreaks(ws.Range("D5").Value)
    ws.Range("D5").Value = cleanInput
    
    ' Read input values from the unlocked data entry range
    staffName = UCase(Trim(cleanInput)) ' Name (already cleaned)
    dept = Trim(RemoveLineBreaks(ws.Range("D6").Value)) ' Department
    availType = UCase(Trim(RemoveLineBreaks(ws.Range("D7").Value))) ' Availability Type
    workDays = Trim(RemoveLineBreaks(ws.Range("D8").Value)) ' Working Days
    percentage = Trim(RemoveLineBreaks(ws.Range("D9").Value)) ' Duties Percentage

    ' Auto-fill logic based on Availability Type
    If availType = "ALL DAYS" Then
        percentage = "100"
        workDays = ""
    ElseIf availType = "SPECIFIC DAYS" Then
        If workDays = "" Then
            MsgBox "Please enter Working Days for Specific Days availability.", vbExclamation
            GoTo ReprotectAndExit
        End If
    End If
    
    ' Validate percentage
    If percentage = "" Or Not IsNumeric(percentage) Or Val(percentage) <= 0 Or Val(percentage) > 100 Then
        MsgBox "Please enter a valid Duties Percentage (1-100).", vbExclamation
        GoTo ReprotectAndExit
    End If

    If Len(Trim(staffName)) = 0 Or Len(Trim(dept)) = 0 Then
        MsgBox "Please fill in Name and Department.", vbExclamation
        GoTo ReprotectAndExit
    End If
    
    ' Check for duplicate names (now handles line breaks properly)
    For checkRow = 1 To tbl.ListRows.count
        Dim existingName As String
        existingName = UCase(Trim(RemoveLineBreaks(tbl.ListRows(checkRow).Range.Cells(1, GetColumnIndex(tbl, "Name")).Value)))
        If existingName = staffName Then
            MsgBox "This staff name already exists.", vbExclamation
            GoTo ReprotectAndExit
        End If
    Next checkRow
    
    Set newRow = tbl.ListRows.Add(AlwaysInsert:=True)
    With newRow.Range
        Dim nameIndex As Long, deptIndex As Long, availIndex As Long
        Dim percIndex As Long, maxIndex As Long, counterIndex As Long
        
        nameIndex = GetColumnIndex(tbl, "Name")
        deptIndex = GetColumnIndex(tbl, "Department")
        availIndex = GetColumnIndex(tbl, "Availability Type")
        percIndex = GetColumnIndex(tbl, "Duties Percentage (%)")
        maxIndex = GetColumnIndex(tbl, "Max Duties")
        counterIndex = GetColumnIndex(tbl, "Duties Counter")
        
        If nameIndex = -1 Or deptIndex = -1 Or counterIndex = -1 Or availIndex = -1 Or percIndex = -1 Or maxIndex = -1 Then
            MsgBox "Required columns not found in '" & tbl.Name & "'.", vbExclamation
            newRow.Delete
            GoTo ReprotectAndExit
        End If
        
        .Cells(1, nameIndex).Value = staffName
        .Cells(1, deptIndex).Value = dept
        .Cells(1, availIndex).Value = availType
        .Cells(1, percIndex).Value = Val(percentage)
        .Cells(1, counterIndex).Value = 0
    End With

    ' Handle specific days workers table
    If availType = "SPECIFIC DAYS" Then
        Set specificRow = specificDaysTbl.ListRows.Add(AlwaysInsert:=True)
        With specificRow.Range
            Dim specNameIndex As Long, specDaysIndex As Long
            specNameIndex = GetColumnIndex(specificDaysTbl, "Name")
            specDaysIndex = GetColumnIndex(specificDaysTbl, "Working Days")
            If specNameIndex = -1 Or specDaysIndex = -1 Then
                MsgBox "Required columns not found in specific days table.", vbExclamation
                specificRow.Delete
                newRow.Delete
                GoTo ReprotectAndExit
            End If
            .Cells(1, specNameIndex).Value = staffName
            .Cells(1, specDaysIndex).Value = workDays
        End With
    End If

    CalculateMaxDuties.CalculateMaxDuties dutyType

    ' Clear input of data entry
    ws.Range("D5:D9").ClearContents

    MsgBox "Staff added successfully for " & dutyType & "!", vbInformation
    GoTo ReprotectAndExit

ReprotectAndExit:
    ' Reprotect the worksheet and lock table ranges
    With ws
        On Error Resume Next
        .Unprotect
        If Not tbl Is Nothing Then .ListObjects(tbl.Name).Range.Locked = True
        If Not specificDaysTbl Is Nothing Then .ListObjects(specificDaysTbl.Name).Range.Locked = True
        .Range("D5:D9").Locked = False
        .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                 AllowFiltering:=True, AllowSorting:=True, AllowUsingPivotTables:=True
        On Error GoTo 0
    End With
    ws.Activate
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    If Not newRow Is Nothing Then newRow.Delete
    If Not specificRow Is Nothing Then specificRow.Delete
    GoTo ReprotectAndExit
End Sub

' NEW: Function to remove line breaks from text
Private Function RemoveLineBreaks(inputText As String) As String
    If Len(inputText) = 0 Then
        RemoveLineBreaks = ""
        Exit Function
    End If
    
    ' Replace various types of line breaks with space
    RemoveLineBreaks = Replace(inputText, Chr(10), " ")  ' Line feed
    RemoveLineBreaks = Replace(RemoveLineBreaks, Chr(13), " ")  ' Carriage return
    RemoveLineBreaks = Replace(RemoveLineBreaks, vbCrLf, " ")  ' CR+LF
    RemoveLineBreaks = Replace(RemoveLineBreaks, vbNewLine, " ")  ' New line
    RemoveLineBreaks = Replace(RemoveLineBreaks, vbLf, " ")  ' Line feed
    RemoveLineBreaks = Replace(RemoveLineBreaks, vbCr, " ")  ' Carriage return
    
    ' Clean up multiple spaces
    RemoveLineBreaks = WorksheetFunction.Trim(RemoveLineBreaks)
End Function

' Helper function to get column index
Private Function GetColumnIndex(tbl As ListObject, columnName As String) As Long
    On Error Resume Next
    GetColumnIndex = tbl.ListColumns(columnName).Index
    If Err.Number <> 0 Then GetColumnIndex = -1
    On Error GoTo 0
End Function

' Wrapper subroutines for different shifts
Sub RunInsertStaffLMB()
    InsertStaff "LoanMailBox"
End Sub

Sub RunInsertStaffMorning()
    InsertStaff "Morning"
End Sub

Sub RunInsertStaffAfternoon()
    InsertStaff "Afternoon"
End Sub

Sub RunInsertStaffAOH()
    InsertStaff "AOH"
End Sub

Sub RunInsertStaffSatAOH()
Attribute RunInsertStaffSatAOH.VB_ProcData.VB_Invoke_Func = " \n14"
    InsertStaff "Sat_AOH"
End Sub

