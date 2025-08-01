Attribute VB_Name = "AssignMorningDuties"
' Declare worksheet and table
Private wsRoster As Worksheet
Private wsSettings As Worksheet
Private wsPersonnel As Worksheet
Private morningtbl As ListObject
Private spectbl As ListObject

Sub AssignMorningDuties()
    Set wsRoster = Sheets("Roster")
    Set wsSettings = Sheets("Settings")
    Set wsPersonnel = Sheets("Morning PersonnelList")
    Set morningtbl = wsPersonnel.ListObjects("MorningMainList")
    Set spectbl = wsPersonnel.ListObjects("MorningSpecificDaysWorkingStaff")
    
    Dim i As Long, j As Long, r As Long
    Dim dateCount As Long
    Dim totalDays As Long
    Dim dayName As String
    Dim maxDuties As Long
    Dim candidates() As String
    Dim staffName As String
    Dim workDays As Variant
    Dim eligibleRows As Collection
    Dim tmpRows() As Long
    Dim assignedCount As Long
    Dim reassignmentAttempts As Long

    totalDays = wsRoster.Range(wsRoster.Cells(START_ROW, DATE_COL), wsRoster.Cells(last_row_roster, DATE_COL)).Rows.count
    Debug.Print "Morning assignment starts here"
    
    ' Step 1: Assign Specific Days Staff
    For i = 1 To spectbl.ListRows.count
        staffName = spectbl.DataBodyRange(i, spectbl.ListColumns("Name").Index).Value
        Debug.Print staffName
        workDays = Split(spectbl.DataBodyRange(i, spectbl.ListColumns("Working Days").Index).Value, ",")
        
        ' Clean up day names (remove spaces)
        For j = 0 To UBound(workDays)
            workDays(j) = Trim(workDays(j))
        Next j
        
        ' Get max duties for this staff from MorningMainList
        For r = 1 To morningtbl.ListRows.count
            If morningtbl.DataBodyRange(r, morningtbl.ListColumns("Name").Index).Value = staffName Then
                maxDuties = morningtbl.DataBodyRange(r, morningtbl.ListColumns("Max Duties").Index).Value
                Exit For
            End If
        Next r
        
        ' Build candidate pool of eligible rows
        Set eligibleRows = GetEligibleRows(totalDays, workDays)
        
        ' Shuffle eligibleRows randomly
        ReDim tmpRows(1 To eligibleRows.count)
        For j = 1 To eligibleRows.count
            tmpRows(j) = eligibleRows(j)
        Next j
        Call ShuffleArray(tmpRows)
        
        ' Assign staff
        assignedCount = 0
        For j = 1 To eligibleRows.count
            If assignedCount >= maxDuties Then Exit For
            If Not IsWorkingOnSameDay(tmpRows(j), staffName) Then
                wsRoster.Cells(tmpRows(j), MOR_COL).Value = staffName
                Call IncrementDutiesCounter(staffName)
                assignedCount = assignedCount + 1
            End If
        Next j
    Next i
    
    ' Step 2: Assign All Days Staff
    For r = START_ROW To last_row_roster
        If wsRoster.Cells(r, DAY_COL).Value = "Sat" Then GoTo SkipDay
        If wsRoster.Cells(r, MOR_COL).Value = "CLOSED" Then GoTo SkipDay
        For i = 1 To morningtbl.ListRows.count
            staffName = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Name").Index).Value
            If UCase(morningtbl.DataBodyRange(i, morningtbl.ListColumns("Availability Type").Index).Value) = "SPECIFIC DAYS" Then
                GoTo SkipStaff
            End If
            
            maxDuties = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Max Duties").Index).Value
            Dim currDuties As Long
            currDuties = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Duties Counter").Index).Value
            ' Check if the staff already reached his max duties
            If currDuties >= maxDuties Then GoTo SkipStaff
            If IsWorkingOnSameDay(r, staffName) Then GoTo SkipStaff
            
            ' Assign from top
            If wsRoster.Cells(r, MOR_COL).Value = "" Then
                wsRoster.Cells(r, MOR_COL).Value = staffName
                Call IncrementDutiesCounter(staffName)
                Exit For
            End If
        
SkipStaff:
        Next i
        
SkipDay:
        Next r
    
    ' Step 3: Reassign for unfilled slots
    reassignmentAttempts = 0
    Const MAX_ATTEMPTS As Long = 10
    Do
        ' Check for unfilled slots
        Dim unfilledSlots As Long
        unfilledSlots = CountUnfilledSlots(START_ROW, last_row_roster)
        Debug.Print "Unfilled slots: " & unfilledSlots
        
        ' Check for eligible staff
        Dim eligibleStaffCount As Long
        eligibleStaffCount = CountEligibleStaff
        Debug.Print "Eligible staff (duties < max): " & eligibleStaffCount
        
        If unfilledSlots > 0 And eligibleStaffCount > 0 Then
            ' Enter reassignment phase
            Call ReassignMorningDuties
        Else
            Exit Do
        End If
        reassignmentAttempts = reassignmentAttempts + 1
        If reassignmentAttempts > MAX_ATTEMPTS Then
            Debug.Print "Max reassignment attempts reached. Assigning remaining staff to unfilled slots."
            ' Fallback: Assign all remaining eligible staff to unfilled slots
            Dim eligibleStaffList As Collection
            Set eligibleStaffList = New Collection
            For i = 1 To morningtbl.ListRows.count
                staffName = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Name").Index).Value
                maxDuties = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Max Duties").Index).Value
                currDuties = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Duties Counter").Index).Value
                If UCase(morningtbl.DataBodyRange(i, morningtbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" And _
                   currDuties < maxDuties Then
                    eligibleStaffList.Add staffName
                End If
            Next i
            
            Dim unfilledSlotsList As Collection
            Set unfilledSlotsList = New Collection
            For r = START_ROW To last_row_roster
                If wsRoster.Cells(r, DAY_COL).Value <> "Sat" And _
                   wsRoster.Cells(r, MOR_COL).Value = "" Then
                    unfilledSlotsList.Add r
                End If
            Next r
            
            ' Assign each eligible staff to an unfilled slot
            For i = 1 To eligibleStaffList.count
                If i <= unfilledSlotsList.count Then
                    staffName = eligibleStaffList(i)
                    r = unfilledSlotsList(i)
                    If Not IsWorkingOnSameDay(r, staffName) Then
                        wsRoster.Cells(r, MOR_COL).Value = staffName
                        wsRoster.Cells(r, MOR_COL).Interior.Color = vbYellow ' Highlight yellow (RGB 255, 255, 0)
                        Call IncrementDutiesCounter(staffName)
                        Debug.Print "Fallback: Assigned " & staffName & " to row " & r & " (highlighted yellow)"
                    Else
                        Debug.Print "Fallback: Skipped " & staffName & " at row " & r & " due to AOH/AFT conflict"
                    End If
                End If
            Next i
            Exit Do
        End If
    Loop
    
    MsgBox "Duties assignment completed!", vbInformation
End Sub

' Helper to count unfilled morning slots
Function CountUnfilledSlots(startRow As Long, endRow As Long) As Long
    Dim r As Long
    Dim count As Long
    count = 0
    For r = startRow To endRow
        If wsRoster.Cells(r, DAY_COL).Value <> "Sat" And _
           wsRoster.Cells(r, MOR_COL).Value = "" Then
            count = count + 1
        End If
    Next r
    CountUnfilledSlots = count
End Function

' Helper to count eligible staff (duties < max)
Function CountEligibleStaff() As Long
    Dim i As Long
    Dim count As Long
    count = 0
    For i = 1 To morningtbl.ListRows.count
        Dim staffName As String
        Dim maxDuties As Long
        Dim currDuties As Long
        staffName = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Name").Index).Value
        maxDuties = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Max Duties").Index).Value
        currDuties = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Duties Counter").Index).Value
        If UCase(morningtbl.DataBodyRange(i, morningtbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" And _
           currDuties < maxDuties Then
            count = count + 1
        End If
    Next i
    CountEligibleStaff = count
End Function
Sub ReassignMorningDuties()
    Dim r As Long
    Dim i As Long
    Dim staffName As String
    Dim maxDuties As Long
    Dim currDuties As Long
    Dim eligibleStaff As String
    Dim swapCandidate As String
    Dim emptyRow As Long ' The empty original slot where eligibleStaff couldn't be assigned
    
    ' Find the first eligible staff (currDuties < maxDuties)
    eligibleStaff = GetFirstEligibleStaff
    If eligibleStaff = "" Then
        Debug.Print "No eligible staff found for Morning reassignment."
        Exit Sub
    End If
    Debug.Print "Eligible staff for Morning reassignment: " & eligibleStaff
    
    ' Assume emptyRow is determined earlier (e.g., the last unfilled slot)
    emptyRow = FindLastUnfilledSlot(START_ROW, last_row_roster)
    If emptyRow = 0 Then Exit Sub ' No empty slot to reassign
    
    ' Loop through all rows to find a swap opportunity
    For r = START_ROW To last_row_roster
        If wsRoster.Cells(r, DAY_COL).Value = "Sat" Or _
           wsRoster.Cells(r, MOR_COL).Value = "CLOSED" Or _
           wsRoster.Cells(r, MOR_COL).Value = "" Then
            GoTo NextRow
        End If
        
        swapCandidate = wsRoster.Cells(r, MOR_COL).Value
        If swapCandidate <> "" And swapCandidate <> eligibleStaff Then
            ' Check if swap is allowed (no conflict in destination and original slot)
            If Not IsWorkingOnSameDay(r, eligibleStaff) And Not IsWorkingOnSameDay(emptyRow, swapCandidate) Then
                ' Perform swap
                wsRoster.Cells(r, MOR_COL).Value = eligibleStaff
                Call IncrementDutiesCounter(eligibleStaff)
                Debug.Print "Assigned " & eligibleStaff & " to row " & r
                
                wsRoster.Cells(emptyRow, MOR_COL).Value = swapCandidate
                Debug.Print "Assigned " & swapCandidate & " to row " & emptyRow
                
                Exit Sub ' Exit after successful swap
            Else
                Debug.Print "Swap not possible at row " & r & " due to conflict (dest: " & r & ", orig: " & emptyRow & ")"
            End If
        End If
NextRow:
    Next r
End Sub

' Helper to find the last unfilled slot
Function FindLastUnfilledSlot(startRow As Long, endRow As Long) As Long
    Dim r As Long
    FindLastUnfilledSlot = 0
    For r = startRow To endRow
        If wsRoster.Cells(r, DAY_COL).Value <> "Sat" And _
           wsRoster.Cells(r, MOR_COL).Value = "" Then
            FindLastUnfilledSlot = r
        End If
    Next r
End Function

' Helper to shuffle array
Sub ShuffleArray(arr() As Long)
    Dim i As Long, j As Long, tmp As Long
    Randomize
    For i = UBound(arr) To LBound(arr) + 1 Step -1
        j = Int(Rnd() * (i - LBound(arr) + 1)) + LBound(arr)
        tmp = arr(i)
        arr(i) = arr(j)
        arr(j) = tmp
    Next i
End Sub

Function GetEligibleRows(totalDays As Long, workDays As Variant) As Collection
    Dim eligibleRows As New Collection
    Dim r As Long, j As Long
    Dim dayName As String

    Debug.Print "=== Checking Eligible Rows ==="
    Debug.Print "WorkDays:"
    For j = LBound(workDays) To UBound(workDays)
        Debug.Print "[" & j & "]: " & workDays(j)
    Next j
    
    For r = START_ROW To last_row_roster
        dayName = Trim(wsRoster.Cells(r, DAY_COL).Value)
        ' Debug: show what day we are checking
        Debug.Print "Row " & r & ": " & dayName
        
        ' Skip if already filled
        If Not IsEmpty(wsRoster.Cells(r, MOR_COL)) Then
            Debug.Print "  -> Skipped (Already Assigned)"
            GoTo SkipRow
        End If
        
        ' Check if the day is in workDays
        For j = LBound(workDays) To UBound(workDays)
            If dayName = workDays(j) Then
                eligibleRows.Add r
                Debug.Print "  -> Added (Matched with " & workDays(j) & ")"
                Exit For
            End If
        Next j
        
SkipRow:
    Next r
    Debug.Print "Total Eligible: " & eligibleRows.count
    Set GetEligibleRows = eligibleRows
End Function

Sub IncrementDutiesCounter(staffName As String)
    Dim rowIdx As Variant
    Dim foundCell As Range

    ' Search for the staff name
    Set foundCell = morningtbl.ListColumns("Name").DataBodyRange.Find( _
        What:=staffName, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        ' Get relative row index in the table
        rowIdx = foundCell.row - morningtbl.HeaderRowRange.row

        ' Increment Duties Counter
        With morningtbl.ListRows(rowIdx).Range.Cells(morningtbl.ListColumns("Duties Counter").Index)
            .Value = .Value + 1
        End With
    Else
        MsgBox "Staff '" & staffName & "' not found in table.", vbExclamation
    End If
End Sub

Sub DecrementDutiesCounter(staffName As String)
    Dim rowIdx As Variant
    Dim foundCell As Range

    ' Search for the staff name
    Set foundCell = morningtbl.ListColumns("Name").DataBodyRange.Find( _
        What:=staffName, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        ' Get relative row index in the table
        rowIdx = foundCell.row - morningtbl.HeaderRowRange.row

        ' Decrement Duties Counter
        With morningtbl.ListRows(rowIdx).Range.Cells(morningtbl.ListColumns("Duties Counter").Index)
            .Value = .Value - 1
            If .Value < 0 Then .Value = 0 ' Prevent negative values
        End With
    Else
        MsgBox "Staff '" & staffName & "' not found in table.", vbExclamation
    End If
End Sub

Function IsWorkingOnSameDay(row As Long, staffName As String) As Boolean
    If wsRoster.Cells(row, AFT_COL).Value = staffName Then
        IsWorkingOnSameDay = True
        Exit Function
    End If
    
    If wsRoster.Cells(row, AOH_COL).Value = staffName Then
        IsWorkingOnSameDay = True
        Exit Function
    End If
        
    IsWorkingOnSameDay = False
End Function

' Helper to get the first eligible staff (duties < max)
Function GetFirstEligibleStaff() As String
    Dim i As Long
    For i = 1 To morningtbl.ListRows.count
        Dim staffName As String
        Dim maxDuties As Long
        Dim currDuties As Long
        staffName = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Name").Index).Value
        maxDuties = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Max Duties").Index).Value
        currDuties = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Duties Counter").Index).Value
        If UCase(morningtbl.DataBodyRange(i, morningtbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" And _
           currDuties < maxDuties Then
            GetFirstEligibleStaff = staffName
            Exit Function
        End If
    Next i
    GetFirstEligibleStaff = ""
End Function



