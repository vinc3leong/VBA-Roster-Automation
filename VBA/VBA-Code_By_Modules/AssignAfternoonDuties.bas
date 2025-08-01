Attribute VB_Name = "AssignAfternoonDuties"
' Declare worksheet and table
Private wsPersonnel As Worksheet
Private afternoontbl As ListObject
Private spectbl As ListObject

Sub AssignAfternoonDuties()
    ' Variable declarations at the top
    Dim i As Long, j As Long, r As Long, k As Long
    Dim dateCount As Long
    Dim totalDays As Long
    Dim dayName As String
    Dim maxDuties As Long
    Dim staffName As String
    Dim workDays As Variant
    Dim eligibleRows As Collection
    Dim tmpRows() As Long
    Dim assignedCount As Long
    Dim weekStart As Long, weekEnd As Long
    Dim dutyCount As Long
    Dim needsReassignment As Boolean
    Dim lastTwoWeeksStart As Long
    Dim staffPool() As String
    Dim tmpStaff() As String
    Dim poolIndex As Long
    Dim assignedInThisPass As Boolean
    Dim poolSize As Long
    Dim initialSize As Long
    Dim currDuties As Long
    Dim dutiesNeeded As Long
    Dim reassignmentAttempts As Long

    Set wsRoster = Sheets("Roster")
    Set wsSettings = Sheets("Settings")
    Set wsPersonnel = Sheets("Afternoon PersonnelList")
    Set afternoontbl = wsPersonnel.ListObjects("AfternoonMainList")
    Set spectbl = wsPersonnel.ListObjects("AfternoonSpecificDaysWorkingStaff")
    
    totalDays = wsRoster.Range(wsRoster.Cells(START_ROW, DATE_COL), wsRoster.Cells(last_row_roster, DATE_COL)).Rows.count
    Debug.Print "Afternoon assignment starts here"
    
    ' Step 1: Assign Specific Days Staff
    For i = 1 To spectbl.ListRows.count
        staffName = spectbl.DataBodyRange(i, spectbl.ListColumns("Name").Index).Value
        workDays = Split(spectbl.DataBodyRange(i, spectbl.ListColumns("Working Days").Index).Value, ",")
        
        ' Clean up day names (remove spaces)
        For j = 0 To UBound(workDays)
            workDays(j) = Trim(workDays(j))
        Next j
        
        ' Get max duties and department for this staff from AfternoonMainList
        Dim dept As String
        For r = 1 To afternoontbl.ListRows.count
            If afternoontbl.DataBodyRange(r, afternoontbl.ListColumns("Name").Index).Value = staffName Then
                maxDuties = afternoontbl.DataBodyRange(r, afternoontbl.ListColumns("Max Duties").Index).Value
                dept = afternoontbl.DataBodyRange(r, afternoontbl.ListColumns("Department").Index).Value
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
            r = tmpRows(j)
            If Not IsWorkingOnSameDay(r, staffName) And wsRoster.Cells(r, AFT_COL).Value = "" Then
                ' Check vacation constraint
                If UCase(Trim(wsRoster.Cells(r, VAC_COL).Value)) = "VACATION" And UCase(dept) <> "APRM" Then
                    Debug.Print "Skipped " & staffName & " for vacation row " & r & " (not APRM)"
                    GoTo NextRowJ
                End If
                wsRoster.Cells(r, AFT_COL).Value = staffName
                Call IncrementDutiesCounter(staffName)
                assignedCount = assignedCount + 1
                Debug.Print "Assigned Specific Days staff " & staffName & " to row " & r
            End If
NextRowJ:
        Next j
    Next i
    
    ' Step 2: Assign All Days Staff
    For r = START_ROW To last_row_roster
        If wsRoster.Cells(r, DAY_COL).Value = "Sat" Then GoTo SkipDay
        If wsRoster.Cells(r, AFT_COL).Value = "CLOSED" Then GoTo SkipDay
        For i = 1 To afternoontbl.ListRows.count
            staffName = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Name").Index).Value
            If UCase(afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Availability Type").Index).Value) = "SPECIFIC DAYS" Then
                GoTo SkipStaff
            End If
            
            maxDuties = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Max Duties").Index).Value
            currDuties = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Duties Counter").Index).Value
            dept = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Department").Index).Value
            ' Check if the staff already reached his max duties
            If currDuties >= maxDuties Then GoTo SkipStaff
            If IsWorkingOnSameDay(r, staffName) Then GoTo SkipStaff
            
            ' Assign from top with vacation constraint
            If wsRoster.Cells(r, AFT_COL).Value = "" Then
                If UCase(Trim(wsRoster.Cells(r, VAC_COL).Value)) = "VACATION" And UCase(dept) <> "APRM" Then
                    Debug.Print "Skipped " & staffName & " for vacation row " & r & " (not APRM)"
                    GoTo SkipStaff
                End If
                wsRoster.Cells(r, AFT_COL).Value = staffName
                Call IncrementDutiesCounter(staffName)
                Debug.Print "Assigned All Days staff " & staffName & " to row " & r
                Exit For
            End If
        
SkipStaff:
        Next i
        
SkipDay:
        Next r
    
    ' Step 3: Reassign All Days Staff for unfilled slots with reassignment
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
            Call ReassignAfternoonDuties
        Else
            Exit Do
        End If
        reassignmentAttempts = reassignmentAttempts + 1
        If reassignmentAttempts > MAX_ATTEMPTS Then
            Debug.Print "Max reassignment attempts reached. Assigning remaining staff to unfilled slots."
            ' Fallback: Assign all remaining eligible staff to unfilled slots
            Dim eligibleStaffList As Collection
            Set eligibleStaffList = New Collection
            For i = 1 To afternoontbl.ListRows.count
                staffName = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Name").Index).Value
                maxDuties = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Max Duties").Index).Value
                currDuties = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Duties Counter").Index).Value
                dept = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Department").Index).Value
                If UCase(afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" And _
                   currDuties < maxDuties Then
                    eligibleStaffList.Add Array(staffName, dept)
                End If
            Next i
            
            Dim unfilledSlotsList As Collection
            Set unfilledSlotsList = New Collection
            For r = START_ROW To last_row_roster
                If wsRoster.Cells(r, DAY_COL).Value <> "Sat" And _
                   wsRoster.Cells(r, AFT_COL).Value = "" Then
                    unfilledSlotsList.Add r
                End If
            Next r
            
            ' Assign each eligible staff to an unfilled slot
            For i = 1 To eligibleStaffList.count
                If i <= unfilledSlotsList.count Then
                    staffName = eligibleStaffList(i)(0)
                    dept = eligibleStaffList(i)(1)
                    r = unfilledSlotsList(i)
                    If Not IsWorkingOnSameDay(r, staffName) Then
                        If UCase(Trim(wsRoster.Cells(r, VAC_COL).Value)) = "VACATION" And UCase(dept) <> "APRM" Then
                            Debug.Print "Fallback: Skipped " & staffName & " for vacation row " & r & " (not APRM)"
                        Else
                            wsRoster.Cells(r, AFT_COL).Value = staffName
                            wsRoster.Cells(r, AFT_COL).Interior.Color = vbYellow ' Highlight yellow (RGB 255, 255, 0)
                            Call IncrementDutiesCounter(staffName)
                            Debug.Print "Fallback: Assigned " & staffName & " to row " & r & " (highlighted yellow)"
                        End If
                    Else
                        Debug.Print "Fallback: Skipped " & staffName & " at row " & r & " due to AOH conflict"
                    End If
                End If
            Next i
            Exit Do
        End If
    Loop
    
    MsgBox "Afternoon duties assignment completed!", vbInformation
End Sub

' Helper to count unfilled Afternoon slots
Function CountUnfilledSlots(startRow As Long, endRow As Long) As Long
    Dim r As Long
    Dim count As Long
    count = 0
    For r = startRow To endRow
        If wsRoster.Cells(r, DAY_COL).Value <> "Sat" And _
           wsRoster.Cells(r, AFT_COL).Value = "" Then
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
    For i = 1 To afternoontbl.ListRows.count
        Dim staffName As String
        Dim maxDuties As Long
        Dim currDuties As Long
        staffName = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Name").Index).Value
        maxDuties = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Max Duties").Index).Value
        currDuties = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Duties Counter").Index).Value
        If UCase(afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" And _
           currDuties < maxDuties Then
            count = count + 1
        End If
    Next i
    CountEligibleStaff = count
End Function

Sub ReassignAfternoonDuties()
    Dim r As Long
    Dim i As Long
    Dim staffName As String
    Dim maxDuties As Long
    Dim currDuties As Long
    Dim eligibleStaff As String
    Dim swapCandidate As String
    Dim emptyRow As Long ' The empty original slot where eligibleStaff couldn't be assigned
    Dim dept As String
    
    ' Find the first eligible staff (currDuties < maxDuties)
    eligibleStaff = GetFirstEligibleStaff
    If eligibleStaff = "" Then
        Debug.Print "No eligible staff found for Afternoon reassignment."
        Exit Sub
    End If
    Debug.Print "Eligible staff for Afternoon reassignment: " & eligibleStaff
    
    ' Get department of eligibleStaff
    For i = 1 To afternoontbl.ListRows.count
        If afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Name").Index).Value = eligibleStaff Then
            dept = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Department").Index).Value
            Exit For
        End If
    Next i
    
    ' Assume emptyRow is determined earlier (e.g., the last unfilled slot)
    emptyRow = FindLastUnfilledSlot(START_ROW, last_row_roster)
    If emptyRow = 0 Then Exit Sub ' No empty slot to reassign
    
    ' Loop through all rows to find a swap opportunity
    For r = START_ROW To last_row_roster
        If wsRoster.Cells(r, DAY_COL).Value = "Sat" Or _
           wsRoster.Cells(r, AFT_COL).Value = "CLOSED" Or _
           wsRoster.Cells(r, AFT_COL).Value = "" Then
            GoTo NextRow
        End If
        
        swapCandidate = wsRoster.Cells(r, AFT_COL).Value
        If swapCandidate <> "" And swapCandidate <> eligibleStaff Then
            ' Get department of swapCandidate
            Dim swapDept As String
            For i = 1 To afternoontbl.ListRows.count
                If afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Name").Index).Value = swapCandidate Then
                    swapDept = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Department").Index).Value
                    Exit For
                End If
            Next i
            
            ' Check if swap is allowed (no conflict in destination and original slot, respect vacation constraint)
            If Not IsWorkingOnSameDay(r, eligibleStaff) And Not IsWorkingOnSameDay(emptyRow, swapCandidate) Then
                If UCase(Trim(wsRoster.Cells(emptyRow, VAC_COL).Value)) = "VACATION" And UCase(swapDept) <> "APRM" Then
                    Debug.Print "Swap skipped for row " & r & " (swapCandidate " & swapCandidate & " not APRM for vacation slot " & emptyRow & ")"
                    GoTo NextRow
                End If
                ' Perform swap
                wsRoster.Cells(r, AFT_COL).Value = eligibleStaff
                Call IncrementDutiesCounter(eligibleStaff)
                Debug.Print "Assigned " & eligibleStaff & " to row " & r
                
                wsRoster.Cells(emptyRow, AFT_COL).Value = swapCandidate
                Call IncrementDutiesCounter(swapCandidate)
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
           wsRoster.Cells(r, AFT_COL).Value = "" Then
            FindLastUnfilledSlot = r
        End If
    Next r
End Function

' Helper to find the original Afternoon duty slot for a staff (excluding the current row)
Function FindOriginalAfternoonDuty(staffName As String, excludeRow As Long) As Long
    Dim r As Long
    For r = START_ROW To last_row_roster
        If r <> excludeRow And wsRoster.Cells(r, AFT_COL).Value = staffName And _
           wsRoster.Cells(r, DAY_COL).Value <> "Sat" Then
            FindOriginalAfternoonDuty = r
            Exit Function
        End If
    Next r
    FindOriginalAfternoonDuty = 0 ' No original slot found
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

    For r = START_ROW To last_row_roster
        dayName = Trim(wsRoster.Cells(r, DAY_COL).Value)
        
        ' Skip if already filled
        If Not IsEmpty(wsRoster.Cells(r, AFT_COL)) Then
            GoTo SkipRow
        End If
        
        ' Check if the day is in workDays
        For j = LBound(workDays) To UBound(workDays)
            If dayName = workDays(j) Then
                eligibleRows.Add r
                Exit For
            End If
        Next j
        
SkipRow:
    Next r
    Set GetEligibleRows = eligibleRows
End Function

Sub IncrementDutiesCounter(staffName As String)
    Dim rowIdx As Variant
    Dim foundCell As Range
    Set foundCell = afternoontbl.ListColumns("Name").DataBodyRange.Find(What:=staffName, LookIn:=xlValues, LookAt:=xlWhole)
    If Not foundCell Is Nothing Then
        rowIdx = foundCell.row - afternoontbl.HeaderRowRange.row
        With afternoontbl.ListRows(rowIdx).Range.Cells(afternoontbl.ListColumns("Duties Counter").Index)
            .Value = .Value + 1
        End With
    Else
        MsgBox "Staff '" & staffName & "' not found in table.", vbExclamation
    End If
End Sub

Sub DecrementDutiesCounter(staffName As String)
    Dim rowIdx As Variant
    Dim foundCell As Range
    Set foundCell = afternoontbl.ListColumns("Name").DataBodyRange.Find(What:=staffName, LookIn:=xlValues, LookAt:=xlWhole)
    If Not foundCell Is Nothing Then
        rowIdx = foundCell.row - afternoontbl.HeaderRowRange.row
        With afternoontbl.ListRows(rowIdx).Range.Cells(afternoontbl.ListColumns("Duties Counter").Index)
            .Value = .Value - 1
            If .Value < 0 Then .Value = 0 ' Prevent negative values
        End With
    Else
        MsgBox "Staff '" & staffName & "' not found in table.", vbExclamation
    End If
End Sub

Function IsWorkingOnSameDay(row As Long, staffName As String) As Boolean
    ' Check if staff is working on AOH on the same day
    If wsRoster.Cells(row, AOH_COL).Value = staffName Then
        IsWorkingOnSameDay = True
        Exit Function
    End If
    IsWorkingOnSameDay = False
End Function

' Helper to get the first eligible staff (duties < max)
Function GetFirstEligibleStaff() As String
    Dim i As Long
    For i = 1 To afternoontbl.ListRows.count
        Dim staffName As String
        Dim maxDuties As Long
        Dim currDuties As Long
        staffName = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Name").Index).Value
        maxDuties = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Max Duties").Index).Value
        currDuties = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Duties Counter").Index).Value
        If UCase(afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" And _
           currDuties < maxDuties Then
            GetFirstEligibleStaff = staffName
            Exit Function
        End If
    Next i
    GetFirstEligibleStaff = ""
End Function



