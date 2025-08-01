Attribute VB_Name = "AssignAOHDuties"
' Declare worksheet and table at module level
Private wsRoster As Worksheet
Private wsSettings As Worksheet
Private wsPersonnel As Worksheet
Private aohtbl As ListObject
Private spectbl As ListObject

Sub AssignAOHDuties()
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
    Dim lastTwoWeeksStart As Long
    Dim currDuties As Long
    Dim priorityPool As Collection
    Dim prioritySize As Long
    Dim tmpStaff() As String
    Dim targetWeek As Long
    Dim anyPriorityAssignments As Boolean
    Dim reassignmentAttempts As Long

    ' Set worksheet and table references
    Set wsRoster = Sheets("Roster")
    Set wsSettings = Sheets("Settings")
    Set wsPersonnel = Sheets("AOH PersonnelList")
    Set aohtbl = wsPersonnel.ListObjects("AOHMainList")
    Set spectbl = wsPersonnel.ListObjects("AOHSpecificDaysWorkingStaff")
    
    totalDays = wsRoster.Range(wsRoster.Cells(START_ROW, DATE_COL), wsRoster.Cells(last_row_roster, DATE_COL)).Rows.count
    
    ' Step 1: Assign Specific Days Staff with weekly limit
    For i = 1 To spectbl.ListRows.count
        staffName = spectbl.DataBodyRange(i, spectbl.ListColumns("Name").Index).Value
        Debug.Print "Processing Specific Days staff: " & staffName
        workDays = Split(spectbl.DataBodyRange(i, spectbl.ListColumns("Working Days").Index).Value, ",")
        
        ' Clean up day names (remove spaces)
        For j = 0 To UBound(workDays)
            workDays(j) = Trim(workDays(j))
        Next j
        
        ' Get max duties for this staff from AOHMainList
        For r = 1 To aohtbl.ListRows.count
            If aohtbl.DataBodyRange(r, aohtbl.ListColumns("Name").Index).Value = staffName Then
                maxDuties = aohtbl.DataBodyRange(r, aohtbl.ListColumns("Max Duties").Index).Value
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
        
        ' Assign staff with weekly limit check
        assigned = 0
        For j = 1 To eligibleRows.count
            If assigned >= maxDuties Then Exit For
            r = tmpRows(j)
            If r = 0 Or r < START_ROW Or r > last_row_roster Then
                Debug.Print "Invalid row number " & r & " for staff " & staffName & " at index " & j
                GoTo NextIteration
            End If
            Debug.Print "Considering row " & r & " for " & staffName & " (Shuffled index: " & j & ")"
            
            weekStart = r - (Weekday(wsRoster.Cells(r, DATE_COL).Value, vbMonday) - 1)
            If weekStart < START_ROW Then weekStart = START_ROW
            weekEnd = weekStart + 6
            If weekEnd >= last_row_roster Then weekEnd = last_row_roster - 1
            Debug.Print "  Week boundaries: Start = " & weekStart & ", End = " & weekEnd
            
            dutyCount = 0
            For k = weekStart To weekEnd
                If k >= START_ROW And k < last_row_roster And wsRoster.Cells(k, AOH_COL).Value = staffName And _
                   UCase(Trim(wsRoster.Cells(k, VAC_COL).Value)) = "SEM TIME" Then
                    dutyCount = dutyCount + 1
                End If
            Next k
            Debug.Print "  Current duty count in week: " & dutyCount
            
            If wsRoster.Cells(r, AOH_COL).Value = "" And CheckWeeklyLimit(staffName, r, START_ROW, last_row_roster) Then
                wsRoster.Cells(r, AOH_COL).Value = staffName
                Call IncrementDutiesCounter(staffName)
                assigned = assigned + 1
                Debug.Print "  Assigned " & staffName & " to row " & r & " (Duty count: " & assigned & ")"
            Else
                Debug.Print "  Skipped row " & r & " for " & staffName & " (Limit reached or slot taken)"
            End If
NextIteration:
        Next j
        Debug.Print "Total assigned to " & staffName & ": " & assigned
    Next i
    
    ' Step 2: Assign All Days Staff with weekly limit
    For r = START_ROW To last_row_roster
        If wsRoster.Cells(r, DAY_COL).Value = "Sat" Then GoTo SkipDay
        If wsRoster.Cells(r, AOH_COL).Value = "CLOSED" Then GoTo SkipDay
        ' Check if the day is sem time (not vacation)
        If UCase(Trim(wsRoster.Cells(r, VAC_COL).Value)) <> "SEM TIME" Then GoTo SkipDay
        
        For i = 1 To aohtbl.ListRows.count
            staffName = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Name").Index).Value
            If UCase(aohtbl.DataBodyRange(i, aohtbl.ListColumns("Availability Type").Index).Value) = "SPECIFIC DAYS" Then
                GoTo SkipStaff
            End If
            
            maxDuties = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Max Duties").Index).Value
            currDuties = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Duties Counter").Index).Value
            ' Check if the staff already reached his max duties
            If currDuties >= maxDuties Then GoTo SkipStaff
            
            ' Assign from top with weekly limit check
            If wsRoster.Cells(r, AOH_COL).Value = "" And CheckWeeklyLimit(staffName, r, START_ROW, last_row_roster) Then
                wsRoster.Cells(r, AOH_COL).Value = staffName
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
            Call ReassignAOHDuties
        Else
            Exit Do
        End If
        reassignmentAttempts = reassignmentAttempts + 1
        If reassignmentAttempts > MAX_ATTEMPTS Then
            Debug.Print "Max reassignment attempts reached. Assigning remaining staff to unfilled slots."
            ' Fallback: Assign all remaining eligible staff to unfilled slots
            Dim eligibleStaffList As Collection
            Set eligibleStaffList = New Collection
            For i = 1 To aohtbl.ListRows.count
                staffName = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Name").Index).Value
                maxDuties = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Max Duties").Index).Value
                currDuties = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Duties Counter").Index).Value
                If UCase(aohtbl.DataBodyRange(i, aohtbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" And _
                   currDuties < maxDuties Then
                    eligibleStaffList.Add staffName
                End If
            Next i
            
            Dim unfilledSlotsList As Collection
            Set unfilledSlotsList = New Collection
            For r = START_ROW To last_row_roster
                If wsRoster.Cells(r, DAY_COL).Value <> "Sat" And _
                   wsRoster.Cells(r, AOH_COL).Value = "" And _
                   UCase(Trim(wsRoster.Cells(r, VAC_COL).Value)) = "SEM TIME" Then
                    unfilledSlotsList.Add r
                End If
            Next r
            
            ' Assign each eligible staff to an unfilled slot
            For i = 1 To eligibleStaffList.count
                If i <= unfilledSlotsList.count Then
                    staffName = eligibleStaffList(i)
                    r = unfilledSlotsList(i)
                    wsRoster.Cells(r, AOH_COL).Value = staffName
                    wsRoster.Cells(r, AOH_COL).Interior.Color = vbYellow ' Highlight yellow (RGB 255, 255, 0)
                    Call IncrementDutiesCounter(staffName)
                    Debug.Print "Fallback: Assigned " & staffName & " to row " & r & " (highlighted yellow)"
                End If
            Next i
            Exit Do
        End If
    Loop
    
    MsgBox "Duties assignment completed!", vbInformation
End Sub

' Helper to count unfilled AOH slots
Function CountUnfilledSlots(startRow As Long, endRow As Long) As Long
    Dim r As Long
    Dim count As Long
    count = 0
    For r = startRow To endRow
        If wsRoster.Cells(r, DAY_COL).Value <> "Sat" And _
           wsRoster.Cells(r, AOH_COL).Value = "" And _
           UCase(Trim(wsRoster.Cells(r, VAC_COL).Value)) = "SEM TIME" Then
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
    For i = 1 To aohtbl.ListRows.count
        Dim staffName As String
        Dim maxDuties As Long
        Dim currDuties As Long
        staffName = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Name").Index).Value
        maxDuties = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Max Duties").Index).Value
        currDuties = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Duties Counter").Index).Value
        If UCase(aohtbl.DataBodyRange(i, aohtbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" And _
           currDuties < maxDuties Then
            count = count + 1
        End If
    Next i
    CountEligibleStaff = count
End Function

' Helper to reassign AOH duties with swap logic
Sub ReassignAOHDuties()
    Dim r As Long
    Dim i As Long
    Dim staffName As String
    Dim maxDuties As Long
    Dim currDuties As Long
    Dim eligibleStaff As String
    Dim weekStart As Long, weekEnd As Long
    Dim swapCandidate As String
    Dim swapRow As Long
    
    ' Find the first eligible staff (currDuties < maxDuties)
    eligibleStaff = GetFirstEligibleStaff
    If eligibleStaff = "" Then
        Debug.Print "No eligible staff found for reassignment."
        Exit Sub
    End If
    Debug.Print "Eligible staff for reassignment: " & eligibleStaff
    
    ' Find a week where eligibleStaff has no AOH duties
    Dim hasNoDutyWeek As Boolean
    hasNoDutyWeek = False
    For r = START_ROW To last_row_roster
        If wsRoster.Cells(r, DAY_COL).Value <> "Sat" And _
           UCase(Trim(wsRoster.Cells(r, VAC_COL).Value)) = "SEM TIME" Then
            weekStart = r - (Weekday(wsRoster.Cells(r, DATE_COL).Value, vbMonday) - 1)
            If weekStart < START_ROW Then weekStart = START_ROW
            weekEnd = weekStart + 6
            If weekEnd >= last_row_roster Then weekEnd = last_row_roster - 1
         
            If Not HasAOHDutyInWeek(eligibleStaff, weekStart, weekEnd) Then
                hasNoDutyWeek = True
                Debug.Print "Found week with no duty for " & eligibleStaff & ": Start = " & weekStart & ", End = " & weekEnd
                Exit For
            End If
        End If
    Next r
    
    If Not hasNoDutyWeek Then
        Debug.Print "No week found where " & eligibleStaff & " has no AOH duty."
        Exit Sub
    End If
    
    ' Loop through work all day staff to attempt swapping
    For i = 1 To aohtbl.ListRows.count
        swapCandidate = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Name").Index).Value
        If UCase(aohtbl.DataBodyRange(i, aohtbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" Then
            ' Find an existing AOH duty for swapCandidate in the identified week
            For r = weekStart To weekEnd
                If wsRoster.Cells(r, AOH_COL).Value = swapCandidate And _
                   UCase(Trim(wsRoster.Cells(r, VAC_COL).Value)) = "SEM TIME" Then
                    swapRow = r
                    Debug.Print "Checking swap for " & swapCandidate & " at row " & swapRow
                    
                    ' Check if swapping is possible for eligibleStaff
                    If CheckWeeklyLimit(eligibleStaff, swapRow, START_ROW, last_row_roster) Then
                        ' Find another duty slot for swapCandidate (if any)
                        Dim otherDutyRow As Long
                        otherDutyRow = FindOtherDutySlot(swapCandidate, swapRow, weekStart, weekEnd)
                        If otherDutyRow > 0 Or (aohtbl.DataBodyRange(i, aohtbl.ListColumns("Duties Counter").Index).Value > 1) Then
                            ' Perform swap
                            ' Remove swapCandidate's duty from swapRow
                            wsRoster.Cells(swapRow, AOH_COL).Value = ""
                            Call DecrementDutiesCounter(swapCandidate)
                            
                            ' Assign eligibleStaff to swapCandidate's original slot
                            wsRoster.Cells(swapRow, AOH_COL).Value = eligibleStaff
                            Call IncrementDutiesCounter(eligibleStaff)
                            Debug.Print "Assigned " & eligibleStaff & " to row " & swapRow
                            
                            ' If another duty exists, move swapCandidate there; otherwise, leave unassigned
                            If otherDutyRow > 0 Then
                                wsRoster.Cells(otherDutyRow, AOH_COL).Value = swapCandidate
                                Call IncrementDutiesCounter(swapCandidate)
                                Debug.Print "Reassigned " & swapCandidate & " to row " & otherDutyRow
                            Else
                                Debug.Print "No other duty slot for " & swapCandidate & "; left unassigned"
                            End If
                            Exit Sub ' Exit after successful swap
                        Else
                            Debug.Print "No other duty slot or multiple duties for " & swapCandidate & " to swap"
                        End If
                    Else
                        Debug.Print "Swap not possible: Weekly limit check failed for " & eligibleStaff
                    End If
                End If
            Next r
        End If
    Next i
End Sub

' Helper to find the first unfilled slot
Function FindFirstUnfilledSlot() As Long
    Dim r As Long
    For r = START_ROW To last_row_roster
        If wsRoster.Cells(r, DAY_COL).Value <> "Sat" And _
           wsRoster.Cells(r, AOH_COL).Value = "" And _
           UCase(Trim(wsRoster.Cells(r, VAC_COL).Value)) = "SEM TIME" Then
            FindFirstUnfilledSlot = r
            Exit Function
        End If
    Next r
    FindFirstUnfilledSlot = 0 ' No eligible slot found
End Function

' Helper to find another duty slot for swapCandidate
Function FindOtherDutySlot(staffName As String, avoidRow As Long, weekStart As Long, weekEnd As Long) As Long
    Dim r As Long
    For r = weekStart To weekEnd
        If r <> avoidRow And wsRoster.Cells(r, AOH_COL).Value = staffName And _
           UCase(Trim(wsRoster.Cells(r, VAC_COL).Value)) = "SEM TIME" Then
            FindOtherDutySlot = r
            Exit Function
        End If
    Next r
    FindOtherDutySlot = 0 ' No other duty slot found
End Function

' Helper to check if a staff has an AOH duty in a given week
Function HasAOHDutyInWeek(staffName As String, weekStart As Long, weekEnd As Long) As Boolean
    Dim r As Long
    For r = weekStart To weekEnd
        If (wsRoster.Cells(r, AOH_COL).Value = staffName Or _
            wsRoster.Cells(r, SAT_AOH_COL1).Value = staffName Or _
            wsRoster.Cells(r, SAT_AOH_COL2).Value = staffName) And _
           UCase(Trim(wsRoster.Cells(r, VAC_COL).Value)) = "SEM TIME" Then
            HasAOHDutyInWeek = True
            Exit Function
        End If
    Next r
    HasAOHDutyInWeek = False
End Function

' Helper to get the first eligible staff (duties < max)
Function GetFirstEligibleStaff() As String
    Dim i As Long
    For i = 1 To aohtbl.ListRows.count
        Dim staffName As String
        Dim maxDuties As Long
        Dim currDuties As Long
        staffName = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Name").Index).Value
        maxDuties = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Max Duties").Index).Value
        currDuties = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Duties Counter").Index).Value
        If UCase(aohtbl.DataBodyRange(i, aohtbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" And _
           currDuties < maxDuties Then
            GetFirstEligibleStaff = staffName
            Exit Function
        End If
    Next i
    GetFirstEligibleStaff = ""
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
        If Not IsEmpty(wsRoster.Cells(r, AOH_COL)) Then
            GoTo SkipRow
        End If
        
        ' Check if the day is within sem time
        If UCase(Trim(wsRoster.Cells(r, VAC_COL).Value)) <> "SEM TIME" Then
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
    Debug.Print "Total Eligible: " & eligibleRows.count
    Set GetEligibleRows = eligibleRows
End Function

Sub IncrementDutiesCounter(staffName As String)
    Dim rowIdx As Variant
    Dim foundCell As Range
    Set foundCell = aohtbl.ListColumns("Name").DataBodyRange.Find(What:=staffName, LookIn:=xlValues, LookAt:=xlWhole)
    If Not foundCell Is Nothing Then
        rowIdx = foundCell.row - aohtbl.HeaderRowRange.row
        With aohtbl.ListRows(rowIdx).Range.Cells(aohtbl.ListColumns("Duties Counter").Index)
            .Value = .Value + 1
        End With
    Else
        MsgBox "Staff '" & staffName & "' not found in table.", vbExclamation
    End If
End Sub

Sub DecrementDutiesCounter(staffName As String)
    Dim rowIdx As Variant
    Dim foundCell As Range
    Set foundCell = aohtbl.ListColumns("Name").DataBodyRange.Find(What:=staffName, LookIn:=xlValues, LookAt:=xlWhole)
    If Not foundCell Is Nothing Then
        rowIdx = foundCell.row - aohtbl.HeaderRowRange.row
        With aohtbl.ListRows(rowIdx).Range.Cells(aohtbl.ListColumns("Duties Counter").Index)
            .Value = .Value - 1
            If .Value < 0 Then .Value = 0
        End With
    Else
        MsgBox "Staff '" & staffName & "' not found in table.", vbExclamation
    End If
End Sub

Function CheckWeeklyLimit(staffName As String, rowNum As Long, startRow As Long, endRow As Long) As Boolean
    Dim ws As Worksheet
    Dim i As Long
    Dim weekStart As Long
    Dim weekEnd As Long
    Dim dutyCount As Long
    Dim dateValue As Variant
    
    Set ws = wsRoster
    
    ' Validate rowNum
    If rowNum = 0 Or rowNum < startRow Or rowNum > endRow Then
        Debug.Print "Invalid rowNum " & rowNum & ". Skipping weekly limit check."
        CheckWeeklyLimit = False
        Exit Function
    End If
    
    ' Validate date
    dateValue = ws.Cells(rowNum, DATE_COL).Value
    If IsEmpty(dateValue) Or Not IsDate(dateValue) Then
        Debug.Print "Invalid date at row " & rowNum & ", column " & DATE_COL & ". Skipping weekly limit check."
        CheckWeeklyLimit = False
        Exit Function
    End If
    
    Debug.Print "Weekday: " & Weekday(dateValue, vbMonday)
    weekStart = rowNum - (Weekday(dateValue, vbMonday) - 1)
    If weekStart < startRow Then weekStart = startRow
    weekEnd = weekStart + 6
    If weekEnd >= endRow Then weekEnd = endRow - 1
    
    Debug.Print "Here i am now" & weekStart & weekEnd
    Debug.Print "Row " & rowNum & ": Week Start = " & weekStart & ", End = " & weekEnd
    
    dutyCount = 0
    For i = weekStart To weekEnd
        If i >= startRow And i < endRow And (ws.Cells(i, AOH_COL).Value = staffName Or _
           ws.Cells(i, SAT_AOH_COL1).Value = staffName Or ws.Cells(i, SAT_AOH_COL2).Value = staffName) And _
           UCase(Trim(ws.Cells(i, VAC_COL).Value)) = "SEM TIME" Then
            dutyCount = dutyCount + 1
            If dutyCount >= 1 Then
                CheckWeeklyLimit = False
                Exit Function
            End If
        End If
    Next i
    CheckWeeklyLimit = True
End Function



