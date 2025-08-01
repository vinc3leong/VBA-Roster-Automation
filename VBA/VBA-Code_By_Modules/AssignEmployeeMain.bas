Attribute VB_Name = "AssignEmployeeMain"
' Global variable
Public last_row_roster As Long
Public wsRoster As Worksheet
Public wsSettings As Worksheet

' Declare roster column numbers
Public Const VAC_COL As Long = 1
Public Const DATE_COL As Long = 2
Public Const DAY_COL As Long = 3
Public Const LMB_COL As Long = 4
Public Const MOR_COL As Long = 6
Public Const AFT_COL As Long = 8
Public Const AOH_COL As Long = 10
Public Const SAT_AOH_COL1 As Long = 12
Public Const SAT_AOH_COL2 As Long = 14
Public Const START_ROW As Long = 6

Sub Main()
Attribute Main.VB_ProcData.VB_Invoke_Func = "M\n14"
    Set wsRoster = Sheets("Roster")
    Set wsSettings = Sheets("Settings")
    
    Dim dateRow As Long
    Dim currDate As Date
    Dim slotCol As Variant
    Dim slotCell As Range
    Dim enteredPassword As String
    Const password As String = "rostering2025"
    
    enteredPassword = InputBox("Please enter the password for roster population:", "Password Authentication")
    If enteredPassword <> password Then
        MsgBox "Incorrect password. Fail to populate roster.", vbCritical
        Exit Sub
    End If
    
    'Find last row of roster
    If wsRoster.Cells(2, 10).Value = "Jan-Jun" And wsRoster.Cells(2, 13).Value Mod 4 = 0 Then
        last_row_roster = 187
    ElseIf wsRoster.Cells(2, 10).Value = "Jan-Jun" Then
        last_row_roster = 186
    Else
        last_row_roster = 189
    End If
    
    
    ' Define your table and corresponding sheet names
    sheetNames = Array("Loan Mail Box PersonnelList", "Morning PersonnelList", _
                       "Afternoon PersonnelList", "AOH PersonnelList", "Sat AOH PersonnelList")
    tblNames = Array("LoanMailBoxMainList", "MorningMainList", "AfternoonMainList", _
                     "AOHMainList", "SatAOHMainList")
                     
    ' Ask for confirmation before clearing
    If MsgBox("This action will clear the current roster table. Are you sure you want to proceed?", vbYesNo + vbQuestion, "Confirm Clear") = vbNo Then
        Exit Sub ' Exit Main if user selects No
    End If
    
    wsRoster.Unprotect
        
    ' Clear table content
    wsRoster.Range("D6:O189").ClearContents
    wsRoster.Range("D6:O189").Rows.AutoFit
    
    
    ' Loop through all specified sheets and unprotect the sheets
    For i = 0 To UBound(sheetNames)
        On Error Resume Next ' Enable error handling
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        If Not ws Is Nothing Then
            Debug.Print "Attempting to unprotect sheet: " & sheetNames(i) & " at " & Now
            ws.Unprotect
            If Err.Number = 0 Then
                Debug.Print "Successfully unprotected sheet: " & sheetNames(i) & ws.ProtectContents
            Else
                Debug.Print "Failed to unprotect sheet: " & sheetNames(i) & " - Error: " & Err.Description
                Err.Clear
            End If
        Else
            Debug.Print "Sheet not found: " & sheetNames(i)
        End If
        On Error GoTo 0 ' Disable error handling
    Next i
       
    ' Loop through each date row
    For dateRow = START_ROW To last_row_roster
        currDate = wsRoster.Cells(dateRow, DATE_COL).Value
        ' Reset formatting for all slots
        For Each slotCol In Array(LMB_COL, MOR_COL, AFT_COL, AOH_COL, SAT_AOH_COL1, SAT_AOH_COL2)
            Set slotCell = wsRoster.Cells(dateRow, slotCol)
            slotCell.Interior.ColorIndex = xlNone ' Reset to no fill (default)
            slotCell.Font.Strikethrough = False
        Next slotCol
        
        'Check for Closed date
        If IsClosedDate(currDate) Then
            Call MarkAllSlotsClosed(dateRow)
        End If
        
    Next dateRow
    
    Call ResetAllCounters.ResetAllCounters
    
    Call AssignSatAOHDuties.AssignSatAOHDuties
    Call AssignAOHDuties.AssignAOHDuties
    Call AssignAfternoonDuties.AssignAfternoonDuties
    Call AssignMorningDuties.AssignMorningDuties
    Call AssignLoanMailBoxDuties.AssignLoanMailBoxDuties
    
    Call DuplicateSystemRoster.DuplicateSystemRoster
    
    ' Reprotect the worksheet and lock table ranges
    For i = 0 To UBound(sheetNames)
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        Set tbl = ws.ListObjects(tblNames(i))
        

        With ws
            If Not tbl Is Nothing Then
                .ListObjects(tbl.Name).Range.Locked = True
            End If
            .Range("D5:D9").Locked = False ' keep data entry remains unlocked
            .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                     AllowFiltering:=True, AllowSorting:=True, AllowUsingPivotTables:=True
        End With
    
    Next i
    
    ' Reprotect the Roster sheet with specified properties
    'With wsRoster
    '    .Protect DrawingObjects:=True, Contents:=True, Scenarios:=False, _
    '             AllowFormattingCells:=True, AllowFormattingColumns:=True, _
    '             AllowFormattingRows:=True
    'End With

End Sub

Private Function IsClosedDate(currDate As Date) As Boolean
    IsClosedDate = (Weekday(currDate, vbMonday) = 7) Or _
        Application.WorksheetFunction.CountIf(wsSettings.Range("Settings_Holidays"), currDate) > 0
End Function

Private Sub MarkAllSlotsClosed(dateRow As Long)
    Dim col As Variant
    For Each col In Array(LMB_COL, MOR_COL, AFT_COL, AOH_COL, SAT_AOH_COL1, SAT_AOH_COL2)
        With wsRoster.Cells(dateRow, col)
            .Value = "CLOSED"
            .Interior.Color = vbRed
        End With
    Next col
End Sub

