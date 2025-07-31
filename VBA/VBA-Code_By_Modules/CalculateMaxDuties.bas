Attribute VB_Name = "CalculateMaxDuties"
Public Sub CalculateMaxDuties(dutyType As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim totalDuties As Long
    Dim totalStaff As Long
    Dim fullDuties As Long
    Dim i As Long
    Dim remaining As Long
    Dim totalAssigned As Long
    Dim dutiesPercentage As Double
    Dim eligibleCount As Long
    Dim eligible100() As Long ' Store the indices of staff with 100% duty
    Dim j As Long
    Dim rounded() As Long

    ' Set worksheet and table based on dutyType
    Select Case UCase(dutyType)
        Case "LOANMAILBOX"
            Set ws = ThisWorkbook.Sheets("Loan Mail Box PersonnelList")
            Set tbl = ws.ListObjects("LoanMailBoxMainList")
        Case "MORNING"
            Set ws = ThisWorkbook.Sheets("Morning PersonnelList")
            Set tbl = ws.ListObjects("MorningMainList")
        Case "AFTERNOON"
            Set ws = ThisWorkbook.Sheets("Afternoon PersonnelList")
            Set tbl = ws.ListObjects("AfternoonMainList")
        Case "AOH"
            Set ws = ThisWorkbook.Sheets("AOH PersonnelList")
            Set tbl = ws.ListObjects("AOHMainList")
        Case "SAT_AOH"
            Set ws = ThisWorkbook.Sheets("Sat AOH PersonnelList")
            Set tbl = ws.ListObjects("SatAOHMainList")
        Case Else
            MsgBox "Invalid duty type. Use 'LoanMailBox', 'Morning', 'Afternoon', 'AOH', or 'Sat_AOH'.", vbExclamation
            Exit Sub
    End Select

    ' Unprotect the worksheet
    On Error Resume Next ' Handle case where sheet is not protected
    ws.Unprotect ' Remove protection (add password if required, e.g., ws.Unprotect "password")
    On Error GoTo 0

    totalStaff = tbl.ListRows.count
    totalDuties = ws.Range("H6").Value
    fullDuties = WorksheetFunction.RoundDown(totalDuties / totalStaff, 0)
    remaining = 0
    eligibleCount = 0
    
    ReDim eligible100(1 To totalStaff)
    ReDim rounded(1 To totalStaff)
    
    ' Calculate initial duties and max cap
    For i = 1 To totalStaff
        dutiesPercentage = tbl.ListRows(i).Range.Cells(GetColumnIndex(tbl, "Duties Percentage (%)")).Value
        
        If dutiesPercentage < 100 Then
            rounded(i) = CLng(fullDuties * (dutiesPercentage / 100))
        Else
            rounded(i) = fullDuties
            ' Mark eligible 100% staff for distribution
            eligibleCount = eligibleCount + 1
            eligible100(eligibleCount) = i
        End If
        
        totalAssigned = totalAssigned + rounded(i)
    Next i
    
    ' Distribute remaining slots to 100% staff
    remaining = totalDuties - totalAssigned
    
    If remaining > 0 Then
        If eligibleCount > 0 Then
            For j = 1 To remaining
                i = eligible100(((j - 1) Mod eligibleCount) + 1) ' Rotate among 100% staff
                rounded(i) = rounded(i) + 1
            Next j
        Else
            MsgBox "No available staff to assign remaining duties for " & dutyType, vbExclamation
        End If
    End If
    
    ' Write results back to sheet
    For i = 1 To totalStaff
        tbl.ListRows(i).Range.Cells(GetColumnIndex(tbl, "Max Duties")).Value = rounded(i)
    Next i
    
    'Reprotects all personnelLists worksheet
    Call ReprotectPersonnelLists.ReprotectPersonnelLists
    Debug.Print "Max Duties calculated for " & dutyType & " with total duties: " & totalDuties & ", total staff: " & totalStaff
End Sub

Private Function GetColumnIndex(tbl As ListObject, columnName As String) As Long
    On Error Resume Next
    GetColumnIndex = tbl.ListColumns(columnName).Index
    If Err.Number <> 0 Then GetColumnIndex = -1
    On Error GoTo 0
End Function

Sub RunMaxDutiesLMB()
    CalculateMaxDuties "LoanMailBox"
End Sub

Sub RunMaxDutiesMorning()
    CalculateMaxDuties "Morning"
End Sub

Sub RunMaxDutiesAfternoon()
    CalculateMaxDuties "Afternoon"
End Sub

Sub RunMaxDutiesAOH()
    CalculateMaxDuties "AOH"
End Sub

Sub RunMaxDutiesSatAOH()
    CalculateMaxDuties "Sat_AOH"
End Sub
