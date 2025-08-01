Attribute VB_Name = "ReprotectSheet"
Sub ReprotectSheet()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim specificDaysTbl As ListObject
    Dim dutyType As String
    
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    
    ' Determine the duty type based on the active sheet
    Select Case UCase(ws.Name)
        Case UCase("Loan Mail Box PersonnelList")
            dutyType = "LOANMAILBOX"
            Set tbl = ws.ListObjects("LoanMailBoxMainList")
            Set specificDaysTbl = ws.ListObjects("LoanMailBoxSpecificDaysWorkingStaff")
        Case UCase("Morning PersonnelList")
            dutyType = "MORNING"
            Set tbl = ws.ListObjects("MorningMainList")
            Set specificDaysTbl = ws.ListObjects("MorningSpecificDaysWorkingStaff")
        Case UCase("Afternoon PersonnelList")
            dutyType = "AFTERNOON"
            Set tbl = ws.ListObjects("AfternoonMainList")
            Set specificDaysTbl = ws.ListObjects("AfternoonSpecificDaysWorkingStaff")
        Case UCase("AOH PersonnelList")
            dutyType = "AOH"
            Set tbl = ws.ListObjects("AOHMainList")
            Set specificDaysTbl = ws.ListObjects("AOHSpecificDaysWorkingStaff")
        Case UCase("Sat AOH PersonnelList")
            dutyType = "SAT_AOH"
            Set tbl = ws.ListObjects("SatAOHMainList")
            ' No specificDaysTbl for Sat AOH
        Case Else
            MsgBox "This sheet is not a personnel list. Reprotection cancelled.", vbExclamation
            Exit Sub
    End Select
    
    ' Validate tables
    If tbl Is Nothing Then
        MsgBox "Main table not found on '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If

    ' Lock the main table range
    With tbl.Range
        .Locked = True
    End With
    
    ' Lock the specific days table range if it exists
    If Not specificDaysTbl Is Nothing Then
        With specificDaysTbl.Range
            .Locked = True
        End With
    End If
    
    ' Ensure data entry range (D5:D9) remains unlocked
    ws.Range("D5:D9").Locked = False
    
    ' Reprotect the worksheet with password, allowing selection of unlocked cells
    ws.Protect , DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                AllowFiltering:=True, AllowSorting:=True, AllowUsingPivotTables:=True
    
    MsgBox "Main table and specific working day staff table have been reprotected successfully for " & dutyType & ".", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "An error occurred while reprotecting the tables: " & Err.Description, vbCritical
    Exit Sub
End Sub

