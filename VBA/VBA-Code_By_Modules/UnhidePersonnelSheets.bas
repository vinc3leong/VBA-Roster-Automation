Attribute VB_Name = "UnhidePersonnelSheets"
Sub UnhidePersonnelSheetsWithPassword()
    Dim ws As Worksheet
    Dim password As String
    Dim enteredPassword As String
    Dim tbl As ListObject
    Dim specificDaysTbl As ListObject
    
    ' match the one used to hide
    password = "rostering2025"
    
    enteredPassword = InputBox("Please enter the password to unhide all Personnel List sheets:", "Password Authentication")
    If enteredPassword <> password Then
        MsgBox "Incorrect password. Unhide operation declined.", vbCritical
        Exit Sub
    End If
    
    ' Loop through all worksheets to unhide and reprotect the personnel lists
    For Each ws In ThisWorkbook.Sheets
        Select Case UCase(ws.Name)
        
            Case UCase("Loan Mail Box PersonnelList")
                On Error Resume Next
                ws.Unprotect password
                If Err.Number = 0 Then
                    ws.Visible = xlSheetVisible
                    Set tbl = ws.ListObjects("LoanMailBoxMainList")
                    Set specificDaysTbl = ws.ListObjects("LoanMailBoxSpecificDaysWorkingStaff")
                    ws.Protect , DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                               AllowFiltering:=True, AllowSorting:=True, AllowUsingPivotTables:=True
                    With ws
                        If Not tbl Is Nothing Then .ListObjects(tbl.Name).Range.Locked = True
                        If Not specificDaysTbl Is Nothing Then .ListObjects(specificDaysTbl.Name).Range.Locked = True
                        .Range("D5:D9").Locked = False
                    End With
                Else
                    MsgBox "Incorrect password for sheet: " & ws.Name, vbExclamation
                End If
                On Error GoTo 0
                
            Case UCase("Morning PersonnelList")
                On Error Resume Next
                ws.Unprotect password
                If Err.Number = 0 Then
                    ws.Visible = xlSheetVisible
                    Set tbl = ws.ListObjects("MorningMainList")
                    Set specificDaysTbl = ws.ListObjects("MorningSpecificDaysWorkingStaff")
                    ws.Protect , DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                               AllowFiltering:=True, AllowSorting:=True, AllowUsingPivotTables:=True
                    With ws
                        If Not tbl Is Nothing Then .ListObjects(tbl.Name).Range.Locked = True
                        If Not specificDaysTbl Is Nothing Then .ListObjects(specificDaysTbl.Name).Range.Locked = True
                        .Range("D5:D9").Locked = False
                    End With
                Else
                    MsgBox "Incorrect password for sheet: " & ws.Name, vbExclamation
                End If
                On Error GoTo 0
                
            Case UCase("Afternoon PersonnelList")
                On Error Resume Next
                ws.Unprotect password
                If Err.Number = 0 Then
                    ws.Visible = xlSheetVisible
                    Set tbl = ws.ListObjects("AfternoonMainList")
                    Set specificDaysTbl = ws.ListObjects("AfternoonSpecificDaysWorkingStaff")
                    ws.Protect , DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                               AllowFiltering:=True, AllowSorting:=True, AllowUsingPivotTables:=True
                    With ws
                        If Not tbl Is Nothing Then .ListObjects(tbl.Name).Range.Locked = True
                        If Not specificDaysTbl Is Nothing Then .ListObjects(specificDaysTbl.Name).Range.Locked = True
                        .Range("D5:D9").Locked = False
                    End With
                Else
                    MsgBox "Incorrect password for sheet: " & ws.Name, vbExclamation
                End If
                On Error GoTo 0
                
            Case UCase("AOH PersonnelList")
                On Error Resume Next
                ws.Unprotect password
                If Err.Number = 0 Then
                    ws.Visible = xlSheetVisible
                    Set tbl = ws.ListObjects("AOHMainList")
                    Set specificDaysTbl = ws.ListObjects("AOHSpecificDaysWorkingStaff")
                    ws.Protect , DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                               AllowFiltering:=True, AllowSorting:=True, AllowUsingPivotTables:=True
                    With ws
                        If Not tbl Is Nothing Then .ListObjects(tbl.Name).Range.Locked = True
                        If Not specificDaysTbl Is Nothing Then .ListObjects(specificDaysTbl.Name).Range.Locked = True
                        .Range("D5:D9").Locked = False
                    End With
                Else
                    MsgBox "Incorrect password for sheet: " & ws.Name, vbExclamation
                End If
                On Error GoTo 0
                
            Case UCase("Sat AOH PersonnelList")
                On Error Resume Next
                ws.Unprotect password
                If Err.Number = 0 Then
                    ws.Visible = xlSheetVisible
                    Set tbl = ws.ListObjects("SatAOHMainList")
                    Set specificDaysTbl = Nothing ' No specific days table for Sat AOH
                    ws.Protect , DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                               AllowFiltering:=True, AllowSorting:=True, AllowUsingPivotTables:=True
                    With ws
                        If Not tbl Is Nothing Then .ListObjects(tbl.Name).Range.Locked = True
                        .Range("D5:D9").Locked = False
                    End With
                Else
                    MsgBox "Incorrect password for sheet: " & ws.Name, vbExclamation
                End If
                On Error GoTo 0
        End Select
    Next ws
    ThisWorkbook.Sheets("Roster").Activate
    MsgBox "Personnel list sheets have been unhidden and reprotected.", vbInformation
    Exit Sub
End Sub

