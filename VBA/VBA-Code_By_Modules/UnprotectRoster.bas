Attribute VB_Name = "UnprotectRoster"
Sub UnprotectRosterSheet()
    Dim ws As Worksheet
    Dim userPassword As String
    Const correctPassword As String = "rostering2025"
    
    ' Reference the Roster sheet
    Set ws = ThisWorkbook.Sheets("Roster")
    
    ' Prompt user for password
    userPassword = InputBox("Enter the password to unprotect the Roster sheet:", "Unprotect Roster")
    
    ' If user cancels or leaves it blank, exit
    If Trim(userPassword) = "" Then
        MsgBox "Unprotect action cancelled.", vbExclamation
        Exit Sub
    End If
    
    ' Check if password is correct
    If userPassword <> correctPassword Then
        MsgBox "Incorrect password. The sheet remains protected.", vbCritical
        Exit Sub
    End If
    
    ' Unprotect the sheet
    ws.Unprotect password:=correctPassword
    
    MsgBox "Roster sheet has been successfully unprotected!", vbInformation
    ws.Activate
End Sub

