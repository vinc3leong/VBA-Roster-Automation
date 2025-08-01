Attribute VB_Name = "ReprotectRoster"
Sub ProtectRosterSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Roster")
    
    ' Unprotect first (in case it is already protected)
    On Error Resume Next
    ws.Unprotect password:="rostering2025"
    On Error GoTo 0
    
    ' Lock all cells
    ws.Cells.Locked = True
    
    ' Protect sheet with password
    ws.Protect password:="rostering2025", _
               AllowSorting:=True, _
               AllowFiltering:=True, _
               AllowFormattingCells:=True
               
    MsgBox "Roster sheet has been protected successfully!", vbInformation
End Sub

