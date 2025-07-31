Attribute VB_Name = "ClearTableContent"
Public Sub ClearTableContent()
    '
    ' ClearTableContent Macro
    '
    ' Ask for confirmation before clearing
    If MsgBox("This action will clear the current roster table. Are you sure you want to clear the content of the roster table ?", vbYesNo + vbQuestion, "Confirm Clear") = vbNo Then
        Exit Sub ' Exit if user selects No
    End If
    
    Range("D6:O189").Select
    Selection.ClearContents
    Selection.Rows.AutoFit
End Sub
