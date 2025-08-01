Attribute VB_Name = "DuplicateActualRoster"
Sub DuplicateActualRoster()
    Dim srcSheet As Worksheet
    Dim copySheet As Worksheet
    Dim sheetName As String
    Dim enteredPassword As String
    Const password As String = "rostering2025"
    
    enteredPassword = InputBox("Please enter the password to duplicate the roster:", "Password Authentication")
    If enteredPassword <> password Then
        MsgBox "Incorrect password. Duplicate operation declined.", vbCritical
        Exit Sub
    End If
    
    Set srcSheet = Sheets("Roster")
    sheetName = "ActualRoster_" & Format(Now, "yyyymmdd_hhnn")
    
    srcSheet.Copy After:=Sheets(Sheets.count)
    Set copySheet = ActiveSheet
    copySheet.Name = sheetName
    
    'With srcSheet.UsedRange
    '    .Copy Destination:=copySheet.Range("A1")
    'End With
    
    ' Remove all shapes (e.g., buttons, form controls, etc.)
    For Each shp In copySheet.Shapes
        shp.Delete
    Next shp

    With copySheet.Cells
        .Locked = True
    End With
    
    copySheet.Protect password:="nuslib2017@52", _
                        AllowSorting:=True, _
                        AllowFiltering:=True, _
                        AllowFormattingCells:=True
                        
    
    srcSheet.Activate
    
End Sub


