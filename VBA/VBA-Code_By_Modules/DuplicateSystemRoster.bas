Attribute VB_Name = "DuplicateSystemRoster"
Sub DuplicateSystemRoster()
    Dim srcSheet As Worksheet
    Dim copySheet As Worksheet
    Dim sheetName As String
    
    Set wsRoster = Sheets("Roster")
    Set srcSheet = wsRoster
    sheetName = "SystemRoster_" & Format(Now, "yymmdd_hhnn")
    
    'Copy the whole sheet
    srcSheet.Copy After:=Sheets(Sheets.count)
    Set copySheet = ActiveSheet
    copySheet.Name = sheetName
    
    'With srcSheet.UsedRange
    '    copySheet.Range("A1").Resize(.Rows.Count, .Columns.Count).Value = .Value
    '    Application.CutCopyMode = False
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
    
    wsRoster.Activate
    
End Sub






