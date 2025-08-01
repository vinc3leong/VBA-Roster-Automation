
Option Explicit

' Main subroutine to handle staff swapping
Sub xxSwapStaff()
    Dim wsRoster As Worksheet
    Dim wsPersonnel As Worksheet
    Dim wsSwap As Worksheet
    Dim slotCols As Variant
    Dim dateRange As Range
    Dim oriName As String
    Dim newName As String
    Dim dateCell As Range
    
    InitializeWorksheets wsRoster, wsPersonnel, wsSwap
    GetSwapNames wsSwap, oriName, newName
    ValidateNames oriName, newName
    
    Set dateRange = GetDateRange
    If dateRange Is Nothing Then Exit Sub
    If Not IsValidDateColumn(dateRange) Then Exit Sub
    
    slotCols = Array(6, 8, 10, 12, 14)
    Dim r As Long
    For Each dateCell In dateRange
        r = dateCell.row
        Dim oriNameFound As Boolean
        CheckOriginalNameExists wsRoster, r, slotCols, oriName, oriNameFound
        If Not oriNameFound Then
            DisplayError "Error: " & oriName & " not found in row " & r & ". Swap not allowed.", vbExclamation
        Else
            Dim nameExists As Boolean
            CheckNewNameExists wsRoster, r, slotCols, newName, nameExists
            If nameExists Then
                DisplayError "Error: " & newName & " already exists in row " & r & ". Swap not allowed.", vbExclamation
            Else
                PerformSwap wsRoster, r, slotCols, oriName, newName, wsPersonnel
            End If
        End If
    Next dateCell
    
    MsgBox "Swap completed.", vbInformation
End Sub

' Initialize worksheet references
Private Sub InitializeWorksheets(wsRoster As Worksheet, wsPersonnel As Worksheet, wsSwap As Worksheet)
    Set wsRoster = Sheets("MasterCopy")
    Set wsPersonnel = Sheets("PersonnelList (AOH & Desk)")
    Set wsSwap = Sheets("Swap")
End Sub

' Get original and new staff names from Swap sheet
Private Sub GetSwapNames(wsSwap As Worksheet, oriName As String, newName As String)
    oriName = UCase(Trim(wsSwap.Range("C4").Value))
    newName = UCase(Trim(wsSwap.Range("C5").Value))
End Sub

' Validate that names are not empty
Private Sub ValidateNames(oriName As String, newName As String)
    If Len(oriName) = 0 Then
        MsgBox "Error: Original staff name is empty. Please enter a valid personnel.", vbCritical
        Exit Sub
    End If
    If Len(newName) = 0 Then
        MsgBox "Error: New staff name is empty. Please enter a valid personnel.", vbCritical
        Exit Sub
    End If
End Sub

' Prompt user to select date range and return it
Private Function GetDateRange() As Range
    On Error Resume Next
    Set GetDateRange = Application.InputBox("Select date cells (Column A)", Type:=8)
    On Error GoTo 0
End Function

' Validate that the selected range is from column A (column 1)
Private Function IsValidDateColumn(dateRange As Range) As Boolean
    If Not dateRange.Columns(1).Column = 2 Then
        MsgBox "Please only select dates from Date column.", vbExclamation
        IsValidDateColumn = False
    Else
        IsValidDateColumn = True
    End If
End Function

' Check if the original name exists in the row
Private Sub CheckOriginalNameExists(wsRoster As Worksheet, r As Long, slotCols As Variant, oriName As String, ByRef oriNameFound As Boolean)
    Dim col As Variant
    Dim cellValue As String
    Dim lines() As String
    oriNameFound = False
    For Each col In slotCols
        cellValue = wsRoster.Cells(r, col).Value
        If InStr(cellValue, vbNewLine) > 0 Then
            If UCase(Trim(Split(cellValue, vbNewLine)(0))) = oriName Then
                oriNameFound = True
            End If
        Else
            If UCase(Trim(cellValue)) = oriName Then
                oriNameFound = True
            End If
        End If
        If oriNameFound Then Exit For
    Next col
End Sub

' Check if the new name exists in the same row
Private Sub CheckNewNameExists(wsRoster As Worksheet, r As Long, slotCols As Variant, newName As String, ByRef nameExists As Boolean)
    Dim col As Variant
    Dim cellValue As String
    Dim lines() As String
    nameExists = False
    For Each col In slotCols
        cellValue = wsRoster.Cells(r, col).Value
        If InStr(cellValue, vbNewLine) > 0 Then
            If UCase(Trim(Split(cellValue, vbNewLine)(0))) = newName Then
                nameExists = True
            End If
        Else
            If UCase(Trim(cellValue)) = newName Then
                nameExists = True
            End If
        End If
        If nameExists Then Exit For
    Next col
End Sub

' Display an error message
Private Sub DisplayError(message As String, messageType As VbMsgBoxStyle)
    MsgBox message, messageType
End Sub

' Perform the swap operation for a given row
Private Sub PerformSwap(wsRoster As Worksheet, r As Long, slotCols As Variant, oriName As String, newName As String, wsPersonnel As Worksheet)
    Dim slotCol As Variant
    Dim currentName As String
    Dim lines() As String
    Dim i As Long
    Dim lastRow As Long
    Dim cumulativeLength As Long
    Dim startPos As Integer
    
    For Each slotCol In slotCols
        With wsRoster.Cells(r, slotCol)
            ' Determine the current name based on whether there is a line break
            If InStr(.Value, vbNewLine) > 0 Then
                currentName = Trim(Split(.Value, vbNewLine)(0)) ' First unstriked line for subsequent swaps
            Else
                currentName = Trim(.Value) ' Entire value for initial swap
            End If
            
            If UCase(currentName) = oriName Then ' Check the current name
                ' Add new name first (unstriked) and preserve existing content
                .Value = newName & vbNewLine & .Value
                .VerticalAlignment = xlTop ' Align text to the top
                .WrapText = True
                
                ' Split into lines to apply strikethrough to all previous names
                lines = Split(.Value, vbNewLine)
                cumulativeLength = Len(newName) + 2 ' Start with newName and its vbNewLine
                
                ' Apply strikethrough to all lines except the first one
                For i = 1 To UBound(lines)
                    startPos = cumulativeLength
                    .Characters(startPos, Len(lines(i)) + 1).Font.Strikethrough = True
                    cumulativeLength = cumulativeLength + Len(lines(i)) + 2 ' Update for next line
                Next i
                
                ' Explicitly increase row height by 15 points per swap
                .RowHeight = .RowHeight + 15
                
                ' Update personnel counter for the new staff
                lastRow = wsPersonnel.Cells(wsPersonnel.Rows.count, "B").End(xlUp).row
                ' Deduct duties from the old staff
                For i = 12 To lastRow
                    If UCase(Trim(wsPersonnel.Cells(i, 2).Value)) = oriName Then
                        wsPersonnel.Cells(i, 5).Value = wsPersonnel.Cells(i, 5).Value - 1 ' Decrement Weekly Duties Counter
                        If slotCol = 10 Or slotCol = 12 Or slotCol = 14 Then ' AOH slots
                            wsPersonnel.Cells(i, 6).Value = wsPersonnel.Cells(i, 6).Value - 1 ' Decrement AOH Counter
                        End If
                        Exit For
                    End If
                Next i
                ' Update duties for the new staff
                For i = 12 To lastRow
                    If UCase(Trim(wsPersonnel.Cells(i, 2).Value)) = newName Then
                        wsPersonnel.Cells(i, 5).Value = wsPersonnel.Cells(i, 5).Value + 1 ' Increment Weekly Duties Counter
                        If slotCol = 10 Or slotCol = 12 Or slotCol = 14 Then ' AOH slots
                            wsPersonnel.Cells(i, 6).Value = wsPersonnel.Cells(i, 6).Value + 1 ' Increment AOH Counter
                        End If
                        Exit For
                    End If
                Next i
            End If
        End With
    Next slotCol
End Sub
Sub AssignFirstEmployeeToFirstSlot()
    Dim wsRoster As Worksheet
    Dim wsPersonnel As Worksheet
    Dim slotCols As Variant
    Dim slotCol As Variant
    Dim slotCell As Range
    Dim staffName As String
    Dim maxDuties As Long
    Dim currDuties As Long 'weekly duties
    Dim currAOH As Long
    Dim lastRow As Long
    Dim currRow As Long
    Dim found As Boolean
    Dim dateRow As Long
    Dim currDateSlotRange As Range
    Dim isAohSlot As Boolean
    Dim alreadyAssigned As Boolean 'already assigned on current day
    Dim canAssign As Boolean


    ' Set references to sheets
    Set wsRoster = Sheets("Master")
    Set wsPersonnel = Sheets("PersonnelList (AOH & Desk)")

    ' Find last row number of the employee list
    lastRow = wsPersonnel.Cells(wsPersonnel.Rows.count, "B").End(xlUp).row
    found = False
    
    ' Array of columns for Morning, Afternoon, AOH
    slotCols = Array(6, 8, 10)  ' F, H, J columns
    
    'Find current date and slot range
    
     ' Loop through each date row
     For dateRow = 6 To 186
        ' Loop through each slot column for this date
        For Each slotCol In slotCols
            Set slotCell = wsRoster.Cells(dateRow, slotCol)
            isAohSlot = (slotCol = 10) ' 10 is column J
            found = False
            
            'Loop through each staff
            For currRow = 12 To lastRow
                staffName = wsPersonnel.Cells(currRow, "B").Value
                maxDuties = wsPersonnel.Cells(currRow, "D").Value
                currDuties = wsPersonnel.Cells(currRow, "E").Value
                currAOH = wsPersonnel.Cells(currRow, "F").Value
                
                'Check if this staff already assigned today
                alreadyAssigned = False
                Set currDateSlotRange = wsRoster.Range("F" & dateRow & ":J" & dateRow)
                For Each cell In currDateSlotRange
                    If cell.Value = staffName Then
                        alreadyAssigned = True
                        Exit For
                    End If
                Next cell
                
                'Determine the criteria
                If isAohSlot Then
                    canAssign = (currAOH < 1) And (currDuties < maxDuties) And Not alreadyAssigned
                Else
                    canAssign = (currDuties < maxDuties) And Not alreadyAssigned
                End If
                    
                'Assign the staff and do counter increment if meet th criteria
                If canAssign Then
                    'Assign staff to a slot
                    slotCell.Value = staffName
                    
                    'Do increment
                    If isAohSlot Then
                        wsPersonnel.Cells(currRow, "F").Value = currAOH + 1
                    End If
                    wsPersonnel.Cells(currRow, "E").Value = currDuties + 1
                    
                    found = True
                    Exit For
                End If
            Next currRow
                
            ' If no staff found who can still take duties
            If Not found Then
                slotCell.Value = "Not Available"
            End If
            
        Next slotCol
    Next dateRow
    
    MsgBox "Roster filled"
    
End Sub

'declare worksheet and table
    Private wsRosterCopy As Worksheet
    Private wsPersonnel As Worksheet
    Private wsSettings As Worksheet
    Private morningtbl As ListObject
    
'declare roster column number
    Private dateCol As Long
    Private dayCol As Long
    Private LMBCol As Long
    Private morCol As Long
    Private aftCol As Long
    Private AOHCol As Long
    Private satAOHCol1 As Long
    Private satAOHCol2 As Long
    
Sub xxAssignFirstEmployeeToFirstSlotCopy()
    Set wsRosterCopy = Sheets("MasterCopy")
    Set wsPersonnel = Sheets("PersonnelList (AOH & Desk)")
    Set wsSettings = Sheets("Settings")

    dateCol = 2
    dayCol = 3
    LMBCol = 4
    morCol = 6
    aftCol = 8
    AOHCol = 10
    satAOHCol1 = 12
    satAOHCol2 = 14
    
    Dim slotCols As Variant
    Dim slotCol As Variant
    Dim slotCell As Range
    Dim staffName As String
    Dim maxDuties As Long
    Dim currDuties As Long 'weekly duties
    Dim currAOH As Long
    Dim lastRow As Long
    Dim currRow As Long
    Dim found As Boolean
    Dim dateRow As Long
    Dim currDateSlotRange As Range
    Dim isAohSlot As Boolean
    Dim alreadyAssigned As Boolean 'already assigned on current day
    Dim canAssign As Boolean
    Dim currDate As Date
    Dim isSaturday As Boolean
    Dim isVacation As Boolean
    Dim lastRowRoster As Integer

    
    ' Find last row number of the employee list
    lastRow = wsPersonnel.Cells(wsPersonnel.Rows.count, "B").End(xlUp).row
    found = False
    
    If wsRosterCopy.Cells(2, 10).Value = "Jan-Jun" And wsRosterCopy.Cells(2, 13).Value Mod 4 = 0 Then
        lastRowRoster = 187
    ElseIf wsRosterCopy.Cells(2, 10).Value = "Jan-Jun" Then
        lastRowRoster = 186
    Else
        lastRowRoster = 189
    End If
    
    
     'Loop through each date row
     For dateRow = 6 To lastRowRoster
     
        currDate = wsRosterCopy.Cells(dateRow, dateCol).Value
        
        If Weekday(currDate, vbMonday) = 7 Or _
            Application.WorksheetFunction.CountIf(wsSettings.Range("Settings_Holidays"), currDate) > 0 Then
            
            ' Skip this date by marking all slots as "CLOSED"
            wsRosterCopy.Cells(dateRow, LMBCol).Value = "CLOSED" ' D column
            wsRosterCopy.Cells(dateRow, LMBCol).Interior.Color = vbRed
            
            wsRosterCopy.Cells(dateRow, morCol).Value = "CLOSED" ' F column
            wsRosterCopy.Cells(dateRow, morCol).Interior.Color = vbRed
            
            wsRosterCopy.Cells(dateRow, aftCol).Value = "CLOSED" ' H column
            wsRosterCopy.Cells(dateRow, aftCol).Interior.Color = vbRed
            
            wsRosterCopy.Cells(dateRow, AOHCol).Value = "CLOSED" ' J column
            wsRosterCopy.Cells(dateRow, AOHCol).Interior.Color = vbRed
            
            wsRosterCopy.Cells(dateRow, satAOHCol1).Value = "CLOSED" ' L column
            wsRosterCopy.Cells(dateRow, satAOHCol1).Interior.Color = vbRed
            
            wsRosterCopy.Cells(dateRow, satAOHCol2).Value = "CLOSED" ' N column
            wsRosterCopy.Cells(dateRow, satAOHCol2).Interior.Color = vbRed
            GoTo NextDate ' Skip to the next date
        End If
        
        For Each slotCol In Array(LMBCol, morCol, aftCol, AOHCol, satAOHCol1, satAOHCol2) ' D, F, H, J, L, N columns
            Set slotCell = wsRosterCopy.Cells(dateRow, slotCol)
            slotCell.Interior.ColorIndex = xlNone ' Reset to no fill (default)
            slotCell.Font.Strikethrough = False
        Next slotCol
        
        isSaturday = (Weekday(currDate, vbMonday) = 6)
        
        isVacation = (wsRosterCopy.Cells(dateRow, 1).Value = "Vacation")
        
        If isSaturday Then
            slotCols = Array(satAOHCol1, satAOHCol2) ' L, N for Saturday
        ElseIf isVacation Then
            slotCols = Array(6, 8) ' F, H only for vacation weekdays (no J AOH)
        Else
            slotCols = Array(6, 8, 10) ' F, H, J for Sem Time weekdays
        End If
            
        
        'ResetAOHCounter.ResetAOHCounter
        
        ' Loop through each slot column for this date
        For Each slotCol In slotCols
            Set slotCell = wsRosterCopy.Cells(dateRow, slotCol)
            isAohSlot = (slotCol = 10 Or isSaturday) And Not isVacation ' J, L, or N as AOH
            found = Falsez
            
            'Loop through each staff
            For currRow = 12 To lastRow
                staffName = wsPersonnel.Cells(currRow, "B").Value
                maxDuties = wsPersonnel.Cells(currRow, "D").Value
                currDuties = wsPersonnel.Cells(currRow, "E").Value
                currAOH = wsPersonnel.Cells(currRow, "F").Value
                
                'Check if this staff already assigned today
                alreadyAssigned = False
                If isSaturday Then
                    Set currDateSlotRange = wsRosterCopy.Range("L" & dateRow & ":N" & dateRow)
                ElseIf isVacation Then
                    Set currDateSlotRange = wsRosterCopy.Range("F" & dateRow & ":H" & dateRow)
                Else
                    Set currDateSlotRange = wsRosterCopy.Range("F" & dateRow & ":J" & dateRow)
                End If
                
                For Each cell In currDateSlotRange
                    If cell.Value = staffName Then
                        alreadyAssigned = True
                        Exit For
                    End If
                Next cell
                
                'Determine the criteria
                If isAohSlot Then
                    canAssign = (currAOH < 1) And (currDuties < maxDuties) And Not alreadyAssigned
                Else
                    canAssign = (currDuties < maxDuties) And Not alreadyAssigned
                End If
                    
                'Assign the staff and do counter increment if meet th criteria
                If canAssign Then
                    'Assign staff to a slot
                    slotCell.Value = staffName
                    
                    'Do increment
                    If isAohSlot Then
                        wsPersonnel.Cells(currRow, "F").Value = currAOH + 1
                    End If
                    wsPersonnel.Cells(currRow, "E").Value = currDuties + 1
                    
                    found = True
                    Exit For
                End If
            Next currRow
                
            ' If no staff found who can still take duties
            If Not found Then
                slotCell.Value = "Not Available"
            End If
            
        Next slotCol
        
NextDate:
    Next dateRow
    
    MsgBox "Roster filled"
    
End Sub

Public Sub ResetAOHCounter()
    Dim ws As Worksheet
    Dim i As Long
    Dim lastRow As Long
    Dim isAllOne As Boolean

    Set ws = ThisWorkbook.Sheets("PersonnelList (AOH & Desk)")

    lastRow = ws.Cells(ws.Rows.count, "B").End(xlUp).row
    isAllOne = True

    For i = 12 To lastRow
        If ws.Cells(i, 6).Value <> 1 Then
            isAllOne = False
            
            Exit For
        End If
    Next i

    ' Reset if all have AOH = 1
    If isAllOne Then
        For i = 12 To lastRow
            ws.Cells(i, 6).Value = 0
        Next i
    End If
End Sub

Sub ResetDutiesAOHCounter()
'
' Reset_Duties_AOH_Counter Macro
'
    Sheets("PersonnelList (AOH & Desk)").Select
    Range("E12").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("F12").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("E12:F12").Select
    Selection.AutoFill Destination:=Range( _
        "Desk_PersonnelList[[Weekly Duties Counter]:[AOH Counter]]")
    Sheets("MasterCopy").Select
End Sub

Public Sub GitSave()
    DeleteAndMake
    ExportModules
    PrintAllCode
    PrintAllContainers
End Sub

Public Sub DeleteAndMake()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim desktopPath As String: desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    Dim parentFolder As String: parentFolder = desktopPath & "\VBA"
    Dim childA As String: childA = parentFolder & "\VBA-Code_Together"
    Dim childB As String: childB = parentFolder & "\VBA-Code_By_Modules"
    
    Debug.Print "Parent Folder Path: " & parentFolder
    
    On Error Resume Next
    fso.DeleteFolder parentFolder, True
    On Error GoTo 0
    
    On Error GoTo MkDirError
    If Not fso.FolderExists(parentFolder) Then MkDir parentFolder
    If Not fso.FolderExists(childA) Then MkDir childA
    If Not fso.FolderExists(childB) Then MkDir childB
    On Error GoTo 0
    Exit Sub

MkDirError:
    MsgBox "Error creating directory: " & Err.Description & " (Path: " & parentFolder & ")", vbCritical
    Exit Sub
End Sub

Public Sub PrintAllCode()
    Dim item As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        If Not item.codeModule Is Nothing Then
            Dim lineCount As Long
            lineCount = item.codeModule.CountOfLines
            If lineCount > 0 Then
                lineToPrint = item.codeModule.lines(1, lineCount)
                Debug.Print lineToPrint
                textToPrint = textToPrint & vbCrLf & lineToPrint
            Else
                Debug.Print item.Name & " has no code lines"
            End If
        Else
            Debug.Print item.Name & " has no accessible code module"
        End If
    Next item
    
    Dim pathToExport As String: pathToExport = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\VBA\VBA-Code_Together\"
    If Dir(pathToExport) <> "" Then Kill pathToExport & "*.*"
    SaveTextToFile textToPrint, pathToExport & "all_code.vb"
End Sub

Public Sub PrintAllContainers()
    Dim item As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        lineToPrint = item.Name
        Debug.Print lineToPrint
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    Dim pathToExport As String: pathToExport = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\VBA\VBA-Code_Together\"
    SaveTextToFile textToPrint, pathToExport & "all_modules.vb"
End Sub

Public Sub ExportModules()
    Dim pathToExport As String: pathToExport = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\VBA\VBA-Code_By_Modules\"
    
    If Dir(pathToExport) <> "" Then
        Kill pathToExport & "*.*"
    End If
     
    Dim wkb As Workbook: Set wkb = Excel.Workbooks(ThisWorkbook.Name)
    
    Dim unitsCount As Long
    Dim filePath As String
    Dim component As Object
    Dim tryExport As Boolean

    For Each component In wkb.VBProject.VBComponents
        tryExport = True
        filePath = component.Name
       
        Select Case component.Type
            Case 3 ' vbext_ct_ClassModule or vbext_ct_MSForm
                filePath = filePath & ".cls"
            Case 1 ' vbext_ct_StdModule
                filePath = filePath & ".bas"
            Case 100 ' vbext_ct_Document
                tryExport = False
        End Select
        
        If tryExport Then
            Debug.Print unitsCount & " exporting " & filePath
            component.Export pathToExport & "\" & filePath
        End If
    Next

    Debug.Print "Exported at " & pathToExport
End Sub

Public Sub SaveTextToFile(dataToPrint As String, pathToExport As String)
    Dim fileSystem As Object
    Dim textObject As Object
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set textObject = fileSystem.CreateTextFile(pathToExport, True)
    
    textObject.WriteLine dataToPrint
    textObject.Close
        
    On Error GoTo 0
    Exit Sub

CreateLogFile_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateLogFile of Sub mod_TDD_Export"
End Sub

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
'@TestModule
'@Folder("Tests")


Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrHandler
    
    Dim table As ListObject
    Dim dutiesCounterCol As ListColumn

    Set table = Me.ListObjects("MorningMainList")
    
    ' Check if the table exists
    If table Is Nothing Then
        Exit Sub
    End If
    
    ' Get the Duties Counter column
    Set dutiesCounterCol = table.ListColumns("Duties Counter")
    If dutiesCounterCol Is Nothing Then
        MsgBox "Column 'Duties Counter' not found in 'MorningMainList'.", vbExclamation
        Exit Sub
    End If
    
    ' Exit if the table is empty, but allow D7 change processing
    If table.ListRows.count = 0 And Not Intersect(Target, Me.Range("D7")) Is Nothing Then
        GoTo ProcessD7Change
    ElseIf table.ListRows.count = 0 Then
        Exit Sub
    End If
    
    ' Sort duties counter by ascending if change is in Duties Counter column
    If Not Intersect(Target, dutiesCounterCol.DataBodyRange) Is Nothing Then
        With table.Sort
            .SortFields.Clear
            .SortFields.Add Key:=dutiesCounterCol.DataBodyRange, _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Header = xlYes
            .Apply
        End With
    End If
    
ProcessD7Change:
    ' Sets Working Days for All Days and show form for Specific Days
    If Not Intersect(Target, Me.Range("D7")) Is Nothing Then
        Dim availabilityType As String
        availabilityType = UCase(Trim(Me.Range("D7").Value))
        
        ' Debug: Log the change
        Debug.Print "Worksheet_Change triggered at " & Now & " for D7: " & availabilityType
        
        Select Case availabilityType
            Case "ALL DAYS"
                ' Set percentage to 100%
                Me.Range("D9").Value = 100
                Me.Range("D8").Value = "Mon, Tue, Wed, Thu, Fri, Sat"
                Debug.Print "Set D9 to 100% for All Days"
                
            Case "SPECIFIC DAYS"
                ' Show the multiselect form
                With frmSpecificDays
                    .targetSheetName = Me.Name
                    .Show
                End With
                Debug.Print "Showed form for Specific Days"
                
            Case Else
                ' Clear D9 if invalid selection
                Me.Range("D9").Value = ""
                Debug.Print "Invalid selection, D9 cleared"
        End Select
    End If
    
    Exit Sub

ErrHandler:
    MsgBox "An error occurred: " & Err.Description & vbCrLf & _
           "Line: " & Erl, vbCritical
    Exit Sub
End Sub

Public targetSheetName As String

Private Sub chkMon_Click()

End Sub



Private Sub chkWed_Click()

End Sub

Private Sub commandCancel_Click()
    Dim currentDays As String
    currentDays = Trim(ThisWorkbook.Sheets(targetSheetName).Range("D8").Value)
    
    If currentDays = "" Then
        MsgBox "Please select at least one day before proceeding.", vbExclamation
    Else
        ' Keep existing selection and hide the form
        Me.Hide
    End If
End Sub
Private Sub commandOK_Click()
    Dim selectedDays As String
    Dim firstDay As Boolean
    
    firstDay = True
    If chkMon.Value Then
        selectedDays = "Mon"
        firstDay = False
    End If
    If chkTues.Value Then
        If Not firstDay Then selectedDays = selectedDays & ", "
        selectedDays = selectedDays & "Tue"
        firstDay = False
    End If
    If chkWed.Value Then
        If Not firstDay Then selectedDays = selectedDays & ", "
        selectedDays = selectedDays & "Wed"
        firstDay = False
    End If
    If chkThurs.Value Then
        If Not firstDay Then selectedDays = selectedDays & ", "
        selectedDays = selectedDays & "Thu"
        firstDay = False
    End If
    If chkFri.Value Then
        If Not firstDay Then selectedDays = selectedDays & ", "
        selectedDays = selectedDays & "Fri"
        firstDay = False
    End If
    
    If selectedDays = "" Then
        MsgBox "Please select at least one day.", vbExclamation
    Else
        ThisWorkbook.Sheets(targetSheetName).Range("D8").Value = selectedDays
        Me.Hide
    End If
End Sub

Private Sub UserForm_Click()

End Sub

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
' Declare worksheet and table
Private wsRoster As Worksheet
Private wsSettings As Worksheet
Private wsPersonnel As Worksheet
Private morningtbl As ListObject
Private spectbl As ListObject

Sub AssignMorningDuties()
    Set wsRoster = Sheets("Roster")
    Set wsSettings = Sheets("Settings")
    Set wsPersonnel = Sheets("Morning PersonnelList")
    Set morningtbl = wsPersonnel.ListObjects("MorningMainList")
    Set spectbl = wsPersonnel.ListObjects("MorningSpecificDaysWorkingStaff")
    
    Dim i As Long, j As Long, r As Long
    Dim dateCount As Long
    Dim totalDays As Long
    Dim dayName As String
    Dim maxDuties As Long
    Dim candidates() As String
    Dim staffName As String
    Dim workDays As Variant
    Dim eligibleRows As Collection
    Dim tmpRows() As Long
    Dim assignedCount As Long
    Dim reassignmentAttempts As Long

    totalDays = wsRoster.Range(wsRoster.Cells(START_ROW, DATE_COL), wsRoster.Cells(last_row_roster, DATE_COL)).Rows.count
    Debug.Print "Morning assignment starts here"
    
    ' Step 1: Assign Specific Days Staff
    For i = 1 To spectbl.ListRows.count
        staffName = spectbl.DataBodyRange(i, spectbl.ListColumns("Name").Index).Value
        Debug.Print staffName
        workDays = Split(spectbl.DataBodyRange(i, spectbl.ListColumns("Working Days").Index).Value, ",")
        
        ' Clean up day names (remove spaces)
        For j = 0 To UBound(workDays)
            workDays(j) = Trim(workDays(j))
        Next j
        
        ' Get max duties for this staff from MorningMainList
        For r = 1 To morningtbl.ListRows.count
            If morningtbl.DataBodyRange(r, morningtbl.ListColumns("Name").Index).Value = staffName Then
                maxDuties = morningtbl.DataBodyRange(r, morningtbl.ListColumns("Max Duties").Index).Value
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
        
        ' Assign staff
        assignedCount = 0
        For j = 1 To eligibleRows.count
            If assignedCount >= maxDuties Then Exit For
            If Not IsWorkingOnSameDay(tmpRows(j), staffName) Then
                wsRoster.Cells(tmpRows(j), MOR_COL).Value = staffName
                Call IncrementDutiesCounter(staffName)
                assignedCount = assignedCount + 1
            End If
        Next j
    Next i
    
    ' Step 2: Assign All Days Staff
    For r = START_ROW To last_row_roster
        If wsRoster.Cells(r, DAY_COL).Value = "Sat" Then GoTo SkipDay
        If wsRoster.Cells(r, MOR_COL).Value = "CLOSED" Then GoTo SkipDay
        For i = 1 To morningtbl.ListRows.count
            staffName = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Name").Index).Value
            If UCase(morningtbl.DataBodyRange(i, morningtbl.ListColumns("Availability Type").Index).Value) = "SPECIFIC DAYS" Then
                GoTo SkipStaff
            End If
            
            maxDuties = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Max Duties").Index).Value
            Dim currDuties As Long
            currDuties = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Duties Counter").Index).Value
            ' Check if the staff already reached his max duties
            If currDuties >= maxDuties Then GoTo SkipStaff
            If IsWorkingOnSameDay(r, staffName) Then GoTo SkipStaff
            
            ' Assign from top
            If wsRoster.Cells(r, MOR_COL).Value = "" Then
                wsRoster.Cells(r, MOR_COL).Value = staffName
                Call IncrementDutiesCounter(staffName)
                Exit For
            End If
        
SkipStaff:
        Next i
        
SkipDay:
        Next r
    
    ' Step 3: Reassign for unfilled slots
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
            Call ReassignMorningDuties
        Else
            Exit Do
        End If
        reassignmentAttempts = reassignmentAttempts + 1
        If reassignmentAttempts > MAX_ATTEMPTS Then
            Debug.Print "Max reassignment attempts reached. Assigning remaining staff to unfilled slots."
            ' Fallback: Assign all remaining eligible staff to unfilled slots
            Dim eligibleStaffList As Collection
            Set eligibleStaffList = New Collection
            For i = 1 To morningtbl.ListRows.count
                staffName = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Name").Index).Value
                maxDuties = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Max Duties").Index).Value
                currDuties = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Duties Counter").Index).Value
                If UCase(morningtbl.DataBodyRange(i, morningtbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" And _
                   currDuties < maxDuties Then
                    eligibleStaffList.Add staffName
                End If
            Next i
            
            Dim unfilledSlotsList As Collection
            Set unfilledSlotsList = New Collection
            For r = START_ROW To last_row_roster
                If wsRoster.Cells(r, DAY_COL).Value <> "Sat" And _
                   wsRoster.Cells(r, MOR_COL).Value = "" Then
                    unfilledSlotsList.Add r
                End If
            Next r
            
            ' Assign each eligible staff to an unfilled slot
            For i = 1 To eligibleStaffList.count
                If i <= unfilledSlotsList.count Then
                    staffName = eligibleStaffList(i)
                    r = unfilledSlotsList(i)
                    If Not IsWorkingOnSameDay(r, staffName) Then
                        wsRoster.Cells(r, MOR_COL).Value = staffName
                        wsRoster.Cells(r, MOR_COL).Interior.Color = vbYellow ' Highlight yellow (RGB 255, 255, 0)
                        Call IncrementDutiesCounter(staffName)
                        Debug.Print "Fallback: Assigned " & staffName & " to row " & r & " (highlighted yellow)"
                    Else
                        Debug.Print "Fallback: Skipped " & staffName & " at row " & r & " due to AOH/AFT conflict"
                    End If
                End If
            Next i
            Exit Do
        End If
    Loop
    
    MsgBox "Duties assignment completed!", vbInformation
End Sub

' Helper to count unfilled morning slots
Function CountUnfilledSlots(startRow As Long, endRow As Long) As Long
    Dim r As Long
    Dim count As Long
    count = 0
    For r = startRow To endRow
        If wsRoster.Cells(r, DAY_COL).Value <> "Sat" And _
           wsRoster.Cells(r, MOR_COL).Value = "" Then
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
    For i = 1 To morningtbl.ListRows.count
        Dim staffName As String
        Dim maxDuties As Long
        Dim currDuties As Long
        staffName = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Name").Index).Value
        maxDuties = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Max Duties").Index).Value
        currDuties = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Duties Counter").Index).Value
        If UCase(morningtbl.DataBodyRange(i, morningtbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" And _
           currDuties < maxDuties Then
            count = count + 1
        End If
    Next i
    CountEligibleStaff = count
End Function
Sub ReassignMorningDuties()
    Dim r As Long
    Dim i As Long
    Dim staffName As String
    Dim maxDuties As Long
    Dim currDuties As Long
    Dim eligibleStaff As String
    Dim swapCandidate As String
    Dim emptyRow As Long ' The empty original slot where eligibleStaff couldn't be assigned
    
    ' Find the first eligible staff (currDuties < maxDuties)
    eligibleStaff = GetFirstEligibleStaff
    If eligibleStaff = "" Then
        Debug.Print "No eligible staff found for Morning reassignment."
        Exit Sub
    End If
    Debug.Print "Eligible staff for Morning reassignment: " & eligibleStaff
    
    ' Assume emptyRow is determined earlier (e.g., the last unfilled slot)
    emptyRow = FindLastUnfilledSlot(START_ROW, last_row_roster)
    If emptyRow = 0 Then Exit Sub ' No empty slot to reassign
    
    ' Loop through all rows to find a swap opportunity
    For r = START_ROW To last_row_roster
        If wsRoster.Cells(r, DAY_COL).Value = "Sat" Or _
           wsRoster.Cells(r, MOR_COL).Value = "CLOSED" Or _
           wsRoster.Cells(r, MOR_COL).Value = "" Then
            GoTo NextRow
        End If
        
        swapCandidate = wsRoster.Cells(r, MOR_COL).Value
        If swapCandidate <> "" And swapCandidate <> eligibleStaff Then
            ' Check if swap is allowed (no conflict in destination and original slot)
            If Not IsWorkingOnSameDay(r, eligibleStaff) And Not IsWorkingOnSameDay(emptyRow, swapCandidate) Then
                ' Perform swap
                wsRoster.Cells(r, MOR_COL).Value = eligibleStaff
                Call IncrementDutiesCounter(eligibleStaff)
                Debug.Print "Assigned " & eligibleStaff & " to row " & r
                
                wsRoster.Cells(emptyRow, MOR_COL).Value = swapCandidate
                Debug.Print "Assigned " & swapCandidate & " to row " & emptyRow
                
                Exit Sub ' Exit after successful swap
            Else
                Debug.Print "Swap not possible at row " & r & " due to conflict (dest: " & r & ", orig: " & emptyRow & ")"
            End If
        End If
NextRow:
    Next r
End Sub

' Helper to find the last unfilled slot
Function FindLastUnfilledSlot(startRow As Long, endRow As Long) As Long
    Dim r As Long
    FindLastUnfilledSlot = 0
    For r = startRow To endRow
        If wsRoster.Cells(r, DAY_COL).Value <> "Sat" And _
           wsRoster.Cells(r, MOR_COL).Value = "" Then
            FindLastUnfilledSlot = r
        End If
    Next r
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

    Debug.Print "=== Checking Eligible Rows ==="
    Debug.Print "WorkDays:"
    For j = LBound(workDays) To UBound(workDays)
        Debug.Print "[" & j & "]: " & workDays(j)
    Next j
    
    For r = START_ROW To last_row_roster
        dayName = Trim(wsRoster.Cells(r, DAY_COL).Value)
        ' Debug: show what day we are checking
        Debug.Print "Row " & r & ": " & dayName
        
        ' Skip if already filled
        If Not IsEmpty(wsRoster.Cells(r, MOR_COL)) Then
            Debug.Print "  -> Skipped (Already Assigned)"
            GoTo SkipRow
        End If
        
        ' Check if the day is in workDays
        For j = LBound(workDays) To UBound(workDays)
            If dayName = workDays(j) Then
                eligibleRows.Add r
                Debug.Print "  -> Added (Matched with " & workDays(j) & ")"
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

    ' Search for the staff name
    Set foundCell = morningtbl.ListColumns("Name").DataBodyRange.Find( _
        What:=staffName, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        ' Get relative row index in the table
        rowIdx = foundCell.row - morningtbl.HeaderRowRange.row

        ' Increment Duties Counter
        With morningtbl.ListRows(rowIdx).Range.Cells(morningtbl.ListColumns("Duties Counter").Index)
            .Value = .Value + 1
        End With
    Else
        MsgBox "Staff '" & staffName & "' not found in table.", vbExclamation
    End If
End Sub

Sub DecrementDutiesCounter(staffName As String)
    Dim rowIdx As Variant
    Dim foundCell As Range

    ' Search for the staff name
    Set foundCell = morningtbl.ListColumns("Name").DataBodyRange.Find( _
        What:=staffName, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        ' Get relative row index in the table
        rowIdx = foundCell.row - morningtbl.HeaderRowRange.row

        ' Decrement Duties Counter
        With morningtbl.ListRows(rowIdx).Range.Cells(morningtbl.ListColumns("Duties Counter").Index)
            .Value = .Value - 1
            If .Value < 0 Then .Value = 0 ' Prevent negative values
        End With
    Else
        MsgBox "Staff '" & staffName & "' not found in table.", vbExclamation
    End If
End Sub

Function IsWorkingOnSameDay(row As Long, staffName As String) As Boolean
    If wsRoster.Cells(row, AFT_COL).Value = staffName Then
        IsWorkingOnSameDay = True
        Exit Function
    End If
    
    If wsRoster.Cells(row, AOH_COL).Value = staffName Then
        IsWorkingOnSameDay = True
        Exit Function
    End If
        
    IsWorkingOnSameDay = False
End Function

' Helper to get the first eligible staff (duties < max)
Function GetFirstEligibleStaff() As String
    Dim i As Long
    For i = 1 To morningtbl.ListRows.count
        Dim staffName As String
        Dim maxDuties As Long
        Dim currDuties As Long
        staffName = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Name").Index).Value
        maxDuties = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Max Duties").Index).Value
        currDuties = morningtbl.DataBodyRange(i, morningtbl.ListColumns("Duties Counter").Index).Value
        If UCase(morningtbl.DataBodyRange(i, morningtbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" And _
           currDuties < maxDuties Then
            GetFirstEligibleStaff = staffName
            Exit Function
        End If
    Next i
    GetFirstEligibleStaff = ""
End Function



Private Sub Worksheet_Calculate()
    Dim rng1 As Range, rng2 As Range, rng3 As Range
    Dim rng4 As Range, rng5 As Range, rng6 As Range

    ' Qualify ranges with the worksheet (Me refers to the worksheet containing this code)
    Set rng1 = Me.Range("B187:O187")
    Set rng2 = Me.Range("B188:O188")
    Set rng3 = Me.Range("B189:O189")
    Set rng4 = Me.Range("D187:O187")
    Set rng5 = Me.Range("D188:O188")
    Set rng6 = Me.Range("D189:O189")

    ' Check if B188 and B189 are empty
    If Me.Range("B188").Value = "" And Me.Range("B189").Value = "" Then
        If Me.Range("B187").Value = "" Then
            
            rng1.Borders.LineStyle = xlNone
            rng2.Borders.LineStyle = xlNone
            rng3.Borders.LineStyle = xlNone
            
            rng1.Interior.ColorIndex = xlNone
            rng2.Interior.ColorIndex = xlNone
            rng3.Interior.ColorIndex = xlNone
        Else
            
            rng1.Borders.LineStyle = xlContinuous
            rng2.Borders.LineStyle = xlNone
            rng3.Borders.LineStyle = xlNone
            
            For Each cell In rng4
                If cell.Value <> "" Then
                    cell.Interior.Color = vbWhite ' White for cells with content
                Else
                    cell.Interior.Color = RGB(128, 128, 128) ' Gray for empty cells
                End If
            Next cell
            
            rng2.Interior.ColorIndex = xlNone
            rng3.Interior.ColorIndex = xlNone
        End If
    Else
       
        rng1.Borders.LineStyle = xlContinuous
        rng2.Borders.LineStyle = xlContinuous
        rng3.Borders.LineStyle = xlContinuous

        ' Check each cell in the ranges and set color based on content
        For Each cell In rng4
            If cell.Value <> "" Then
                cell.Interior.Color = vbWhite ' White for cells with content
            Else
                cell.Interior.Color = RGB(128, 128, 128) ' Gray for empty cells
            End If
        Next cell
        For Each cell In rng5
            If cell.Value <> "" Then
                cell.Interior.Color = vbWhite ' White for cells with content
            Else
                cell.Interior.Color = RGB(128, 128, 128) ' Gray for empty cells
            End If
        Next cell
        For Each cell In rng6
            If cell.Value <> "" Then
                cell.Interior.Color = vbWhite ' White for cells with content
            Else
                cell.Interior.Color = RGB(128, 128, 128) ' Gray for empty cells
            End If
        Next cell
    End If
    
    ' Reprotect the worksheet with specified properties
    'With Me
    '    .Protect DrawingObjects:=True, Contents:=True, Scenarios:=False, _
    '             AllowFormattingCells:=True, AllowFormattingColumns:=True, _
    '             AllowFormattingRows:=True
    'End With
End Sub
' Helper function to check toggle condition for J2
Private Function IsToggleConditionJ2(previousValue As String, currentValue As String) As Boolean
    IsToggleConditionJ2 = ((previousValue = "JAN-JUN" And currentValue = "JUL-DEC") Or _
                           (previousValue = "JUL-DEC" And currentValue = "JAN-JUN"))
End Function

' Helper function to check change condition for M2
Private Function IsChangeConditionM2(previousValue As String, currentValue As String) As Boolean
    IsChangeConditionM2 = (previousValue <> currentValue)
End Function

' Procedure to call CalculateMaxDuties for multiple categories
Private Sub CallCalculateMaxDuties()
    Dim categories() As String
    ReDim categories(0 To 4)
    categories(0) = "LoanMailBox"
    categories(1) = "Morning"
    categories(2) = "Afternoon"
    categories(3) = "AOH"
    categories(4) = "Sat_AOH"
    
    Dim category As Variant
    For Each category In categories
        Debug.Print "Calling CalculateMaxDuties for " & category
        CalculateMaxDuties.CalculateMaxDuties (category)
    Next category
End Sub

' Abstracted procedure to handle changes
Private Sub HandleChange(toggleCell As Range, previousValue As String, currentValue As String, _
                        conditionCheck As String, targetProcedure As String, outputCell As Range)
    Dim result As Long
    Debug.Print "Handling change for " & toggleCell.Address & ", Procedure: " & targetProcedure & ", Condition: " & conditionCheck
    
    ' Since we want any change to trigger, skip specific condition checks
    Debug.Print "Change detected for " & toggleCell.Address & ": Prev=" & previousValue & ", Curr=" & currentValue
    ' Disable events to prevent recursion
    Application.EnableEvents = False
    Range("D6:O189").Select
    Selection.ClearContents
    Selection.Rows.AutoFit
    Application.EnableEvents = True
    
    ' Handle Sub calls with enhanced debugging
    If targetProcedure = "countMorningOrAfternoonOrLMBSlotsSub" Then
        Debug.Print "Calling countMorningOrAfternoonOrLMBSlotsSub"
        Call countMorningOrAfternoonOrLMBSlotsSub(Me.Name, result)
    ElseIf targetProcedure = "countAOHslotsSub" Then
        Debug.Print "Calling countAOHslotsSub"
        Call countAOHslotsSub(Me.Name, result)
    ElseIf targetProcedure = "countSatAOHSub" Then
        Debug.Print "Calling countSatAOHSub"
        Call countSatAOHSub(Me.Name, result)
    Else
        Debug.Print "Unknown target procedure: " & targetProcedure
    End If
    Debug.Print "Result for " & targetProcedure & ": " & result
    ' Disable events to prevent recursion during assignment
    Application.EnableEvents = False
    outputCell.Value = result
    Application.EnableEvents = True
    
    ' Call CalculateMaxDuties for all sheets
    CallCalculateMaxDuties
End Sub

' Worksheet_Change event handler
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Debug: Log that the event has started
    Debug.Print "Worksheet_Change event triggered for Target: " & Target.Address
    
    ' Configuration for trigger cells
    Dim triggers(1 To 6) As Variant
    ' Triggers for J2
    triggers(1) = Array(Me.Range("J2"), "previousValueJ2", "", "countMorningOrAfternoonOrLMBSlotsSub", Me.Range("E191"))
    triggers(2) = Array(Me.Range("J2"), "previousValueJ2", "", "countAOHslotsSub", Me.Range("K191"))
    triggers(3) = Array(Me.Range("J2"), "previousValueJ2", "", "countSatAOHSub", Me.Range("M191"))
    ' Triggers for M2
    triggers(4) = Array(Me.Range("M2"), "previousValueM2", "", "countMorningOrAfternoonOrLMBSlotsSub", Me.Range("E191"))
    triggers(5) = Array(Me.Range("M2"), "previousValueM2", "", "countAOHslotsSub", Me.Range("K191"))
    triggers(6) = Array(Me.Range("M2"), "previousValueM2", "", "countSatAOHSub", Me.Range("M191"))
    
    Dim i As Long
    ' Capture initial values before processing all triggers
    Dim initialCurrentValue As String
    Dim initialPreviousValue As String
    
    For i = 1 To UBound(triggers)
        Dim toggleCell As Range
        Dim previousValueName As String
        Dim conditionCheck As String
        Dim targetProcedure As String
        Dim outputCell As Range
        Set toggleCell = triggers(i)(0)
        previousValueName = triggers(i)(1)
        conditionCheck = triggers(i)(2) ' Now empty, as we handle change directly
        targetProcedure = triggers(i)(3)
        Set outputCell = triggers(i)(4)
        
        If Not Intersect(Target, toggleCell) Is Nothing Then
            If initialCurrentValue = "" Then
                initialCurrentValue = UCase(Trim(toggleCell.Value))
                ' Use a static variable indirectly via a dictionary for persistence
                Static previousValues As Object
                If previousValues Is Nothing Then Set previousValues = CreateObject("Scripting.Dictionary")
                If Not previousValues.Exists(previousValueName) Then previousValues(previousValueName) = ""
                initialPreviousValue = previousValues(previousValueName)
                Debug.Print "Initial values - Toggle cell (" & toggleCell.Address & ") current value: " & initialCurrentValue
                Debug.Print "Initial values - Toggle cell (" & toggleCell.Address & ") previous value: " & initialPreviousValue
                Debug.Print "Initial values - Target value before trim: |" & Target.Value & "|"
            End If
            
            ' Trigger HandleChange if values differ, regardless of specific condition
            If initialPreviousValue <> initialCurrentValue Then
                HandleChange toggleCell, initialPreviousValue, initialCurrentValue, conditionCheck, targetProcedure, outputCell
            Else
                Debug.Print "No change detected for " & toggleCell.Address & ": Prev=" & initialPreviousValue & ", Curr=" & initialCurrentValue
            End If
            ' Update previous value after all triggers are processed
        End If
    Next i
    
    ' Update previous values after all triggers are handled
    If initialCurrentValue <> "" Then
        For i = 1 To UBound(triggers)
            Dim previousValueNameUpdate As String
            Set toggleCell = triggers(i)(0)
            previousValueNameUpdate = triggers(i)(1)
            If Not Intersect(Target, toggleCell) Is Nothing Then
                previousValues(previousValueNameUpdate) = initialCurrentValue
            End If
        Next i
    End If
End Sub

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

' Declare worksheet and table
Private wsPersonnel As Worksheet
Private afternoontbl As ListObject
Private spectbl As ListObject

Sub AssignAfternoonDuties()
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
    Dim needsReassignment As Boolean
    Dim lastTwoWeeksStart As Long
    Dim staffPool() As String
    Dim tmpStaff() As String
    Dim poolIndex As Long
    Dim assignedInThisPass As Boolean
    Dim poolSize As Long
    Dim initialSize As Long
    Dim currDuties As Long
    Dim dutiesNeeded As Long
    Dim reassignmentAttempts As Long

    Set wsRoster = Sheets("Roster")
    Set wsSettings = Sheets("Settings")
    Set wsPersonnel = Sheets("Afternoon PersonnelList")
    Set afternoontbl = wsPersonnel.ListObjects("AfternoonMainList")
    Set spectbl = wsPersonnel.ListObjects("AfternoonSpecificDaysWorkingStaff")
    
    totalDays = wsRoster.Range(wsRoster.Cells(START_ROW, DATE_COL), wsRoster.Cells(last_row_roster, DATE_COL)).Rows.count
    Debug.Print "Afternoon assignment starts here"
    
    ' Step 1: Assign Specific Days Staff
    For i = 1 To spectbl.ListRows.count
        staffName = spectbl.DataBodyRange(i, spectbl.ListColumns("Name").Index).Value
        workDays = Split(spectbl.DataBodyRange(i, spectbl.ListColumns("Working Days").Index).Value, ",")
        
        ' Clean up day names (remove spaces)
        For j = 0 To UBound(workDays)
            workDays(j) = Trim(workDays(j))
        Next j
        
        ' Get max duties and department for this staff from AfternoonMainList
        Dim dept As String
        For r = 1 To afternoontbl.ListRows.count
            If afternoontbl.DataBodyRange(r, afternoontbl.ListColumns("Name").Index).Value = staffName Then
                maxDuties = afternoontbl.DataBodyRange(r, afternoontbl.ListColumns("Max Duties").Index).Value
                dept = afternoontbl.DataBodyRange(r, afternoontbl.ListColumns("Department").Index).Value
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
        
        ' Assign staff
        assignedCount = 0
        For j = 1 To eligibleRows.count
            If assignedCount >= maxDuties Then Exit For
            r = tmpRows(j)
            If Not IsWorkingOnSameDay(r, staffName) And wsRoster.Cells(r, AFT_COL).Value = "" Then
                ' Check vacation constraint
                If UCase(Trim(wsRoster.Cells(r, VAC_COL).Value)) = "VACATION" And UCase(dept) <> "APRM" Then
                    Debug.Print "Skipped " & staffName & " for vacation row " & r & " (not APRM)"
                    GoTo NextRowJ
                End If
                wsRoster.Cells(r, AFT_COL).Value = staffName
                Call IncrementDutiesCounter(staffName)
                assignedCount = assignedCount + 1
                Debug.Print "Assigned Specific Days staff " & staffName & " to row " & r
            End If
NextRowJ:
        Next j
    Next i
    
    ' Step 2: Assign All Days Staff
    For r = START_ROW To last_row_roster
        If wsRoster.Cells(r, DAY_COL).Value = "Sat" Then GoTo SkipDay
        If wsRoster.Cells(r, AFT_COL).Value = "CLOSED" Then GoTo SkipDay
        For i = 1 To afternoontbl.ListRows.count
            staffName = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Name").Index).Value
            If UCase(afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Availability Type").Index).Value) = "SPECIFIC DAYS" Then
                GoTo SkipStaff
            End If
            
            maxDuties = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Max Duties").Index).Value
            currDuties = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Duties Counter").Index).Value
            dept = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Department").Index).Value
            ' Check if the staff already reached his max duties
            If currDuties >= maxDuties Then GoTo SkipStaff
            If IsWorkingOnSameDay(r, staffName) Then GoTo SkipStaff
            
            ' Assign from top with vacation constraint
            If wsRoster.Cells(r, AFT_COL).Value = "" Then
                If UCase(Trim(wsRoster.Cells(r, VAC_COL).Value)) = "VACATION" And UCase(dept) <> "APRM" Then
                    Debug.Print "Skipped " & staffName & " for vacation row " & r & " (not APRM)"
                    GoTo SkipStaff
                End If
                wsRoster.Cells(r, AFT_COL).Value = staffName
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
            Call ReassignAfternoonDuties
        Else
            Exit Do
        End If
        reassignmentAttempts = reassignmentAttempts + 1
        If reassignmentAttempts > MAX_ATTEMPTS Then
            Debug.Print "Max reassignment attempts reached. Assigning remaining staff to unfilled slots."
            ' Fallback: Assign all remaining eligible staff to unfilled slots
            Dim eligibleStaffList As Collection
            Set eligibleStaffList = New Collection
            For i = 1 To afternoontbl.ListRows.count
                staffName = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Name").Index).Value
                maxDuties = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Max Duties").Index).Value
                currDuties = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Duties Counter").Index).Value
                dept = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Department").Index).Value
                If UCase(afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" And _
                   currDuties < maxDuties Then
                    eligibleStaffList.Add Array(staffName, dept)
                End If
            Next i
            
            Dim unfilledSlotsList As Collection
            Set unfilledSlotsList = New Collection
            For r = START_ROW To last_row_roster
                If wsRoster.Cells(r, DAY_COL).Value <> "Sat" And _
                   wsRoster.Cells(r, AFT_COL).Value = "" Then
                    unfilledSlotsList.Add r
                End If
            Next r
            
            ' Assign each eligible staff to an unfilled slot
            For i = 1 To eligibleStaffList.count
                If i <= unfilledSlotsList.count Then
                    staffName = eligibleStaffList(i)(0)
                    dept = eligibleStaffList(i)(1)
                    r = unfilledSlotsList(i)
                    If Not IsWorkingOnSameDay(r, staffName) Then
                        If UCase(Trim(wsRoster.Cells(r, VAC_COL).Value)) = "VACATION" And UCase(dept) <> "APRM" Then
                            Debug.Print "Fallback: Skipped " & staffName & " for vacation row " & r & " (not APRM)"
                        Else
                            wsRoster.Cells(r, AFT_COL).Value = staffName
                            wsRoster.Cells(r, AFT_COL).Interior.Color = vbYellow ' Highlight yellow (RGB 255, 255, 0)
                            Call IncrementDutiesCounter(staffName)
                            Debug.Print "Fallback: Assigned " & staffName & " to row " & r & " (highlighted yellow)"
                        End If
                    Else
                        Debug.Print "Fallback: Skipped " & staffName & " at row " & r & " due to AOH conflict"
                    End If
                End If
            Next i
            Exit Do
        End If
    Loop
    
    MsgBox "Afternoon duties assignment completed!", vbInformation
End Sub

' Helper to count unfilled Afternoon slots
Function CountUnfilledSlots(startRow As Long, endRow As Long) As Long
    Dim r As Long
    Dim count As Long
    count = 0
    For r = startRow To endRow
        If wsRoster.Cells(r, DAY_COL).Value <> "Sat" And _
           wsRoster.Cells(r, AFT_COL).Value = "" Then
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
    For i = 1 To afternoontbl.ListRows.count
        Dim staffName As String
        Dim maxDuties As Long
        Dim currDuties As Long
        staffName = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Name").Index).Value
        maxDuties = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Max Duties").Index).Value
        currDuties = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Duties Counter").Index).Value
        If UCase(afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" And _
           currDuties < maxDuties Then
            count = count + 1
        End If
    Next i
    CountEligibleStaff = count
End Function

Sub ReassignAfternoonDuties()
    Dim r As Long
    Dim i As Long
    Dim staffName As String
    Dim maxDuties As Long
    Dim currDuties As Long
    Dim eligibleStaff As String
    Dim swapCandidate As String
    Dim emptyRow As Long ' The empty original slot where eligibleStaff couldn't be assigned
    Dim dept As String
    
    ' Find the first eligible staff (currDuties < maxDuties)
    eligibleStaff = GetFirstEligibleStaff
    If eligibleStaff = "" Then
        Debug.Print "No eligible staff found for Afternoon reassignment."
        Exit Sub
    End If
    Debug.Print "Eligible staff for Afternoon reassignment: " & eligibleStaff
    
    ' Get department of eligibleStaff
    For i = 1 To afternoontbl.ListRows.count
        If afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Name").Index).Value = eligibleStaff Then
            dept = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Department").Index).Value
            Exit For
        End If
    Next i
    
    ' Assume emptyRow is determined earlier (e.g., the last unfilled slot)
    emptyRow = FindLastUnfilledSlot(START_ROW, last_row_roster)
    If emptyRow = 0 Then Exit Sub ' No empty slot to reassign
    
    ' Loop through all rows to find a swap opportunity
    For r = START_ROW To last_row_roster
        If wsRoster.Cells(r, DAY_COL).Value = "Sat" Or _
           wsRoster.Cells(r, AFT_COL).Value = "CLOSED" Or _
           wsRoster.Cells(r, AFT_COL).Value = "" Then
            GoTo NextRow
        End If
        
        swapCandidate = wsRoster.Cells(r, AFT_COL).Value
        If swapCandidate <> "" And swapCandidate <> eligibleStaff Then
            ' Get department of swapCandidate
            Dim swapDept As String
            For i = 1 To afternoontbl.ListRows.count
                If afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Name").Index).Value = swapCandidate Then
                    swapDept = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Department").Index).Value
                    Exit For
                End If
            Next i
            
            ' Check if swap is allowed (no conflict in destination and original slot, respect vacation constraint)
            If Not IsWorkingOnSameDay(r, eligibleStaff) And Not IsWorkingOnSameDay(emptyRow, swapCandidate) Then
                If UCase(Trim(wsRoster.Cells(emptyRow, VAC_COL).Value)) = "VACATION" And UCase(swapDept) <> "APRM" Then
                    Debug.Print "Swap skipped for row " & r & " (swapCandidate " & swapCandidate & " not APRM for vacation slot " & emptyRow & ")"
                    GoTo NextRow
                End If
                ' Perform swap
                wsRoster.Cells(r, AFT_COL).Value = eligibleStaff
                Call IncrementDutiesCounter(eligibleStaff)
                Debug.Print "Assigned " & eligibleStaff & " to row " & r
                
                wsRoster.Cells(emptyRow, AFT_COL).Value = swapCandidate
                Call IncrementDutiesCounter(swapCandidate)
                Debug.Print "Assigned " & swapCandidate & " to row " & emptyRow
                
                Exit Sub ' Exit after successful swap
            Else
                Debug.Print "Swap not possible at row " & r & " due to conflict (dest: " & r & ", orig: " & emptyRow & ")"
            End If
        End If
NextRow:
    Next r
End Sub

' Helper to find the last unfilled slot
Function FindLastUnfilledSlot(startRow As Long, endRow As Long) As Long
    Dim r As Long
    FindLastUnfilledSlot = 0
    For r = startRow To endRow
        If wsRoster.Cells(r, DAY_COL).Value <> "Sat" And _
           wsRoster.Cells(r, AFT_COL).Value = "" Then
            FindLastUnfilledSlot = r
        End If
    Next r
End Function

' Helper to find the original Afternoon duty slot for a staff (excluding the current row)
Function FindOriginalAfternoonDuty(staffName As String, excludeRow As Long) As Long
    Dim r As Long
    For r = START_ROW To last_row_roster
        If r <> excludeRow And wsRoster.Cells(r, AFT_COL).Value = staffName And _
           wsRoster.Cells(r, DAY_COL).Value <> "Sat" Then
            FindOriginalAfternoonDuty = r
            Exit Function
        End If
    Next r
    FindOriginalAfternoonDuty = 0 ' No original slot found
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
        If Not IsEmpty(wsRoster.Cells(r, AFT_COL)) Then
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
    Set GetEligibleRows = eligibleRows
End Function

Sub IncrementDutiesCounter(staffName As String)
    Dim rowIdx As Variant
    Dim foundCell As Range
    Set foundCell = afternoontbl.ListColumns("Name").DataBodyRange.Find(What:=staffName, LookIn:=xlValues, LookAt:=xlWhole)
    If Not foundCell Is Nothing Then
        rowIdx = foundCell.row - afternoontbl.HeaderRowRange.row
        With afternoontbl.ListRows(rowIdx).Range.Cells(afternoontbl.ListColumns("Duties Counter").Index)
            .Value = .Value + 1
        End With
    Else
        MsgBox "Staff '" & staffName & "' not found in table.", vbExclamation
    End If
End Sub

Sub DecrementDutiesCounter(staffName As String)
    Dim rowIdx As Variant
    Dim foundCell As Range
    Set foundCell = afternoontbl.ListColumns("Name").DataBodyRange.Find(What:=staffName, LookIn:=xlValues, LookAt:=xlWhole)
    If Not foundCell Is Nothing Then
        rowIdx = foundCell.row - afternoontbl.HeaderRowRange.row
        With afternoontbl.ListRows(rowIdx).Range.Cells(afternoontbl.ListColumns("Duties Counter").Index)
            .Value = .Value - 1
            If .Value < 0 Then .Value = 0 ' Prevent negative values
        End With
    Else
        MsgBox "Staff '" & staffName & "' not found in table.", vbExclamation
    End If
End Sub

Function IsWorkingOnSameDay(row As Long, staffName As String) As Boolean
    ' Check if staff is working on AOH on the same day
    If wsRoster.Cells(row, AOH_COL).Value = staffName Then
        IsWorkingOnSameDay = True
        Exit Function
    End If
    IsWorkingOnSameDay = False
End Function

' Helper to get the first eligible staff (duties < max)
Function GetFirstEligibleStaff() As String
    Dim i As Long
    For i = 1 To afternoontbl.ListRows.count
        Dim staffName As String
        Dim maxDuties As Long
        Dim currDuties As Long
        staffName = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Name").Index).Value
        maxDuties = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Max Duties").Index).Value
        currDuties = afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Duties Counter").Index).Value
        If UCase(afternoontbl.DataBodyRange(i, afternoontbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" And _
           currDuties < maxDuties Then
            GetFirstEligibleStaff = staffName
            Exit Function
        End If
    Next i
    GetFirstEligibleStaff = ""
End Function



Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrHandler
    
    Dim table As ListObject
    Dim dutiesCounterCol As ListColumn

    Set table = Me.ListObjects("AfternoonMainList")
    
    ' Check if the table exists
    If table Is Nothing Then
        Exit Sub
    End If
    
    ' Get the Duties Counter column
    Set dutiesCounterCol = table.ListColumns("Duties Counter")
    If dutiesCounterCol Is Nothing Then
        MsgBox "Column 'Duties Counter' not found in 'AfternoonMainList'.", vbExclamation
        Exit Sub
    End If
    
    ' Exit if the table is empty, but allow D7 change processing
    If table.ListRows.count = 0 And Not Intersect(Target, Me.Range("D7")) Is Nothing Then
        GoTo ProcessD7Change
    ElseIf table.ListRows.count = 0 Then
        Exit Sub
    End If
    
    ' Sort duties counter by ascending if change is in Duties Counter column
    If Not Intersect(Target, dutiesCounterCol.DataBodyRange) Is Nothing Then
        With table.Sort
            .SortFields.Clear
            .SortFields.Add Key:=dutiesCounterCol.DataBodyRange, _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Header = xlYes
            .Apply
        End With
    End If
    
ProcessD7Change:
    ' Sets Working Days for All Days and show form for Specific Days
    If Not Intersect(Target, Me.Range("D7")) Is Nothing Then
        Dim availabilityType As String
        availabilityType = UCase(Trim(Me.Range("D7").Value))
        
        ' Debug: Log the change
        Debug.Print "Worksheet_Change triggered at " & Now & " for D7: " & availabilityType
        
        Select Case availabilityType
            Case "ALL DAYS"
                ' Set percentage to 100%
                Me.Range("D9").Value = 100
                Me.Range("D8").Value = "Mon, Tue, Wed, Thu, Fri, Sat"
                Debug.Print "Set D9 to 100% for All Days"
                
            Case "SPECIFIC DAYS"
                ' Show the multiselect form
                With frmSpecificDays
                    .targetSheetName = Me.Name
                    .Show
                End With
                Debug.Print "Showed form for Specific Days"
                
            Case Else
                ' Clear D9 if invalid selection
                Me.Range("D9").Value = ""
                Debug.Print "Invalid selection, D9 cleared"
        End Select
    End If
    
    Exit Sub

ErrHandler:
    MsgBox "An error occurred: " & Err.Description & vbCrLf & _
           "Line: " & Erl, vbCritical
    Exit Sub
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrHandler
    
    Dim table As ListObject
    Dim dutiesCounterCol As ListColumn

    Set table = Me.ListObjects("AOHMainList")
    
    ' Check if the table exists
    If table Is Nothing Then
        Exit Sub
    End If
    
    ' Get the Duties Counter column
    Set dutiesCounterCol = table.ListColumns("Duties Counter")
    If dutiesCounterCol Is Nothing Then
        MsgBox "Column 'Duties Counter' not found in 'AOHMainList'.", vbExclamation
        Exit Sub
    End If
    
    ' Exit if the table is empty, but allow D7 change processing
    If table.ListRows.count = 0 And Not Intersect(Target, Me.Range("D7")) Is Nothing Then
        GoTo ProcessD7Change
    ElseIf table.ListRows.count = 0 Then
        Exit Sub
    End If
    
    ' Sort duties counter by ascending if change is in Duties Counter column
    If Not Intersect(Target, dutiesCounterCol.DataBodyRange) Is Nothing Then
        With table.Sort
            .SortFields.Clear
            .SortFields.Add Key:=dutiesCounterCol.DataBodyRange, _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Header = xlYes
            .Apply
        End With
    End If
    
ProcessD7Change:
    ' Sets Working Days for All Days and show form for Specific Days
    If Not Intersect(Target, Me.Range("D7")) Is Nothing Then
        Dim availabilityType As String
        availabilityType = UCase(Trim(Me.Range("D7").Value))
        
        ' Debug: Log the change
        Debug.Print "Worksheet_Change triggered at " & Now & " for D7: " & availabilityType
        
        Select Case availabilityType
            Case "ALL DAYS"
                ' Set percentage to 100%
                Me.Range("D9").Value = 100
                Me.Range("D8").Value = "Mon, Tue, Wed, Thu, Fri, Sat"
                Debug.Print "Set D9 to 100% for All Days"
                
            Case "SPECIFIC DAYS"
                ' Show the multiselect form
                With frmSpecificDays
                    .targetSheetName = Me.Name
                    .Show
                End With
                Debug.Print "Showed form for Specific Days"
                
            Case Else
                ' Clear D9 if invalid selection
                Me.Range("D9").Value = ""
                Debug.Print "Invalid selection, D9 cleared"
        End Select
    End If
    
    Exit Sub

ErrHandler:
    MsgBox "An error occurred: " & Err.Description & vbCrLf & _
           "Line: " & Erl, vbCritical
    Exit Sub
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrHandler
    
    Dim table As ListObject
    Dim dutiesCounterCol As ListColumn

    Set table = Me.ListObjects("SatAOHMainList")
    
    ' Check if the table exists
    If table Is Nothing Then
        Exit Sub
    End If
    
    ' Get the Duties Counter column
    Set dutiesCounterCol = table.ListColumns("Duties Counter")
    If dutiesCounterCol Is Nothing Then
        MsgBox "Column 'Duties Counter' not found in 'SatAOHMainList'.", vbExclamation
        Exit Sub
    End If
    
    ' Exit if the table is empty, but allow D7 change processing
    If table.ListRows.count = 0 And Not Intersect(Target, Me.Range("D7")) Is Nothing Then
        GoTo ProcessD7Change
    ElseIf table.ListRows.count = 0 Then
        Exit Sub
    End If
    
    ' Sort duties counter by ascending if change is in Duties Counter column
    If Not Intersect(Target, dutiesCounterCol.DataBodyRange) Is Nothing Then
        With table.Sort
            .SortFields.Clear
            .SortFields.Add Key:=dutiesCounterCol.DataBodyRange, _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Header = xlYes
            .Apply
        End With
    End If
    
ProcessD7Change:
    ' Sets Working Days for All Days
    If Not Intersect(Target, Me.Range("D7")) Is Nothing Then
        Dim availabilityType As String
        availabilityType = UCase(Trim(Me.Range("D7").Value))
        
        ' Debug: Log the change
        Debug.Print "Worksheet_Change triggered at " & Now & " for D7: " & availabilityType
        
        Select Case availabilityType
            Case "ALL DAYS"
                ' Set percentage to 100%
                Me.Range("D9").Value = 100
                Me.Range("D8").Value = "Sat"
                Debug.Print "Set D9 to 100% for All Days"
        End Select
    End If
    
    Exit Sub

ErrHandler:
    MsgBox "An error occurred: " & Err.Description & vbCrLf & _
           "Line: " & Erl, vbCritical
    Exit Sub
End Sub

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



' Declare worksheet and table
Private wsRoster As Worksheet
Private wsSettings As Worksheet
Private wsPersonnel As Worksheet
Private aohtbl As ListObject
Private spectbl As ListObject

Sub AssignSatAOHDuties()
    Set wsRoster = Sheets("Roster")
    Set wsSettings = Sheets("Settings")
    Set wsPersonnel = Sheets("Sat AOH PersonnelList")
    Set aohtbl = wsPersonnel.ListObjects("SatAOHMainList")
    
    Dim i As Long, r As Long
    Dim maxDuties As Long
    Dim staffName As String
    Dim assignedStaff1 As String
    Dim prevSatRow As Long
    Dim prevStaff1 As String
    Dim prevStaff2 As String
    
    ' Pass 1: Assign staff to SAT_AOH_COL1
    For r = START_ROW To last_row_roster
        Dim dayValue As String
        dayValue = Trim(wsRoster.Cells(r, DAY_COL).Text)
        If dayValue = "Sat" Then
            If wsRoster.Cells(r, SAT_AOH_COL1).Value = "" Then
                prevSatRow = r - 7 ' Previous Saturday (exactly 7 days back)
                Debug.Print "prevsatrow: " & prevSatRow
                If prevSatRow >= START_ROW Then
                    If wsRoster.Cells(prevSatRow, DAY_COL).Text = "Sat" Then
                        prevStaff1 = wsRoster.Cells(prevSatRow, SAT_AOH_COL1).Value
                        prevStaff2 = wsRoster.Cells(prevSatRow, SAT_AOH_COL2).Value
                    End If
                Else
                    prevStaff1 = "" ' No previous Saturday or unassigned
                    prevStaff2 = ""
                End If
                
                For i = 1 To aohtbl.ListRows.count
                    staffName = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Name").Index).Value
                    maxDuties = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Max Duties").Index).Value
                    Dim currDuties As Long
                    currDuties = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Duties Counter").Index).Value

                    If currDuties < maxDuties And (prevStaff1 = "" Or prevStaff2 = "" Or (staffName <> prevStaff1 And staffName <> prevStaff2)) Then
                        wsRoster.Cells(r, SAT_AOH_COL1).Value = staffName
                        Call IncrementDutiesCounter(staffName)
                        assignedStaff1 = staffName
                        Exit For
                    End If
                Next i
                If wsRoster.Cells(r, SAT_AOH_COL1).Value = "" Then
                    Debug.Print "Warning: No eligible staff for SAT_AOH_COL1 at row " & r & " due to consecutive SAT AOH constraint or insufficient staff."
                End If
            End If
        End If
    Next r

    ' Pass 2: Assign different staff to SAT_AOH_COL2
    For r = START_ROW To last_row_roster
        Dim dayValue2 As String
        dayValue2 = Trim(wsRoster.Cells(r, DAY_COL).Text)
        If dayValue2 = "Sat" And wsRoster.Cells(r, SAT_AOH_COL1).Value <> "" And wsRoster.Cells(r, SAT_AOH_COL2).Value = "" Then
            prevSatRow = r - 7 ' Previous Saturday
            If prevSatRow >= START_ROW Then
                If wsRoster.Cells(prevSatRow, DAY_COL).Text = "Sat" Then
                    prevStaff1 = wsRoster.Cells(prevSatRow, SAT_AOH_COL1).Value
                    prevStaff2 = wsRoster.Cells(prevSatRow, SAT_AOH_COL2).Value
                End If
            Else
                prevStaff1 = "" ' No previous Saturday or unassigned
                prevStaff2 = ""
            End If
            
            For i = 1 To aohtbl.ListRows.count
                staffName = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Name").Index).Value
                maxDuties = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Max Duties").Index).Value
                Dim currDuties2 As Long
                currDuties2 = aohtbl.DataBodyRange(i, aohtbl.ListColumns("Duties Counter").Index).Value
                If currDuties2 < maxDuties And staffName <> wsRoster.Cells(r, SAT_AOH_COL1).Value And _
                   (prevStaff1 = "" Or prevStaff2 = "" Or (staffName <> prevStaff1 And staffName <> prevStaff2)) Then
                    wsRoster.Cells(r, SAT_AOH_COL2).Value = staffName
                    Call IncrementDutiesCounter(staffName)
                    Exit For
                End If
            Next i
            If wsRoster.Cells(r, SAT_AOH_COL2).Value = "" Then
                Debug.Print "Warning: No eligible staff for SAT_AOH_COL2 at row " & r & " due to consecutive SAT AOH constraint or insufficient staff."
            End If
        End If
    Next r

    MsgBox "Sat AOH duties assignment completed!", vbInformation
End Sub

Sub IncrementDutiesCounter(staffName As String)
    Dim rowIdx As Variant
    Dim foundCell As Range

    ' Search for the staff name
    Set foundCell = aohtbl.ListColumns("Name").DataBodyRange.Find( _
        What:=staffName, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        ' Get relative row index in the table
        rowIdx = foundCell.row - aohtbl.HeaderRowRange.row
        Debug.Print "Checking worksheet status: " & wsPersonnel.ProtectContents
        ' Increment Duties Counter
        With aohtbl.ListRows(rowIdx).Range.Cells(aohtbl.ListColumns("Duties Counter").Index)
            .Value = .Value + 1
        End With
    Else
        MsgBox "Staff '" & staffName & "' not found in table.", vbExclamation
    End If
End Sub



Sub GenerateRosterComparison(systemSheetName As String, actualSheetName As String)
    Dim wsSystem As Worksheet, wsActual As Worksheet, wsReport As Worksheet
    Dim lastRow As Long, r As Long, col As Long
    Dim colStart As Long: colStart = 6 ' Start from MOR_COL (Morning), adjust as needed
    Dim colEnd As Long: colEnd = 14 ' Up to SatAOH2

    Set wsSystem = Sheets(systemSheetName)
    Set wsActual = Sheets(actualSheetName)
    
    Set wsReport = Sheets.Add(After:=Sheets(Sheets.count))
    wsReport.Name = "RosterComparison_" & Format(Now, "yyyymmdd_hhnn")
    
    lastRow = wsSystem.Cells(wsSystem.Rows.count, 1).End(xlUp).row
    
    ' Header
    wsReport.Cells(1, 1).Value = "Date"
    wsReport.Cells(1, 2).Value = "Day"
    wsReport.Cells(1, 3).Value = "Slot"
    wsReport.Cells(1, 4).Value = "System"
    wsReport.Cells(1, 5).Value = "Actual"
    wsReport.Cells(1, 6).Value = "Match?"
    
    Dim reportRow As Long: reportRow = 2
    
    For r = START_ROW To lastRow
        For col = colStart To colEnd Step 2 ' Assuming each slot has 2-col spacing
            Dim slotName As String
            Select Case col
                Case MOR_COL: slotName = "Morning"
                Case AFT_COL: slotName = "Afternoon"
                Case AOH_COL: slotName = "AOH"
                Case SAT_AOH_COL1: slotName = "Sat AOH 1"
                Case SAT_AOH_COL2: slotName = "Sat AOH 2"
                Case Else: slotName = "Slot " & col
            End Select
            
            Dim systemVal As String, actualVal As String
            systemVal = Trim(wsSystem.Cells(r, col).Value)
            actualVal = Trim(wsActual.Cells(r, col).Value)
            
            wsReport.Cells(reportRow, 1).Value = wsSystem.Cells(r, DATE_COL).Value
            wsReport.Cells(reportRow, 2).Value = wsSystem.Cells(r, DAY_COL).Value
            wsReport.Cells(reportRow, 3).Value = slotName
            wsReport.Cells(reportRow, 4).Value = systemVal
            wsReport.Cells(reportRow, 5).Value = actualVal
            wsReport.Cells(reportRow, 6).Value = IIf(systemVal = actualVal, "?", "?")
            
            reportRow = reportRow + 1
        Next col
    Next r
    
    MsgBox "Roster comparison completed. See '" & wsReport.Name & "'.", vbInformation
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrHandler
    
    Dim table As ListObject
    Dim dutiesCounterCol As ListColumn
    
    Set table = Me.ListObjects("LoanMailBoxMainList")
    
    ' Check if the table exists
    If table Is Nothing Then
        Exit Sub
    End If
    
    ' Get the Duties Counter column
    Set dutiesCounterCol = table.ListColumns("Duties Counter")
    If dutiesCounterCol Is Nothing Then
        MsgBox "Column 'Duties Counter' not found in 'LoanMailBoxMainList'.", vbExclamation
        Exit Sub
    End If
    
    ' Exit if the table is empty, but allow D7 change processing
    If table.ListRows.count = 0 And Not Intersect(Target, Me.Range("D7")) Is Nothing Then
        GoTo ProcessD7Change
    ElseIf table.ListRows.count = 0 Then
        Exit Sub
    End If
    
    ' Sort duties counter by ascending if change is in Duties Counter column
    If Not Intersect(Target, dutiesCounterCol.DataBodyRange) Is Nothing Then
        With table.Sort
            .SortFields.Clear
            .SortFields.Add Key:=dutiesCounterCol.DataBodyRange, _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .Header = xlYes
            .Apply
        End With
    End If
    
ProcessD7Change:
    ' Sets Working Days for All Days and show form for Specific Days
    If Not Intersect(Target, Me.Range("D7")) Is Nothing Then
        Dim availabilityType As String
        availabilityType = UCase(Trim(Me.Range("D7").Value))
        
        ' Debug: Log the change
        Debug.Print "Worksheet_Change triggered at " & Now & " for D7: " & availabilityType
        
        Select Case availabilityType
            Case "ALL DAYS"
                ' Set percentage to 100%
                Me.Range("D9").Value = 100
                Me.Range("D8").Value = "Mon, Tue, Wed, Thu, Fri, Sat"
                Debug.Print "Set D9 to 100% for All Days"
                
            Case "SPECIFIC DAYS"
                ' Show the multiselect form
                With frmSpecificDays
                    .targetSheetName = Me.Name
                    .Show
                End With
                Debug.Print "Showed form for Specific Days"
                
            Case Else
                ' Clear D9 if invalid selection
                Me.Range("D9").Value = ""
                Debug.Print "Invalid selection, D9 cleared"
        End Select
    End If
    
    Exit Sub

ErrHandler:
    MsgBox "An error occurred: " & Err.Description & vbCrLf & _
           "Line: " & Erl, vbCritical
    ' Ensure events are re-enabled if they were disabled
    If Not Application.EnableEvents Then Application.EnableEvents = True
    Exit Sub
End Sub

' Declare worksheet and table
Private wsPersonnel As Worksheet
Private lmbtbl As ListObject
Private spectbl As ListObject

Sub AssignLoanMailBoxDuties()
    Set wsRoster = Sheets("Roster")
    Set wsSettings = Sheets("Settings")
    Set wsPersonnel = Sheets("Loan Mail Box PersonnelList")
    Set lmbtbl = wsPersonnel.ListObjects("LoanMailBoxMainList")
    Set spectbl = wsPersonnel.ListObjects("LoanMailBoxSpecificDaysWorkingStaff")
    
    Dim i As Long, j As Long, r As Long
    Dim dateCount As Long
    Dim totalDays As Long
    Dim dayName As String
    Dim maxDuties As Long
    Dim candidates() As String
    Dim staffName As String
    Dim workDays As Variant
    Dim eligibleRows As Collection
    Dim tmpRows() As Long
    Dim assignedCount As Long
    Dim reassignmentAttempts As Long

    totalDays = wsRoster.Range(wsRoster.Cells(START_ROW, DATE_COL), wsRoster.Cells(last_row_roster, DATE_COL)).Rows.count
    Debug.Print "Loan Mailbox assignment starts here"
    
    ' Step 1: Assign Specific Days Staff
    For i = 1 To spectbl.ListRows.count
        staffName = spectbl.DataBodyRange(i, spectbl.ListColumns("Name").Index).Value
        Debug.Print staffName
        workDays = Split(spectbl.DataBodyRange(i, spectbl.ListColumns("Working Days").Index).Value, ",")
        
        ' Clean up day names (remove spaces)
        For j = 0 To UBound(workDays)
            workDays(j) = Trim(workDays(j))
        Next j
        
        ' Get max duties for this staff from LoanMailBoxMainList
        For r = 1 To lmbtbl.ListRows.count
            If lmbtbl.DataBodyRange(r, lmbtbl.ListColumns("Name").Index).Value = staffName Then
                maxDuties = lmbtbl.DataBodyRange(r, lmbtbl.ListColumns("Max Duties").Index).Value
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
        
        ' Assign staff
        assignedCount = 0
        For j = 1 To eligibleRows.count
            If assignedCount >= maxDuties Then Exit For
            If Not IsWorkingOnSameDay(tmpRows(j), staffName) Then
                wsRoster.Cells(tmpRows(j), LMB_COL).Value = staffName
                Call IncrementDutiesCounter(staffName)
                assignedCount = assignedCount + 1
            End If
        Next j
    Next i
    
    ' Step 2: Assign All Days Staff
    For r = START_ROW To last_row_roster
        If wsRoster.Cells(r, DAY_COL).Value = "Sat" Then GoTo SkipDay
        If wsRoster.Cells(r, LMB_COL).Value = "CLOSED" Then GoTo SkipDay
        For i = 1 To lmbtbl.ListRows.count
            staffName = lmbtbl.DataBodyRange(i, lmbtbl.ListColumns("Name").Index).Value
            If UCase(lmbtbl.DataBodyRange(i, lmbtbl.ListColumns("Availability Type").Index).Value) = "SPECIFIC DAYS" Then
                GoTo SkipStaff
            End If
            
            maxDuties = lmbtbl.DataBodyRange(i, lmbtbl.ListColumns("Max Duties").Index).Value
            Dim currDuties As Long
            currDuties = lmbtbl.DataBodyRange(i, lmbtbl.ListColumns("Duties Counter").Index).Value
            ' Check if the staff already reached his max duties
            If currDuties >= maxDuties Then GoTo SkipStaff
            If IsWorkingOnSameDay(r, staffName) Then GoTo SkipStaff
            
            ' Assign from top
            If wsRoster.Cells(r, LMB_COL).Value = "" Then
                wsRoster.Cells(r, LMB_COL).Value = staffName
                Call IncrementDutiesCounter(staffName)
                Exit For
            End If
        
SkipStaff:
        Next i
        
SkipDay:
        Next r
    
    ' Step 3: Reassign for unfilled slots
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
            Call ReassignLoanMailBoxDuties
        Else
            Exit Do
        End If
        reassignmentAttempts = reassignmentAttempts + 1
        If reassignmentAttempts > MAX_ATTEMPTS Then
            Debug.Print "Max reassignment attempts reached. Assigning remaining staff to unfilled slots."
            ' Fallback: Assign all remaining eligible staff to unfilled slots
            Dim eligibleStaffList As Collection
            Set eligibleStaffList = New Collection
            For i = 1 To lmbtbl.ListRows.count
                staffName = lmbtbl.DataBodyRange(i, lmbtbl.ListColumns("Name").Index).Value
                maxDuties = lmbtbl.DataBodyRange(i, lmbtbl.ListColumns("Max Duties").Index).Value
                currDuties = lmbtbl.DataBodyRange(i, lmbtbl.ListColumns("Duties Counter").Index).Value
                If UCase(lmbtbl.DataBodyRange(i, lmbtbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" And _
                   currDuties < maxDuties Then
                    eligibleStaffList.Add staffName
                End If
            Next i
            
            Dim unfilledSlotsList As Collection
            Set unfilledSlotsList = New Collection
            For r = START_ROW To last_row_roster
                If wsRoster.Cells(r, DAY_COL).Value <> "Sat" And _
                   wsRoster.Cells(r, LMB_COL).Value = "" Then
                    unfilledSlotsList.Add r
                End If
            Next r
            
            ' Assign each eligible staff to an unfilled slot
            For i = 1 To eligibleStaffList.count
                If i <= unfilledSlotsList.count Then
                    staffName = eligibleStaffList(i)
                    r = unfilledSlotsList(i)
                    If Not IsWorkingOnSameDay(r, staffName) Then
                        wsRoster.Cells(r, LMB_COL).Value = staffName
                        wsRoster.Cells(r, LMB_COL).Interior.Color = vbYellow ' Highlight yellow (RGB 255, 255, 0)
                        Call IncrementDutiesCounter(staffName)
                        Debug.Print "Fallback: Assigned " & staffName & " to row " & r & " (highlighted yellow)"
                    Else
                        Debug.Print "Fallback: Skipped " & staffName & " at row " & r & " due to MOR/AFT/AOH conflict"
                    End If
                End If
            Next i
            Exit Do
        End If
    Loop
    
    MsgBox "Duties assignment completed!", vbInformation
End Sub

' Helper to count unfilled Loan Mailbox slots
Function CountUnfilledSlots(startRow As Long, endRow As Long) As Long
    Dim r As Long
    Dim count As Long
    count = 0
    For r = startRow To endRow
        If wsRoster.Cells(r, DAY_COL).Value <> "Sat" And _
           wsRoster.Cells(r, LMB_COL).Value = "" Then
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
    For i = 1 To lmbtbl.ListRows.count
        Dim staffName As String
        Dim maxDuties As Long
        Dim currDuties As Long
        staffName = lmbtbl.DataBodyRange(i, lmbtbl.ListColumns("Name").Index).Value
        maxDuties = lmbtbl.DataBodyRange(i, lmbtbl.ListColumns("Max Duties").Index).Value
        currDuties = lmbtbl.DataBodyRange(i, lmbtbl.ListColumns("Duties Counter").Index).Value
        If UCase(lmbtbl.DataBodyRange(i, lmbtbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" And _
           currDuties < maxDuties Then
            count = count + 1
        End If
    Next i
    CountEligibleStaff = count
End Function

Sub ReassignLoanMailBoxDuties()
    Dim r As Long
    Dim i As Long
    Dim staffName As String
    Dim maxDuties As Long
    Dim currDuties As Long
    Dim eligibleStaff As String
    Dim swapCandidate As String
    Dim emptyRow As Long ' The empty original slot where eligibleStaff couldn't be assigned
    
    ' Find the first eligible staff (currDuties < maxDuties)
    eligibleStaff = GetFirstEligibleStaff
    If eligibleStaff = "" Then
        Debug.Print "No eligible staff found for LMB reassignment."
        Exit Sub
    End If
    Debug.Print "Eligible staff for LMB reassignment: " & eligibleStaff
    
    ' Assume emptyRow is determined earlier (e.g., the last unfilled slot)
    emptyRow = FindLastUnfilledSlot(START_ROW, last_row_roster)
    If emptyRow = 0 Then Exit Sub ' No empty slot to reassign
    
    ' Loop through all rows to find a swap opportunity
    For r = START_ROW To last_row_roster
        If wsRoster.Cells(r, DAY_COL).Value = "Sat" Or _
           wsRoster.Cells(r, LMB_COL).Value = "CLOSED" Or _
           wsRoster.Cells(r, LMB_COL).Value = "" Then
            GoTo NextRow
        End If
        
        swapCandidate = wsRoster.Cells(r, LMB_COL).Value
        If swapCandidate <> "" And swapCandidate <> eligibleStaff Then
            ' Check if swap is allowed (no conflict in destination and original slot)
            If Not IsWorkingOnSameDay(r, eligibleStaff) And Not IsWorkingOnSameDay(emptyRow, swapCandidate) Then
                ' Perform swap
                wsRoster.Cells(r, LMB_COL).Value = eligibleStaff
                Call IncrementDutiesCounter(eligibleStaff)
                Debug.Print "Assigned " & eligibleStaff & " to row " & r
                
                wsRoster.Cells(emptyRow, LMB_COL).Value = swapCandidate
                Debug.Print "Assigned " & swapCandidate & " to row " & emptyRow
                
                Exit Sub ' Exit after successful swap
            Else
                Debug.Print "Swap not possible at row " & r & " due to conflict (dest: " & r & ", orig: " & emptyRow & ")"
            End If
        End If
NextRow:
    Next r
End Sub

' Helper to find the last unfilled slot
Function FindLastUnfilledSlot(startRow As Long, endRow As Long) As Long
    Dim r As Long
    FindLastUnfilledSlot = 0
    For r = startRow To endRow
        If wsRoster.Cells(r, DAY_COL).Value <> "Sat" And _
           wsRoster.Cells(r, LMB_COL).Value = "" Then
            FindLastUnfilledSlot = r
        End If
    Next r
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

    Debug.Print "=== Checking Eligible Rows ==="
    Debug.Print "WorkDays:"
    For j = LBound(workDays) To UBound(workDays)
        Debug.Print "[" & j & "]: " & workDays(j)
    Next j
    
    For r = START_ROW To last_row_roster
        dayName = Trim(wsRoster.Cells(r, DAY_COL).Value)
        ' Debug: show what day we are checking
        Debug.Print "Row " & r & ": " & dayName
        
        ' Skip if already filled
        If Not IsEmpty(wsRoster.Cells(r, LMB_COL)) Then
            Debug.Print "  -> Skipped (Already Assigned)"
            GoTo SkipRow
        End If
        
        ' Check if the day is in workDays
        For j = LBound(workDays) To UBound(workDays)
            If dayName = workDays(j) Then
                eligibleRows.Add r
                Debug.Print "  -> Added (Matched with " & workDays(j) & ")"
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

    ' Search for the staff name
    Set foundCell = lmbtbl.ListColumns("Name").DataBodyRange.Find( _
        What:=staffName, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        ' Get relative row index in the table
        rowIdx = foundCell.row - lmbtbl.HeaderRowRange.row

        ' Increment Duties Counter
        With lmbtbl.ListRows(rowIdx).Range.Cells(lmbtbl.ListColumns("Duties Counter").Index)
            .Value = .Value + 1
        End With
    Else
        MsgBox "Staff '" & staffName & "' not found in table.", vbExclamation
    End If
End Sub

Sub DecrementDutiesCounter(staffName As String)
    Dim rowIdx As Variant
    Dim foundCell As Range

    ' Search for the staff name
    Set foundCell = lmbtbl.ListColumns("Name").DataBodyRange.Find( _
        What:=staffName, LookIn:=xlValues, LookAt:=xlWhole)

    If Not foundCell Is Nothing Then
        ' Get relative row index in the table
        rowIdx = foundCell.row - lmbtbl.HeaderRowRange.row

        ' Decrement Duties Counter
        With lmbtbl.ListRows(rowIdx).Range.Cells(lmbtbl.ListColumns("Duties Counter").Index)
            .Value = .Value - 1
            If .Value < 0 Then .Value = 0 ' Prevent negative values
        End With
    Else
        MsgBox "Staff '" & staffName & "' not found in table.", vbExclamation
    End If
End Sub

Function IsWorkingOnSameDay(row As Long, staffName As String) As Boolean
    If wsRoster.Cells(row, MOR_COL).Value = staffName Then
        IsWorkingOnSameDay = True
        Exit Function
    End If
    
    If wsRoster.Cells(row, AFT_COL).Value = staffName Then
        IsWorkingOnSameDay = True
        Exit Function
    End If
    
    If wsRoster.Cells(row, AOH_COL).Value = staffName Then
        IsWorkingOnSameDay = True
        Exit Function
    End If
        
    IsWorkingOnSameDay = False
End Function

' Helper to get the first eligible staff (duties < max)
Function GetFirstEligibleStaff() As String
    Dim i As Long
    For i = 1 To lmbtbl.ListRows.count
        Dim staffName As String
        Dim maxDuties As Long
        Dim currDuties As Long
        staffName = lmbtbl.DataBodyRange(i, lmbtbl.ListColumns("Name").Index).Value
        maxDuties = lmbtbl.DataBodyRange(i, lmbtbl.ListColumns("Max Duties").Index).Value
        currDuties = lmbtbl.DataBodyRange(i, lmbtbl.ListColumns("Duties Counter").Index).Value
        If UCase(lmbtbl.DataBodyRange(i, lmbtbl.ListColumns("Availability Type").Index).Value) <> "SPECIFIC DAYS" And _
           currDuties < maxDuties Then
            GetFirstEligibleStaff = staffName
            Exit Function
        End If
    Next i
    GetFirstEligibleStaff = ""
End Function



Sub InsertStaffOld()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim morningtbl As ListObject
    Dim newRow As ListRow
    Dim staffName As String, dept As String
    Dim availType As String, workDays As String, percentage As String
    Dim checkRow As Long
    Dim specificDaysTbl As ListObject
    Dim specificRow As ListRow
    
    ' Set worksheet and tables
    Set ws = ThisWorkbook.Sheets("Morning PersonnelList")
    If ws Is Nothing Then
        MsgBox "Worksheet 'Morning PersonnelList' not found.", vbExclamation
        Exit Sub
    End If
    On Error Resume Next
    Set morningtbl = ws.ListObjects("MorningMainList")
    Set specificDaysTbl = ws.ListObjects("MorningSpecificDaysWorkingStaff")
    On Error GoTo ErrHandler
    If morningtbl Is Nothing Then
        MsgBox "Table 'MorningMainList' not found.", vbExclamation
        Exit Sub
    End If
    If specificDaysTbl Is Nothing And UCase(Trim(ws.Range("D7").Value)) = "SPECIFIC DAYS" Then
        MsgBox "Table 'MorningSpecificDaysWorkingStaff' not found.", vbExclamation
        Exit Sub
    End If

    ' Read input values
    staffName = UCase(Trim(ws.Range("D5").Value)) ' Name
    dept = Trim(ws.Range("D6").Value)             ' Department
    availType = UCase(Trim(ws.Range("D7").Value)) ' Availability Type
    workDays = Trim(ws.Range("D8").Value)         ' Working Days
    percentage = Trim(ws.Range("D9").Value)       ' Duties Percentage

    ' Auto-fill logic based on Availability Type
    If availType = "ALL DAYS" Then
        percentage = "100"
        workDays = ""
    ElseIf availType = "SPECIFIC DAYS" Then
        If workDays = "" Then
            MsgBox "Please enter Working Days for Specific Days availability.", vbExclamation
            Exit Sub
        End If
        If percentage = "" Or Not IsNumeric(percentage) Or Val(percentage) <= 0 Or Val(percentage) > 100 Then
            MsgBox "Please enter a valid Duties Percentage (1-100).", vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "Availability Type must be 'All Days' or 'Specific Days'.", vbExclamation
        Exit Sub
    End If

    ' Validation
    If Len(Trim(staffName)) = 0 Or Len(Trim(dept)) = 0 Then
        MsgBox "Please fill in Name and Department.", vbExclamation
        Exit Sub
    End If
    
    ' Check for duplicate names
    For checkRow = 1 To morningtbl.ListRows.count
        If UCase(Trim(morningtbl.ListRows(checkRow).Range.Cells(1, GetColumnIndex(morningtbl, "Name")).Value)) = staffName Then
            MsgBox "This staff name already exists.", vbExclamation
            Exit Sub
        End If
    Next checkRow

    ' Insert new row based on Availability Type
    If availType = "SPECIFIC DAYS" Then
        Set newRow = morningtbl.ListRows.Add(AlwaysInsert:=True) ' Insert at top
    Else ' ALL DAYS
        Set newRow = morningtbl.ListRows.Add(AlwaysInsert:=True) ' Insert at bottom
    End If
    
    ' Populate the new row
    With newRow.Range
        Dim nameIndex As Long, deptIndex As Long, availIndex As Long
        Dim percIndex As Long, maxIndex As Long, counterIndex As Long
        
        nameIndex = GetColumnIndex(morningtbl, "Name")
        deptIndex = GetColumnIndex(morningtbl, "Department")
        availIndex = GetColumnIndex(morningtbl, "Availability Type")
        percIndex = GetColumnIndex(morningtbl, "Duties Percentage (%)")
        maxIndex = GetColumnIndex(morningtbl, "Max Duties")
        counterIndex = GetColumnIndex(morningtbl, "Duties Counter")
        
        If nameIndex = -1 Or deptIndex = -1 Or availIndex = -1 Or percIndex = -1 Or maxIndex = -1 Or counterIndex = -1 Then
            MsgBox "Required columns not found in 'MorningMainList'.", vbExclamation
            newRow.Delete
            Exit Sub
        End If
        
        .Cells(1, nameIndex).Value = staffName
        .Cells(1, deptIndex).Value = dept
        .Cells(1, availIndex).Value = availType
        .Cells(1, percIndex).Value = Val(percentage)
        .Cells(1, counterIndex).Value = 0
        ' Do not set Max Duties here
    End With

    ' Handle specific days workers
    If availType = "SPECIFIC DAYS" Then
        Set specificRow = specificDaysTbl.ListRows.Add(AlwaysInsert:=True)
        With specificRow.Range
            Dim specNameIndex As Long, specDaysIndex As Long
            specNameIndex = GetColumnIndex(specificDaysTbl, "Name")
            specDaysIndex = GetColumnIndex(specificDaysTbl, "Working Days")
            If specNameIndex = -1 Or specDaysIndex = -1 Then
                MsgBox "Columns 'Name' or 'Working Days' not found in 'MorningSpecificDaysWorkingStaff'.", vbExclamation
                specificRow.Delete
                newRow.Delete
                Exit Sub
            End If
            .Cells(1, specNameIndex).Value = staffName
            .Cells(1, specDaysIndex).Value = workDays
        End With
    End If

    ' Update Max Duties for all rows
    RunMaxDutiesMorning

    ' Delete last empty row if it exists (for SPECIFIC DAYS)
    If availType = "SPECIFIC DAYS" Then
        Dim lastRow As ListRow
        If morningtbl.ListRows.count > 0 Then ' Check even if only one row
            Set lastRow = morningtbl.ListRows(morningtbl.ListRows.count)
            ' Check if key columns (Name, Department, Availability Type) are empty
            nameIndex = GetColumnIndex(morningtbl, "Name")
            deptIndex = GetColumnIndex(morningtbl, "Department")
            availIndex = GetColumnIndex(morningtbl, "Availability Type")
            If nameIndex <> -1 And deptIndex <> -1 And availIndex <> -1 Then
                Dim lastRowContent As String
                lastRowContent = "Last row contents: Name='" & Trim(lastRow.Range.Cells(1, nameIndex).Value) & _
                                 "', Department='" & Trim(lastRow.Range.Cells(1, deptIndex).Value) & _
                                 "', Availability Type='" & Trim(lastRow.Range.Cells(1, availIndex).Value) & "'"
                If Trim(lastRow.Range.Cells(1, nameIndex).Value) = "" And _
                   Trim(lastRow.Range.Cells(1, deptIndex).Value) = "" And _
                   Trim(lastRow.Range.Cells(1, availIndex).Value) = "" Then
                    lastRow.Delete
                    Debug.Print "Deleted last row: " & lastRowContent
                Else
                    Debug.Print "Last row not deleted (not empty): " & lastRowContent
                End If
            Else
                Debug.Print "Column index error: Name=" & nameIndex & ", Dept=" & deptIndex & ", Avail=" & availIndex
            End If
        End If
    End If

    ' Clear input
    ws.Range("D5:D9").ClearContents

    MsgBox "Staff added and Max Duties calculated successfully!", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    If Not newRow Is Nothing Then newRow.Delete
    If Not specificRow Is Nothing Then specificRow.Delete
    Exit Sub
End Sub
' Helper function to get column index
Private Function GetColumnIndex(tbl As ListObject, columnName As String) As Long
    On Error Resume Next
    GetColumnIndex = tbl.ListColumns(columnName).Index
    If Err.Number <> 0 Then GetColumnIndex = -1
    On Error GoTo 0
End Function


Sub InsertStaff(dutyType As String)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim staffName As String, dept As String
    Dim availType As String, workDays As String, percentage As String
    Dim checkRow As Long
    Dim specificDaysTbl As ListObject
    Dim specificRow As ListRow
    
    ' Set worksheet and tables based on dutyType
    Select Case UCase(dutyType)
        Case "LOANMAILBOX"
            Set ws = ThisWorkbook.Sheets("Loan Mail Box PersonnelList")
            Set tbl = ws.ListObjects("LoanMailBoxMainList")
            Set specificDaysTbl = ws.ListObjects("LoanMailBoxSpecificDaysWorkingStaff")
        Case "MORNING"
            Set ws = ThisWorkbook.Sheets("Morning PersonnelList")
            Set tbl = ws.ListObjects("MorningMainList")
            Set specificDaysTbl = ws.ListObjects("MorningSpecificDaysWorkingStaff")
        Case "AFTERNOON"
            Set ws = ThisWorkbook.Sheets("Afternoon PersonnelList")
            Set tbl = ws.ListObjects("AfternoonMainList")
            Set specificDaysTbl = ws.ListObjects("AfternoonSpecificDaysWorkingStaff")
        Case "AOH"
            Set ws = ThisWorkbook.Sheets("AOH PersonnelList")
            Set tbl = ws.ListObjects("AOHMainList")
            Set specificDaysTbl = ws.ListObjects("AOHSpecificDaysWorkingStaff")
        Case "SAT_AOH"
            Set ws = ThisWorkbook.Sheets("Sat AOH PersonnelList")
            Set tbl = ws.ListObjects("SatAOHMainList")
            ' No specificDaysTbl for Sat AOH
        Case Else
            MsgBox "Invalid duty type. Use 'LoanMailBox', 'Morning', 'Afternoon', 'AOH', or 'Sat_AOH'.", vbExclamation
            Exit Sub
    End Select
    
    ' Validate worksheet and table
    If ws Is Nothing Then
        MsgBox "Worksheet for " & dutyType & " not found.", vbExclamation
        Exit Sub
    End If
    If tbl Is Nothing Then
        MsgBox "Table 'MainList' not found on '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If
    If specificDaysTbl Is Nothing And UCase(Trim(ws.Range("D7").Value)) = "SPECIFIC DAYS" And UCase(dutyType) <> "SAT_AOH" Then
        MsgBox "Table 'SpecificDaysWorkingStaff' not found on '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If

    ' Unprotect the worksheet (unlock)
    ws.Unprotect
    
    ' to remove the line breaks before checking for name duplication
    Dim cleanInput As String
    cleanInput = RemoveLineBreaks(ws.Range("D5").Value)
    ws.Range("D5").Value = cleanInput
    
    ' Read input values from the unlocked data entry range
    staffName = UCase(Trim(cleanInput)) ' Name (already cleaned)
    dept = Trim(RemoveLineBreaks(ws.Range("D6").Value)) ' Department
    availType = UCase(Trim(RemoveLineBreaks(ws.Range("D7").Value))) ' Availability Type
    workDays = Trim(RemoveLineBreaks(ws.Range("D8").Value)) ' Working Days
    percentage = Trim(RemoveLineBreaks(ws.Range("D9").Value)) ' Duties Percentage

    ' Auto-fill logic based on Availability Type
    If availType = "ALL DAYS" Then
        percentage = "100"
        workDays = ""
    ElseIf availType = "SPECIFIC DAYS" Then
        If workDays = "" Then
            MsgBox "Please enter Working Days for Specific Days availability.", vbExclamation
            GoTo ReprotectAndExit
        End If
    End If
    
    ' Validate percentage
    If percentage = "" Or Not IsNumeric(percentage) Or Val(percentage) <= 0 Or Val(percentage) > 100 Then
        MsgBox "Please enter a valid Duties Percentage (1-100).", vbExclamation
        GoTo ReprotectAndExit
    End If

    If Len(Trim(staffName)) = 0 Or Len(Trim(dept)) = 0 Then
        MsgBox "Please fill in Name and Department.", vbExclamation
        GoTo ReprotectAndExit
    End If
    
    ' Check for duplicate names (now handles line breaks properly)
    For checkRow = 1 To tbl.ListRows.count
        Dim existingName As String
        existingName = UCase(Trim(RemoveLineBreaks(tbl.ListRows(checkRow).Range.Cells(1, GetColumnIndex(tbl, "Name")).Value)))
        If existingName = staffName Then
            MsgBox "This staff name already exists.", vbExclamation
            GoTo ReprotectAndExit
        End If
    Next checkRow
    
    Set newRow = tbl.ListRows.Add(AlwaysInsert:=True)
    With newRow.Range
        Dim nameIndex As Long, deptIndex As Long, availIndex As Long
        Dim percIndex As Long, maxIndex As Long, counterIndex As Long
        
        nameIndex = GetColumnIndex(tbl, "Name")
        deptIndex = GetColumnIndex(tbl, "Department")
        availIndex = GetColumnIndex(tbl, "Availability Type")
        percIndex = GetColumnIndex(tbl, "Duties Percentage (%)")
        maxIndex = GetColumnIndex(tbl, "Max Duties")
        counterIndex = GetColumnIndex(tbl, "Duties Counter")
        
        If nameIndex = -1 Or deptIndex = -1 Or counterIndex = -1 Or availIndex = -1 Or percIndex = -1 Or maxIndex = -1 Then
            MsgBox "Required columns not found in '" & tbl.Name & "'.", vbExclamation
            newRow.Delete
            GoTo ReprotectAndExit
        End If
        
        .Cells(1, nameIndex).Value = staffName
        .Cells(1, deptIndex).Value = dept
        .Cells(1, availIndex).Value = availType
        .Cells(1, percIndex).Value = Val(percentage)
        .Cells(1, counterIndex).Value = 0
    End With

    ' Handle specific days workers table
    If availType = "SPECIFIC DAYS" Then
        Set specificRow = specificDaysTbl.ListRows.Add(AlwaysInsert:=True)
        With specificRow.Range
            Dim specNameIndex As Long, specDaysIndex As Long
            specNameIndex = GetColumnIndex(specificDaysTbl, "Name")
            specDaysIndex = GetColumnIndex(specificDaysTbl, "Working Days")
            If specNameIndex = -1 Or specDaysIndex = -1 Then
                MsgBox "Required columns not found in specific days table.", vbExclamation
                specificRow.Delete
                newRow.Delete
                GoTo ReprotectAndExit
            End If
            .Cells(1, specNameIndex).Value = staffName
            .Cells(1, specDaysIndex).Value = workDays
        End With
    End If

    CalculateMaxDuties.CalculateMaxDuties dutyType

    ' Clear input of data entry
    ws.Range("D5:D9").ClearContents

    MsgBox "Staff added successfully for " & dutyType & "!", vbInformation
    GoTo ReprotectAndExit

ReprotectAndExit:
    ' Reprotect the worksheet and lock table ranges
    With ws
        On Error Resume Next
        .Unprotect
        If Not tbl Is Nothing Then .ListObjects(tbl.Name).Range.Locked = True
        If Not specificDaysTbl Is Nothing Then .ListObjects(specificDaysTbl.Name).Range.Locked = True
        .Range("D5:D9").Locked = False
        .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                 AllowFiltering:=True, AllowSorting:=True, AllowUsingPivotTables:=True
        On Error GoTo 0
    End With
    ws.Activate
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    If Not newRow Is Nothing Then newRow.Delete
    If Not specificRow Is Nothing Then specificRow.Delete
    GoTo ReprotectAndExit
End Sub

' NEW: Function to remove line breaks from text
Private Function RemoveLineBreaks(inputText As String) As String
    If Len(inputText) = 0 Then
        RemoveLineBreaks = ""
        Exit Function
    End If
    
    ' Replace various types of line breaks with space
    RemoveLineBreaks = Replace(inputText, Chr(10), " ")  ' Line feed
    RemoveLineBreaks = Replace(RemoveLineBreaks, Chr(13), " ")  ' Carriage return
    RemoveLineBreaks = Replace(RemoveLineBreaks, vbCrLf, " ")  ' CR+LF
    RemoveLineBreaks = Replace(RemoveLineBreaks, vbNewLine, " ")  ' New line
    RemoveLineBreaks = Replace(RemoveLineBreaks, vbLf, " ")  ' Line feed
    RemoveLineBreaks = Replace(RemoveLineBreaks, vbCr, " ")  ' Carriage return
    
    ' Clean up multiple spaces
    RemoveLineBreaks = WorksheetFunction.Trim(RemoveLineBreaks)
End Function

' Helper function to get column index
Private Function GetColumnIndex(tbl As ListObject, columnName As String) As Long
    On Error Resume Next
    GetColumnIndex = tbl.ListColumns(columnName).Index
    If Err.Number <> 0 Then GetColumnIndex = -1
    On Error GoTo 0
End Function

' Wrapper subroutines for different shifts
Sub RunInsertStaffLMB()
    InsertStaff "LoanMailBox"
End Sub

Sub RunInsertStaffMorning()
    InsertStaff "Morning"
End Sub

Sub RunInsertStaffAfternoon()
    InsertStaff "Afternoon"
End Sub

Sub RunInsertStaffAOH()
    InsertStaff "AOH"
End Sub

Sub RunInsertStaffSatAOH()
    InsertStaff "Sat_AOH"
End Sub

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


Sub GenerateMorningShiftAnalysis()
    Dim wsPersonnel As Worksheet, wsRoster As Worksheet, wsAnalysis As Worksheet
    Dim tbl As ListObject
    Dim nameList As Range, dutyCounterList As Range
    Dim lastRow As Long, i As Long
    Dim dict As Object
    Dim empName As String
    Dim MOR_COL As Long: MOR_COL = 6 ' Morning Shift column
    Dim START_ROW As Long: START_ROW = 6 ' Data start row

    ' Set sheets
    Set wsPersonnel = Sheets("Morning PersonnelList")
    Set wsRoster = Sheets("Roster")
    
    ' Create or clear analysis sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("MorningAnalysis").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsAnalysis = Sheets.Add(After:=Sheets(Sheets.count))
    wsAnalysis.Name = "MorningAnalysis"
    
    ' Get the table
    Set tbl = wsPersonnel.ListObjects("MorningMainList")
    Set nameList = tbl.ListColumns("Name").DataBodyRange
    Set dutyCounterList = tbl.ListColumns("Duties Counter").DataBodyRange

    ' Header row for analysis
    With wsAnalysis
        .Range("A1").Value = "Name"
        .Range("B1").Value = "System Counter"
        .Range("C1").Value = "Actual Counter"
        .Range("D1").Value = "Difference"
    End With

    ' Copy names and system counter to analysis sheet
    For i = 1 To nameList.Rows.count
        wsAnalysis.Cells(i + 1, 1).Value = nameList.Cells(i, 1).Value
        wsAnalysis.Cells(i + 1, 2).Value = dutyCounterList.Cells(i, 1).Value
    Next i

    ' Create dictionary to count actual appearances
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To nameList.Rows.count
        empName = nameList.Cells(i, 1).Value
        dict(empName) = 0
    Next i

    ' Walk through the roster and count appearances
    lastRow = wsRoster.Cells(wsRoster.Rows.count, MOR_COL).End(xlUp).row
    For i = START_ROW To lastRow
        empName = Trim(wsRoster.Cells(i, MOR_COL).Value)
        If dict.Exists(empName) Then
            dict(empName) = dict(empName) + 1
        End If
    Next i

    ' Write actual counter and difference to sheet
    For i = 2 To nameList.Rows.count + 1
        empName = wsAnalysis.Cells(i, 1).Value
        wsAnalysis.Cells(i, 3).Value = dict(empName) ' Actual Counter
        wsAnalysis.Cells(i, 4).FormulaR1C1 = "=RC[-2]-RC[-1]" ' System - Actual
    Next i

    MsgBox "Morning shift analysis generated in 'MorningAnalysis' sheet.", vbInformation
End Sub
Sub ResetAllCounters()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim counterColumn As ListColumn
    Dim tblNames As Variant
    Dim sheetNames As Variant
    Dim i As Integer
    Dim wsRoster As Worksheet
    
    ' Define your table and corresponding sheet names
    sheetNames = Array("Loan Mail Box PersonnelList", "Morning PersonnelList", _
                       "Afternoon PersonnelList", "AOH PersonnelList", "Sat AOH PersonnelList")
    tblNames = Array("LoanMailBoxMainList", "MorningMainList", "AfternoonMainList", _
                     "AOHMainList", "SatAOHMainList")
                     
    'if table is empty exit sub
    
    ' Loop through all specified tables and reset Duties Counter column to 0
    For i = 0 To UBound(sheetNames)
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        'ws.Unprotect
        Set tbl = ws.ListObjects(tblNames(i))
        
        ' Check if the table is empty
        If tbl.ListRows.count = 0 Then
            Debug.Print "Table " & tblNames(i) & " on sheet " & sheetNames(i) & " is empty. Skipping reset."
            GoTo NextSheet
        End If
        
        Set counterColumn = tbl.ListColumns("Duties Counter")
        
        Dim j As Long
        For j = 1 To counterColumn.DataBodyRange.Rows.count
            counterColumn.DataBodyRange.Cells(j, 1).Value = 0
        Next j
NextSheet:
    Next i
        ' Reprotect the worksheet and lock table ranges
        'With ws
        '    If Not tbl Is Nothing Then
        '        .ListObjects(tbl.Name).Range.Locked = True
        '    End If
        '    .Range("D5:D9").Locked = False ' keep data entry remains unlocked
        '    .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                     AllowFiltering:=True, AllowSorting:=True, AllowUsingPivotTables:=True
        'End With
    
    ' Activate Roster sheet
    Set wsRoster = Sheets("Roster")
    wsRoster.Activate


End Sub




Public Sub countMorningOrAfternoonOrLMBSlotsSub(Worksheet As String, ByRef result As Long)
    Dim startDate As Date
    Dim endDate As Date
    Dim currentDate As Date
    Dim r As Long
    Dim holidayCell As Range
    Dim isHoliday As Boolean
    Dim ws As Worksheet
    Dim last_row_roster As Long
    
    Set ws = ThisWorkbook.Sheets(Worksheet)
    
    ' Initialize counter
    result = 0
    
    ' Get the fixed start and end dates from H3 and K3 on the provided worksheet
    startDate = ws.Range("H3").Value
    endDate = ws.Range("K3").Value
    If Not IsDate(startDate) Or Not IsDate(endDate) Then
        Debug.Print "Invalid dates in H3 or K3: " & ws.Range("H3").Value & ", " & ws.Range("K3").Value
        Exit Sub
    End If
    
    ' Ensure startDate is before or equal to endDate
    If startDate > endDate Then
        Dim tempDate As Date
        tempDate = startDate
        startDate = endDate
        endDate = tempDate
    End If
    
    'Find last row of roster
    If ws.Cells(2, 10).Value = "Jan-Jun" And ws.Cells(2, 13).Value Mod 4 = 0 Then
        last_row_roster = 187
    ElseIf ws.Cells(2, 10).Value = "Jan-Jun" Then
        last_row_roster = 186
    Else
        last_row_roster = 189
    End If
    
    ' Loop through each row in column B
    For r = 6 To last_row_roster
        currentDate = ws.Cells(r, 2).Value ' Date from column B
        If IsDate(Trim(currentDate)) Then
            ' Check if the date is within the custom period
            If currentDate >= startDate And currentDate <= endDate Then
                ' Check if it's not Sunday (1) or Saturday (7)
                If Weekday(currentDate) <> 1 And Weekday(currentDate) <> 7 Then
                    ' Check if it's not a public holiday using the named range
                    isHoliday = False
                    For Each holidayCell In Range("Settings_Holidays")
                        If IsDate(holidayCell.Value) Then
                            If dateValue(currentDate) = dateValue(holidayCell.Value) Then
                                isHoliday = True
                                Exit For
                            End If
                        End If
                    Next holidayCell
                    ' If not a holiday, increment counter
                    If Not isHoliday Then
                        result = result + 1
                    End If
                End If
            End If
        End If
    Next r
End Sub
Public Sub countAOHslotsSub(Worksheet As String, ByRef result As Long)
    Dim startDate As Date
    Dim endDate As Date
    Dim currentDate As Date
    Dim r As Long
    Dim holidayCell As Range
    Dim isHoliday As Boolean
    Dim ws As Worksheet
    Dim last_row_roster As Long
    
    ' Initialize counter
    result = 0
    
    ' Set worksheet reference
    Set ws = ThisWorkbook.Sheets(Worksheet)
    
    ' Get the fixed start and end dates from H3 and K3
    startDate = ws.Range("H3").Value
    endDate = ws.Range("K3").Value
    If Not IsDate(startDate) Or Not IsDate(endDate) Then Exit Sub ' Exit if dates are invalid
    
    ' Ensure startDate is before or equal to endDate
    If startDate > endDate Then
        Dim tempDate As Date
        tempDate = startDate
        startDate = endDate
        endDate = tempDate
    End If
    
    ' Find last row of roster
    If ws.Cells(2, 10).Value = "Jan-Jun" And ws.Cells(2, 13).Value Mod 4 = 0 Then
        last_row_roster = 187
    ElseIf ws.Cells(2, 10).Value = "Jan-Jun" Then
        last_row_roster = 186
    Else
        last_row_roster = 189
    End If
    
    ' Loop through each row in column B
    For r = 6 To last_row_roster
        currentDate = ws.Cells(r, 2).Value ' Date from column B
        If IsDate(Trim(currentDate)) Then
            ' Check if the date is within the custom period
            If currentDate >= startDate And currentDate <= endDate Then
                ' Check if it's not Sunday (1) or Saturday (7)
                If Weekday(currentDate) <> 1 And Weekday(currentDate) <> 7 Then
                    ' Check if it's not a public holiday using the named range
                    isHoliday = False
                    For Each holidayCell In Range("Settings_Holidays")
                        If IsDate(holidayCell.Value) Then
                            If dateValue(currentDate) = dateValue(holidayCell.Value) Then
                                isHoliday = True
                                Exit For
                            End If
                        End If
                    Next holidayCell
                    ' Check if the corresponding marker in column A is "sem time"
                    If Not isHoliday And LCase(Trim(ws.Cells(r, 1).Value)) = "sem time" Then
                        result = result + 1
                    End If
                End If
            End If
        End If
    Next r
End Sub
Public Sub countSatAOHSub(Worksheet As String, ByRef result As Long)
    Dim startDate As Date
    Dim endDate As Date
    Dim currentDate As Date
    Dim ws As Worksheet
    Dim holidayCell As Range
    Dim holidaySaturdays As Long
    
    ' Initialize counter
    result = 0
    
    ' Set worksheet reference
    Set ws = ThisWorkbook.Sheets(Worksheet)
    
    ' Get the start and end dates from H3 and K3
    startDate = ws.Range("H3").Value
    endDate = ws.Range("K3").Value
    If Not IsDate(startDate) Or Not IsDate(endDate) Then Exit Sub ' Exit if dates are invalid
    
    ' Ensure startDate is before or equal to endDate
    If startDate > endDate Then
        Dim tempDate As Date
        tempDate = startDate
        startDate = endDate
        endDate = tempDate
    End If
    
    ' Count Saturdays in the date range
    currentDate = startDate
    Do While currentDate <= endDate
        If Weekday(currentDate) = 7 Then ' 7 = Saturday
            result = result + 1
        End If
        currentDate = currentDate + 1
    Loop
    
    ' Subtract Saturdays that are public holidays
    holidaySaturdays = 0
    For Each holidayCell In Range("Settings_Holidays")
        If IsDate(holidayCell.Value) Then
            If dateValue(holidayCell.Value) >= startDate And dateValue(holidayCell.Value) <= endDate And Weekday(holidayCell.Value) = 7 Then
                holidaySaturdays = holidaySaturdays + 1
            End If
        End If
    Next holidayCell
    result = result - holidaySaturdays
End Sub

Option Explicit
Const password As String = "swapswap"
' Main subroutine to handle staff swapping
Sub SwapStaff()
    
    Dim wsPersonnel As Worksheet
    Dim wsSwap As Worksheet
    Dim slotCols As Variant
    Dim dateRange As Range
    Dim oriName As String
    Dim newName As String
    Dim dateCell As Range
    Dim enteredPassword As String
    
    enteredPassword = InputBox("Please enter the password to proceed with the swap:", "Password Authentication")
    If enteredPassword <> password Then
        MsgBox "Incorrect password. Swap operation cancelled.", vbCritical
        Exit Sub
    End If
    
    InitializeWorksheets wsPersonnel, wsSwap
    
    ' Unprotect worksheets with password
    On Error Resume Next
    wsRoster.Unprotect "rostering2025"
    wsPersonnel.Unprotect "rostering2025"
    On Error GoTo 0
    
    GetSwapNames wsSwap, oriName, newName
    ValidateNames oriName, newName
    
    Set dateRange = GetDateRange
    If dateRange Is Nothing Then Exit Sub
    If Not IsValidDateColumn(dateRange) Then Exit Sub
    
    slotCols = Array(LMB_COL, MOR_COL, AFT_COL, AOH_COL, SAT_AOH_COL1, SAT_AOH_COL2)
    Dim r As Long
    For Each dateCell In dateRange
        r = dateCell.row
        Dim oriNameFound As Boolean
        CheckOriginalNameExists wsRoster, r, slotCols, oriName, oriNameFound
        If Not oriNameFound Then
            DisplayError "Error: " & oriName & " not found on date " & wsRoster.Cells(r, DATE_COL).Value & ". Swap not allowed.", vbExclamation
        Else
            Dim nameExists As Boolean
            CheckNewNameExists wsRoster, r, slotCols, newName, nameExists
            If nameExists Then
                DisplayError "Error: " & newName & " already exists on date " & wsRoster.Cells(r, DATE_COL).Value & ". Swap not allowed.", vbExclamation
            Else
                PerformSwap wsRoster, r, slotCols, oriName, newName, wsPersonnel
                MsgBox "Swap completed.", vbInformation
            End If
        End If
    Next dateCell
    
    wsRoster.Activate
    
    ' Reprotect worksheets with password
    On Error Resume Next
    wsRoster.Protect "rostering2025", DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                     AllowFiltering:=True, AllowSorting:=True, AllowUsingPivotTables:=True
    wsPersonnel.Protect "rostering2025", DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                        AllowFiltering:=True, AllowSorting:=True, AllowUsingPivotTables:=True
    On Error GoTo 0
End Sub

' Initialize worksheet references
Private Sub InitializeWorksheets(wsPersonnel As Worksheet, wsSwap As Worksheet)
    Set wsRoster = Sheets("Roster")
    Set wsPersonnel = Sheets("PersonnelList (AOH & Desk)")
    Set wsSwap = Sheets("Swap")
End Sub

' Get original and new staff names from Swap sheet
Private Sub GetSwapNames(wsSwap As Worksheet, oriName As String, newName As String)
    oriName = UCase(Trim(wsSwap.Range("C4").Value))
    newName = UCase(Trim(wsSwap.Range("C5").Value))
End Sub

' Validate that names are not empty
Private Sub ValidateNames(oriName As String, newName As String)
    If Len(oriName) = 0 Then
        MsgBox "Error: Original staff name is empty. Please enter a valid personnel.", vbCritical
        Exit Sub
    End If
    If Len(newName) = 0 Then
        MsgBox "Error: New staff name is empty. Please enter a valid personnel.", vbCritical
        Exit Sub
    End If
End Sub

' Prompt user to select date range and return it
Private Function GetDateRange() As Range
    On Error Resume Next
    Set GetDateRange = Application.InputBox("Select date cells (Column B)", Type:=8)
    On Error GoTo 0
End Function

' Validate that the selected range is from column A (column 1)
Private Function IsValidDateColumn(dateRange As Range) As Boolean
    If Not dateRange.Columns.count = 1 Or dateRange.Column <> 2 Then
        MsgBox "Please only select dates from Date column.", vbExclamation
        IsValidDateColumn = False
    Else
        IsValidDateColumn = True
    End If
End Function

' Check if the original name exists in the row
Private Sub CheckOriginalNameExists(wsRoster As Worksheet, r As Long, slotCols As Variant, oriName As String, ByRef oriNameFound As Boolean)
    Dim col As Variant
    Dim cellValue As String
    Dim lines() As String
    Dim currStaff As String
    oriNameFound = False
    For Each col In slotCols
        cellValue = wsRoster.Cells(r, col).Value
        If InStr(cellValue, vbNewLine) > 0 Then
            currStaff = UCase(Trim(Replace(Split(cellValue, vbNewLine)(0), Chr(160), " ")))
        Else
            currStaff = UCase(Trim(cellValue))
        End If
        If currStaff = oriName Then
            oriNameFound = True
            Exit For
        End If
    Next col
End Sub

' Check if the new name exists in the same row
Private Sub CheckNewNameExists(wsRoster As Worksheet, r As Long, slotCols As Variant, newName As String, ByRef nameExists As Boolean)
    Dim col As Variant
    Dim cellValue As String
    Dim lines() As String
    nameExists = False
    For Each col In slotCols
        cellValue = wsRoster.Cells(r, col).Value
        If InStr(cellValue, vbNewLine) > 0 Then
            If UCase(Trim(Split(cellValue, vbNewLine)(0))) = newName Then
                nameExists = True
            End If
        Else
            If UCase(Trim(cellValue)) = newName Then
                nameExists = True
            End If
        End If
        If nameExists Then Exit For
    Next col
End Sub

' Display an error message
Private Sub DisplayError(message As String, messageType As VbMsgBoxStyle)
    MsgBox message, messageType
End Sub

' Perform the swap operation for a given row
Private Sub PerformSwap(wsRoster As Worksheet, r As Long, slotCols As Variant, oriName As String, newName As String, wsPersonnel As Worksheet)
    Dim slotCol As Variant
    Dim currentName As String
    Dim lines() As String
    Dim i As Long
    Dim lastRow As Long
    Dim cumulativeLength As Long
    Dim startPos As Integer
    
    For Each slotCol In slotCols
        With wsRoster.Cells(r, slotCol)
            ' Determine the current name based on whether there is a line break
            If InStr(.Value, vbNewLine) > 0 Then
                currentName = Trim(Split(.Value, vbNewLine)(0)) ' First unstriked line for subsequent swaps
            Else
                currentName = Trim(.Value) ' Entire value for initial swap
            End If
            
            If UCase(currentName) = oriName Then ' Check the current name
                ' Add new name first (unstriked) and preserve existing content
                .Value = newName & vbNewLine & .Value
                .VerticalAlignment = xlTop ' Align text to the top
                .WrapText = True
                
                ' Split into lines to apply strikethrough to all previous names
                lines = Split(.Value, vbNewLine)
                cumulativeLength = Len(newName) + 2 ' Start with newName and its vbNewLine
                
                ' Apply strikethrough to all lines except the first one
                For i = 1 To UBound(lines)
                    startPos = cumulativeLength
                    .Characters(startPos, Len(lines(i)) + 1).Font.Strikethrough = True
                    cumulativeLength = cumulativeLength + Len(lines(i)) + 2 ' Update for next line
                Next i
                
                ' Explicitly increase row height by 15 points per swap
                .RowHeight = .RowHeight + 15
                
                ' Update personnel counter for the new staff
                lastRow = wsPersonnel.Cells(wsPersonnel.Rows.count, "B").End(xlUp).row
                ' Deduct duties from the old staff
                For i = 12 To lastRow
                    If UCase(Trim(wsPersonnel.Cells(i, 2).Value)) = oriName Then
                        wsPersonnel.Cells(i, 5).Value = wsPersonnel.Cells(i, 5).Value - 1 ' Decrement Weekly Duties Counter
                        If slotCol = 10 Or slotCol = 12 Or slotCol = 14 Then ' AOH slots
                            wsPersonnel.Cells(i, 6).Value = wsPersonnel.Cells(i, 6).Value - 1 ' Decrement AOH Counter
                        End If
                        Exit For
                    End If
                Next i
                ' Update duties for the new staff
                For i = 12 To lastRow
                    If UCase(Trim(wsPersonnel.Cells(i, 2).Value)) = newName Then
                        wsPersonnel.Cells(i, 5).Value = wsPersonnel.Cells(i, 5).Value + 1 ' Increment Weekly Duties Counter
                        If slotCol = 10 Or slotCol = 12 Or slotCol = 14 Then ' AOH slots
                            wsPersonnel.Cells(i, 6).Value = wsPersonnel.Cells(i, 6).Value + 1 ' Increment AOH Counter
                        End If
                        Exit For
                    End If
                Next i
            End If
        End With
    Next slotCol
End Sub

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

Sub DeleteStaff()
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim selectedCell As Range
    Dim rowToDelete As ListRow
    Dim dutyType As String
    Dim specificDaysTbl As ListObject
    Dim availIndex As Long
    Const password As String = "rostering2025" ' Password just for deletion authorization
    Dim enteredPassword As String
    Dim sdRow As ListRow
    Dim nameIndex As Long, rowIndex As Long, sdNameIndex As Long
    
    ' Set the active worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then
        MsgBox "No active worksheet found.", vbExclamation
        Exit Sub
    End If
    
    On Error Resume Next
    ws.Unprotect
    On Error GoTo ErrHandler
    
    ' Determine the duty type based on the active sheet with trimming
    Select Case Trim(UCase(ws.Name))
        Case "LOAN MAIL BOX PERSONNELLIST"
            dutyType = "LOANMAILBOX"
            Set tbl = ws.ListObjects("LoanMailBoxMainList")
            Set specificDaysTbl = ws.ListObjects("LoanMailBoxSpecificDaysWorkingStaff")
        Case "MORNING PERSONNELLIST"
            dutyType = "MORNING"
            Set tbl = ws.ListObjects("MorningMainList")
            Set specificDaysTbl = ws.ListObjects("MorningSpecificDaysWorkingStaff")
        Case "AFTERNOON PERSONNELLIST"
            dutyType = "AFTERNOON"
            Set tbl = ws.ListObjects("AfternoonMainList")
            Set specificDaysTbl = ws.ListObjects("AfternoonSpecificDaysWorkingStaff")
        Case "AOH PERSONNELLIST"
            dutyType = "AOH"
            Set tbl = ws.ListObjects("AOHMainList")
            Set specificDaysTbl = ws.ListObjects("AOHSpecificDaysWorkingStaff")
        Case "SAT AOH PERSONNELLIST"
            dutyType = "SAT_AOH"
            Set tbl = ws.ListObjects("SatAOHMainList")
            Set specificDaysTbl = Nothing
        Case Else
            MsgBox "This sheet is not a recognized personnel list: '" & ws.Name & "'. Please select a valid personnel list sheet.", vbExclamation
            GoTo ReprotectAndExit
    End Select
    
    ' Validate main table exists
    If tbl Is Nothing Then
        MsgBox "Main table not found on '" & ws.Name & "'.", vbExclamation
        GoTo ReprotectAndExit
    End If
    
    ' Get the selected cell
    Set selectedCell = ActiveCell
    If Not Intersect(selectedCell, tbl.Range) Is Nothing Then
        ' Check if the selected cell is in the "Name" column
        nameIndex = GetColumnIndex(tbl, "Name")
        If nameIndex = -1 Then
            MsgBox "Column 'Name' not found in main table.", vbExclamation
            GoTo ReprotectAndExit
        End If
        If selectedCell.Column <> tbl.Range.Cells(1, nameIndex).Column Then
            MsgBox "Please select a cell in the 'Name' column to delete the staff.", vbExclamation
            GoTo ReprotectAndExit
        End If
        
        ' Find the row in the table
        rowIndex = selectedCell.row - tbl.Range.row
        If rowIndex <= 0 Or rowIndex > tbl.ListRows.count Then
            MsgBox "Selected cell is not within a valid table row.", vbExclamation
            GoTo ReprotectAndExit
        End If
        Set rowToDelete = tbl.ListRows(rowIndex)
        
        ' Only check Availability Type if this isn't Sat AOH sheet
        If dutyType <> "SAT_AOH" Then
            availIndex = GetColumnIndex(tbl, "Availability Type")
            If availIndex = -1 Then
                MsgBox "Column 'Availability Type' not found in main table.", vbExclamation
                GoTo ReprotectAndExit
            End If
            
            ' Clear filters if any (for validation purposes)
            If tbl.ShowAutoFilter Then
                tbl.AutoFilter.ShowAllData
            End If
            If Not specificDaysTbl Is Nothing And specificDaysTbl.ShowAutoFilter Then
                specificDaysTbl.AutoFilter.ShowAllData
            End If
        End If
    Else
        MsgBox "Please select a cell within the main table to delete a staff.", vbExclamation
        GoTo ReprotectAndExit
    End If
    
    ' Only ask for password AFTER all validations pass
    enteredPassword = InputBox("Please enter the password for deletion:", "Password Authentication")
    If enteredPassword <> password Then
        MsgBox "Incorrect password. Delete operation declined.", vbCritical
        GoTo ReprotectAndExit
    End If
    
    ' Perform deletion within the current sheets contextt
    If dutyType = "SAT_AOH" Then
        rowToDelete.Delete
    Else
        ' For other sheets, handle specific days workers if needed
        If UCase(Trim(rowToDelete.Range.Cells(1, availIndex).Value)) = "SPECIFIC DAYS" And Not specificDaysTbl Is Nothing Then
            sdNameIndex = GetColumnIndex(specificDaysTbl, "Name")
            If sdNameIndex <> -1 Then
                ' Clear filters on specific days table
                If specificDaysTbl.ShowAutoFilter Then
                    specificDaysTbl.AutoFilter.ShowAllData
                End If
                ' Find and delete the corresponding row in the same sheets specific days tablee
                For Each sdRow In specificDaysTbl.ListRows
                    If UCase(Trim(sdRow.Range.Cells(1, sdNameIndex).Value)) = UCase(Trim(selectedCell.Value)) Then
                        sdRow.Delete
                        Exit For
                    End If
                Next sdRow
            End If
        End If
        rowToDelete.Delete
    End If
    
    CalculateMaxDuties.CalculateMaxDuties dutyType
    MsgBox "Staff deleted and Max Duties recalculated successfully for " & dutyType & " on " & ws.Name & ".", vbInformation

ReprotectAndExit:
    ' Reprotect the worksheet WITHOUT password but keep D5:D9 unlocked
    If Not ws Is Nothing Then
        On Error Resume Next ' In case sheet is already protected
        With ws
            .Unprotect ' Ensure sheet is unprotected before setting locked ranges
            If Not tbl Is Nothing Then
                .ListObjects(tbl.Name).Range.Locked = True
            End If
            If Not specificDaysTbl Is Nothing Then
                .ListObjects(specificDaysTbl.Name).Range.Locked = True
            End If
            .Range("D5:D9").Locked = False ' Keep data entry unlocked
            .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                     AllowFiltering:=True, AllowSorting:=True, AllowUsingPivotTables:=True
        End With
        On Error GoTo 0
    End If
    ws.Activate
    Exit Sub

ErrHandler:
    MsgBox "An error occurred: " & Err.Description & vbCrLf & _
           "Line: " & Erl & vbCrLf & _
           "Duty Type: " & dutyType & vbCrLf & _
           "Worksheet: " & IIf(ws Is Nothing, "Not Set", ws.Name), vbCritical
    GoTo ReprotectAndExit
End Sub

' Helper function to get column index
Private Function GetColumnIndex(tbl As ListObject, columnName As String) As Long
    On Error Resume Next
    GetColumnIndex = tbl.ListColumns(columnName).Index
    If Err.Number <> 0 Then GetColumnIndex = -1
    On Error GoTo 0
End Function

Sub GenerateMorningShiftAnalysis()
    Dim wsPersonnel As Worksheet
    Dim wsAnalysis As Worksheet
    Dim wsRosterCopy As Worksheet
    Dim tbl As ListObject
    Dim nameList As Range
    Dim dutyCounterList As Range
    Dim lastRow As Long, i As Long
    Dim dict As Object
    Dim empName As String
    Dim latestRosterName As String
    Dim sht As Worksheet
    Dim newestDate As Date
    Dim START_ROW As Long: START_ROW = 6
    Dim MOR_COL As Long: MOR_COL = 6
    Dim cell As Range, cellValue As String, currStaff As String
    Dim NextRow As Long, tableStartRow As Long

    ' Find latest ActualRoster_* sheet
    newestDate = 0
    For Each sht In ThisWorkbook.Sheets
        If sht.Name Like "ActualRoster_*" Then
            Dim dtPart As String
            dtPart = Replace(Mid(sht.Name, 14), "_", " ")
            On Error Resume Next
            Dim parsedDate As Date
            parsedDate = CDate(Left(dtPart, 4) & "/" & Mid(dtPart, 5, 2) & "/" & Mid(dtPart, 7, 2) & " " & Mid(dtPart, 10, 2) & ":" & Mid(dtPart, 12, 2))
            If Err.Number = 0 Then
                If parsedDate > newestDate Then
                    newestDate = parsedDate
                    latestRosterName = sht.Name
                End If
            End If
            On Error GoTo 0
        End If
    Next sht
    
    If latestRosterName = "" Then
        MsgBox "No ActualRoster_* sheet found.", vbExclamation
        Exit Sub
    End If

    ' Add big title at row 1
    With wsAnalysis.Range("A1:E1")
        .Merge
        .Value = "Analysis Report"
        .Interior.Color = RGB(255, 199, 206) ' Light red
        .Font.Size = 16
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Set wsRosterCopy = Sheets(latestRosterName)
    Set wsPersonnel = Sheets("Morning PersonnelList")

    ' Create or clear analysis sheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Sheets("MorningAnalysis").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsAnalysis = Sheets.Add(After:=Sheets(Sheets.count))
    wsAnalysis.Name = "MorningAnalysis"

    ' Write title
    wsAnalysis.Range("A1").Value = "Morning Slot Analysis"
    wsAnalysis.Range("A1").Font.Bold = True
    wsAnalysis.Range("A1").Font.Size = 14
    tableStartRow = 4 ' header starts here

    ' Header row
    With wsAnalysis
        .Range("A" & tableStartRow).Value = "Name"
        .Range("B" & tableStartRow).Value = "System Counter"
        .Range("C" & tableStartRow).Value = "Actual Counter"
        .Range("D" & tableStartRow).Value = "Difference"
        .Range("E" & tableStartRow).Value = "% Difference"
    End With

    ' Get personnel table
    Set tbl = wsPersonnel.ListObjects("MorningMainList")
    Set nameList = tbl.ListColumns("Name").DataBodyRange
    Set dutyCounterList = tbl.ListColumns("Duties Counter").DataBodyRange

    ' Create dictionary and fill initial system counter
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To nameList.Rows.count
        empName = UCase(Trim(nameList.Cells(i, 1).Value))
        wsAnalysis.Cells(i + tableStartRow, 1).Value = empName
        wsAnalysis.Cells(i + tableStartRow, 2).Value = dutyCounterList.Cells(i, 1).Value
        dict(empName) = 0
    Next i

    ' Count actual appearances from roster
    For i = START_ROW To 186
        Set cell = wsRosterCopy.Cells(i, MOR_COL)
        cellValue = cell.Value

        If InStr(cellValue, vbNewLine) > 0 Then
            currStaff = UCase(Trim(Replace(Split(cellValue, vbNewLine)(0), Chr(160), " ")))
        Else
            currStaff = UCase(Trim(cellValue))
        End If

        If Len(currStaff) > 0 And currStaff <> "CLOSED" Then
            If dict.Exists(currStaff) Then
                dict(currStaff) = dict(currStaff) + 1
            Else
                ' New staff found
                NextRow = wsAnalysis.Cells(wsAnalysis.Rows.count, 1).End(xlUp).row + 1
                wsAnalysis.Cells(NextRow, 1).Value = currStaff
                wsAnalysis.Cells(NextRow, 2).Value = 0
                wsAnalysis.Cells(NextRow, 3).Value = 1
                wsAnalysis.Cells(NextRow, 4).FormulaR1C1 = "=RC[-1]-RC[-2]"
                wsAnalysis.Cells(NextRow, 5).FormulaR1C1 = "=IF(RC[-3]=0,"""",RC[-1]/RC[-3]*100)"
                dict(currStaff) = 1

                ' Highlight new row
                wsAnalysis.Range(wsAnalysis.Cells(NextRow, 1), wsAnalysis.Cells(NextRow, 5)).Interior.Color = RGB(255, 255, 153)
            End If
        End If
    Next i

    ' Fill actual count and compute difference + percentage
    For i = tableStartRow + 1 To wsAnalysis.Cells(wsAnalysis.Rows.count, 1).End(xlUp).row
        empName = UCase(Trim(wsAnalysis.Cells(i, 1).Value))
        If dict.Exists(empName) Then
            wsAnalysis.Cells(i, 3).Value = dict(empName)
            wsAnalysis.Cells(i, 4).FormulaR1C1 = "=RC[-1]-RC[-2]"
            wsAnalysis.Cells(i, 5).FormulaR1C1 = "=IF(RC[-3]=0,0,RC[-1]/RC[-3]*100)"
        End If
    Next i

    ' Format as Table
    Dim tableRange As Range
    Dim analysisTable As ListObject

    lastRow = wsAnalysis.Cells(wsAnalysis.Rows.count, 1).End(xlUp).row
    Set tableRange = wsAnalysis.Range("A" & tableStartRow & ":E" & lastRow)
    Set analysisTable = wsAnalysis.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    analysisTable.Name = "MorningShiftTable"

    ' Format % column to 2 decimal places
    analysisTable.ListColumns("% Difference").DataBodyRange.NumberFormat = "0.00"


    ' Add small section title at row 3
    With wsAnalysis.Range("A3:C3")
        .Merge
        .Value = "Morning Slot Analysis"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = analysisTable.HeaderRowRange.Interior.Color
    End With

    MsgBox "Morning shift analysis generated using '" & latestRosterName & "'.", vbInformation
End Sub

Private last_row_roster As Long
' Reusable function to generate side-by-side shift analysis blocks
Sub GenerateShiftAnalysisBlock(wsAnalysis As Worksheet, rosterSheet As Worksheet, _
                                personnelSheetName As String, tableName As String, _
                                slotTitle As String, rosterCol1 As Long, startCol As Long, _
                                Optional rosterCol2 As Variant)

    Dim wsPersonnel As Worksheet
    Dim tbl As ListObject
    Dim nameList As Range, dutyCounterList As Range
    Dim dict As Object
    Dim empName As String
    Dim cell As Range, cellValue As String, currStaff As String
    Dim rowOffset As Long: rowOffset = 4
    Dim lastRow As Long, NextRow As Long
    Dim tableRange As Range, analysisTable As ListObject
    Dim i As Long, tableWidth As Long: tableWidth = 5
    Dim enteredPassword As String

    Set wsPersonnel = Sheets(personnelSheetName)
    Set tbl = wsPersonnel.ListObjects(tableName)
    Set nameList = tbl.ListColumns("Name").DataBodyRange
    Set dutyCounterList = tbl.ListColumns("Duties Counter").DataBodyRange

    ' Add small section title
    With wsAnalysis.Cells(3, startCol).Resize(1, 3)
        .Merge
        .Value = slotTitle
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(184, 204, 228)
    End With

    ' Add header
    With wsAnalysis
        .Cells(rowOffset, startCol).Value = "Name"
        .Cells(rowOffset, startCol + 1).Value = "System Counter"
        .Cells(rowOffset, startCol + 2).Value = "Actual Counter"
        .Cells(rowOffset, startCol + 3).Value = "Difference"
        .Cells(rowOffset, startCol + 4).Value = "% Difference"
    End With

    ' Check if table exists and if it is empty
    On Error Resume Next
    Set tbl = wsPersonnel.ListObjects(tableName)
    On Error GoTo 0
    
    If tbl Is Nothing Or tbl.ListRows.count = 0 Then
        ' Create an empty analysis table with only headers
        Set tableRange = wsAnalysis.Range(wsAnalysis.Cells(rowOffset, startCol), _
                                          wsAnalysis.Cells(rowOffset, startCol + tableWidth - 1))
        
        Set analysisTable = wsAnalysis.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
        analysisTable.Name = Replace(slotTitle, " ", "") & "Table"
        
        ' No data rows, so skip formatting % Difference
        Exit Sub
    End If


    Set dict = CreateObject("Scripting.Dictionary")
    ' Load initial staff counters
    For i = 1 To nameList.Rows.count
        empName = UCase(Trim(nameList.Cells(i, 1).Value))
        wsAnalysis.Cells(rowOffset + i, startCol).Value = empName
        wsAnalysis.Cells(rowOffset + i, startCol + 1).Value = dutyCounterList.Cells(i, 1).Value
        dict(empName) = 0
    Next i

    Dim colIndex As Variant
    Dim colList As Variant
    
    If IsMissing(rosterCol2) Then
        colList = Array(rosterCol1)
    Else
        colList = Array(rosterCol1, rosterCol2)
    End If
    
    For Each colIndex In colList
        For i = 6 To last_row_roster
            Set cell = rosterSheet.Cells(i, colIndex)
            cellValue = cell.Value
    
            If InStr(cellValue, vbNewLine) > 0 Then
                currStaff = UCase(Trim(Replace(Split(cellValue, vbNewLine)(0), Chr(160), " ")))
            Else
                currStaff = UCase(Trim(cellValue))
            End If
    
            If Len(currStaff) > 0 And currStaff <> "CLOSED" Then
                If dict.Exists(currStaff) Then
                    dict(currStaff) = dict(currStaff) + 1
                Else
                    NextRow = wsAnalysis.Cells(wsAnalysis.Rows.count, startCol).End(xlUp).row + 1
                    wsAnalysis.Cells(NextRow, startCol).Value = currStaff
                    wsAnalysis.Cells(NextRow, startCol + 1).Value = 0
                    wsAnalysis.Cells(NextRow, startCol + 2).Value = 1
                    wsAnalysis.Cells(NextRow, startCol + 3).FormulaR1C1 = "=RC[-1]-RC[-2]"
                    wsAnalysis.Cells(NextRow, startCol + 4).FormulaR1C1 = "=IF(RC[-3]=0,"""",RC[-1]/RC[-3]*100)"
                    dict(currStaff) = 1
                    wsAnalysis.Range(wsAnalysis.Cells(NextRow, startCol), wsAnalysis.Cells(NextRow, startCol + 4)).Interior.Color = RGB(255, 255, 153)
                End If
            End If
        Next i
    Next colIndex

    For i = rowOffset + 1 To wsAnalysis.Cells(wsAnalysis.Rows.count, startCol).End(xlUp).row
        empName = UCase(Trim(wsAnalysis.Cells(i, startCol).Value))
        If dict.Exists(empName) Then
            wsAnalysis.Cells(i, startCol + 2).Value = dict(empName)
            wsAnalysis.Cells(i, startCol + 3).FormulaR1C1 = "=RC[-1]-RC[-2]"
            wsAnalysis.Cells(i, startCol + 4).FormulaR1C1 = "=IF(RC[-3]=0,0,RC[-1]/RC[-3]*100)"
        End If
    Next i

    ' Format as Table
    lastRow = wsAnalysis.Cells(wsAnalysis.Rows.count, startCol).End(xlUp).row
    Set tableRange = wsAnalysis.Range(wsAnalysis.Cells(rowOffset, startCol), wsAnalysis.Cells(lastRow, startCol + tableWidth - 1))
    Set analysisTable = wsAnalysis.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    analysisTable.Name = Replace(slotTitle, " ", "") & "Table"
    analysisTable.ListColumns("% Difference").DataBodyRange.NumberFormat = "0.00"
End Sub
Sub GenerateTotalSummaryTable(wsAnalysis As Worksheet)
    Dim summaryDict As Object
    Set summaryDict = CreateObject("Scripting.Dictionary")
    
    Dim tableNames As Variant
    tableNames = Array("LoanMailBoxSlotAnalysisTable", "MorningSlotAnalysisTable", _
                       "AfternoonSlotAnalysisTable", "AOHSlotAnalysisTable", "SatAOHSlotAnalysisTable")

    Dim tbl As ListObject, row As ListRow
    Dim empName As String
    Dim sysCount As Long, actCount As Long
    Dim tblName As Variant

    ' Loop through each analysis table
    For Each tblName In tableNames
        On Error Resume Next
        Set tbl = wsAnalysis.ListObjects(tblName)
        On Error GoTo 0

        ' Skip if table doesn't exist or is empty
        If Not tbl Is Nothing Then
            If tbl.ListRows.count > 0 Then
                ' Loop through each row in the table
                For Each row In tbl.ListRows
                    empName = UCase(Trim(row.Range.Cells(1, 1).Value))
                    sysCount = Val(row.Range.Cells(1, 2).Value)
                    actCount = Val(row.Range.Cells(1, 3).Value)
                    
                    If summaryDict.Exists(empName) Then
                        Dim counts As Variant
                        counts = summaryDict(empName)
                        counts(0) = counts(0) + sysCount
                        counts(1) = counts(1) + actCount
                        summaryDict(empName) = counts
                    Else
                        summaryDict(empName) = Array(sysCount, actCount)
                    End If
                Next row
            End If
        End If
        Set tbl = Nothing
    Next tblName

    ' Determine where to place the summary table
    Dim startCol As Long
    startCol = wsAnalysis.Cells(4, wsAnalysis.Columns.count).End(xlToLeft).Column + 2
    Dim rowOffset As Long: rowOffset = 4

    ' Add table header
    With wsAnalysis
        .Cells(3, startCol).Resize(1, 5).Merge
        .Cells(3, startCol).Value = "Total Summary"
        .Cells(3, startCol).Font.Bold = True
        .Cells(3, startCol).Interior.Color = RGB(184, 204, 228)
        .Cells(3, startCol).HorizontalAlignment = xlCenter

        .Cells(rowOffset, startCol).Value = "Name"
        .Cells(rowOffset, startCol + 1).Value = "System Counter"
        .Cells(rowOffset, startCol + 2).Value = "Actual Counter"
        .Cells(rowOffset, startCol + 3).Value = "Difference"
        .Cells(rowOffset, startCol + 4).Value = "% Difference"
        .Range(.Cells(rowOffset, startCol), .Cells(rowOffset, startCol + 4)).Font.Bold = True
    End With

    ' Write data into sheet
    Dim i As Long: i = rowOffset + 1
    Dim diff As Long, pctDiff As Double
    Dim sysVal As Long, actVal As Long
    Dim staff As Variant

    For Each staff In summaryDict.Keys
        sysVal = summaryDict(staff)(0)
        actVal = summaryDict(staff)(1)
        diff = actVal - sysVal
        
        If sysVal <> 0 Then
            pctDiff = diff / sysVal * 100
        Else
            pctDiff = 0
        End If

        wsAnalysis.Cells(i, startCol).Value = staff
        wsAnalysis.Cells(i, startCol + 1).Value = sysVal
        wsAnalysis.Cells(i, startCol + 2).Value = actVal
        wsAnalysis.Cells(i, startCol + 3).Value = diff
        wsAnalysis.Cells(i, startCol + 4).Value = pctDiff
        i = i + 1
    Next staff

    ' Create an empty table if no data exists
    Dim lastRow As Long
    If summaryDict.count = 0 Then
        lastRow = rowOffset
    Else
        lastRow = wsAnalysis.Cells(wsAnalysis.Rows.count, startCol).End(xlUp).row
    End If

    ' Format as table
    Dim tableRange As Range
    Set tableRange = wsAnalysis.Range(wsAnalysis.Cells(rowOffset, startCol), wsAnalysis.Cells(lastRow, startCol + 4))

    Dim summaryTable As ListObject
    Set summaryTable = wsAnalysis.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    summaryTable.Name = "TotalSummaryTable"
    If summaryTable.ListRows.count > 0 Then
        summaryTable.ListColumns("% Difference").DataBodyRange.NumberFormat = "0.00"
    End If
End Sub



Sub MasterGenerateAllAnalyses()
    Dim wsAnalysis As Worksheet
    Dim rosterSheet As Worksheet
    Dim userRange As Range
    Dim selectedSheet As Worksheet
    Const password As String = "rostering2025"
    
    enteredPassword = InputBox("Please enter the password to generate analysis report:", "Password Authentication")
    If enteredPassword <> password Then
        MsgBox "Incorrect password. Unable to generate report.", vbCritical
        Exit Sub
    End If

    ' Prompt user to click on any cell in the target ActualRoster_* sheet
    On Error Resume Next
    Set userRange = Application.InputBox( _
        Prompt:="Please choose one 'ActualRoster' sheet to analyse." & vbCrLf & _
                "After that, click on any cell on the selected 'ActualRoster' sheet." & vbCrLf & _
                "The sheet name must start with 'ActualRoster_'", _
        title:="Select Actual Roster Sheet", _
        Type:=8)
    On Error GoTo 0

    If userRange Is Nothing Then Exit Sub ' User cancelled

    Set selectedSheet = userRange.Worksheet
    If selectedSheet.Name Like "ActualRoster_*" = False Then
        MsgBox "Invalid selection. Please choose a sheet that starts with 'ActualRoster_'.", vbExclamation
        Exit Sub
    End If
    Set rosterSheet = selectedSheet

    ' Create ReportAnalysis sheet
    Dim dtNamePart As String, formattedName As String
    dtNamePart = Mid(selectedSheet.Name, 14)
    formattedName = "AnalysisReport_" & dtNamePart
    
    Set wsAnalysis = Sheets.Add(After:=Sheets(Sheets.count))
    On Error Resume Next
    wsAnalysis.Name = formattedName
    If Err.Number <> 0 Then
        MsgBox "Could not name sheet as " & formattedName & ". It may already exist.", vbExclamation
        wsAnalysis.Name = "AnalysisReport_" & Format(Now, "yyyymmdd_hhnnss")
    End If
    On Error GoTo 0

    'Find last row of roster
    If selectedSheet.Cells(2, 10).Value = "Jan-Jun" And selectedSheet.Cells(2, 13).Value Mod 4 = 0 Then
        last_row_roster = 187
    ElseIf selectedSheet.Cells(2, 10).Value = "Jan-Jun" Then
        last_row_roster = 186
    Else
        last_row_roster = 189
    End If
    
    ' Big title
    With wsAnalysis.Range("A1:Z1")
        .Merge
        .Value = "Analysis Report"
        .Font.Size = 16
        .Font.Bold = True
        .Interior.Color = RGB(255, 199, 206)
        .HorizontalAlignment = xlCenter
    End With

    ' Show selected ActualRoster sheet name in row 2
    With wsAnalysis.Range("A2:Z2")
        .Merge
        .Value = "Based on: " & rosterSheet.Name
        .Font.Italic = True
        .HorizontalAlignment = xlCenter
    End With

    ' Generate all 5 analyses side by side
    GenerateShiftAnalysisBlock wsAnalysis, rosterSheet, "Loan Mail Box PersonnelList", "LoanMailBoxMainList", "Loan Mail Box Slot Analysis", LMB_COL, 1
    GenerateShiftAnalysisBlock wsAnalysis, rosterSheet, "Morning PersonnelList", "MorningMainList", "Morning Slot Analysis", MOR_COL, 7
    GenerateShiftAnalysisBlock wsAnalysis, rosterSheet, "Afternoon PersonnelList", "AfternoonMainList", "Afternoon Slot Analysis", AFT_COL, 13
    GenerateShiftAnalysisBlock wsAnalysis, rosterSheet, "AOH PersonnelList", "AOHMainList", "AOH Slot Analysis", AOH_COL, 19
    GenerateShiftAnalysisBlock wsAnalysis, rosterSheet, "Sat AOH PersonnelList", "SatAOHMainList", _
    "Sat AOH Slot Analysis", SAT_AOH_COL1, 25, SAT_AOH_COL2
    GenerateTotalSummaryTable wsAnalysis
    
    With wsAnalysis.Cells
        .Locked = True
    End With
    
    wsAnalysis.Protect password:="nuslib2017@52", _
                        AllowSorting:=True, _
                        AllowFiltering:=True, _
                        AllowFormattingCells:=True
                        
    

    MsgBox "All shift analyses completed for '" & rosterSheet.Name & "'!", vbInformation
End Sub


Sub TestGenerateSummary()
    Call GenerateTotalSummaryTable(Sheets("AnalysisReport_20250724_0306"))
End Sub


Sub HidePersonnelSheetsWithPassword()
    Dim ws As Worksheet
    Dim enteredPassword As String
    Const password As String = "rostering2025"
    
    enteredPassword = InputBox("Please enter the password hide all Personnel List sheets:", "Password Authentication")
    If enteredPassword <> password Then
        MsgBox "Incorrect password. Hiding operation declined.", vbCritical
        Exit Sub
    End If
    
    ' Loop through all worksheets and hide personnel lists
    For Each ws In ThisWorkbook.Sheets
        Select Case UCase(ws.Name)
            Case UCase("AOH PersonnelList"), UCase("Sat AOH PersonnelList"), _
                 UCase("Loan Mail Box PersonnelList"), UCase("Morning PersonnelList"), _
                 UCase("Afternoon PersonnelList")
                ' Protect the entire sheet
                ws.Protect , DrawingObjects:=True, Contents:=True, Scenarios:=True
                ' Set to "Very Hidden" (not visible in UI, only via VBA)
                ws.Visible = xlSheetVeryHidden
        End Select
    Next ws
    
    MsgBox "Personnel list sheets have been hidden.", vbInformation
End Sub
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

Sub ReprotectPersonnelLists()
    Dim ws As Worksheet
    Dim sheetNames As Variant
    Dim i As Long
    Dim tbl As ListObject
    Dim specificDaysTbl As ListObject

    ' Array of sheet names to reprotect
    sheetNames = Array("Loan Mail Box PersonnelList", "Morning PersonnelList", "Afternoon PersonnelList", _
                       "AOH PersonnelList", "Sat AOH PersonnelList")

    ' Loop through each sheet
    For i = LBound(sheetNames) To UBound(sheetNames)
        On Error Resume Next ' Handle case where sheet or table might not exist
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        If Not ws Is Nothing Then
            ' Unprotect the worksheet to allow changes to Locked property
            On Error Resume Next
            ws.Unprotect ' Remove protection (add password if required, e.g., ws.Unprotect "password")
            On Error GoTo 0
            ' Set tables based on sheet name
            Select Case sheetNames(i)
                Case "Loan Mail Box PersonnelList"
                    Set tbl = ws.ListObjects("LoanMailBoxMainList")
                    Set specificDaysTbl = ws.ListObjects("LoanMailBoxSpecificDaysWorkingStaff")
                Case "Morning PersonnelList"
                    Set tbl = ws.ListObjects("MorningMainList")
                    Set specificDaysTbl = ws.ListObjects("MorningSpecificDaysWorkingStaff")
                Case "Afternoon PersonnelList"
                    Set tbl = ws.ListObjects("AfternoonMainList")
                    Set specificDaysTbl = ws.ListObjects("AfternoonSpecificDaysWorkingStaff")
                Case "AOH PersonnelList"
                    Set tbl = ws.ListObjects("AOHMainList")
                    Set specificDaysTbl = ws.ListObjects("AOHSpecificDaysWorkingStaff")
                Case "Sat AOH PersonnelList"
                    Set tbl = ws.ListObjects("SatAOHMainList")
                    Set specificDaysTbl = Nothing
                    ' No specificDaysTbl for Sat AOH, so it remains Nothing
            End Select
            On Error GoTo 0

            With ws
                ' Lock table ranges if they exist
                If Not tbl Is Nothing Then
                    .ListObjects(tbl.Name).Range.Locked = True
                End If
                If Not specificDaysTbl Is Nothing Then
                    Debug.Print specificDaysTbl
                    .ListObjects(specificDaysTbl.Name).Range.Locked = True
                End If
                ' Unlock D5:D9 for data entry
                .Range("D5:D9").Locked = False
                ' Apply protection with specified allowances
                .Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
                         AllowFiltering:=True, AllowSorting:=True, AllowUsingPivotTables:=True
            End With
        Else
            Debug.Print "Sheet not found: " & sheetNames(i)
        End If
        On Error GoTo 0
    Next i

    Debug.Print "Reprotection completed for all personnel lists."
End Sub



