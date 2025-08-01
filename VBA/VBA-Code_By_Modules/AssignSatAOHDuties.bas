Attribute VB_Name = "AssignSatAOHDuties"
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



