Attribute VB_Name = "CountSlotsSub"
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

