Attribute VB_Name = "ReprotectPersonnelLists"
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



