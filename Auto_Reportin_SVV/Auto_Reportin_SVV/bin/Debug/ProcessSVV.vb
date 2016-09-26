'This VBA is For the CycleTime(SVV) report

'Created By : Sambhav Patni for AutoReportSVV

'This VBA is For the CycleTime(SVV) report

'a. Run the CycleTime(SVV).SQL in NMAP  and ifnmap (merg)in the SQL file referenced above.
'b. Copy results to Excel.
'c. Make a backup copy of the "Table" tab,   & remove the blank spaces
'd. Run ProcessSVVTab, which deletes PL/SQL column if necessary, formats, adds headers, inserts
'     columns, and computes adjustments
'e. rename the "Table" tab to "SVV_all-data".

'Final Cleanup:
'
'A. Save spreadsheet under a 'for distro' name.
'B. Delete backup copy tabs (all tabs except "SVV_all-data") and this VBA before sending out.

Dim CurrSheet As String
Dim LastRowPivot As Integer
Dim LastSVVRowToProcess As Integer
Dim strReasons As String
Const FirstRowToProcess = 2

Const ArrivedInESRTCol = "F"
Const AssignedToResearcherCol = "G"
Const ElapsedTimeArrivedToAssignedCol = "H"
Const ComputedTimeArrivedToAssignedWithAdjCol = "I"
Const ReasonsForAdjCol1 = "J"
Const DQSubmittedCol = "K"
Const ElapsedTimeAssignedToDQCol = "L"
Const ComputedTimeAssignedToDQCol = "M"
Const ReasonsForAdjCol2 = "N"
Const TotalElapsedTimeCol = "O"
Const ComputedTotalTimeWithAdjCol = "P"
Const ReasonsForAdjCol3 = "Q"

Dim LastRow As Integer

Sub DetermineLastRowToValidate(sheetName As String)
Dim nRowCounter As Integer
nRowCounter = 1
Do Until Worksheets(sheetName).Cells(nRowCounter, "B").Value = ""
    nRowCounter = nRowCounter + 1
Loop
LastRow = nRowCounter - 1
End Sub

Sub CreateSVV()

	ClearBlanks
	ProcessSVVTab
	MidWork
	CreatePivot_alldata
	CreateSummary
	colors
End Sub
Sub CreateSummary()
    
    CurrSheet = "SVV_All-data_Summary-Pivot"
    Sheets(CurrSheet).Select
    DetermineLastRowPivot CurrSheet
    CreateSummary_Header
    AddSummaryData
	
	CurrSheet = "SVV_Classic_Summary-Pivot"
    Sheets(CurrSheet).Select
    DetermineLastRowPivot CurrSheet
    CreateSummary_Header
    AddSummaryData
	
	CurrSheet = "SVV_MOTORCYCLE_Summary-Pivot"
    Sheets(CurrSheet).Select
    DetermineLastRowPivot CurrSheet
    CreateSummary_Header
    AddSummaryData
	
	CurrSheet = "SVV_TRUCK_Summary-Pivot"
    Sheets(CurrSheet).Select
    DetermineLastRowPivot CurrSheet
    CreateSummary_Header
    AddSummaryData
	
	CurrSheet = "SVV_UTILITY_Summary-Pivot"
    Sheets(CurrSheet).Select
    DetermineLastRowPivot CurrSheet
    CreateSummary_Header
    AddSummaryData
	
	CurrSheet = "SVV_RECREATIONAL_Summary-Pivot"
    Sheets(CurrSheet).Select
    DetermineLastRowPivot CurrSheet
    CreateSummary_Header
    AddSummaryData
	
	CurrSheet = "SVV_OTHER_Summary-Pivot"
    Sheets(CurrSheet).Select
    DetermineLastRowPivot CurrSheet
    CreateSummary_Header
    AddSummaryData

End Sub

Sub ClearBlanks()
'
' ClearBlanks Macro
' Created By Sambhav

'
	Dim Range_Temp As String
	DetermineLastRowToValidate "Table"
    Selection.AutoFilter
	Range_Temp = "$A$1:$L$" & LastRow
    ActiveSheet.Range(Range_Temp).AutoFilter Field:=8, Criteria1:="="
	DetermineLastRowToValidate "Table"
	Range_Temp = "2:" & LastRow
    Rows(Range_Temp).Select
    Selection.Delete Shift:=xlUp
    Selection.AutoFilter
End Sub

Sub Clear103WeekEnd()
'
' Clear103WeekEnd Macro
' Created By Sambhav

'
	Dim Range_Temp As String
	DetermineLastRowToValidate "Table"
    Selection.AutoFilter
	Range_Temp = "$A$1:$L$" & LastRow
    ActiveSheet.Range(Range_Temp).AutoFilter Field:=18, Criteria1:="=*--Spanned 103 Weekend Day(s)*"
	DetermineLastRowToValidate "Table"
	Range_Temp = "2:" & LastRow
    Rows(Range_Temp).Select
    Selection.Delete Shift:=xlUp
    'Selection.AutoFilter
End Sub

Sub MidWork()
'
' MidWork Macro
'

'
    Sheets("Table").Select
    Sheets("Table").Name = "SVV_all-data"    
End Sub

Sub CreatePivot_alldata()
	DetermineLastRowToValidate "SVV_all-data"
	Range_Temp = "SVV_all-data!R1C1:R" & LastRow & "C18"
    Sheets("SVV_all-data").Select    
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Range_Temp, Version:=xlPivotTableVersion10). _
        CreatePivotTable TableDestination:="Sheet7!R3C1", TableName:="PivotTable1" _
        , DefaultVersion:=xlPivotTableVersion10
    Sheets("Sheet7").Select    
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Cycle Time"), "Count of Cycle Time", xlCount
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO CD")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Compliance T/F ")
        .Orientation = xlColumnField
        .Position = 1
    End With
    Sheets("Sheet7").Select
    Sheets("Sheet7").Name = "SVV_All-data_Summary-Pivot"    
	CreatePivot_CLASSICdata
End Sub

Sub CreatePivot_CLASSICdata()
	DetermineLastRowToValidate "SVV_CLASSIC-data"
	Range_Temp = "SVV_CLASSIC-data!R1C1:R" & LastRow & "C18"
    Sheets("SVV_CLASSIC-data").Select    
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Range_Temp, Version:=xlPivotTableVersion10). _
        CreatePivotTable TableDestination:="Sheet8!R3C1", TableName:="PivotTable1" _
        , DefaultVersion:=xlPivotTableVersion10
    Sheets("Sheet8").Select    
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Cycle Time"), "Count of Cycle Time", xlCount
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO CD")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Compliance T/F ")
        .Orientation = xlColumnField
        .Position = 1
    End With
    Sheets("Sheet8").Select
    Sheets("Sheet8").Name = "SVV_Classic_Summary-Pivot"    
	CreatePivot_MOTORCYCLEdata
End Sub

Sub CreatePivot_MOTORCYCLEdata()
	DetermineLastRowToValidate "SVV_MOTORCYCLE-data"
	Range_Temp = "SVV_MOTORCYCLE-data!R1C1:R" & LastRow & "C18"
    Sheets("SVV_MOTORCYCLE-data").Select    
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Range_Temp, Version:=xlPivotTableVersion10). _
        CreatePivotTable TableDestination:="Sheet9!R3C1", TableName:="PivotTable1" _
        , DefaultVersion:=xlPivotTableVersion10
    Sheets("Sheet9").Select    
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Cycle Time"), "Count of Cycle Time", xlCount
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO CD")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Compliance T/F ")
        .Orientation = xlColumnField
        .Position = 1
    End With
    Sheets("Sheet9").Select
    Sheets("Sheet9").Name = "SVV_MOTORCYCLE_Summary-Pivot"    
	CreatePivot_TRUCKdata
End Sub

Sub CreatePivot_TRUCKdata()
	DetermineLastRowToValidate "SVV_TRUCK-data"
	Range_Temp = "SVV_TRUCK-data!R1C1:R" & LastRow & "C18"
    Sheets("SVV_TRUCK-data").Select    
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Range_Temp, Version:=xlPivotTableVersion10). _
        CreatePivotTable TableDestination:="Sheet10!R3C1", TableName:="PivotTable1" _
        , DefaultVersion:=xlPivotTableVersion10
    Sheets("Sheet10").Select    
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Cycle Time"), "Count of Cycle Time", xlCount
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO CD")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Compliance T/F ")
        .Orientation = xlColumnField
        .Position = 1
    End With
    Sheets("Sheet10").Select
    Sheets("Sheet10").Name = "SVV_TRUCK_Summary-Pivot"
	CreatePivot_TRAILERdata
End Sub

Sub CreatePivot_TRAILERdata()
	DetermineLastRowToValidate "SVV_UTILITY TRAILER-data"
	Range_Temp = "SVV_UTILITY TRAILER-data!R1C1:R" & LastRow & "C18"
    Sheets("SVV_UTILITY TRAILER-data").Select    
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Range_Temp, Version:=xlPivotTableVersion10). _
        CreatePivotTable TableDestination:="Sheet11!R3C1", TableName:="PivotTable1" _
        , DefaultVersion:=xlPivotTableVersion10
    Sheets("Sheet11").Select
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Cycle Time"), "Count of Cycle Time", xlCount
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO CD")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Compliance T/F ")
        .Orientation = xlColumnField
        .Position = 1
    End With
    Sheets("Sheet11").Select
    Sheets("Sheet11").Name = "SVV_UTILITY_Summary-Pivot"
	CreatePivot_RECREATIONALdata
End Sub

Sub CreatePivot_RECREATIONALdata()
	DetermineLastRowToValidate "SVV_RECREATIONAL-data"
	Range_Temp = "SVV_RECREATIONAL-data!R1C1:R" & LastRow & "C18"
    Sheets("SVV_RECREATIONAL-data").Select    
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Range_Temp, Version:=xlPivotTableVersion10). _
        CreatePivotTable TableDestination:="Sheet12!R3C1", TableName:="PivotTable1" _
        , DefaultVersion:=xlPivotTableVersion10
    Sheets("Sheet12").Select
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Cycle Time"), "Count of Cycle Time", xlCount
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO CD")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Compliance T/F ")
        .Orientation = xlColumnField
        .Position = 1
    End With
    Sheets("Sheet12").Select
    Sheets("Sheet12").Name = "SVV_RECREATIONAL_Summary-Pivot"
	CreatePivot_OTHERdata
End Sub

Sub CreatePivot_OTHERdata()
	DetermineLastRowToValidate "SVV_OTHER-data"
	Range_Temp = "SVV_OTHER-data!R1C1:R" & LastRow & "C18"
    Sheets("SVV_OTHER-data").Select    
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Range_Temp, Version:=xlPivotTableVersion10). _
        CreatePivotTable TableDestination:="Sheet13!R3C1", TableName:="PivotTable1" _
        , DefaultVersion:=xlPivotTableVersion10
    Sheets("Sheet13").Select
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Cycle Time"), "Count of Cycle Time", xlCount
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("CO CD")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Compliance T/F ")
        .Orientation = xlColumnField
        .Position = 1
    End With
    Sheets("Sheet13").Select
    Sheets("Sheet13").Name = "SVV_OTHER_Summary-Pivot"
End Sub


'--------------------------------------------------------------------------------------------------

Sub ProcessSVVTab()
'delete PL/SQL Numbering column if not already deleted; this assumes sort order on sheet hasn't been changed
If Worksheets("Table").Cells(2, "A").Value = "1" Then
    Worksheets("Table").Columns("A").Delete
End If
'determine LastSVVRowToProcess
Dim nRowCounter As Integer
nRowCounter = 1
Do Until Worksheets("Table").Cells(nRowCounter, "A").Value = ""
    nRowCounter = nRowCounter + 1
Loop
LastSVVRowToProcess = nRowCounter - 1
'insert columns for adjusted times and reasons; two more columns will be used, but they are at
'the end of the sheet so no inserts are necessary for them
Worksheets("Table").Columns("I").Insert
Worksheets("Table").Columns("I").Insert
Worksheets("Table").Columns("M").Insert
Worksheets("Table").Columns("M").Insert
'add column headers
Worksheets("Table").Cells(1, "B").Value = "VALREQ ID"
Worksheets("Table").Cells(1, "C").Value = "CLAIM-EXP NUMBER"
Worksheets("Table").Cells(1, "D").Value = "SVV TYPE"
Worksheets("Table").Cells(1, "E").Value = "CREATED BY"
Worksheets("Table").Cells(1, "F").Value = "DATE/TIME ASSIGNMENT ARRIVED IN ESRT"
Worksheets("Table").Cells(1, "G").Value = "DATE/TIME ASSIGNMENT WAS ASSIGNED TO RESEARCHER"
Worksheets("Table").Cells(1, "H").Value = "ELAPSED TIME FROM ARRIVING IN ESRT TO BEING ASSIGNED (dd:hh:mm:ss)"
Worksheets("Table").Cells(1, "I").Value = "COMPUTED CYCLE TIME FROM ARRIVING IN ESRT TO BEING ASSIGNED, WITH ADJUSTMENTS (dd:hh:mm:ss)"
Worksheets("Table").Cells(1, "J").Value = "REASONS FOR ADJUSTMENTS (previous column)"
Worksheets("Table").Cells(1, "K").Value = "DATE/TIME DEALER QUOTES SUBMITTED"
Worksheets("Table").Cells(1, "L").Value = "ELAPSED TIME FROM BEING ASSIGNED TO RESEARCHER TO DEALER QUOTES BEING SUBMITTED (dd:hh:mm:ss)"
Worksheets("Table").Cells(1, "M").Value = "COMPUTED CYCLE TIME FROM BEING ASSIGNED TO DEALER QUOTES BEING SUBMITTED, WITH ADJUSTMENTS (dd:hh:mm:ss)"
Worksheets("Table").Cells(1, "N").Value = "REASONS FOR ADJUSTMENTS (previous column)"
Worksheets("Table").Cells(1, "O").Value = "TOTAL ELAPSED TIME IN ESRT (FROM ARRIVING IN ESRT TO DQS SUBMITTED) (dd:hh:mm:ss)"
Worksheets("Table").Cells(1, "P").Value = "Cycle Time"
Worksheets("Table").Cells(1, "Q").Value = "REASONS FOR ADJUSTMENTS (previous column)"
'center, bold, and wrap
Worksheets("Table").Range("A1:T1").HorizontalAlignment = xlCenter
Worksheets("Table").Range("A1:T1").WrapText = True
Worksheets("Table").Range("A1:T1").Font.Bold = True
FormatSVVColumns  'separate proc in case it needs to be called separately
ExcludeSVVTime
ProcessCompliance
Clear103WeekEnd
CreateSubTab
End Sub

Sub FormatSVVColumns()
Worksheets("Table").Activate
Worksheets("Table").Columns(ComputedTimeArrivedToAssignedWithAdjCol).Select
Selection.NumberFormat = "dd:hh:mm:ss"
Worksheets("Table").Columns(ComputedTimeAssignedToDQCol).Select
Selection.NumberFormat = "dd:hh:mm:ss"
Worksheets("Table").Columns(ComputedTotalTimeWithAdjCol).Select
Selection.NumberFormat = "dd:hh:mm:ss"
Worksheets("Table").Columns(ElapsedTimeArrivedToAssignedCol).Select
Selection.NumberFormat = "dd:hh:mm:ss"
Worksheets("Table").Columns(ElapsedTimeAssignedToDQCol).Select
Selection.NumberFormat = "dd:hh:mm:ss"
Worksheets("Table").Columns(TotalElapsedTimeCol).Select
Selection.NumberFormat = "dd:hh:mm:ss"
'adjust widths
Worksheets("Table").Columns("A").ColumnWidth = 7  'co cd
Worksheets("Table").Columns("B").ColumnWidth = 8  'Valuation Request ID
Worksheets("Table").Columns("C").ColumnWidth = 18 'claimExp num
Worksheets("Table").Columns("D").ColumnWidth = 14  'SVV type
Worksheets("Table").Columns("E").ColumnWidth = 13  'created by
Worksheets("Table").Columns("F").ColumnWidth = 20
Worksheets("Table").Columns("G").ColumnWidth = 20
Worksheets("Table").Columns("J").ColumnWidth = 18
Worksheets("Table").Columns("K").ColumnWidth = 20
Worksheets("Table").Columns("L").ColumnWidth = 18
Worksheets("Table").Columns("M").ColumnWidth = 18
Worksheets("Table").Columns("N").ColumnWidth = 18
Worksheets("Table").Columns("O").ColumnWidth = 18
Worksheets("Table").Columns("P").ColumnWidth = 18
Worksheets("Table").Columns("Q").ColumnWidth = 18
'format two new columns at end in MS San Serif 8
Worksheets("Table").Columns("P:T").Select   'Cells.Select
    With Selection.Font
        .Name = "MS Sans Serif"
        .Size = 9
    End With
Selection.RowHeight = 12.75
Worksheets("Table").Range("A1:T1").Cells.Select
Selection.RowHeight = 90
'freeze headers
Worksheets("Table").Rows("2:2").Select
ActiveWindow.FreezePanes = True
End Sub


Sub ExcludeSVVTime()
Dim ComputedElapsedTimeCol As String
Dim ReasonsCol As String
Dim nStartDateMonthPart As Integer
Dim nStartDateDayPart As Integer
Dim nEndDateMonthPart As Integer
Dim nEndDateDayPart As Integer
Dim nWeekendCounter As Integer
Dim nHolidayCounter As Integer
Dim dtStartDate As Date
Dim dtEndDate As Date
Dim dtEffectiveStartDateTime As Date
Dim dtEffectiveEndDateTime As Date
Dim StartDateDayofYear As Integer
Dim EndDateDayofYear As Integer
Dim j As Integer
Dim k As Integer
For j = FirstRowToProcess To LastSVVRowToProcess
    For k = 1 To 3
        Select Case k
            Case 1   'arrived in ESRT to assigned to researcher
                dtStartDate = Worksheets("Table").Cells(j, ArrivedInESRTCol).Value
                dtEndDate = Worksheets("Table").Cells(j, AssignedToResearcherCol).Value
                ComputedElapsedTimeCol = ComputedTimeArrivedToAssignedWithAdjCol
                ReasonsCol = ReasonsForAdjCol1
            Case 2   'assigned to researcher to DQ submitted
                dtStartDate = Worksheets("Table").Cells(j, AssignedToResearcherCol).Value
                dtEndDate = Worksheets("Table").Cells(j, DQSubmittedCol).Value
                ComputedElapsedTimeCol = ComputedTimeAssignedToDQCol
                ReasonsCol = ReasonsForAdjCol2
            Case 3   'arrived in ESRT to DQ submitted (total time in ESRT)
                dtStartDate = Worksheets("Table").Cells(j, ArrivedInESRTCol).Value
                dtEndDate = Worksheets("Table").Cells(j, DQSubmittedCol).Value
                ComputedElapsedTimeCol = ComputedTotalTimeWithAdjCol
                ReasonsCol = ReasonsForAdjCol3
        End Select
        'determine effective start and end dates
        dtEffectiveStartDateTime = DetermineEffectiveStartDateTime(dtStartDate)
        dtEffectiveEndDateTime = DetermineEffectiveEndDateTime(dtEndDate)
        'subtract effective dates
        tempdiff = dtEffectiveEndDateTime - dtEffectiveStartDateTime
        'count weekends and holiday spans; at this point, there should be no effective start or end days on weekends or holidays
        StartDateDayofYear = DatePart("y", dtEffectiveStartDateTime)   'this would be like 341 for Dec 7
        EndDateDayofYear = DatePart("y", dtEffectiveEndDateTime)
 If (StartDateDayofYear > EndDateDayofYear) Then
        ndaysbetween = Abs(366 + EndDateDayofYear - StartDateDayofYear) - 1 'days between doesn't count 'bookend days'
        Else
        ndaysbetween = EndDateDayofYear - StartDateDayofYear - 1
        End If
        nWeekendCounter = 0
        nHolidayCounter = 0
        'this doesn't account for Memorial's Day, but there weren't any that started that day, so logic not added;
        '  if ever a report is done across the year-end break, that wouldn't be included, but the logic for
        '  day of year wouldn't work anyway so special logic would have to be written.
        For i = 1 To ndaysbetween
            If Weekday(dtEffectiveStartDateTime + i) = vbSaturday Or Weekday(dtEffectiveStartDateTime + i) = vbSunday Then
                nWeekendCounter = nWeekendCounter + 1
            Else
               If DateValue((dtEffectiveStartDateTime + i)) = "11/24/2011" Or _
                  DateValue((dtEffectiveStartDateTime + i)) = "11/25/2011" Then
                  nHolidayCounter = nHolidayCounter + 1
                  strReasons = strReasons & "--Spanned Thanksgiving holiday"
                Else
                    If DateValue((dtEffectiveStartDateTime + i)) = "12/23/2011" Or _
                       DateValue((dtEffectiveStartDateTime + i)) = "12/26/2011" Then
                       nHolidayCounter = nHolidayCounter + 1
                       strReasons = strReasons & "--Spanned Christmas holiday"
                       Else
                       If DateValue((dtEffectiveStartDateTime + i)) = "05/28/2012" Then
                            nHolidayCounter = nHolidayCounter + 1
                            strReasons = strReasons & "--Spanned Memorial Day's holiday"
                       End If
                    End If
               End If
            End If
        Next i
        'subtract weekends and holidays
        If nWeekendCounter > 0 Then
            strReasons = strReasons & "--Spanned " & nWeekendCounter & " Weekend Day(s)"
        End If
        tempdiff = tempdiff - (nWeekendCounter) - (nHolidayCounter)
        If tempdiff < 0 Then tempdiff = 0
        Worksheets("Table").Cells(j, ComputedElapsedTimeCol).Value = tempdiff
        If strReasons = "--No adjustment to Start Date--No adjustment to End Date" Then
            Worksheets("Table").Cells(j, ReasonsCol).Value = "No adjustments"
        Else
            Worksheets("Table").Cells(j, ReasonsCol).Value = strReasons
        End If
        strReasons = ""
    Next k
Next j
End Sub

Function DetermineEffectiveStartDateTime(inStartDate As Date) As Date
'see note earlier about Memorial's Day not being here
   If DateValue(inStartDate) = "11/24/2011" Or DateValue(inStartDate) = "11/25/2011" Then 'Thanksgiving and Friday after it
      strReasons = strReasons & "--Started during Thanksgiving holiday"
      DetermineEffectiveStartDateTime = "11/28/2011 00:00:00 AM"  'effective date is following Monday
   Else
      If DateValue(inStartDate) = "12/23/2011" Or DateValue(inStartDate) = "12/26/2011" Then 'Christmas Eve and Christmas
        strReasons = strReasons & "--Started during Christmas holiday"
        DetermineEffectiveStartDateTime = "12/27/2011 00:00:00 AM"   'effective date is following Friday
      Else
        If DateValue(inStartDate) = "05/28/2012" Then  'Memorial's Day
          strReasons = strReasons & "--Started during Memorial Day's Holiday"
          DetermineEffectiveStartDateTime = "05/29/2012 00:00:00 AM"   'effective date is following Friday
      Else
         Select Case Weekday(inStartDate)
            Case vbSaturday  'effective date is following Monday at 0:00:00 AM
                strReasons = strReasons & "--Started on a Saturday"
                DetermineEffectiveStartDateTime = (DateValue(inStartDate) + 2)
            Case vbSunday    'effective date is following Monday at 0:00:00 AM
                strReasons = strReasons & "--Started on a Sunday"
                DetermineEffectiveStartDateTime = (DateValue(inStartDate) + 1)
            Case Else
                strReasons = strReasons & "--No adjustment to Start Date"
                DetermineEffectiveStartDateTime = inStartDate
        End Select
      End If
    End If
    End If
End Function

Function DetermineEffectiveEndDateTime(inEndDate As Date) As Date
'see note earlier about Memorial's Day not being here
   If DateValue(inEndDate) = "11/24/2011" Or DateValue(inEndDate) = "11/25/2011" Or _
        DateValue(inEndDate) = "11/26/2011" Or DateValue(inEndDate) = "11/27/2011" Then  'Thanksgiving Thursday, Friday, Sat or Sun
           strReasons = strReasons & "--Ended during Thanksgiving holiday"
           DetermineEffectiveEndDateTime = "11/23/2011 11:59:59 PM"   'midnight Wed night
   Else
       If DateValue(inEndDate) = "12/23/2011" Or DateValue(inEndDate) = "12/26/2011" Then   ' Ch Eve, and Christmas
           strReasons = strReasons & "--Ended during Christmas holiday"
           DetermineEffectiveEndDateTime = "12/22/2011 11:59:59 PM"
       Else
            If DateValue(inEndDate) = "05/28/2012" Then   ' Memorial's Day
                strReasons = strReasons & "--Ended during Memorial Day's Holiday"
                DetermineEffectiveEndDateTime = "05/27/2012 11:59:59 PM"
       Else
            Select Case Weekday(inEndDate)
                Case vbSaturday   'effective date is that previous Friday midnight  11:59:59 PM
                    strReasons = strReasons & "--Ended on a Saturday"
                    DetermineEffectiveEndDateTime = (DateValue(inEndDate) - 0.00001)
                Case vbSunday     'effective date is that previous Friday midnight
                    strReasons = strReasons & "--Ended on a Sunday"
                    DetermineEffectiveEndDateTime = (DateValue(inEndDate) - 1.00001)
                Case Else
                    strReasons = strReasons & "--No adjustment to End Date"
                    DetermineEffectiveEndDateTime = inEndDate
            End Select
       End If
   End If
   End If
End Function

Sub ProcessCompliance()

'insert columns for "Compliance T/F"
Worksheets("Table").Columns("Q").Insert

'add column headers
Worksheets("Table").Cells(1, "Q").Value = "Compliance T/F "

'sort sheet by Cycle Time TRUCK
Worksheets("Table").Range("A2:T" & LastSVVRowToProcess).Sort Key1:=Worksheets("Table").Range("P" & "2")

Dim c As Integer
For c = FirstRowToProcess To LastSVVRowToProcess

    If Worksheets("Table").Cells(c, "D").Value = "OTHER" Or Worksheets("Table").Cells(c, "D").Value = "MOTORCYCLE" _
    Or Worksheets("Table").Cells(c, "D").Value = "UTILITY TRAILER" Or Worksheets("Table").Cells(c, "D").Value = "CLASSIC" Then
            If Worksheets("Table").Cells(c, "P").Value <= 1 Then
            Worksheets("Table").Cells(c, "Q").Value = "TRUE"
           Else
            Worksheets("Table").Cells(c, "Q").Value = "FALSE"
           End If
    Else
    If Worksheets("Table").Cells(c, "D").Value = "RECREATIONAL" Or Worksheets("Table").Cells(c, "D").Value = "TRUCK" _
	Or Worksheets("Table").Cells(c, "D").Value = "MARINE" Then
            If Worksheets("Table").Cells(c, "P").Value <= 2 Then
               Worksheets("Table").Cells(c, "Q").Value = "TRUE"
             Else
               Worksheets("Table").Cells(c, "Q").Value = "FALSE"
            End If
    End If
  End If
 Next c

End Sub

Sub CreateSubTab()

'create a tab called SVV_CLASSIC-data
Dim wrkSheet As Worksheet
Set wrkSheet = ActiveWorkbook.Worksheets.Add
wrkSheet.Name = "SVV_CLASSIC-data"
Format "SVV_CLASSIC-data"

'create a tab called SVV_MOTORCYCLE-data
Set wrkSheet = ActiveWorkbook.Worksheets.Add
wrkSheet.Name = "SVV_MOTORCYCLE-data"
Format "SVV_MOTORCYCLE-data"

'create a tab called SVV_TRUCK-data
Set wrkSheet = ActiveWorkbook.Worksheets.Add
wrkSheet.Name = "SVV_TRUCK-data"
Format "SVV_TRUCK-data"

'create a tab called SVV_UTILITY TRAILER-data
Set wrkSheet = ActiveWorkbook.Worksheets.Add
wrkSheet.Name = "SVV_UTILITY TRAILER-data"
Format "SVV_UTILITY TRAILER-data"

'create a tab called SVV_RECREATIONAL-data
Set wrkSheet = ActiveWorkbook.Worksheets.Add
wrkSheet.Name = "SVV_RECREATIONAL-data"
Format "SVV_RECREATIONAL-data"

'create a tab called SVV_OTHER-data
Set wrkSheet = ActiveWorkbook.Worksheets.Add
wrkSheet.Name = "SVV_OTHER-data"
Format "SVV_OTHER-data"

Dim MAIN As Integer
Dim CLASSIC As Integer
Dim MOTORCYCLE As Integer
Dim TRUCK As Integer
Dim UTILITY As Integer
Dim RECREATIONAL As Integer
Dim OTHER As Integer

CLASSIC = 2
MOTORCYCLE = 2
TRUCK = 2
UTILITY = 2
RECREATIONAL = 2
OTHER = 2

For MAIN = 2 To LastSVVRowToProcess
            If Worksheets("Table").Cells(MAIN, "D").Value = "CLASSIC" Then
                    Sheets("SVV_CLASSIC-data").Range("A" & CLASSIC, "T" & CLASSIC).Value = Sheets("Table").Range("A" & MAIN, "T" & MAIN).Value
                    CLASSIC = CLASSIC + 1
               
            Else
            If Worksheets("Table").Cells(MAIN, "D").Value = "MOTORCYCLE" Then
                    Sheets("SVV_MOTORCYCLE-data").Range("A" & MOTORCYCLE, "T" & MOTORCYCLE).Value = Sheets("Table").Range("A" & MAIN, "T" & MAIN).Value
                    MOTORCYCLE = MOTORCYCLE + 1
                   
            Else
            If Worksheets("Table").Cells(MAIN, "D").Value = "TRUCK" Then
                    Sheets("SVV_TRUCK-data").Range("A" & TRUCK, "T" & TRUCK).Value = Sheets("Table").Range("A" & MAIN, "T" & MAIN).Value
                    TRUCK = TRUCK + 1
                    
                
            Else
            If Worksheets("Table").Cells(MAIN, "D").Value = "UTILITY TRAILER" Then
                    Sheets("SVV_UTILITY TRAILER-data").Range("A" & UTILITY, "T" & UTILITY).Value = Sheets("Table").Range("A" & MAIN, "T" & MAIN).Value
                    UTILITY = UTILITY + 1
                    
            Else
            If Worksheets("Table").Cells(MAIN, "D").Value = "RECREATIONAL" Then
                    Sheets("SVV_RECREATIONAL-data").Range("A" & RECREATIONAL, "T" & RECREATIONAL).Value = Sheets("Table").Range("A" & MAIN, "T" & MAIN).Value
                    RECREATIONAL = RECREATIONAL + 1
                   
            Else
            If Worksheets("Table").Cells(MAIN, "D").Value = "OTHER" Then
                    Sheets("SVV_OTHER-data").Range("A" & OTHER, "T" & OTHER).Value = Sheets("Table").Range("A" & MAIN, "T" & MAIN).Value
                    OTHER = OTHER + 1
                    End If
                End If
               End If
            End If
        End If
      End If
    Next MAIN

End Sub

Sub Format(sheetName As String)

Sheets(sheetName).Range("A1:T1").Value = Sheets("Table").Range("A1:T1").Value

Worksheets(sheetName).Activate
Worksheets(sheetName).Columns(ComputedTimeArrivedToAssignedWithAdjCol).Select
Selection.NumberFormat = "dd:hh:mm:ss"
Worksheets(sheetName).Columns(ComputedTimeAssignedToDQCol).Select
Selection.NumberFormat = "dd:hh:mm:ss"
Worksheets(sheetName).Columns(ComputedTotalTimeWithAdjCol).Select
Selection.NumberFormat = "dd:hh:mm:ss"
Worksheets(sheetName).Columns(ElapsedTimeArrivedToAssignedCol).Select
Selection.NumberFormat = "dd:hh:mm:ss"
Worksheets(sheetName).Columns(ElapsedTimeAssignedToDQCol).Select
Selection.NumberFormat = "dd:hh:mm:ss"
Worksheets(sheetName).Columns(TotalElapsedTimeCol).Select
Selection.NumberFormat = "dd:hh:mm:ss"
'Make column headers bold, wrap text in them, and center
Worksheets(sheetName).Range("A1:T1").HorizontalAlignment = xlCenter
Worksheets(sheetName).Range("A1:T1").WrapText = True
Worksheets(sheetName).Range("A1:T1").Font.Bold = True
'format sheet
Worksheets(sheetName).Activate

'adjust widths
Worksheets(sheetName).Columns("A").ColumnWidth = 7  'co cd
Worksheets(sheetName).Columns("B").ColumnWidth = 8  'Valuation Request ID
Worksheets(sheetName).Columns("C").ColumnWidth = 18 'claimExp num
Worksheets(sheetName).Columns("D").ColumnWidth = 14  'SVV type
Worksheets(sheetName).Columns("E").ColumnWidth = 13  'created by
Worksheets(sheetName).Columns("F").ColumnWidth = 18
Worksheets(sheetName).Columns("G").ColumnWidth = 18
Worksheets(sheetName).Columns("H").ColumnWidth = 18
Worksheets(sheetName).Columns("I").ColumnWidth = 18
Worksheets(sheetName).Columns("J").ColumnWidth = 18
Worksheets(sheetName).Columns("K").ColumnWidth = 18
Worksheets(sheetName).Columns("L").ColumnWidth = 18
Worksheets(sheetName).Columns("M").ColumnWidth = 18
Worksheets(sheetName).Columns("N").ColumnWidth = 18
Worksheets(sheetName).Columns("O").ColumnWidth = 18
Worksheets(sheetName).Columns("P").ColumnWidth = 18
Worksheets(sheetName).Columns("Q").ColumnWidth = 18
Worksheets(sheetName).Columns("R").ColumnWidth = 24
'format two new columns at end in MS San Serif 9
Worksheets(sheetName).Columns("A:T").Select   'Cells.Select
    With Selection.Font
        .Name = "MS Sans Serif"
        .Size = 9
    End With
Selection.RowHeight = 12.75
Worksheets(sheetName).Range("A1:T1").Cells.Select
Selection.RowHeight = 75
'freeze headers
Worksheets(sheetName).Rows("2:2").Select
ActiveWindow.FreezePanes = True

End Sub


'------------------------------------------------------------------Find Last Row for Pivot-----------------------------
Sub DetermineLastRowPivot(sheetName As String)
    Dim nRowCounter As Integer
    nRowCounter = 3
    Do Until Worksheets(sheetName).Cells(nRowCounter, "A").Value = ""
        nRowCounter = nRowCounter + 1
    Loop
    LastRowPivot = nRowCounter - 1
End Sub    

'-------------------------------------------------------------Create Summary Header------------------------------------------
Sub CreateSummary_Header()
'
' CreateSummaryHeader Macro
'

'
    Dim HeadRow1 As Integer
    Dim HeadRow2 As Integer
    HeadRow1 = LastRowPivot + 5
    HeadRow2 = LastRowPivot + 6
    Dim TempRange As String
    TempRange = "A" & HeadRow1 & ":B" & HeadRow1
    Sheets(CurrSheet).Range(TempRange).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = " SVV-All"
    TempRange = "C" & HeadRow1 & ":D" & HeadRow1
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveCell.FormulaR1C1 = "Average Cycle time"
    TempRange = "F" & HeadRow1
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveCell.FormulaR1C1 = "Not in Compliance with SLA"
    TempRange = "I" & HeadRow1
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveCell.FormulaR1C1 = "In Compliance with SLA"
    TempRange = "A" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveCell.FormulaR1C1 = "Companies Serviced"
    TempRange = "B" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveCell.FormulaR1C1 = "Total Number of Calls"
    TempRange = "C" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveCell.FormulaR1C1 = "In Days"
    TempRange = "D" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveCell.FormulaR1C1 = "In Hours"
    TempRange = "E" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveCell.FormulaR1C1 = "Avg. Cycle Time in Days"
    TempRange = "F" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveCell.FormulaR1C1 = "Avg. Cycle Time in Hours"
    TempRange = "G" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveCell.FormulaR1C1 = "Occurrence"
    TempRange = "H" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveCell.FormulaR1C1 = "Avg. Cycle Time in Days"
    TempRange = "I" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveCell.FormulaR1C1 = "Avg. Cycle Time in Hours"
    TempRange = "J" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveCell.FormulaR1C1 = "Occurrence"
    TempRange = "K" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveCell.FormulaR1C1 = "% in Compliance"
    TempRange = "A" & HeadRow1 & ":B" & HeadRow1
    Sheets(CurrSheet).Range(TempRange).Select
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Font.Bold = True
    With Selection.Font
        .Name = "Arial"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    TempRange = "C" & HeadRow1
    Sheets(CurrSheet).Range(TempRange).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Font.Bold = True
    With Selection.Font
        .Name = "Arial"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    TempRange = "C" & HeadRow1 & ":D" & HeadRow1
    Sheets(CurrSheet).Range(TempRange).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    TempRange = "F" & HeadRow1
    Sheets(CurrSheet).Range(TempRange).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Font.Bold = True
    With Selection.Font
        .Name = "Arial"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    TempRange = "I" & HeadRow1
    Sheets(CurrSheet).Range(TempRange).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Font.Bold = True
    With Selection.Font
        .Name = "Arial"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    TempRange = "A" & HeadRow1 & ":K" & HeadRow1
    Sheets(CurrSheet).Range(TempRange).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark2
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    TempRange = "A" & HeadRow1 & ":K" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    TempRange = "A" & HeadRow1 & ":B" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    TempRange = "C" & HeadRow1 & ":D" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    TempRange = "E" & HeadRow1 & ":G" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    TempRange = "H" & HeadRow1 & ":J" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    TempRange = "K" & HeadRow1 & ":K" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    TempRange = "A" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    TempRange = "B" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    TempRange = "C" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    TempRange = "D" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    TempRange = "E" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    TempRange = "F" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    TempRange = "G" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    TempRange = "H" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    TempRange = "I" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    TempRange = "J" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    TempRange = "K" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    TempRange = "A" & HeadRow2 & ":K" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    TempRange = "A" & HeadRow2 & ":B" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    TempRange = "C" & HeadRow2 & ":D" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    TempRange = "E" & HeadRow2 & ":G" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    TempRange = "H" & HeadRow2 & ":J" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    TempRange = "K" & HeadRow2
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Sheets(CurrSheet).Columns("E:E").ColumnWidth = 17.71
    Sheets(CurrSheet).Columns("H:H").ColumnWidth = 18.29
    ActiveWindow.ScrollColumn = 2
    ActiveWindow.ScrollColumn = 1
    Sheets(CurrSheet).Columns("B:B").ColumnWidth = 16.29
    Sheets(CurrSheet).Columns("C:C").ColumnWidth = 15.14
    Sheets(CurrSheet).Columns("D:D").ColumnWidth = 12.71
    Sheets(CurrSheet).Columns("E:E").ColumnWidth = 17.14
    Sheets(CurrSheet).Columns("H:H").ColumnWidth = 17.71
    Sheets(CurrSheet).Columns("J:J").ColumnWidth = 11.14
    Sheets(CurrSheet).Columns("G:G").ColumnWidth = 12.86
    Sheets(CurrSheet).Columns("F:F").ColumnWidth = 14.43
    TempRange = HeadRow2 & ":" & HeadRow2
    Sheets(CurrSheet).Rows(TempRange).RowHeight = 27
    Sheets(CurrSheet).Columns("I:I").ColumnWidth = 14
    Sheets(CurrSheet).Columns("K:K").ColumnWidth = 13
End Sub


'-----------------------------------------------------Add Data to Summary Table----------------------------------------------


Sub AddSummaryData()
'
' AddSummaryData Macro
'

'
    Dim TempRange As String
    Dim SummaryRow1 As Integer
    SummaryRow1 = LastRowPivot + 7
    TempRange = "A5:A" & LastRowPivot
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Copy
    TempRange = "A" & SummaryRow1
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveSheet.Paste
    'ActiveWindow.SmallScroll Down:=3
    TempRange = "D5:D" & LastRowPivot
    Sheets(CurrSheet).Range(TempRange).Select
    Application.CutCopyMode = False
    Selection.Copy
    TempRange = "B" & SummaryRow1
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveSheet.Paste
    TempRange = "B5:B" & LastRowPivot
    Sheets(CurrSheet).Range(TempRange).Select
    Application.CutCopyMode = False
    Selection.Copy
    TempRange = "G" & SummaryRow1
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveSheet.Paste
    TempRange = "C5:C" & LastRowPivot
    Sheets(CurrSheet).Range(TempRange).Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveWindow.ScrollColumn = 2
    TempRange = "J" & SummaryRow1
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveSheet.Paste
    ActiveWindow.ScrollColumn = 1
    TempRange = "C" & SummaryRow1
    Sheets(CurrSheet).Range(TempRange).Select
    Application.CutCopyMode = False
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Count of Cycle Time"). _
        Function = xlAverage
    TempRange = "D5:D" & LastRowPivot
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Copy
    TempRange = "C" & SummaryRow1
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveSheet.Paste
    TempRange = "B5:B" & LastRowPivot
    Sheets(CurrSheet).Range(TempRange).Select
    Application.CutCopyMode = False
    Selection.Copy
    TempRange = "E" & SummaryRow1
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveSheet.Paste
    TempRange = "C5:C" & LastRowPivot
    Sheets(CurrSheet).Range(TempRange).Select
    Application.CutCopyMode = False
    Selection.Copy
    TempRange = "H" & SummaryRow1
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveSheet.Paste
    'TempRange="D" & SummaryRow1
    'Sheets(CurrSheet).Range(TempRange).Select
    'Application.CutCopyMode = False
    'ActiveCell.FormulaR1C1 = "=(RC[-1]*24)"
    'TempRange="D" & SummaryRow1
    'Sheets(CurrSheet).Range(TempRange).Select
    'ActiveWindow.SmallScroll Down:=9
    
    Dim LastRowSum As Integer
    Dim nRowCounter As Integer
    nRowCounter = SummaryRow1
    Do Until Worksheets(CurrSheet).Cells(nRowCounter, "A").Value = ""
        nRowCounter = nRowCounter + 1
    Loop
    LastRowSum = nRowCounter - 1
    
    TempRange = "D" & SummaryRow1 & ":D" & LastRowSum
    Sheets(CurrSheet).Range(TempRange).Select
    Application.CutCopyMode = False
    Selection.FormulaR1C1 = "=(RC[-1]*24)"
    
    TempRange = "D" & SummaryRow1 & ":D" & LastRowSum
    'Selection.AutoFill Destination:=Range(TempRange), Type:=xlFillDefault
    Sheets(CurrSheet).Range(TempRange).Select
    TempRange = "D" & SummaryRow1
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Copy
    TempRange = "F" & SummaryRow1
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    TempRange = "F" & SummaryRow1 & ":F" & LastRowSum
    Selection.AutoFill Destination:=Sheets(CurrSheet).Range(TempRange), Type:=xlFillDefault
    Sheets(CurrSheet).Range(TempRange).Select
    TempRange = "F" & SummaryRow1
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Copy
    TempRange = "I" & SummaryRow1 & ":I" & LastRowSum
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveSheet.Paste
    TempRange = "K" & SummaryRow1
    Sheets(CurrSheet).Range(TempRange).Select
    
    'Application.WindowState = xlMinimized
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    'Windows("SVV_Backup_1.xls").Activate
    'Application.WindowState = xlMinimized
    'Windows("CycleTime(SVV) bi-weekly 1 November- 15 November 2014.xls").Activate
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    'ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    'Sheets("SVV_TRUCK_Summary-Pivot").Select
    'Sheets(CurrSheet).Range("K30").Select
    'Windows("SVV_Backup_1.xls").Activate
    'Application.CutCopyMode = False
    
    ActiveCell.FormulaR1C1 = "=(RC[-1]/RC[-9])"
    TempRange = "K" & SummaryRow1
    Sheets(CurrSheet).Range(TempRange).Select
    TempRange = "K" & SummaryRow1 & ":K" & LastRowSum
    Selection.AutoFill Destination:=Sheets(CurrSheet).Range(TempRange), Type:=xlFillDefault
    Sheets(CurrSheet).Range(TempRange).Select
    TempRange = "A" & LastRowSum & ":K" & LastRowSum
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    TempRange = "A" & SummaryRow1 & ":K" & LastRowSum - 1
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlDot
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlDot
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlDot
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlDot
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    TempRange = "B" & SummaryRow1 & ":K" & LastRowSum
    Sheets(CurrSheet).Range(TempRange).Select
    'Selection.NumberFormat = "0.00"
    'Sheets(CurrSheet).Range("B46:K46").Select
    Selection.NumberFormat = "0.00"
    'Sheets(CurrSheet).Range("A46").Select
    'Sheets(CurrSheet).Range(Selection, Selection.End(xlToRight)).Select
    TempRange = "A" & LastRowSum & ":K" & LastRowSum
    Sheets(CurrSheet).Range(TempRange).Select
    ActiveWindow.LargeScroll ToRight:=-1
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    TempRange = "A" & SummaryRow1 & ":K" & LastRowSum
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    TempRange = "A" & SummaryRow1 & ":K" & LastRowSum - 1
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    TempRange = "A" & SummaryRow1
    Sheets(CurrSheet).Range(TempRange).Select
    Sheets(CurrSheet).Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    TempRange = "B" & SummaryRow1
    Sheets(CurrSheet).Range(TempRange).Select
    Sheets(CurrSheet).Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Sheets(CurrSheet).Range(TempRange).Select
    Sheets(CurrSheet).Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    TempRange = "C" & SummaryRow1
    Sheets(CurrSheet).Range(TempRange).Select
    Sheets(CurrSheet).Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    TempRange = "D" & SummaryRow1
    Sheets(CurrSheet).Range(TempRange).Select
    Sheets(CurrSheet).Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    TempRange = "E" & SummaryRow1 & ":E" & LastRowSum
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    TempRange = "F" & SummaryRow1 & ":F" & LastRowSum
    Sheets(CurrSheet).Range(TempRange).Select
    'Sheets(CurrSheet).Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    TempRange = "G" & SummaryRow1 & ":G" & LastRowSum
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    TempRange = "H" & SummaryRow1 & ":H" & LastRowSum
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    TempRange = "I" & SummaryRow1 & ":I" & LastRowSum
    Sheets(CurrSheet).Range(TempRange).Select
    'Sheets(CurrSheet).Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    TempRange = "J" & SummaryRow1 & ":J" & LastRowSum
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Average of Cycle Time"). _
        Function = xlCount
	TempRange="K" & SummaryRow1 & ":K" & LastRowSum
	Sheets(CurrSheet).Range(TempRange).Select
    Selection.NumberFormat = "0%"
	TempRange="B" & LastRowSum
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.NumberFormat = "General"
	TempRange="G" & LastRowSum
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.NumberFormat = "General"
	TempRange="J" & LastRowSum
    Sheets(CurrSheet).Range(TempRange).Select
    Selection.NumberFormat = "General"
End Sub

'-------------------------------------------------------Coloring Tabs-------------------------------------------
Sub colors()
'
' colors Macro
'

'   
    Sheets("SVV_all-data").Select
    With ActiveWorkbook.Sheets("SVV_all-data").Tab
        .Color = 192
        .TintAndShade = 0
    End With
    Sheets("SVV_All-data_Summary-Pivot").Select
    With ActiveWorkbook.Sheets("SVV_All-data_Summary-Pivot").Tab
        .Color = 192
        .TintAndShade = 0
    End With
    Sheets("SVV_CLASSIC-data").Select
    With ActiveWorkbook.Sheets("SVV_CLASSIC-data").Tab
        .Color = 49407
        .TintAndShade = 0
    End With
    Sheets("SVV_Classic_Summary-Pivot").Select
    With ActiveWorkbook.Sheets("SVV_Classic_Summary-Pivot").Tab
        .Color = 49407
        .TintAndShade = 0
    End With   
    Sheets("SVV_MOTORCYCLE-data").Select
    With ActiveWorkbook.Sheets("SVV_MOTORCYCLE-data").Tab
        .Color = 5296274
        .TintAndShade = 0
    End With
    Sheets("SVV_MOTORCYCLE_Summary-Pivot").Select
    With ActiveWorkbook.Sheets("SVV_MOTORCYCLE_Summary-Pivot").Tab
        .Color = 5296274
        .TintAndShade = 0
    End With    
    Sheets("SVV_TRUCK-data").Select
    With ActiveWorkbook.Sheets("SVV_TRUCK-data").Tab
        .Color = 15773696
        .TintAndShade = 0
    End With
    Sheets("SVV_TRUCK_Summary-Pivot").Select
    With ActiveWorkbook.Sheets("SVV_TRUCK_Summary-Pivot").Tab
        .Color = 15773696
        .TintAndShade = 0
    End With 
    Sheets("SVV_UTILITY TRAILER-data").Select
    With ActiveWorkbook.Sheets("SVV_UTILITY TRAILER-data").Tab
        .Color = 6299648
        .TintAndShade = 0
    End With
    Sheets("SVV_UTILITY_Summary-Pivot").Select
    With ActiveWorkbook.Sheets("SVV_UTILITY_Summary-Pivot").Tab
        .Color = 6299648
        .TintAndShade = 0
    End With    
    Sheets("SVV_RECREATIONAL-data").Select
    With ActiveWorkbook.Sheets("SVV_RECREATIONAL-data").Tab
        .Color = 10498160
        .TintAndShade = 0
    End With
    Sheets("SVV_RECREATIONAL_Summary-Pivot").Select
    With ActiveWorkbook.Sheets("SVV_RECREATIONAL_Summary-Pivot").Tab
        .Color = 10498160
        .TintAndShade = 0
    End With
    ActiveWindow.ScrollWorkbookTabs Sheets:=-1
    Sheets("SVV_OTHER-data").Select
    With ActiveWorkbook.Sheets("SVV_OTHER-data").Tab
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
    End With
    Sheets("SVV_OTHER_Summary-Pivot").Select
    With ActiveWorkbook.Sheets("SVV_OTHER_Summary-Pivot").Tab
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0
    End With    
End Sub
