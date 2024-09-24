Attribute VB_Name = "Module16"
Sub AFilterByMonth()
    Dim wsCost As Worksheet
    Dim wsRework As Worksheet
    Dim wsNCR As Worksheet
    Dim monthName As String
    Dim startDate As Date
    Dim endDate As Date
    Dim currentYear As Integer
    
    ' Set worksheets
    Set wsCost = ThisWorkbook.Sheets("Cost of Poor Quality")
    Set wsRework = ThisWorkbook.Sheets("Rework Data")
    Set wsNCR = ThisWorkbook.Sheets("NCR data")
    
    ' Get the month name from cell E7
    monthName = wsCost.Range("C7").value
    
    ' Set the current year
    currentYear = Year(Date)
    
    ' Calculate start and end dates for the given month in the current year
    startDate = DateSerial(currentYear, Month(dateValue("01 " & monthName)), 1)
    endDate = DateSerial(currentYear, Month(dateValue("01 " & monthName)) + 1, 0)
    
    ' Clear any existing filters
    On Error Resume Next
    wsRework.ListObjects("rework").AutoFilter.ShowAllData
    wsNCR.ListObjects("ncr").AutoFilter.ShowAllData
    On Error GoTo 0
    
    ' Apply filter to "rework" table
    With wsRework.ListObjects("rework").Range
        .AutoFilter Field:=2, Criteria1:=">=" & startDate, Operator:=xlAnd, Criteria2:="<=" & endDate
    End With
    
    ' Apply filter to "ncr" table
    With wsNCR.ListObjects("ncr").Range
        .AutoFilter Field:=2, Criteria1:=">=" & startDate, Operator:=xlAnd, Criteria2:="<=" & endDate
    End With
    
    Call SetQuarterPickerText
    
End Sub

Sub ApplyQuarterFilter()
    Dim wsCPQ As Worksheet
    Dim wsRework As Worksheet
    Dim wsNCR As Worksheet
    Dim qtRework As ListObject
    Dim qtNCR As ListObject
    Dim quarter As String
    Dim startMonth As Integer
    Dim endMonth As Integer
    Dim currentYear As Integer
    Dim startDate As Date
    Dim endDate As Date
    
    ' Set worksheet references
    Set wsCPQ = Worksheets("Cost of Poor Quality")
    Set wsRework = Worksheets("Rework Data")
    Set wsNCR = Worksheets("NCR data")
    
    ' Set table references
    Set qtRework = wsRework.ListObjects("rework")
    Set qtNCR = wsNCR.ListObjects("ncr")
    
    ' Get the quarter from cell E16
    quarter = wsCPQ.Range("C16").value
    
    ' Determine the start and end months based on the quarter
    Select Case quarter
        Case "Quarter 1"
            startMonth = 1
            endMonth = 3
        Case "Quarter 2"
            startMonth = 4
            endMonth = 6
        Case "Quarter 3"
            startMonth = 7
            endMonth = 9
        Case "Quarter 4"
            startMonth = 10
            endMonth = 12
        Case Else
            MsgBox "Invalid quarter in cell C16"
            Exit Sub
    End Select
    
    ' Get the current year
    currentYear = Year(Date)
    
    ' Set the start and end dates
    startDate = DateSerial(currentYear, startMonth, 1)
    endDate = DateSerial(currentYear, endMonth + 1, 1) - 1
    
    ' Clear any existing filters
    If qtRework.AutoFilter.FilterMode Then qtRework.AutoFilter.ShowAllData
    If qtNCR.AutoFilter.FilterMode Then qtNCR.AutoFilter.ShowAllData
    
    ' Apply the filter for "Rework Data" table
    qtRework.Range.AutoFilter Field:=2, _
        Criteria1:=">=" & startDate, _
        Operator:=xlAnd, _
        Criteria2:="<=" & endDate
    
    ' Apply the filter for "NCR data" table
    qtNCR.Range.AutoFilter Field:=2, _
        Criteria1:=">=" & startDate, _
        Operator:=xlAnd, _
        Criteria2:="<=" & endDate
    
    Call SetMonthPickerText
    
End Sub


Sub SetMonthPickerText()
    Dim wsCPQ As Worksheet
    
    ' Set worksheet reference
    Set wsCPQ = Worksheets("Cost of Poor Quality")
    
    ' Set cell E7 to the specified text
    wsCPQ.Range("E7").value = "Click here to pick a month"
End Sub


Sub SetQuarterPickerText()
    Dim wsCPQ As Worksheet
    
    ' Set worksheet reference
    Set wsCPQ = Worksheets("Cost of Poor Quality")
    
    ' Set cell E16 to the specified text
    wsCPQ.Range("C16").value = "Click here to pick a quarter"
End Sub

Sub SumVisibleReworkData()
    Dim wsRework As Worksheet
    Dim wsCPQ As Worksheet
    Dim cell As Range
    Dim sum As Double
    
    ' Set worksheet references
    Set wsRework = Worksheets("Rework Data")
    Set wsCPQ = Worksheets("Cost of Poor Quality")
    
    ' Initialize the sum
    sum = 0
    
    ' Loop through each visible cell in column C and sum them if they are not blank
    On Error Resume Next ' Handle case where no cells are visible
    For Each cell In wsRework.Range("C2:C" & wsRework.Cells(wsRework.Rows.Count, "C").End(xlUp).Row).SpecialCells(xlCellTypeVisible)
        If IsNumeric(cell.value) And cell.value <> "" Then
            sum = sum + cell.value
        End If
    Next cell
    On Error GoTo 0 ' Turn off error handling
    
    ' Output the sum to cell P8 on the "Cost of Poor Quality" sheet
    wsCPQ.Range("P8").value = sum
End Sub



Sub SumVisibleCells()
    Dim wsNCR As Worksheet
    Dim wsCPQ As Worksheet
    Dim rngC As Range
    Dim rngD As Range
    Dim cell As Range
    Dim sumC As Double
    Dim sumD As Double
    
    ' Set worksheet references
    Set wsNCR = Worksheets("NCR Data")
    Set wsCPQ = Worksheets("Cost of Poor Quality")
    
    ' Set range references for columns C and D in "NCR Data"
    Set rngC = wsNCR.Range("C2:C" & wsNCR.Cells(wsNCR.Rows.Count, "C").End(xlUp).Row)
    Set rngD = wsNCR.Range("D2:D" & wsNCR.Cells(wsNCR.Rows.Count, "D").End(xlUp).Row)
    
    ' Initialize sums
    sumC = 0
    sumD = 0
    
    ' Sum visible cells in column C, skipping blanks
    For Each cell In rngC
        If cell.EntireRow.Hidden = False And IsNumeric(cell.value) And cell.value <> "" Then
            sumC = sumC + cell.value
        End If
    Next cell
    
    ' Sum visible cells in column D, skipping blanks
    For Each cell In rngD
        If cell.EntireRow.Hidden = False And IsNumeric(cell.value) And cell.value <> "" Then
            sumD = sumD + cell.value
        End If
    Next cell
    
    ' Output sums to "Cost of Poor Quality" sheet
    wsCPQ.Range("P11").value = sumC
    wsCPQ.Range("P14").value = sumD
End Sub


Sub Mastermonthqual()
    Call AFilterByMonth
    Call SumVisibleReworkData
    Call SumVisibleCells
    Call SetCellC16
End Sub


Sub Masterquarterqual()
    Call ApplyQuarterFilter
    Call SumVisibleReworkData
    Call SumVisibleCells
    Call SetCellC7
End Sub

Sub SetCellC16()
    ' Set the value of cell C16 to "Click here to pick a quarter"
    ThisWorkbook.Sheets("Cost of Poor Quality").Range("C16").value = "Click here to pick a quarter"
End Sub

Sub SetCellC7()
    ' Set the value of cell C16 to "Click here to pick a quarter"
    ThisWorkbook.Sheets("Cost of Poor Quality").Range("C7").value = "Click here to pick a month"
End Sub
