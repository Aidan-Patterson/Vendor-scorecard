Attribute VB_Name = "Module1"
Sub MatchAndCopyData()
    Call UniqueCompaniesToMaster
    Call ClearFilters
    Call FilterByWeek
    Call OutputWeekRange
    Call CopyVisibleIDsCompanyNamesAndAdditionalColumn
    Call MatchIDsAndOutputNameDatendOutputNameDate
    Call CompareDatesAndOutputStatus
    Call TallyOnTimeStatusForCompanies
    Call TallyCompanyOccurrences


    
    Dim wsPrintout As Worksheet
    Dim wsData As Worksheet
    Dim wsOutput As Worksheet
    Dim lookupValue As String
    Dim matchRow As Long
    Dim cell As Range
    Dim found As Boolean

    ' Set worksheets
    Set wsPrintout = ThisWorkbook.Sheets("Printout")
    Set wsData = ThisWorkbook.Sheets("Data")
    Set wsOutput = ThisWorkbook.Sheets("Output")

    ' Get the lookup value from cell B3 on the Printout sheet
    lookupValue = Trim(wsPrintout.Range("A3").value)
    


    ' Initialize found flag
    found = False

    ' Loop through each cell in column A to find a match
    For Each cell In wsData.Range("A1:A" & wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row)
        If Trim(cell.value) = lookupValue Then
            matchRow = cell.Row
            found = True
            Exit For
        End If
    Next cell

    ' Check if a match is found
    If found Then
        ' Copy the range from the Data sheet to the Output sheet
        wsOutput.Range("A2:F2").value = wsData.Range("A" & matchRow & ":F" & matchRow).value
    Else
        
    End If







Call ClearCharts
Call MoveChart
Call ColorN4BasedOnPercentageRange
Call selectcell


End Sub




Sub MoveChart()
    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim chartObject As chartObject

    ' Set references to the source and destination sheets
    Set wsSource = ThisWorkbook.Sheets("Speedometer")
    Set wsDestination = ThisWorkbook.Sheets("Printout")

    ' Check if there's at least one chart in the source sheet
    If wsSource.ChartObjects.Count > 0 Then
        ' Set reference to the first chart object on the source sheet
        Set chartObject = wsSource.ChartObjects(1)

        ' Copy the chart object
        chartObject.Copy

        ' Paste the chart object into the destination sheet at cell J5
        wsDestination.Paste Destination:=wsDestination.Range("J9")

        ' Optional: Delete the original chart from the source sheet
        
    Else
        MsgBox "No chart found on the Speedometer sheet."
    End If
End Sub

Sub ClearCharts()
    Dim ws As Worksheet
    Dim chartObject As chartObject

    ' Set reference to the "Printout" sheet
    Set ws = ThisWorkbook.Sheets("Printout")

    ' Loop through all chart objects and delete them
    For Each chartObject In ws.ChartObjects
        chartObject.Delete
    Next chartObject
End Sub

Sub ColorN4BasedOnPercentageRange()
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim targetCell2 As Range
    Dim range1 As Range
    Dim range2 As Range
    Dim range3 As Range
    Dim percentage As Double
    Dim range1Lower As Double
    Dim range2Lower As Double
    Dim range3Upper As Double
    Dim parts As Variant
    
    ' Set the worksheet and the relevant cells
    Set ws = ThisWorkbook.Sheets("Printout")
    Set targetCell = ws.Range("L4")
    Set targetCell2 = ws.Range("L3")
    Set range1 = ws.Range("L13")
    Set range2 = ws.Range("M13")
    Set range3 = ws.Range("N13")
    
    ' Get the percentage value from the target cell
    On Error Resume Next
    percentage = targetCell.value * 100 ' Assuming N4 is in percentage format (e.g., 0.95 for 95%)
    On Error GoTo 0
    
    If IsNumeric(percentage) Then
        ' Parse the range values safely
        On Error Resume Next
        parts = Split(range1.value, "-")
        If UBound(parts) = 1 Then
            range1Lower = CDbl(Replace(parts(1), "%", ""))
        End If
        
        parts = Split(range2.value, "-")
        If UBound(parts) = 1 Then
            range2Lower = CDbl(Replace(parts(1), "%", ""))
        End If
        
        range3Upper = CDbl(Replace(Replace(range3.value, "<", ""), "%", ""))
        On Error GoTo 0
        
        ' Determine which range the percentage falls into and color the target cells accordingly
        If percentage <= 100 And percentage >= range1Lower Then
            targetCell.Interior.Color = range1.Interior.Color
            targetCell2.Interior.Color = range1.Interior.Color
        ElseIf percentage < range1Lower And percentage >= range2Lower Then
            targetCell.Interior.Color = range2.Interior.Color
            targetCell2.Interior.Color = range2.Interior.Color
        ElseIf percentage < range2Lower Then
            targetCell.Interior.Color = range3.Interior.Color
            targetCell2.Interior.Color = range3.Interior.Color
        Else
            ' If percentage does not fall into any of the ranges, set a default color
            targetCell.Interior.Color = RGB(255, 255, 255) ' Default to white color or any color you prefer
            targetCell2.Interior.Color = RGB(255, 255, 255) ' Default to white color or any color you prefer
        End If
    Else
        ' If the value in N4 is not numeric, set a default color
        targetCell.Interior.Color = RGB(255, 255, 255) ' Default to white color or any color you prefer
        targetCell2.Interior.Color = RGB(255, 255, 255) ' Default to white color or any color you prefer
    End If
End Sub

Sub ClearFilters()
    Dim wsOnTime As Worksheet
    Dim wsAmount As Worksheet
    Dim ontimeTable As ListObject
    Dim amountTable As ListObject
    
    ' Set worksheets
    Set wsOnTime = ThisWorkbook.Sheets("datar")
    Set wsAmount = ThisWorkbook.Sheets("datap")
    
    ' Set the tables
    On Error Resume Next
    Set ontimeTable = wsOnTime.ListObjects("datar")
    Set amountTable = wsAmount.ListObjects("datap")
    On Error GoTo 0
    
    If ontimeTable Is Nothing Then
        MsgBox "Table 'ontime' not found on sheet 'OnTime'.", vbExclamation
        Exit Sub
    End If
    
    If amountTable Is Nothing Then
        MsgBox "Table 'amount' not found on sheet 'Amount'.", vbExclamation
        Exit Sub
    End If
    
    ' Clear filters in "ontime" table
    If ontimeTable.AutoFilter.FilterMode Then
        ontimeTable.AutoFilter.ShowAllData
    End If
    
    ' Clear filters in "amount" table
    If amountTable.AutoFilter.FilterMode Then
        amountTable.AutoFilter.ShowAllData
    End If
    
End Sub

Sub OutputWeekRange()
    Dim wsPrintout As Worksheet
    Dim dateToFilter As String
    Dim inputDate As Date
    Dim startDate As Date
    Dim endDate As Date
    Dim weekRange As String
    
    ' Set the worksheet
    Set wsPrintout = ThisWorkbook.Sheets("Printout")
    
    ' Get the value from cell A4 as a string
    dateToFilter = wsPrintout.Range("A4").value
    
    ' Check if the input is a valid date
    On Error GoTo InvalidDate
    inputDate = dateValue(dateToFilter)
    On Error GoTo 0
    
    ' Calculate the start and end dates of the week
    startDate = inputDate - Weekday(inputDate, vbSunday) + 1
    endDate = startDate + 6
    
    ' Format the week range as MM/DD/YYYY - MM/DD/YYYY
    weekRange = Format(startDate, "MM/DD/YYYY") & " - " & Format(endDate, "MM/DD/YYYY")
    
    ' Output the week range to cell A5
    wsPrintout.Range("A4").value = weekRange
    
    Exit Sub
    
InvalidDate:
    MsgBox "Please enter a valid date in MM/DD/YYYY format in cell A4.", vbExclamation
End Sub




