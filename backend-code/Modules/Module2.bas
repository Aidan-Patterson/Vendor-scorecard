Attribute VB_Name = "Module2"
Sub selectcell()
Attribute selectcell.VB_ProcData.VB_Invoke_Func = " \n14"
'
' selectcell Macro
'

'
    Range("O4:P5").Select
End Sub

Sub FilterByWeek()
    Dim wsPrintout As Worksheet
    Dim wsOnTime As Worksheet
    Dim wsAmount As Worksheet
    Dim dateToFilter As Date
    Dim startDate As Date
    Dim endDate As Date
    Dim ontimeTable As ListObject
    Dim amountTable As ListObject
    
    ' Set worksheets
    Set wsPrintout = ThisWorkbook.Sheets("Printout")
    Set wsOnTime = ThisWorkbook.Sheets("datar")
    Set wsAmount = ThisWorkbook.Sheets("datap")
    
    ' Get the date from cell A4
    dateToFilter = wsPrintout.Range("A4").value
    
    ' Calculate the start and end dates of the week
    startDate = dateToFilter - Weekday(dateToFilter, vbSunday) + 1
    endDate = startDate + 6
    
    ' Set the tables
    Set ontimeTable = wsOnTime.ListObjects("datar")
    Set amountTable = wsAmount.ListObjects("datap")
    
    ' Apply filter to "ontime" table in column C
    With ontimeTable.Range
        .AutoFilter Field:=3, Criteria1:=">=" & startDate, Operator:=xlAnd, Criteria2:="<=" & endDate
    End With
    
    ' Apply filter to "amount" table in column E
    With amountTable.Range
        .AutoFilter Field:=5, Criteria1:=">=" & startDate, Operator:=xlAnd, Criteria2:="<=" & endDate
    End With
    
End Sub

Sub MatchAndCopyData2()
    Call UniqueCompaniesToMaster
    Call ClearFilters
    Call FilterByQuarter
    
    
    
    Call CopyVisibleIDsCompanyNamesAndAdditionalColumn
    Call MatchIDsAndOutputNameDate
    Call CompareDatesAndOutputStatus
    Call TallyOnTimeStatusForCompanies1
    
    Call TallyCompanyOccurrences1
    Call SumValuesByCompany
    Call UpdateMasterSheet
    
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
        MsgBox "No match found.", vbExclamation
    End If







Call ClearCharts
Call MoveChart
Call ColorN4BasedOnPercentageRange
Call selectcell
End Sub

Sub FilterByQuarter()
    Dim wsPrintout As Worksheet
    Dim wsOnTime As Worksheet
    Dim wsAmount As Worksheet
    Dim quarterToFilter As String
    Dim startDate As Date
    Dim endDate As Date
    Dim currentYear As Integer
    Dim ontimeTable As ListObject
    Dim amountTable As ListObject
    
    ' Set worksheets
    Set wsPrintout = ThisWorkbook.Sheets("Printout")
    Set wsOnTime = ThisWorkbook.Sheets("datar")
    Set wsAmount = ThisWorkbook.Sheets("datap")
    
    ' Get the quarter from cell A5
    quarterToFilter = wsPrintout.Range("A5").value
    
    ' Get the current year
    currentYear = Year(Date)
    
    ' Determine the start and end dates based on the quarter
    Select Case quarterToFilter
        Case "Quarter 1"
            startDate = DateSerial(currentYear, 1, 1)
            endDate = DateSerial(currentYear, 3, 31)
        Case "Quarter 2"
            startDate = DateSerial(currentYear, 4, 1)
            endDate = DateSerial(currentYear, 6, 30)
        Case "Quarter 3"
            startDate = DateSerial(currentYear, 7, 1)
            endDate = DateSerial(currentYear, 9, 30)
        Case "Quarter 4"
            startDate = DateSerial(currentYear, 10, 1)
            endDate = DateSerial(currentYear, 12, 31)
        Case Else
            MsgBox "Invalid quarter specified. Please enter Quarter 1, Quarter 2, Quarter 3, or Quarter 4.", vbExclamation
            Exit Sub
    End Select
    
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
    
    ' Apply filter to "ontime" table in column C
    ontimeTable.Range.AutoFilter Field:=3, Criteria1:=">=" & startDate, Operator:=xlAnd, Criteria2:="<=" & endDate
    
    ' Apply filter to "amount" table in column E
    amountTable.Range.AutoFilter Field:=5, Criteria1:=">=" & startDate, Operator:=xlAnd, Criteria2:="<=" & endDate
    
End Sub

Sub RefreshQueries()
    Dim wsOnTime As Worksheet
    Dim wsAmount As Worksheet
    Dim qt As QueryTable
    Dim lo As ListObject
    
    ' Set worksheets
    Set wsOnTime = ThisWorkbook.Sheets("datar")
    Set wsAmount = ThisWorkbook.Sheets("datap")
    
    ' Refresh queries on "OnTime" sheet
    For Each qt In wsOnTime.QueryTables
        qt.Refresh
    Next qt
    
    ' Refresh queries on "Amount" sheet (assuming they are ListObjects)
    For Each lo In wsAmount.ListObjects
        lo.Refresh
    Next lo
    
    MsgBox "Queries refreshed successfully.", vbInformation
End Sub

Sub SetTextValues()
    ' Set values in cells A3, A4, and A5
    Set ws = ThisWorkbook.Sheets("Printout")
    Range("A3").value = "Pick the vendor"
    Range("A4").value = "Pick Date if applicable"
    Range("A5").value = "Choose a quarter"
End Sub

Sub GoToPrintoutSheet()
    Dim ws As Worksheet
    
    ' Check if "Printout" sheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Printout")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet 'Printout' not found.", vbExclamation
    Else
        ' Activate the "Printout" sheet
        ws.Activate
    End If
End Sub


Sub SetTextValues2()
    Dim ws As Worksheet
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Input")  ' Replace "Sheet1" with your sheet name
    
    ' Set B7 to "Click to pick the vendor"
    ws.Range("B7").value = "Click to pick the vendor"
    
    ' Set D7 to "Click to add date"
    ws.Range("D7").value = "Click to add date"
    
    
End Sub


Sub ClearAndUncheck()
    Dim ws As Worksheet
    Dim chkN As checkbox
    Dim chkO As checkbox
    Dim chkR As checkbox
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Input")
    
    ' Uncheck checkboxes
    Set chkN = ws.CheckBoxes("ncheck")
    Set chkO = ws.CheckBoxes("ocheck")
    Set chkR = ws.CheckBoxes("ocr")
    
    If Not chkN Is Nothing Then
        chkN.value = xlOff
    End If
    
    If Not chkO Is Nothing Then
        chkO.value = xlOff
    End If
    
    If Not chkR Is Nothing Then
        chkR.value = xlOff
    End If
    
    ' Clear specific cells while keeping formatting
    With ws
        .Range("J8").ClearContents
        .Range("K8").ClearContents
        .Range("L11").ClearContents
        .Range("L15").ClearContents
    End With

    ' Call the SetTextValues2 subroutine
    Call SetTextValues2
End Sub

