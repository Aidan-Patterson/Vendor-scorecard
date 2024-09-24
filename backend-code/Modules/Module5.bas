Attribute VB_Name = "Module5"
Sub EnterReworkData()
    Dim inputSheet As Worksheet
    Dim reworkSheet As Worksheet
    Dim otherWorkbook As Workbook
    Dim otherReworkSheet As Worksheet
    Dim companyName As String
    Dim entryDate As Variant ' Use Variant to handle potential type conversions
    Dim cost As Double
    Dim additionalValue1 As Variant
    Dim additionalValue2 As Variant
    Dim nextRow As Long
    Dim nextRowOther As Long
    Dim workbookName As String
    Dim otherWorkbookWasOpen As Boolean

    ' Set input and output sheets
    Set inputSheet = ThisWorkbook.Sheets("Input")
    Set reworkSheet = ThisWorkbook.Sheets("Rework Data")
    
   
    ' Read input values
    companyName = inputSheet.Range("B7").value
    
    ' Handle date conversion
    If IsDate(inputSheet.Range("D7").value) Then
        entryDate = CDate(inputSheet.Range("D7").value) ' Convert to Date if it's a valid date format
    Else
        entryDate = "" ' Handle case where date is not valid
    End If
    
    ' Read cost as double
    If IsNumeric(inputSheet.Range("L8").value) Then
        cost = CDbl(inputSheet.Range("L8").value) ' Convert to Double if it's numeric
    Else
        cost = 0 ' Default to 0 if not numeric
    End If

    ' Read additional values from J8 and K8
    additionalValue1 = inputSheet.Range("J8").value
    additionalValue2 = inputSheet.Range("K8").value
    
    ' Find the next available row in column A of "Rework Data" sheet starting from row 2
    nextRow = reworkSheet.Cells(reworkSheet.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Output values in "Rework Data" sheet
    reworkSheet.Cells(nextRow, 1).value = companyName ' Company name in column A
    reworkSheet.Cells(nextRow, 2).value = entryDate   ' Date in column B
    reworkSheet.Cells(nextRow, 3).value = cost        ' Cost in column C
    reworkSheet.Cells(nextRow, 5).value = additionalValue1 ' Value from J8 in column E
    reworkSheet.Cells(nextRow, 6).value = additionalValue2 ' Value from K8 in column F
    
    
    ' Clean up
    Set inputSheet = Nothing
    Set reworkSheet = Nothing
    Set otherWorkbook = Nothing
    Set otherReworkSheet = Nothing
    Call UFillSequentialNumbersRework
End Sub



Sub SumValuesByCompany1()
    Dim ws As Worksheet, wsOutput As Worksheet
    Dim lastRow As Long
    Dim companyList As Object
    Dim company As Range
    Dim companyName As String
    Dim totalC As Double
    Dim visibleCells As Range
    Dim cell As Range

    ' Set the worksheets
    Set ws = ThisWorkbook.Sheets("Rework Data") ' Replace "Rework Data" with your sheet name
    Set wsOutput = ThisWorkbook.Sheets("Rework DataOutput") ' Replace "Rework DataOutput" with your output sheet name

    ' Initialize a dictionary to store unique company names
    Set companyList = CreateObject("Scripting.Dictionary")

    ' Determine the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Get the visible cells in column A
    On Error Resume Next ' In case there are no visible cells
    Set visibleCells = ws.Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    ' Loop through visible cells in column A to get unique company names
    If Not visibleCells Is Nothing Then
        For Each company In visibleCells
            companyName = company.value
            If Not companyList.Exists(companyName) Then
                companyList.Add companyName, 0 ' Initialize company with total 0
            End If
        Next company
    End If

    ' Clear the output sheet starting from A2
    wsOutput.Range("A2:B" & wsOutput.Rows.Count).ClearContents

    ' Output unique company names in column A of the output sheet
    wsOutput.Range("A2").Resize(companyList.Count).value = WorksheetFunction.Transpose(companyList.Keys)

    ' Calculate totals for column C and output in column B of the output sheet
    If Not visibleCells Is Nothing Then
        For Each company In wsOutput.Range("A2:A" & companyList.Count + 1) ' Iterate through unique company names in column A of the output sheet
            totalC = 0

            ' Loop through visible cells in column A to find matching company names and sum values from column C
            For Each cell In visibleCells
                If cell.value = company.value Then
                    If IsNumeric(cell.Offset(0, 2).value) And cell.Offset(0, 2).value <> "" Then ' Check if cell in column C is not blank and is numeric
                        totalC = totalC + cell.Offset(0, 2).value ' Column C
                    End If
                End If
            Next cell

            ' Output totals in column B next to the corresponding company name in column A
            company.Offset(0, 1).value = totalC ' Column B
        Next company
    End If

    ' Clean up
    Set companyList = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing
End Sub






Sub totalCost()
    Dim wsPO As Worksheet, wsData As Worksheet, wsReworkOutput As Worksheet
    Dim poData As Variant, dataData As Variant
    Dim dict As Object
    Dim i As Long, lastRowReworkOutput As Long

    ' Set worksheets
    Set wsPO = ThisWorkbook.Sheets("PO Data")
    Set wsData = ThisWorkbook.Sheets("datap")
    Set wsReworkOutput = ThisWorkbook.Sheets("Rework DataOutput")

    ' Clear columns F and G in Rework DataOutput sheet
    wsReworkOutput.Range("F:G").ClearContents

    ' Find the last row with data in PO Data and datap sheets
    lastRowPO = wsPO.Cells(wsPO.Rows.Count, "A").End(xlUp).Row
    lastRowData = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    ' Load data into arrays
    poData = wsPO.Range("A2:A" & lastRowPO).value
    dataData = wsData.Range("A2:F" & lastRowData).value

    ' Create a dictionary for faster lookup
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(dataData, 1)
        If Not dict.Exists(dataData(i, 1)) Then
            dict(dataData(i, 1)) = Array(dataData(i, 2), dataData(i, 6)) ' Vendor name and Cost
        End If
    Next i

    ' Loop through each ID in the PO Data array
    For i = 1 To UBound(poData, 1)
        If Not IsEmpty(poData(i, 1)) Then
            If dict.Exists(poData(i, 1)) Then
                lastRowReworkOutput = wsReworkOutput.Cells(wsReworkOutput.Rows.Count, "F").End(xlUp).Row + 1
                wsReworkOutput.Cells(lastRowReworkOutput, "F").value = dict(poData(i, 1))(0) ' Vendor name
                wsReworkOutput.Cells(lastRowReworkOutput, "G").value = dict(poData(i, 1))(1) ' Cost
            End If
        End If
    Next i

    ' Autofit columns F and G in Rework DataOutput sheet
    wsReworkOutput.Columns("F:G").AutoFit

    ' Clean up
    Set dict = Nothing
    Set wsPO = Nothing
    Set wsData = Nothing
    Set wsReworkOutput = Nothing
End Sub

Sub FilterAndExtractDataMONTH()
    Dim wsPrintout As Worksheet
    Dim wsDatap As Worksheet
    Dim wsOutput As Worksheet
    Dim filterMonth As String
    Dim currentYear As String
    Dim lastRow As Long
    Dim cell As Range
    Dim outputRow As Long
    
    ' Set the worksheets
    Set wsPrintout = ThisWorkbook.Sheets("Printout")
    Set wsDatap = ThisWorkbook.Sheets("datap")
    Set wsOutput = ThisWorkbook.Sheets("Rework DataOutput")
    
    ' Get the filter month and current year
    filterMonth = wsPrintout.Range("A4").value
    currentYear = Year(Date)
    
    ' Convert filterMonth to a number
    Dim monthNumber As Integer
    monthNumber = Month(dateValue("1 " & filterMonth & " " & currentYear))
    
    ' Apply the filter to column E in the "datap" sheet
    With wsDatap.ListObjects("datap")
        ' Remove any existing filters
        If .AutoFilter.FilterMode Then .AutoFilter.ShowAllData
        ' Apply the filter
        .Range.AutoFilter Field:=5, Criteria1:=">=" & DateSerial(currentYear, monthNumber, 1), _
                                 Operator:=xlAnd, Criteria2:="<" & DateSerial(currentYear, monthNumber + 1, 1)
    End With
    
    ' Initialize the output row
    outputRow = 2
    
    ' Loop through the filtered data and copy to "ReworkDataOutput" sheet
    ' Only consider visible cells
    On Error Resume Next
    For Each cell In wsDatap.ListObjects("datap").ListColumns(2).DataBodyRange.SpecialCells(xlCellTypeVisible)
        If Not cell.EntireRow.Hidden Then
            wsOutput.Cells(outputRow, 6).value = cell.value
            wsOutput.Cells(outputRow, 7).value = cell.Offset(0, 4).value
            outputRow = outputRow + 1
        End If
    Next cell
    On Error GoTo 0
    
    ' Remove the filter from the "datap" sheet
    With wsDatap.ListObjects("datap")
        If .AutoFilter.FilterMode Then .AutoFilter.ShowAllData
    End With
End Sub
Sub FilterAndExtractDataByQUARTER()
    Dim wsPrintout As Worksheet
    Dim wsDatap As Worksheet
    Dim wsOutput As Worksheet
    Dim filterQuarter As String
    Dim currentYear As Integer
    Dim startDate As Date
    Dim endDate As Date
    Dim lastRow As Long
    Dim cell As Range
    Dim outputRow As Long
    
    ' Set the worksheets
    Set wsPrintout = ThisWorkbook.Sheets("Printout")
    Set wsDatap = ThisWorkbook.Sheets("datap")
    Set wsOutput = ThisWorkbook.Sheets("Rework DataOutput")
    
    ' Get the filter quarter and current year
    filterQuarter = wsPrintout.Range("A5").value
    currentYear = Year(Date)
    
    ' Determine the start and end dates based on the quarter
    Select Case filterQuarter
        Case "Quarter 1"
            startDate = DateSerial(currentYear, 1, 1)
            endDate = DateSerial(currentYear, 4, 1)
        Case "Quarter 2"
            startDate = DateSerial(currentYear, 4, 1)
            endDate = DateSerial(currentYear, 7, 1)
        Case "Quarter 3"
            startDate = DateSerial(currentYear, 7, 1)
            endDate = DateSerial(currentYear, 10, 1)
        Case "Quarter 4"
            startDate = DateSerial(currentYear, 10, 1)
            endDate = DateSerial(currentYear + 1, 1, 1)
        Case Else
            MsgBox "Invalid quarter format. Please use 'Quarter 1', 'Quarter 2', 'Quarter 3', or 'Quarter 4'."
            Exit Sub
    End Select
    
    ' Apply the filter to column E in the "datap" sheet
    With wsDatap.ListObjects("datap")
        ' Remove any existing filters
        If .AutoFilter.FilterMode Then .AutoFilter.ShowAllData
        ' Apply the filter
        .Range.AutoFilter Field:=5, Criteria1:=">=" & startDate, _
                                 Operator:=xlAnd, Criteria2:="<" & endDate
    End With
    
    ' Initialize the output row
    outputRow = 2
    
    ' Loop through the filtered data and copy to "ReworkDataOutput" sheet
    ' Only consider visible cells
    On Error Resume Next
    For Each cell In wsDatap.ListObjects("datap").ListColumns(2).DataBodyRange.SpecialCells(xlCellTypeVisible)
        If Not cell.EntireRow.Hidden Then
            wsOutput.Cells(outputRow, 6).value = cell.value
            wsOutput.Cells(outputRow, 7).value = cell.Offset(0, 4).value
            outputRow = outputRow + 1
        End If
    Next cell
    On Error GoTo 0
    
    ' Remove the filter from the "datap" sheet
    With wsDatap.ListObjects("datap")
        If .AutoFilter.FilterMode Then .AutoFilter.ShowAllData
    End With
End Sub

Sub SumMatchedCompanies()
    Dim ws As Worksheet
    Dim companyDict As Object
    Dim lastRowF As Long
    Dim lastRowA As Long
    Dim i As Long
    Dim companyName As String
    Dim companyValue As Double
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Rework DataOutput")
    
    ' Create a dictionary to store the sum of values for each company
    Set companyDict = CreateObject("Scripting.Dictionary")
    
    ' Find the last rows of columns F and A
    lastRowF = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through the companies in column F and sum their corresponding values in column G
    For i = 2 To lastRowF
        companyName = ws.Cells(i, "F").value
        companyValue = ws.Cells(i, "G").value
        If companyDict.Exists(companyName) Then
            companyDict(companyName) = companyDict(companyName) + companyValue
        Else
            companyDict.Add companyName, companyValue
        End If
    Next i
    
    ' Output the summed values in column C next to their corresponding company in column A
    For i = 2 To lastRowA
        companyName = ws.Cells(i, "A").value
        If companyDict.Exists(companyName) Then
            ws.Cells(i, "C").value = companyDict(companyName)
        End If
    Next i
    
    ' Clean up
    Set companyDict = Nothing
End Sub








Sub SumCompanyValues()
    Dim ws As Worksheet, wsOutput As Worksheet
    Dim lastRowE As Long, lastRowJ As Long, lastRowOutput As Long
    Dim companyDict As Object, costDict As Object
    Dim i As Long, outputRow As Long

    ' Set the worksheets
    Set ws = ThisWorkbook.Sheets("Rework Data") ' Adjust the sheet name as needed
    Set wsOutput = ThisWorkbook.Sheets("Rework DataOutput") ' Adjust the output sheet name as needed

    ' Find the last rows in columns E and J
    lastRowE = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    lastRowJ = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row

    ' Initialize the dictionaries
    Set companyDict = CreateObject("Scripting.Dictionary")
    Set costDict = CreateObject("Scripting.Dictionary")

    ' Loop through each company in column J and sum the values in column K
    For i = 2 To lastRowJ
        If Not IsEmpty(ws.Cells(i, "J").value) Then
            If Not companyDict.Exists(ws.Cells(i, "J").value) Then
                companyDict(ws.Cells(i, "J").value) = ws.Cells(i, "K").value
            Else
                companyDict(ws.Cells(i, "J").value) = companyDict(ws.Cells(i, "J").value) + ws.Cells(i, "K").value
            End If
        End If
    Next i

    ' Loop through each company in column E and store the costs in the cost dictionary
    For i = 2 To lastRowE
        If Not IsEmpty(ws.Cells(i, "E").value) Then
            costDict(ws.Cells(i, "E").value) = ws.Cells(i, "F").value
        End If
    Next i

    ' Clear the output sheet starting from A2
    wsOutput.Range("A2:C" & wsOutput.Rows.Count).ClearContents

    ' Set the starting row for output
    outputRow = 2

    ' Loop through each company in the output sheet and output the total cost and summed value
    lastRowOutput = wsOutput.Cells(wsOutput.Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRowOutput
        If Not IsEmpty(wsOutput.Cells(i, "A").value) Then
            If costDict.Exists(wsOutput.Cells(i, "A").value) Then
                wsOutput.Cells(i, 2).value = costDict(wsOutput.Cells(i, "A").value)
            End If
            If companyDict.Exists(wsOutput.Cells(i, "A").value) Then
                wsOutput.Cells(i, 3).value = companyDict(wsOutput.Cells(i, "A").value)
            End If
        End If
    Next i

End Sub

Sub SumCompanyCostsTotal()
    Dim ws As Worksheet, wsOutput As Worksheet
    Dim lastRowJ As Long, lastRowOutput As Long
    Dim companyDict As Object
    Dim i As Long
    Dim companyName As String
    Dim totalCost As Double

    ' Set the worksheets
    Set ws = ThisWorkbook.Sheets("Rework Data") ' Adjust the sheet name as needed
    Set wsOutput = ThisWorkbook.Sheets("Rework DataOutput") ' Adjust the output sheet name as needed

    ' Find the last row in columns J and A
    lastRowJ = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
    lastRowOutput = wsOutput.Cells(wsOutput.Rows.Count, "A").End(xlUp).Row

    ' Initialize the dictionary to store total costs by company
    Set companyDict = CreateObject("Scripting.Dictionary")

    ' Loop through each company in column J and sum the values in column K
    For i = 2 To lastRowJ
        companyName = ws.Cells(i, "J").value
        If Not IsEmpty(companyName) Then
            If Not companyDict.Exists(companyName) Then
                companyDict.Add companyName, ws.Cells(i, "K").value
            Else
                companyDict(companyName) = companyDict(companyName) + ws.Cells(i, "K").value
            End If
        End If
    Next i

    ' Clear the contents of column C on the output sheet
    wsOutput.Range("C2:C" & lastRowOutput).ClearContents

    ' Loop through each company in column A of the output sheet and output the total cost in column C
    For i = 2 To lastRowOutput
        companyName = wsOutput.Cells(i, "A").value
        If companyDict.Exists(companyName) Then
            wsOutput.Cells(i, "C").value = companyDict(companyName)
        End If
    Next i

    ' Clean up
    Set companyDict = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing
End Sub

