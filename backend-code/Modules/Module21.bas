Attribute VB_Name = "Module21"
Sub AFilterAndOutputDataQUARTER()
    Dim wsPrintout As Worksheet
    Dim wsDatar As Worksheet
    Dim wsPOData As Worksheet
    Dim datarTable As ListObject
    Dim quarter As String
    Dim startDate As Date
    Dim endDate As Date
    Dim currentYear As Integer
    
    ' Set references to sheets and table
    Set wsPrintout = Worksheets("Printout")
    Set wsDatar = Worksheets("datar")
    Set wsPOData = Worksheets("PO Data")
    Set datarTable = wsDatar.ListObjects("datar")

    ' Get the quarter from cell A5 on the Printout sheet
    quarter = wsPrintout.Range("A5").value
    
    ' Get the current year
    currentYear = Year(Date)
    
    ' Determine the date range based on the quarter and current year
    Select Case quarter
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
    End Select

    ' Apply the filter to the "datar" table based on the date range
    datarTable.Range.AutoFilter Field:=4, Criteria1:=">=" & startDate, Criteria2:="<=" & endDate

    ' Clear existing data in PO Data
    wsPOData.Range("A2:B" & wsPOData.Cells(wsPOData.Rows.Count, "A").End(xlUp).Row).ClearContents

    ' Copy filtered values from column E to PO Data sheet column A
    If datarTable.Range.SpecialCells(xlCellTypeVisible).Count > 1 Then
        datarTable.ListColumns(5).DataBodyRange.SpecialCells(xlCellTypeVisible).Copy
        wsPOData.Range("A2").PasteSpecial Paste:=xlPasteValues
    End If

    ' Copy filtered values from column J to PO Data sheet column B
    If datarTable.Range.SpecialCells(xlCellTypeVisible).Count > 1 Then
        datarTable.ListColumns(10).DataBodyRange.SpecialCells(xlCellTypeVisible).Copy
        wsPOData.Range("B2").PasteSpecial Paste:=xlPasteValues
    End If

    ' Clear the filter
    datarTable.AutoFilter.ShowAllData

    ' Optional: Clear clipboard
    Application.CutCopyMode = False
End Sub

Sub AFilterByMonthAndOutputDataMONTH()
    Dim wsPrintout As Worksheet
    Dim wsDatar As Worksheet
    Dim wsPOData As Worksheet
    Dim datarTable As ListObject
    Dim monthName As String
    Dim startDate As Date
    Dim endDate As Date
    Dim currentYear As Integer
    Dim monthNumber As Integer
    
    ' Set references to sheets and table
    Set wsPrintout = Worksheets("Printout")
    Set wsDatar = Worksheets("datar")
    Set wsPOData = Worksheets("PO Data")
    Set datarTable = wsDatar.ListObjects("datar")

    ' Get the month from cell A4 on the Printout sheet
    monthName = wsPrintout.Range("A4").value
    
    ' Get the current year
    currentYear = Year(Date)
    
    ' Determine the start and end date for the given month
    monthNumber = Month(dateValue("01 " & monthName & " " & currentYear))
    startDate = DateSerial(currentYear, monthNumber, 1)
    endDate = DateSerial(currentYear, monthNumber + 1, 0) ' Last day of the month

    ' Apply the filter to the "datar" table based on the date range
    datarTable.Range.AutoFilter Field:=4, Criteria1:=">=" & startDate, Criteria2:="<=" & endDate

    ' Clear existing data in PO Data
    wsPOData.Range("A2:B" & wsPOData.Cells(wsPOData.Rows.Count, "A").End(xlUp).Row).ClearContents

    ' Copy filtered values from column E to PO Data sheet column A
    If datarTable.Range.SpecialCells(xlCellTypeVisible).Count > 1 Then
        datarTable.ListColumns(5).DataBodyRange.SpecialCells(xlCellTypeVisible).Copy
        wsPOData.Range("A2").PasteSpecial Paste:=xlPasteValues
    End If

    ' Copy filtered values from column J to PO Data sheet column B
    If datarTable.Range.SpecialCells(xlCellTypeVisible).Count > 1 Then
        datarTable.ListColumns(10).DataBodyRange.SpecialCells(xlCellTypeVisible).Copy
        wsPOData.Range("B2").PasteSpecial Paste:=xlPasteValues
    End If

    ' Clear the filter
    datarTable.AutoFilter.ShowAllData

    ' Optional: Clear clipboard
    Application.CutCopyMode = False
End Sub

Sub AAProcessCompanyData()
    Dim wsPOData As Worksheet
    Dim wsPODataOutput As Worksheet
    Dim companyData As Variant
    Dim outputData As Variant
    Dim companyDict As Object
    Dim lastRow As Long
    Dim i As Long
    Dim companyName As Variant
    Dim statusValue As String
    Dim totalCount As Long
    Dim earlyOnTimeCount As Long
    Dim outputRow As Long
    
    ' Set references to sheets
    Set wsPOData = Worksheets("PO Data")
    Set wsPODataOutput = Worksheets("PO DataOutput")
    
    ' Find the last row in column A of PO Data
    lastRow = wsPOData.Cells(wsPOData.Rows.Count, "A").End(xlUp).Row
    
    ' Read the data from PO Data into an array
    companyData = wsPOData.Range("A2:B" & lastRow).value
    
    ' Create a dictionary to hold unique company names and counts
    Set companyDict = CreateObject("Scripting.Dictionary")
    
    ' Loop through the array to populate the dictionary
    For i = 1 To UBound(companyData, 1)
        companyName = companyData(i, 1)
        statusValue = companyData(i, 2)
        
        ' If the company is not in the dictionary, add it
        If Not companyDict.Exists(companyName) Then
            companyDict.Add companyName, Array(0, 0) ' (earlyOnTimeCount, totalCount)
        End If
        
        ' Update the counts
        totalCount = companyDict(companyName)(1) + 1
        earlyOnTimeCount = companyDict(companyName)(0)
        If statusValue = "Early" Or statusValue = "On Time" Then
            earlyOnTimeCount = earlyOnTimeCount + 1
        End If
        
        ' Update the dictionary
        companyDict(companyName) = Array(earlyOnTimeCount, totalCount)
    Next i
    
    ' Prepare the output array
    outputData = wsPODataOutput.Range("A2:C" & wsPODataOutput.Cells(wsPODataOutput.Rows.Count, "A").End(xlUp).Row).value
    
    ' Output the results to PO DataOutput
    outputRow = 2
    ReDim outputData(1 To companyDict.Count, 1 To 3)
    i = 1
    For Each companyName In companyDict.Keys
        outputData(i, 1) = companyName
        outputData(i, 2) = companyDict(companyName)(0) ' Early/On-Time count
        outputData(i, 3) = companyDict(companyName)(1) ' Total count
        i = i + 1
    Next companyName
    
    ' Write the output array to the sheet
    wsPODataOutput.Range("A2").Resize(UBound(outputData, 1), UBound(outputData, 2)).value = outputData
End Sub

Sub ShowVendorInputForm()
    ' Display a message box before showing the UserForm


    ' Show the UserForm
    UserForm5.Show
    
    Call CalculateAndOutput
    Call EDITReworkData
End Sub

Sub EDITReworkData()
    Dim wsInput As Worksheet
    Dim wsReworkData As Worksheet
    Dim wsTestReworkData As Worksheet
    Dim numerator As Long
    Dim reworkRow As Long
    Dim testReworkRow As Long
    Dim fraction As String
    Dim J27Value As Variant
    Dim K27Value As Variant
    Dim L27Value As Variant
    Dim testWorkbook As Workbook

    ' Set references to sheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsReworkData = ThisWorkbook.Sheets("Rework Data")
    

    
    ' Extract the fraction from cell B22 on the "Input" sheet
    fraction = wsInput.Range("B22").value
    
    ' Extract the numerator (value before the "/")
    numerator = Val(Split(Split(fraction, " ")(2), "/")(0))
    
    ' Get the values from cells J27, K27, and L27 on the "Input" sheet
    J27Value = wsInput.Range("J27").value
    K27Value = wsInput.Range("K27").value
    L27Value = wsInput.Range("L27").value
    
    ' Find the row in "Rework Data" where the value in column D matches the numerator
    On Error Resume Next
    reworkRow = wsReworkData.Columns("D").Find(What:=numerator, LookIn:=xlValues, LookAt:=xlWhole).Row
    testReworkRow = wsTestReworkData.Columns("D").Find(What:=numerator, LookIn:=xlValues, LookAt:=xlWhole).Row
    On Error GoTo 0
    
    ' If a matching row is found in "Rework Data" sheet in the current workbook, update the corresponding cells
    If reworkRow > 0 Then
        wsReworkData.Cells(reworkRow, 5).value = J27Value ' Output to column E
        wsReworkData.Cells(reworkRow, 6).value = K27Value ' Output to column F
        wsReworkData.Cells(reworkRow, 3).value = L27Value ' Output to column C
    Else
        MsgBox "No matching row found in 'Rework Data' for the numerator value: " & numerator, vbExclamation
    End If
    

End Sub


Sub CalculateAndOutput()
    Dim wsInput As Worksheet
    Dim value1 As Double
    Dim value2 As Double
    Dim result As Double
    
    ' Set reference to the "Input" sheet
    Set wsInput = ThisWorkbook.Sheets("Input")
    
    ' Get the values from cells J27 and K27
    value1 = wsInput.Range("J27").value
    value2 = wsInput.Range("K27").value
    
    ' Calculate the result: (J27 * K27) * 108
    result = value1 * value2 * 108
    
    ' Output the result in cell L27
    wsInput.Range("L27").value = result
End Sub

