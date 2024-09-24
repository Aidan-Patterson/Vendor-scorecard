Attribute VB_Name = "Module6"
Sub OutputResponseData()
    Dim wsInput As Worksheet
    Dim wsResponse As Worksheet
    Dim otherWorkbook As Workbook
    Dim otherResponseSheet As Worksheet
    Dim nextRow As Long
    Dim nextRowOther As Long
    Dim companyName As String
    Dim dateValue As Variant
    Dim number1 As Double
    Dim number2 As Double
    Dim chkBox As checkbox
    Dim workbookName As String
    Dim otherWorkbookWasOpen As Boolean

    ' Set worksheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsResponse = ThisWorkbook.Sheets("Response Data")
    
   
    ' Retrieve values from the Input sheet
    companyName = wsInput.Range("B7").value
    dateValue = wsInput.Range("D7").value
    Set chkBox = wsInput.CheckBoxes("ocr") ' Assuming the checkbox name is "ocr"
    
    ' Check the status of the checkbox
    If chkBox.value = xlOn Then
        number1 = 1
    Else
        number1 = 0
    End If
    
    number2 = wsInput.Range("L15").value

    ' Find the next available row in Response Data sheet
    nextRow = wsResponse.Cells(wsResponse.Rows.Count, "A").End(xlUp).Row + 1

    ' Output the values in the Response Data sheet
    wsResponse.Cells(nextRow, 1).value = companyName
    wsResponse.Cells(nextRow, 2).value = dateValue
    wsResponse.Cells(nextRow, 3).value = number1
    wsResponse.Cells(nextRow, 4).value = number2


    ' Clean up
    Set wsInput = Nothing
    Set wsResponse = Nothing
    Set otherWorkbook = Nothing
    Set otherResponseSheet = Nothing
    
    ' Output success message
    MsgBox "Information submitted!", vbInformation
    Call UFillSequentialNumbersResponse
End Sub





Sub AverageCompanyValuesUsingArrays()
    Dim ws As Worksheet, wsOutput As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim companyName As String
    Dim uniqueCompanies As Collection
    Dim companyIndex As Object
    Dim companyArray() As String
    Dim sumC() As Double, sumD() As Double
    Dim countC() As Long, countD() As Long
    Dim outputRow As Long
    Dim visibleCells As Range
    Dim cell As Range

    ' Set the worksheets
    Set ws = ThisWorkbook.Sheets("Response Data")
    Set wsOutput = ThisWorkbook.Sheets("Response DataOutput")

    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Initialize the collection and dictionary
    Set uniqueCompanies = New Collection
    Set companyIndex = CreateObject("Scripting.Dictionary")

    ' Get the visible cells in the range
    On Error Resume Next
    Set visibleCells = ws.Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    ' Loop through each visible company in column A to identify unique companies
    If Not visibleCells Is Nothing Then
        For Each cell In visibleCells
            companyName = cell.value
            If companyName <> "" Then
                If Not companyIndex.Exists(companyName) Then
                    uniqueCompanies.Add companyName, companyName
                    companyIndex.Add companyName, uniqueCompanies.Count
                End If
            End If
        Next cell
    End If

    ' Initialize arrays based on the number of unique companies
    ReDim companyArray(1 To uniqueCompanies.Count)
    ReDim sumC(1 To uniqueCompanies.Count)
    ReDim sumD(1 To uniqueCompanies.Count)
    ReDim countC(1 To uniqueCompanies.Count)
    ReDim countD(1 To uniqueCompanies.Count)

    ' Populate companyArray with unique company names
    For i = 1 To uniqueCompanies.Count
        companyArray(i) = uniqueCompanies(i)
    Next i

    ' Loop through each visible company in column A again to sum values and count occurrences
    If Not visibleCells Is Nothing Then
        For Each cell In visibleCells
            companyName = cell.value
            If companyName <> "" Then
                j = companyIndex(companyName)
                sumC(j) = sumC(j) + cell.Offset(0, 2).value ' Column C
                countC(j) = countC(j) + 1
                sumD(j) = sumD(j) + cell.Offset(0, 3).value ' Column D
                countD(j) = countD(j) + 1
            End If
        Next cell
    End If

    ' Clear columns A, B, and C on the output sheet
    wsOutput.Range("A2:C" & wsOutput.Rows.Count).ClearContents

    ' Output the unique companies and their averages on the output sheet
    outputRow = 2
    For i = 1 To uniqueCompanies.Count
        wsOutput.Cells(outputRow, 1).value = companyArray(i)
        wsOutput.Cells(outputRow, 2).value = sumC(i) / countC(i) ' Average of column C
        wsOutput.Cells(outputRow, 3).value = sumD(i) / countD(i) ' Average of column D
        outputRow = outputRow + 1
    Next i

    ' Clean up
    Set uniqueCompanies = Nothing
    Set companyIndex = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing
End Sub

