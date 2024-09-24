Attribute VB_Name = "Module8"



Sub CopyPODataToMasterSheet()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Set the worksheets
    Set wsSource = ThisWorkbook.Sheets("Po DataOutput") ' Adjust the source sheet name as needed
    Set wsDest = ThisWorkbook.Sheets("Master Sheet") ' Adjust the destination sheet name as needed

    ' Find the last row with data in column A on the source sheet
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    ' Clear the destination sheet columns A, B, and C before inputting anything
    wsDest.Range("A2:C" & wsDest.Rows.Count).ClearContents

    ' Add headers to the destination sheet
    wsDest.Cells(1, 1).value = "Vendor"
    wsDest.Cells(1, 2).value = "On-Time POs"
    wsDest.Cells(1, 3).value = "Total POs"

    ' Copy the data from the source sheet to the destination sheet starting in row 2
    For i = 1 To lastRow
        wsDest.Cells(i + 1, 1).value = wsSource.Cells(i, 1).value
        wsDest.Cells(i + 1, 2).value = wsSource.Cells(i, 2).value
        wsDest.Cells(i + 1, 3).value = wsSource.Cells(i, 3).value
    Next i

    ' Autofit the columns in the destination sheet
    wsDest.Columns("A:C").AutoFit

    ' Clean up
    Set wsSource = Nothing
    Set wsDest = Nothing
End Sub


Sub CopyNCRDataToMasterSheet()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim lastRowSource As Long, lastRowDest As Long
    Dim companyName As String
    Dim i As Long, j As Long
    Dim companyDict As Object

    ' Set the worksheets
    Set wsSource = ThisWorkbook.Sheets("NCR DataOutput") ' Adjust the source sheet name as needed
    Set wsDest = ThisWorkbook.Sheets("Master Sheet") ' Adjust the destination sheet name as needed

    ' Find the last row with data in column A on both sheets
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row

    ' Initialize the dictionary to store values by company
    Set companyDict = CreateObject("Scripting.Dictionary")

    ' Loop through each company in column A on the source sheet and store values in the dictionary
    For i = 2 To lastRowSource
        companyName = wsSource.Cells(i, 1).value
        If Not IsEmpty(companyName) Then
            If Not companyDict.Exists(companyName) Then
                companyDict.Add companyName, Array(wsSource.Cells(i, 2).value, wsSource.Cells(i, 3).value)
            End If
        End If
    Next i

    ' Clear columns D and E on the destination sheet before inputting anything
    wsDest.Range("D2:E" & wsDest.Rows.Count).ClearContents

    ' Add headers to columns D and E on the destination sheet
    wsDest.Cells(1, 4).value = "Total NCRs"
    wsDest.Cells(1, 5).value = "Total Occurrences"

    ' Loop through each company in column A on the destination sheet and output the corresponding values
    For j = 2 To lastRowDest
        companyName = wsDest.Cells(j, 1).value
        If companyDict.Exists(companyName) Then
            wsDest.Cells(j, 4).value = companyDict(companyName)(0) ' Total NCRs
            wsDest.Cells(j, 5).value = companyDict(companyName)(1) ' Total Occurrences
        End If
    Next j

    ' Autofit columns D and E on the destination sheet
    wsDest.Columns("D:E").AutoFit

    ' Clean up
    Set companyDict = Nothing
    Set wsSource = Nothing
    Set wsDest = Nothing
End Sub


Sub CopyReworkDataToMasterSheet()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim lastRowSource As Long, lastRowDest As Long
    Dim companyName As String
    Dim i As Long, j As Long
    Dim companyDict As Object

    ' Set the worksheets
    Set wsSource = ThisWorkbook.Sheets("Rework DataOutput") ' Adjust the source sheet name as needed
    Set wsDest = ThisWorkbook.Sheets("Master Sheet") ' Adjust the destination sheet name as needed

    ' Find the last row with data in column A on both sheets
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row

    ' Initialize the dictionary to store values by company
    Set companyDict = CreateObject("Scripting.Dictionary")

    ' Loop through each company in column A on the source sheet and store values in the dictionary
    For i = 2 To lastRowSource
        companyName = wsSource.Cells(i, 1).value
        If Not IsEmpty(companyName) Then
            If Not companyDict.Exists(companyName) Then
                companyDict.Add companyName, Array(wsSource.Cells(i, 2).value, wsSource.Cells(i, 3).value)
            End If
        End If
    Next i

    ' Clear columns F and G on the destination sheet before inputting anything
    wsDest.Range("F2:G" & wsDest.Rows.Count).ClearContents

    ' Add headers to columns F and G on the destination sheet
    wsDest.Cells(1, 6).value = "Rework Cost"
    wsDest.Cells(1, 7).value = "Total Cost"

    ' Loop through each company in column A on the destination sheet and output the corresponding values
    For j = 2 To lastRowDest
        companyName = wsDest.Cells(j, 1).value
        If companyDict.Exists(companyName) Then
            wsDest.Cells(j, 6).value = companyDict(companyName)(0) ' Rework Cost
            wsDest.Cells(j, 7).value = companyDict(companyName)(1) ' Total Cost
        End If
    Next j

    ' Autofit columns F and G on the destination sheet
    wsDest.Columns("F:G").AutoFit

    ' Clean up
    Set companyDict = Nothing
    Set wsSource = Nothing
    Set wsDest = Nothing
End Sub


Sub CopyResponseDataToMasterSheet()
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim lastRowSource As Long, lastRowDest As Long
    Dim companyName As String
    Dim i As Long, j As Long
    Dim companyDict As Object

    ' Set the worksheets
    Set wsSource = ThisWorkbook.Sheets("Response DataOutput") ' Adjust the source sheet name as needed
    Set wsDest = ThisWorkbook.Sheets("Master Sheet") ' Adjust the destination sheet name as needed

    ' Find the last row with data in column A on both sheets
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row

    ' Initialize the dictionary to store values by company
    Set companyDict = CreateObject("Scripting.Dictionary")

    ' Loop through each company in column A on the source sheet and store values in the dictionary
    For i = 2 To lastRowSource
        companyName = wsSource.Cells(i, 1).value
        If Not IsEmpty(companyName) Then
            If Not companyDict.Exists(companyName) Then
                companyDict.Add companyName, Array(wsSource.Cells(i, 2).value, wsSource.Cells(i, 3).value)
            End If
        End If
    Next i

    ' Clear columns H and I on the destination sheet before inputting anything
    wsDest.Range("H2:I" & wsDest.Rows.Count).ClearContents

    ' Add headers to columns H and I on the destination sheet
    wsDest.Cells(1, 8).value = "Time Until Order Confirmation Received"
    wsDest.Cells(1, 9).value = "Time Until Quality Issue Response"

    ' Loop through each company in column A on the destination sheet and output the corresponding values
    For j = 2 To lastRowDest
        companyName = wsDest.Cells(j, 1).value
        If companyDict.Exists(companyName) Then
            wsDest.Cells(j, 8).value = companyDict(companyName)(0) ' Time Until Order Confirmation Received
            wsDest.Cells(j, 9).value = companyDict(companyName)(1) ' Time Until Quality Issue Response
        End If
    Next j

    ' Autofit columns H and I on the destination sheet
    wsDest.Columns("H:I").AutoFit

    ' Clean up
    Set companyDict = Nothing
    Set wsSource = Nothing
    Set wsDest = Nothing
End Sub

