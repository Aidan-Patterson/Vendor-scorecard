Attribute VB_Name = "Module20"
Sub OpenVendorScorecard()
    MsgBox ("Please also open up workbook Vendor Scorecard TEST.xlsm")
    
End Sub

Sub FindAndOutputCompanyData()
    Dim wsInput As Worksheet
    Dim wsNCR As Worksheet
    Dim companyName As String
    Dim lastRowNCR As Long
    Dim lastRowOutput As Long
    Dim i As Long
    Dim outputRow As Long
    
    ' Set references to the worksheets
    Set wsInput = ThisWorkbook.Sheets("Input Finder")
    Set wsNCR = ThisWorkbook.Sheets("NCR Data")
    
    ' Get the company name from cell G6 on the Input Finder sheet
    companyName = wsInput.Range("G6").value
    
    ' Find the last row in column A on the NCR Data sheet
    lastRowNCR = wsNCR.Cells(wsNCR.Rows.Count, "A").End(xlUp).Row
    
    ' Start output in row 11 on Input Finder sheet
    outputRow = 11
    
    ' Clear any previous data in the output range
    wsInput.Range("F11:J" & wsInput.Rows.Count).ClearContents
    
    ' Loop through each row in NCR Data sheet to find matching company names
    For i = 1 To lastRowNCR
        If wsNCR.Cells(i, 1).value = companyName Then
            ' Output the company name in column G starting at row 11
            wsInput.Cells(outputRow, 7).value = wsNCR.Cells(i, 1).value
            
            ' Output the corresponding values from columns B, C, and D
            wsInput.Cells(outputRow, 8).value = wsNCR.Cells(i, 2).value
            wsInput.Cells(outputRow, 9).value = wsNCR.Cells(i, 3).value
            wsInput.Cells(outputRow, 10).value = wsNCR.Cells(i, 4).value
            
            ' Output the numbering in column F
            wsInput.Cells(outputRow, 6).value = outputRow - 10 ' Start numbering from 1 in row 11
            
            ' Move to the next row for output
            outputRow = outputRow + 1
        End If
    Next i
    
    ' Check if no matches were found

End Sub

Sub FindAndOutputCompanyDataWithRework()
    Dim wsInput As Worksheet
    Dim wsNCR As Worksheet
    Dim wsRework As Worksheet
    Dim companyName As String
    Dim lastRowNCR As Long
    Dim lastRowOutput As Long
    Dim i As Long
    Dim outputRow As Long
    Dim reworkValue As Variant
    
    ' Set references to the worksheets
    Set wsInput = ThisWorkbook.Sheets("Input Finder")
    Set wsNCR = ThisWorkbook.Sheets("NCR Data")
    Set wsRework = ThisWorkbook.Sheets("Rework Data")
    
    ' Get the company name from cell G6 on the Input Finder sheet
    companyName = wsInput.Range("G6").value
    
    ' Find the last row in column A on the NCR Data sheet
    lastRowNCR = wsNCR.Cells(wsNCR.Rows.Count, "A").End(xlUp).Row
    
    ' Start output in row 11 on Input Finder sheet
    outputRow = 11
    
    ' Clear any previous data in the output range
    wsInput.Range("F11:K" & wsInput.Rows.Count).ClearContents
    
    ' Loop through each row in NCR Data sheet to find matching company names
    For i = 1 To lastRowNCR
        If wsNCR.Cells(i, 1).value = companyName Then
            ' Output the company name in column G starting at row 11
            wsInput.Cells(outputRow, 7).value = wsNCR.Cells(i, 1).value
            
            ' Output the corresponding values from columns B, C, and D
            wsInput.Cells(outputRow, 8).value = wsNCR.Cells(i, 2).value
            wsInput.Cells(outputRow, 9).value = wsNCR.Cells(i, 3).value
            wsInput.Cells(outputRow, 10).value = wsNCR.Cells(i, 4).value
            
            ' Output the numbering in column F
            wsInput.Cells(outputRow, 6).value = outputRow - 10 ' Start numbering from 1 in row 11
            
            ' Retrieve the corresponding value from column C on the Rework Data sheet
            reworkValue = wsRework.Cells(i, 3).value
            
            ' Output the rework value in column K starting at row 11
            wsInput.Cells(outputRow, 11).value = reworkValue
            
            ' Move to the next row for output
            outputRow = outputRow + 1
        End If
    Next i
    
    ' Check if no matches were found
    If outputRow = 11 Then
        MsgBox "No matching company names found in NCR Data.", vbInformation
    End If
End Sub


Sub FindAndOutputCompanyDataWithResponse()
    Dim wsInput As Worksheet
    Dim wsNCR As Worksheet
    Dim wsRework As Worksheet
    Dim wsResponse As Worksheet
    Dim companyName As String
    Dim lastRowNCR As Long
    Dim i As Long
    Dim outputRow As Long
    Dim reworkValue As Variant
    Dim responseValue As Variant
    Dim responseOutput As String
    Dim responseDate As Variant
    
    ' Set references to the worksheets
    Set wsInput = ThisWorkbook.Sheets("Input Finder")
    Set wsNCR = ThisWorkbook.Sheets("NCR Data")
    Set wsRework = ThisWorkbook.Sheets("Rework Data")
    Set wsResponse = ThisWorkbook.Sheets("Response Data")
    
    ' Get the company name from cell G6 on the Input Finder sheet
    companyName = wsInput.Range("G6").value
    
    ' Find the last row in column A on the NCR Data sheet
    lastRowNCR = wsNCR.Cells(wsNCR.Rows.Count, "A").End(xlUp).Row
    
    ' Start output in row 11 on Input Finder sheet
    outputRow = 11
    
    ' Clear any previous data in the output range
    wsInput.Range("F11:M" & wsInput.Rows.Count).ClearContents
    
    ' Loop through each row in NCR Data sheet to find matching company names
    For i = 1 To lastRowNCR
        If wsNCR.Cells(i, 1).value = companyName Then
            ' Output the company name in column G starting at row 11
            wsInput.Cells(outputRow, 7).value = wsNCR.Cells(i, 1).value
            
            ' Output the corresponding values from columns B, C, and D
            wsInput.Cells(outputRow, 8).value = wsNCR.Cells(i, 2).value
            wsInput.Cells(outputRow, 9).value = wsNCR.Cells(i, 3).value
            wsInput.Cells(outputRow, 10).value = wsNCR.Cells(i, 4).value
            
            ' Output the numbering in column F
            wsInput.Cells(outputRow, 6).value = outputRow - 10 ' Start numbering from 1 in row 11
            
            ' Retrieve the corresponding value from column C on the Rework Data sheet
            reworkValue = wsRework.Cells(i, 3).value
            
            ' Output the rework value in column K starting at row 11
            wsInput.Cells(outputRow, 11).value = reworkValue
            
            ' Retrieve the corresponding value from column C on the Response Data sheet
            responseValue = wsResponse.Cells(i, 3).value
            responseDate = wsResponse.Cells(i, 4).value
            
            ' Convert the response value (1 or not 1) to Yes/No
            If responseValue = 1 Then
                responseOutput = "Yes"
            Else
                responseOutput = "No"
            End If
            
            ' Output the response value in column L starting at row 11
            wsInput.Cells(outputRow, 12).value = responseOutput
            
            ' Output the corresponding value from column D on the Response Data sheet in column M
            wsInput.Cells(outputRow, 13).value = responseDate
            
            ' Move to the next row for output
            outputRow = outputRow + 1
        End If
    Next i
    
    ' Check if no matches were found
    If outputRow = 11 Then
        MsgBox "No matching company names found in NCR Data.", vbInformation
    End If
End Sub

Sub FindAndOutputCompanyDataWithFinalChanges()
    Dim wsInput As Worksheet
    Dim wsNCR As Worksheet
    Dim wsRework As Worksheet
    Dim wsResponse As Worksheet
    Dim companyName As String
    Dim lastRowNCR As Long
    Dim i As Long
    Dim outputRow As Long
    Dim reworkValue As Variant
    Dim responseValue As Variant
    Dim responseOutput As String
    Dim valueB As Variant
    Dim valueC As Variant
    Dim valueD As Variant
    Dim valueI As String
    Dim valueJ As String
    
    ' Set references to the worksheets
    Set wsInput = ThisWorkbook.Sheets("Input Finder")
    Set wsNCR = ThisWorkbook.Sheets("NCR Data")
    Set wsRework = ThisWorkbook.Sheets("Rework Data")
    Set wsResponse = ThisWorkbook.Sheets("Response Data")
    
    ' Get the company name from cell G6 on the Input Finder sheet
    companyName = wsInput.Range("G6").value
    
    ' Find the last row in column A on the NCR Data sheet
    lastRowNCR = wsNCR.Cells(wsNCR.Rows.Count, "A").End(xlUp).Row
    
    ' Start output in row 11 on Input Finder sheet
    outputRow = 12
    
    ' Clear any previous data in the output range
    wsInput.Range("F12:M" & wsInput.Rows.Count).ClearContents
    
    ' Loop through each row in NCR Data sheet to find matching company names
    For i = 1 To lastRowNCR
        If wsNCR.Cells(i, 1).value = companyName Then
            ' Output the company name in column G starting at row 11
            wsInput.Cells(outputRow, 7).value = wsNCR.Cells(i, 1).value
            
            ' Convert the values in columns B, C, and D to Yes/No for columns I and J
            valueB = wsNCR.Cells(i, 2).value
            valueC = wsNCR.Cells(i, 3).value
            valueD = wsNCR.Cells(i, 4).value
            
            If valueC = 1 Then
                valueI = "Yes"
            Else
                valueI = "No"
            End If
            
            If valueD = 1 Then
                valueJ = "Yes"
            Else
                valueJ = "No"
            End If
            
            ' Output the corresponding values
            wsInput.Cells(outputRow, 8).value = valueB ' Column H
            wsInput.Cells(outputRow, 9).value = valueI ' Column I
            wsInput.Cells(outputRow, 10).value = valueJ ' Column J
            
            ' Output the numbering in column F
            wsInput.Cells(outputRow, 6).value = outputRow - 11 ' Start numbering from 1 in row 11
            
            ' Retrieve the corresponding value from column C on the Rework Data sheet
            reworkValue = wsRework.Cells(i, 3).value
            
            ' Output the rework value in column K starting at row 11
            wsInput.Cells(outputRow, 11).value = reworkValue
            
            ' Retrieve the corresponding value from column C on the Response Data sheet
            responseValue = wsResponse.Cells(i, 3).value
            responseOutput = IIf(responseValue = 1, "Yes", "No")
            
            ' Output the response value in column L starting at row 11
            wsInput.Cells(outputRow, 12).value = responseOutput
            
            ' Output the corresponding value from column D on the Response Data sheet in column M
            wsInput.Cells(outputRow, 13).value = wsResponse.Cells(i, 4).value
            
            ' Move to the next row for output
            outputRow = outputRow + 1
        End If
    Next i
    
    ' Check if no matches were found
    If outputRow = 11 Then
        MsgBox "No matching company names found in NCR Data.", vbInformation
    End If
End Sub

