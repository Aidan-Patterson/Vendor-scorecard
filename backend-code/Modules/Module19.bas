Attribute VB_Name = "Module19"
Sub UpdateNCRDataTEST()
    Dim wsInputReal As Worksheet
    Dim wsInputTest As Worksheet
    Dim wsNCRData As Worksheet
    Dim fractionValue As String
    Dim numerator As Long
    Dim lastRow As Long
    Dim parts() As String
    Dim fractionParts() As String
    Dim wbReal As Workbook
    Dim wbTest As Workbook
    
    ' Set references to the workbooks
    Set wbReal = Workbooks("Vendor Scorecard EXAMPLE.xlsm")
  
    
    ' Set references to the sheets
    Set wsInputReal = wbReal.Sheets("Input")
    Set wsInputTest = wbTest.Sheets("Input")
    Set wsNCRData = wbRE.Sheets("NCR Data")
    
    ' Get the fraction value from cell B22 on the Input sheet in Vendor Scorecard REAL.xlsm
    fractionValue = wsInputReal.Range("B22").value
    
    ' Extract the numerator from the fraction (assuming the fraction format is "Input No. n/m")
    parts = Split(fractionValue, " ")
    fractionParts = Split(parts(2), "/")
    numerator = CLng(fractionParts(0))
    
    ' Add 1 to the numerator
    numerator = numerator + 1
    
    ' Ensure the numerator is within a valid range
    lastRow = wsNCRData.Cells(wsNCRData.Rows.Count, "A").End(xlUp).Row
    If numerator > lastRow Or numerator < 1 Then
        MsgBox "Row number " & numerator & " is out of range.", vbExclamation
        Exit Sub
    End If
    
    ' Delete the specified row in NCR Data sheet
    wsNCRData.Rows(numerator).Delete
    
    ' Clean up
    Set wsInputReal = Nothing
    Set wsInputTest = Nothing
    Set wsNCRData = Nothing
    Set wbTest = Nothing
    Set wbReal = Nothing
    
    ' Call the UFillSequentialNumbersNCR macro in the Vendor Scorecard REAL.xlsm workbook
    Call UFillSequentialNumbersNCRTEST
End Sub
Sub UFillSequentialNumbersNCRTEST()
    Dim wsNCRData As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim wbTest As Workbook

    ' Set reference to the Vendor Scorecard TEST workbook
    Set wbTest = Workbooks("Vendor Scorecard TEST.xlsm")
    
    ' Set reference to the NCR Data sheet in Vendor Scorecard TEST workbook
    Set wsNCRData = wbTest.Sheets("NCR Data")

    ' Get the last row in column A
    lastRow = wsNCRData.Cells(wsNCRData.Rows.Count, "A").End(xlUp).Row
    
    ' Fill column E with sequential numbers starting from 1 in row 2
    For i = 2 To lastRow
        If wsNCRData.Cells(i, "A").value <> "" Then
            wsNCRData.Cells(i, "E").value = i - 1
        Else
            Exit For
        End If
    Next i

    ' Clean up
    Set wsNCRData = Nothing
    Set wbTest = Nothing
End Sub


Sub UpdateReworkDataTEST()
    Dim wsInputReal As Worksheet
    Dim wsInputTest As Worksheet
    Dim wsReworkData As Worksheet
    Dim fractionValue As String
    Dim numerator As Long
    Dim lastRow As Long
    Dim parts() As String
    Dim fractionParts() As String
    Dim wbReal As Workbook
    Dim wbTest As Workbook
    
    ' Set references to the workbooks
    Set wbReal = Workbooks("Vendor Scorecard REAL.xlsm")
    Set wbTest = Workbooks("Vendor Scorecard TEST.xlsm")
    
    ' Set references to the sheets
    Set wsInputReal = wbReal.Sheets("Input")
    Set wsInputTest = wbTest.Sheets("Input")
    Set wsReworkData = wbTest.Sheets("Rework Data")
    
    ' Get the fraction value from cell B22 on the Input sheet in Vendor Scorecard REAL.xlsm
    fractionValue = wsInputReal.Range("B22").value
    
    ' Extract the numerator from the fraction (assuming the fraction format is "Input No. n/m")
    parts = Split(fractionValue, " ")
    fractionParts = Split(parts(2), "/")
    numerator = CLng(fractionParts(0))
    
    ' Add 1 to the numerator
    numerator = numerator + 1
    
    ' Ensure the numerator is within a valid range
    lastRow = wsReworkData.Cells(wsReworkData.Rows.Count, "A").End(xlUp).Row
    If numerator > lastRow Or numerator < 1 Then
        MsgBox "Row number " & numerator & " is out of range.", vbExclamation
        Exit Sub
    End If
    
    ' Delete the specified row in Rework Data sheet
    wsReworkData.Rows(numerator).Delete
    
    ' Clean up
    Set wsInputReal = Nothing
    Set wsInputTest = Nothing
    Set wsReworkData = Nothing
    Set wbTest = Nothing
    Set wbReal = Nothing
    
    ' Call the UFillSequentialNumbersRework macro in the Vendor Scorecard REAL.xlsm workbook
    Call UFillSequentialNumbersReworkTEST
End Sub

Sub UFillSequentialNumbersReworkTEST()
    Dim wsReworkData As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim wbTest As Workbook

    ' Set reference to the Vendor Scorecard TEST workbook
    Set wbTest = Workbooks("Vendor Scorecard TEST.xlsm")
    
    ' Set reference to the Rework Data sheet in Vendor Scorecard TEST workbook
    Set wsReworkData = wbTest.Sheets("Rework Data")

    ' Get the last row in column A
    lastRow = wsReworkData.Cells(wsReworkData.Rows.Count, "A").End(xlUp).Row
    
    ' Fill column D with sequential numbers starting from 1 in row 2
    For i = 2 To lastRow
        If wsReworkData.Cells(i, "A").value <> "" Then
            wsReworkData.Cells(i, "D").value = i - 1
        Else
            Exit For
        End If
    Next i

    ' Clean up
    Set wsReworkData = Nothing
    Set wbTest = Nothing
End Sub


Sub UpdateResponseDataTEST()
    Dim wsInputReal As Worksheet
    Dim wsInputTest As Worksheet
    Dim wsResponseData As Worksheet
    Dim fractionValue As String
    Dim numerator As Long
    Dim lastRow As Long
    Dim parts() As String
    Dim fractionParts() As String
    Dim wbReal As Workbook
    Dim wbTest As Workbook
    
    ' Set references to the workbooks
    Set wbReal = Workbooks("Vendor Scorecard REAL.xlsm")
    Set wbTest = Workbooks("Vendor Scorecard TEST.xlsm")
    
    ' Set references to the sheets
    Set wsInputReal = wbReal.Sheets("Input")
    Set wsInputTest = wbTest.Sheets("Input")
    Set wsResponseData = wbTest.Sheets("Response Data")
    
    ' Get the fraction value from cell B22 on the Input sheet in Vendor Scorecard REAL.xlsm
    fractionValue = wsInputReal.Range("B22").value
    
    ' Extract the numerator from the fraction (assuming the fraction format is "Input No. n/m")
    parts = Split(fractionValue, " ")
    fractionParts = Split(parts(2), "/")
    numerator = CLng(fractionParts(0))
    
    ' Add 1 to the numerator
    numerator = numerator + 1
    
    ' Ensure the numerator is within a valid range
    lastRow = wsResponseData.Cells(wsResponseData.Rows.Count, "A").End(xlUp).Row
    If numerator > lastRow Or numerator < 1 Then
        MsgBox "Row number " & numerator & " is out of range.", vbExclamation
        Exit Sub
    End If
    
    ' Delete the specified row in Response Data sheet
    wsResponseData.Rows(numerator).Delete
    
    ' Clean up
    Set wsInputReal = Nothing
    Set wsInputTest = Nothing
    Set wsResponseData = Nothing
    Set wbTest = Nothing
    Set wbReal = Nothing
    
    ' Call the UFillSequentialNumbersResponse macro in the Vendor Scorecard REAL.xlsm workbook
    Call UFillSequentialNumbersResponseTEST
End Sub

Sub UFillSequentialNumbersResponseTEST()
    Dim wsResponseData As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim wbTest As Workbook

    ' Set reference to the Vendor Scorecard TEST workbook
    Set wbTest = Workbooks("Vendor Scorecard TEST.xlsm")
    
    ' Set reference to the Response Data sheet in Vendor Scorecard TEST workbook
    Set wsResponseData = wbTest.Sheets("Response Data")

    ' Get the last row in column A
    lastRow = wsResponseData.Cells(wsResponseData.Rows.Count, "A").End(xlUp).Row
    
    ' Fill column E with sequential numbers starting from 1 in row 2
    For i = 2 To lastRow
        If wsResponseData.Cells(i, "A").value <> "" Then
            wsResponseData.Cells(i, "E").value = i - 1
        Else
            Exit For
        End If
    Next i

    ' Clean up
    Set wsResponseData = Nothing
    Set wbTest = Nothing
End Sub

Sub DELETETEST()
Call UpdateNCRDataTEST
Call UpdateReworkDataTEST
Call UpdateResponseDataTEST

End Sub


Sub CheckAndProceedWithDeletion()
    Dim wbTest As Workbook
    Dim isOpen As Boolean
    Dim wbName As String

    ' Define the workbook name
    wbName = "Vendor Scorecard TEST.xlsm"
    isOpen = False

    ' Loop through all open workbooks to check if the specified workbook is open
    For Each wbTest In Workbooks
        If wbTest.name = wbName Then
            isOpen = True
            Exit For
        End If
    Next wbTest

    ' If the workbook is open, call the ProceedWithDeletion macro
    If isOpen Then
        Call NEWDELETE
        Call ClearFiltersInputTEST
    Else
        ' If the workbook is not open, output a message
        MsgBox wbName & " must be open in order to proceed with deletion.", vbExclamation
    End If
End Sub


Sub CheckAndProceedWithEDIT()
    Dim wbTest As Workbook
    Dim isOpen As Boolean
    Dim wbName As String

    ' Define the workbook name
    wbName = "Vendor Scorecard TEST.xlsm"
    isOpen = False

    ' Loop through all open workbooks to check if the specified workbook is open
    For Each wbTest In Workbooks
        If wbTest.name = wbName Then
            isOpen = True
            Exit For
        End If
    Next wbTest

    ' If the workbook is open, call the ProceedWithDeletion macro
    If isOpen Then
        Call ShowVendorInputForm
    Else
        ' If the workbook is not open, output a message
        MsgBox wbName & " must be open in order to proceed with editing.", vbExclamation
    End If
End Sub


