Attribute VB_Name = "Module18"
Sub FindLastFilledCellAndCorrespondingValues()
    Dim wsNCR As Worksheet
    Dim wsInput As Worksheet
    Dim lastRow As Long
    Dim targetRow As Long
    Dim lastValueA As Variant
    Dim correspondingValueB As Variant
    Dim correspondingValueC As Variant
    Dim correspondingValueD As Variant
    Dim correspondingValueE As Variant
    Dim ncheck2 As checkbox
    Dim ocheck2 As checkbox
    Dim totalRows As Long
    
    ' Set the worksheets
    Set wsNCR = ThisWorkbook.Sheets("NCR Data")
    Set wsInput = ThisWorkbook.Sheets("Input")
    
    ' Find the last filled cell in column A of the "NCR Data" sheet
    With wsNCR
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        targetRow = lastRow ' Get the row one above the last filled cell
        If targetRow < 1 Then targetRow = 1 ' Ensure the target row is at least 1
        
        lastValueA = .Cells(targetRow, "A").value
        correspondingValueB = .Cells(targetRow, "B").value
        correspondingValueC = .Cells(targetRow, "C").value
        correspondingValueD = .Cells(targetRow, "D").value
        correspondingValueE = .Cells(targetRow, "E").value
        
        ' Get the total number of values in column A starting from row 2
        totalRows = Application.WorksheetFunction.CountA(.Range("A2:A" & .Rows.Count))
    End With
    
    ' Output the values in cells B26 and D26 on the "Input" sheet
    wsInput.Range("B26").value = lastValueA
    wsInput.Range("D26").value = correspondingValueB
    
    ' Output "Input No." and the value, formatted as "Input No. X/Y" in cell B22
    wsInput.Range("B22").value = "Input No. " & correspondingValueE & "/" & totalRows
    
    ' Check or uncheck the "ncheck2" checkbox based on the value in column C
    Set ncheck2 = wsInput.CheckBoxes("ncheck2")
    If correspondingValueC = 1 Then
        ncheck2.value = xlOn
    Else
        ncheck2.value = xlOff
    End If
    
    ' Check or uncheck the "ocheck2" checkbox based on the value in column D
    Set ocheck2 = wsInput.CheckBoxes("ocheck2")
    If correspondingValueD = 1 Then
        ocheck2.value = xlOn
    Else
        ocheck2.value = xlOff
    End If
End Sub





Sub OutputLastValue()
    Dim wsReworkData As Worksheet
    Dim wsInput As Worksheet
    Dim lastRow As Long
    Dim lastValue As Variant
    
    ' Set references to the worksheets
    Set wsReworkData = ThisWorkbook.Sheets("Rework Data")
    Set wsInput = ThisWorkbook.Sheets("Input")
    
    ' Find the last non-empty row in column A of the "rework" table
    With wsReworkData
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    
    ' Retrieve the corresponding value from column C
    lastValue = wsReworkData.Cells(lastRow, "C").value
    lastValue1 = wsReworkData.Cells(lastRow, "E").value
    lastValue2 = wsReworkData.Cells(lastRow, "F").value
    ' Output the value to cell J27 on the "Input" sheet
    wsInput.Range("L27").value = lastValue
    wsInput.Range("J27").value = lastValue1
    wsInput.Range("K27").value = lastValue2
    

End Sub



Sub FindLastResponse()
    Dim wsNCR As Worksheet
    Dim wsInput As Worksheet
    Dim lastRow As Long
    Dim lastValueA As Variant
    Dim correspondingValueB As Variant
    Dim correspondingValueC As Variant
    Dim correspondingValueD As Variant
    Dim ncheck2 As checkbox
    Dim ocheck2 As checkbox
    
    ' Set the worksheets
    Set wsNCR = ThisWorkbook.Sheets("Response Data")
    Set wsInput = ThisWorkbook.Sheets("Input")
    
    ' Find the last filled cell in column A of the "NCR Data" sheet
    With wsNCR
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        lastValueA = .Cells(lastRow, "A").value

        correspondingValueC = .Cells(lastRow, "C").value
        correspondingValueD = .Cells(lastRow, "D").value
    End With
    
    ' Output the values in cells B26 and D26 on the "Input" sheet

    wsInput.Range("L33").value = correspondingValueD
    
    ' Check or uncheck the "ncheck2" checkbox based on the value in column C
    Set ncheck2 = wsInput.CheckBoxes("ocrcheck2")
    If correspondingValueC = 1 Then
        ncheck2.value = xlOn
    Else
        ncheck2.value = xlOff
    End If
    
    ' Check or uncheck the "ocheck2" checkbox based on the value in column D

End Sub




Sub PREVIOUS()
Call ClearFiltersInput
Call FindLastFilledCellAndCorrespondingValues
Call OutputLastValue
Call FindLastResponse

End Sub

Sub DeleteLastNCRRecord()
    Dim wsNCR As Worksheet
    Dim wsInput As Worksheet
    Dim numerator As Long
    Dim fractionStr As String
    Dim foundCell As Range
    Dim searchRange As Range

    ' Set the worksheets
    Set wsNCR = ThisWorkbook.Sheets("NCR Data")
    Set wsInput = ThisWorkbook.Sheets("Input")

    ' Get the fraction from cell B22 and extract the numerator
    fractionStr = wsInput.Range("B22").value
    fractionStr = Replace(fractionStr, "Input No. ", "")
    numerator = CLng(Split(fractionStr, "/")(0))
    
    ' Define the search range in column E
    Set searchRange = wsNCR.Range("E:E")
    
    ' Find the numerator in column E on the "NCR Data" sheet
    Set foundCell = searchRange.Find(What:=numerator, LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    
    ' Check if the numerator was found
    If Not foundCell Is Nothing Then
        Dim lastRow As Long
        lastRow = foundCell.Row
        
        ' Delete the entire row and shift cells up
        wsNCR.Rows(lastRow).Delete xlShiftUp
    Else
        MsgBox "The numerator " & numerator & " was not found in column E on the 'NCR Data' sheet.", vbExclamation
    End If
End Sub




Sub DeleteLastReworkRecord()
    Dim wsNCR As Worksheet
    Dim lastRow As Long
    
    ' Set the worksheet
    Set wsNCR = ThisWorkbook.Sheets("Rework Data")
    
    ' Find the last filled row in column A of the "NCR Data" sheet
    lastRow = wsNCR.Cells(wsNCR.Rows.Count, "A").End(xlUp).Row
    
    ' Check if there's data to delete
    If lastRow > 1 Then ' Assuming row 1 is headers
        ' Delete the value in column A, C, and D of the last filled row
        wsNCR.Cells(lastRow, "A").ClearContents
        wsNCR.Cells(lastRow, "B").ClearContents
        wsNCR.Cells(lastRow, "C").ClearContents
 
        
        ' Optionally, delete entire row if you want to remove all columns data in the row
        ' wsNCR.Rows(lastRow).Delete
    Else
        MsgBox "No data to delete.", vbExclamation
    End If
End Sub

Sub DeleteLastResponseRecord()
    Dim wsNCR As Worksheet
    Dim lastRow As Long
    
    ' Set the worksheet
    Set wsNCR = ThisWorkbook.Sheets("Response Data")
    
    ' Find the last filled row in column A of the "NCR Data" sheet
    lastRow = wsNCR.Cells(wsNCR.Rows.Count, "A").End(xlUp).Row
    
    ' Check if there's data to delete
    If lastRow > 1 Then ' Assuming row 1 is headers
        ' Delete the value in column A, C, and D of the last filled row
        wsNCR.Cells(lastRow, "A").ClearContents
        wsNCR.Cells(lastRow, "B").ClearContents
        wsNCR.Cells(lastRow, "C").ClearContents
        wsNCR.Cells(lastRow, "D").ClearContents
        
        ' Optionally, delete entire row if you want to remove all columns data in the row
        ' wsNCR.Rows(lastRow).Delete
    Else
        MsgBox "No data to delete.", vbExclamation
    End If
End Sub

Sub DeleteLast()

Call DeleteLastNCRRecord
Call DeleteLastReworkRecord
Call DeleteLastResponseRecord
Call DeleteLastEntryInVendorScorecardTEST
Call DeleteLastEntryInVendorScorecardTESTrework
Call DeleteLastEntryInVendorScorecardTESTresponse

Call PREVIOUS
End Sub

Sub ConfirmAndDelete()
    Dim wsInput As Worksheet
    Dim valueB26 As Variant
    Dim valueD26 As Variant
    Dim valueJ27 As Variant
    Dim valueL33 As Variant
    Dim userResponse As VbMsgBoxResult
    Dim message As String
    Dim ncheck2 As Boolean
    Dim ocheck2 As Boolean
    Dim ocrcheck2 As Boolean
    
    ' Set the worksheet
    Set wsInput = ThisWorkbook.Sheets("Input")
    
    ' Get the values from cells B26, D26, J27, and L33
    valueB26 = wsInput.Range("B26").value
    valueD26 = wsInput.Range("D26").value
    valueL27 = wsInput.Range("L27").value
    valueL33 = wsInput.Range("L33").value
    
    ' Check the state of the form control checkboxes
    ncheck2 = wsInput.CheckBoxes("ncheck2").value = 1
    ocheck2 = wsInput.CheckBoxes("ocheck2").value = 1
    ocrcheck2 = wsInput.CheckBoxes("ocrcheck2").value = 1
    
    ' Build the message
    message = "Are you sure you would like to delete this input?" & vbCrLf & _
              valueB26 & vbCrLf & _
              valueD26 & vbCrLf
    
    ' Append checkbox states to the message
    If ncheck2 Then
        message = message & "NCR" & vbCrLf
    End If
    
    If ocheck2 Then
        message = message & "Occurrence" & vbCrLf
    End If
    
    ' Append additional values to the message
    message = message & "$" & valueL27 & " " & "rework cost" & vbCrLf
    
    If ocrcheck2 Then
        message = message & "Order Confirmed" & vbCrLf
    End If
    
    message = message & valueL33 & " " & "day(s) until quality issue response"
    
    ' Prompt the user for confirmation
    userResponse = MsgBox(message, vbYesNo + vbQuestion, "Confirm Deletion")
    
    ' If the user clicks "Yes", proceed with the macro
    If userResponse = vbYes Then
        Call NEWDELETE
    Else
        MsgBox "Deletion canceled.", vbInformation
    End If
End Sub


Sub DeleteLastEntryInVendorScorecardTEST()
    Dim wbTest As Workbook
    Dim wsNCR As Worksheet
    Dim lastRow As Long
    
    ' Set the workbook and worksheet
    On Error Resume Next
    Set wbTest = Workbooks("Vendor Scorecard TEST.xlsm") ' Ensure the correct file extension
    On Error GoTo 0
    
    If wbTest Is Nothing Then
        MsgBox "Vendor Scorecard TEST workbook is not open.", vbExclamation
        Exit Sub
    End If
    
    Set wsNCR = wbTest.Sheets("NCR Data")
    
    ' Find the last filled row in column A of the "NCR Data" sheet
    lastRow = wsNCR.Cells(wsNCR.Rows.Count, "A").End(xlUp).Row
    
    ' Check if there's data to delete
    If lastRow > 1 Then ' Assuming row 1 is headers
        ' Delete the values in columns A, B, C, and D of the last filled row
        wsNCR.Cells(lastRow, "A").ClearContents
        wsNCR.Cells(lastRow, "B").ClearContents
        wsNCR.Cells(lastRow, "C").ClearContents
        wsNCR.Cells(lastRow, "D").ClearContents
    Else
        MsgBox "No data to delete.", vbExclamation
    End If
End Sub

Sub DeleteLastEntryInVendorScorecardTESTrework()
    Dim wbTest As Workbook
    Dim wsNCR As Worksheet
    Dim lastRow As Long
    
    ' Set the workbook and worksheet
    On Error Resume Next
    Set wbTest = Workbooks("Vendor Scorecard TEST.xlsm") ' Ensure the correct file extension
    On Error GoTo 0
    

    
    Set wsNCR = wbTest.Sheets("Rework Data")
    
    ' Find the last filled row in column A of the "NCR Data" sheet
    lastRow = wsNCR.Cells(wsNCR.Rows.Count, "A").End(xlUp).Row
    
    ' Check if there's data to delete
    If lastRow > 1 Then ' Assuming row 1 is headers
        ' Delete the values in columns A, B, C, and D of the last filled row
        wsNCR.Cells(lastRow, "A").ClearContents
        wsNCR.Cells(lastRow, "B").ClearContents
        wsNCR.Cells(lastRow, "C").ClearContents

    Else
        MsgBox "No data to delete.", vbExclamation
    End If
End Sub

Sub DeleteLastEntryInVendorScorecardTESTresponse()
    Dim wbTest As Workbook
    Dim wsNCR As Worksheet
    Dim lastRow As Long
    
    ' Set the workbook and worksheet
    On Error Resume Next
    Set wbTest = Workbooks("Vendor Scorecard TEST.xlsm") ' Ensure the correct file extension
    On Error GoTo 0
    

    
    Set wsNCR = wbTest.Sheets("Response Data")
    
    ' Find the last filled row in column A of the "NCR Data" sheet
    lastRow = wsNCR.Cells(wsNCR.Rows.Count, "A").End(xlUp).Row
    
    ' Check if there's data to delete
    If lastRow > 1 Then ' Assuming row 1 is headers
        ' Delete the values in columns A, B, C, and D of the last filled row
        wsNCR.Cells(lastRow, "A").ClearContents
        wsNCR.Cells(lastRow, "B").ClearContents
        wsNCR.Cells(lastRow, "C").ClearContents
        wsNCR.Cells(lastRow, "D").ClearContents
    Else
        MsgBox "No data to delete.", vbExclamation
    End If
End Sub

Sub MatchCompaniesAndUpdate()
    Dim wsResponse As Worksheet
    Dim wsMaster As Worksheet
    Dim responseRange As Range
    Dim masterRange As Range
    Dim cell As Range
    Dim companyFound As Boolean
    
    ' Set the worksheets
    Set wsResponse = ThisWorkbook.Sheets("Response DataOutput")
    Set wsMaster = ThisWorkbook.Sheets("Master Sheet")
    
    ' Set the ranges
    Set responseRange = wsResponse.Range("A:A")
    Set masterRange = wsMaster.Range("A2:A" & wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row)
    
    ' Loop through each company in the master sheet
    For Each cell In masterRange
        companyFound = False
        ' Check if the company exists in the response sheet
        If Not IsError(Application.Match(cell.value, responseRange, 0)) Then
            companyFound = True
        End If
        
        ' If the company is not found, output "1" in column H
        If Not companyFound Then
            wsMaster.Cells(cell.Row, "H").value = 1
        End If
    Next cell
End Sub

Sub ProcessFractionInputNCR()
    Dim wsInput As Worksheet
    Dim wsNCRData As Worksheet
    Dim fractionStr As String
    Dim X As Long, Y As Long
    Dim rowNum As Long
    
    ' Define the worksheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsNCRData = ThisWorkbook.Sheets("NCR Data")
    
    ' Get the fraction from cell B22
    fractionStr = wsInput.Range("B22").value
    fractionStr = Replace(fractionStr, "Input No. ", "")
    
    ' Split the fraction into X and Y
    X = CLng(Split(fractionStr, "/")(0))
    Y = CLng(Split(fractionStr, "/")(1))
    
    ' Set the row number from X
    rowNum = X
    
    ' Get the values from NCR Data sheet
    Dim valueA As String
    Dim valueB As String
    Dim valueC As String
    Dim valueD As String
    Dim valueE As Long
    
    valueA = wsNCRData.Cells(rowNum, 1).value
    valueB = wsNCRData.Cells(rowNum, 2).value
    valueC = wsNCRData.Cells(rowNum, 3).value
    valueD = wsNCRData.Cells(rowNum, 4).value
    valueE = wsNCRData.Cells(rowNum, 5).value
    
    ' Display the values in the Input sheet
    wsInput.Range("B26").value = valueA
    wsInput.Range("D26").value = valueB
    
    ' Check the checkboxes based on values in column C and D
    wsInput.Shapes("ncheck2").ControlFormat.value = IIf(valueC = "1", 1, 0)
    wsInput.Shapes("ocheck2").ControlFormat.value = IIf(valueD = "1", 1, 0)
    
    ' Update the fraction in cell B22 with the new X value
    wsInput.Range("B22").value = "Input No. " & valueE & "/" & Y
End Sub


Sub GetValueFromReworkData()

    Dim wsInput As Worksheet
    Dim wsRework As Worksheet
    Dim inputValue As String
    Dim fraction() As String
    Dim rowNumber As Long
    Dim valueToDisplay As Variant

    ' Set references to the worksheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsRework = ThisWorkbook.Sheets("Rework Data")

    ' Get the value from cell B22 on the "Input" sheet
    inputValue = wsInput.Range("B22").value

    ' Check if the value contains "Input No."
    If InStr(inputValue, "Input No.") = 0 Then
        MsgBox "Cell B22 does not contain 'Input No.'"
        Exit Sub
    End If

    ' Extract the fraction part
    inputValue = Replace(inputValue, "Input No.", "")
    inputValue = Trim(inputValue)
    
    ' Split the fraction into X and Y
    fraction = Split(inputValue, "/")
    
    ' Get the row number (value of X)
    If IsNumeric(fraction(0)) Then
        rowNumber = CLng(fraction(0))
    Else
        MsgBox "Invalid fraction format in cell B22."
        Exit Sub
    End If

    ' Get the value in column C of the specified row on "Rework Data" sheet
    valueToDisplay = wsRework.Cells(rowNumber, 3).value
    valueToDisplay2 = wsRework.Cells(rowNumber, 5).value
    valueToDisplay3 = wsRework.Cells(rowNumber, 6).value

    ' Display the value in cell J27 on the "Input" sheet
    wsInput.Range("L27").value = valueToDisplay
    wsInput.Range("J27").value = valueToDisplay2
    wsInput.Range("K27").value = valueToDisplay3
    ' Inform the user that the operation is completed
   

End Sub
Sub ProcessFractionInputResponse()
    Dim wsInput As Worksheet
    Dim wsResponseData As Worksheet
    Dim fractionStr As String
    Dim X As Long, Y As Long
    Dim rowNum As Long
    
    ' Define the worksheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsNCRData = ThisWorkbook.Sheets("Response Data")
    
    ' Get the fraction from cell B22
    fractionStr = wsInput.Range("B22").value
    fractionStr = Replace(fractionStr, "Input No. ", "")
    
    ' Split the fraction into X and Y
    X = CLng(Split(fractionStr, "/")(0))
    Y = CLng(Split(fractionStr, "/")(1))
    
    ' Set the row number from X
    rowNum = X
    
    ' Get the values from NCR Data sheet

    Dim valueC As String
    Dim valueD As String

    

    valueC = wsNCRData.Cells(rowNum, 3).value
    valueD = wsNCRData.Cells(rowNum, 4).value

    
    ' Display the values in the Input sheet
    wsInput.Range("L33").value = valueD
  
    
    ' Check the checkboxes based on values in column C and D
    wsInput.Shapes("ocrcheck2").ControlFormat.value = IIf(valueC = "1", 1, 0)



End Sub



Sub ARROWBACK()

Call GetValueFromReworkData
Call ProcessFractionInputResponse
Call ProcessFractionInputNCR
Call CheckDenominatorMinusNumerator

End Sub


Sub ProcessFractionAndUpdateNCR()

    Dim wsInput As Worksheet
    Dim wsNCRData As Worksheet
    Dim fractionStr As String
    Dim X As Long, Y As Long
    Dim rowNum As Long
    
    ' Define the worksheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsNCRData = ThisWorkbook.Sheets("NCR Data")
    
    ' Get the fraction from cell B22
    fractionStr = wsInput.Range("B22").value
    fractionStr = Replace(fractionStr, "Input No. ", "")
    
    ' Split the fraction into X and Y
    X = CLng(Split(fractionStr, "/")(0))
    Y = CLng(Split(fractionStr, "/")(1))
    
    ' Set the row number from X
    rowNum = X + 2
    
    ' Get the values from NCR Data sheet
    Dim valueA As String
    Dim valueB As String
    Dim valueC As String
    Dim valueD As String

    valueA = wsNCRData.Cells(rowNum, 1).value
    valueB = wsNCRData.Cells(rowNum, 2).value
    valueC = wsNCRData.Cells(rowNum, 3).value
    valueD = wsNCRData.Cells(rowNum, 4).value

    
    ' Display the values in the Input sheet
    wsInput.Range("B26").value = valueA
    wsInput.Range("D26").value = valueB

    ' Check the checkboxes based on values in column C and D
    wsInput.Shapes("ncheck2").ControlFormat.value = IIf(valueC = "1", 1, 0)
    wsInput.Shapes("ocheck2").ControlFormat.value = IIf(valueD = "1", 1, 0)



End Sub

Sub ProcessFractionAndUpdateRework()

    Dim wsInput As Worksheet
    Dim wsReworkData As Worksheet
    Dim fractionStr As String
    Dim X As Long, Y As Long
    Dim rowNum As Long
    
    ' Define the worksheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsReworkData = ThisWorkbook.Sheets("Rework Data")
    
    ' Get the fraction from cell B22
    fractionStr = wsInput.Range("B22").value
    fractionStr = Replace(fractionStr, "Input No. ", "")
    
    ' Split the fraction into X and Y
    X = CLng(Split(fractionStr, "/")(0))
    Y = CLng(Split(fractionStr, "/")(1))
    
    ' Set the row number from X
    rowNum = X + 2
    
    ' Get the values from NCR Data sheet

    Dim valueC As String


    valueC = wsReworkData.Cells(rowNum, 3).value
    valueE = wsReworkData.Cells(rowNum, 5).value
    valueF = wsReworkData.Cells(rowNum, 6).value
    
    ' Display the values in the Input sheet
    wsInput.Range("L27").value = valueC
    wsInput.Range("J27").value = valueE
    wsInput.Range("K27").value = valueF

End Sub

Sub ProcessFractionAndUpdateResponse()

    Dim wsInput As Worksheet
    Dim wsResponseData As Worksheet
    Dim fractionStr As String
    Dim X As Long, Y As Long
    Dim rowNum As Long
    
    ' Define the worksheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsResponseData = ThisWorkbook.Sheets("Response Data")
    
    ' Get the fraction from cell B22
    fractionStr = wsInput.Range("B22").value
    fractionStr = Replace(fractionStr, "Input No. ", "")
    
    ' Split the fraction into X and Y
    X = CLng(Split(fractionStr, "/")(0))
    Y = CLng(Split(fractionStr, "/")(1))
    
    ' Set the row number from X
    rowNum = X + 2
    
    ' Get the values from NCR Data sheet

    Dim valueC As String
    Dim valueD As String


    valueC = wsResponseData.Cells(rowNum, 3).value
    valueD = wsResponseData.Cells(rowNum, 4).value
    valueE = wsResponseData.Cells(rowNum, 5).value
    
    ' Display the values in the Input sheet
    wsInput.Range("L33").value = valueD

    ' Check the checkboxes based on values in column C and D
    wsInput.Shapes("ocrcheck2").ControlFormat.value = IIf(valueC = "1", 1, 0)

    wsInput.Range("B22").value = "Input No. " & valueE & "/" & Y


End Sub

Sub ARROWFORWARD()

Call ProcessFractionAndUpdateNCR
Call ProcessFractionAndUpdateRework
Call ProcessFractionAndUpdateResponse

End Sub

Sub CheckFractionAndCallMacro()
    Dim inputSheet As Worksheet
    Dim cellContent As String
    Dim fraction As String
    Dim splitFraction() As String
    Dim numerator As Long
    Dim denominator As Long
    Dim result As Double
    Dim startPos As Long
    Dim shapeToChange As Shape
    Dim darkGreenAccent3 As Long

    
    ' Define the sheet
    Set inputSheet = ThisWorkbook.Sheets("Input")
    
    ' Define the RGB value for dark green accent 3
    darkGreenAccent3 = RGB(0, 97, 0)
    
    ' Get the content from cell B22
    cellContent = inputSheet.Range("B22").value
    
    ' Find the position of "Input No."
    startPos = InStr(cellContent, "Input No.")
    
    ' Check if "Input No." is found in the cell content
    If startPos > 0 Then
        ' Extract the fraction after "Input No."
        fraction = Trim(Mid(cellContent, startPos + Len("Input No.")))
        
        ' Split the fraction to get numerator and denominator
        splitFraction = Split(fraction, "/")
        
        ' Check if the fraction is in the correct format
        If UBound(splitFraction) = 1 And IsNumeric(splitFraction(0)) And IsNumeric(splitFraction(1)) Then
            numerator = CLng(splitFraction(0))
            denominator = CLng(splitFraction(1))
            
            ' Calculate the result of the fraction
            If denominator <> 0 Then
                result = denominator - numerator
                
                ' Get the shape
                Set shapeToChange = inputSheet.Shapes("forwards")
                
                ' Check if the fraction is less than 1
                If result >= 1 Then
                    ' Change the fill color of the shape to dark green accent 3
                    shapeToChange.FILL.ForeColor.RGB = darkGreenAccent3
                    ' Call your specific macro
                    Call ARROWFORWARD
                Else
                    ' Change the fill color of the shape to red
                    shapeToChange.FILL.ForeColor.RGB = RGB(255, 0, 0)
                    MsgBox "This is the last input available, cannot cycle through this way any further.", vbInformation
                End If
                
            Else
                MsgBox "Denominator cannot be zero.", vbExclamation
            End If
        Else
            MsgBox "The fraction is not in the correct format 'X/Y' or is not numeric.", vbExclamation
        End If
    Else
        MsgBox "'Input No.' not found in cell B22.", vbExclamation
    End If
    
    Call CheckDenominatorMinusNumerator
    Call CheckNumeratorAndChangeShapeColor
End Sub

Sub CheckDenominatorMinusNumerator()
    Dim inputSheet As Worksheet
    Dim cellContent As String
    Dim fraction As String
    Dim splitFraction() As String
    Dim numerator As Long
    Dim denominator As Long
    Dim shapeToChange As Shape
    Dim result As Long
    Dim startPos As Long
    Dim darkGreenAccent3 As Long

    ' Define the sheet
    Set inputSheet = ThisWorkbook.Sheets("Input")
    
    ' Define the RGB value for dark green accent 3
    darkGreenAccent3 = RGB(0, 97, 0)
    
    ' Get the content from cell B22
    cellContent = inputSheet.Range("B22").value
    
    ' Find the position of "Input No."
    startPos = InStr(cellContent, "Input No.")
    
    ' Check if "Input No." is found in the cell content
    If startPos > 0 Then
        ' Extract the fraction after "Input No."
        fraction = Trim(Mid(cellContent, startPos + Len("Input No.")))
        
        ' Split the fraction to get numerator and denominator
        splitFraction = Split(fraction, "/")
        
        ' Check if the fraction is in the correct format
        If UBound(splitFraction) = 1 And IsNumeric(splitFraction(0)) And IsNumeric(splitFraction(1)) Then
            numerator = CLng(splitFraction(0))
            denominator = CLng(splitFraction(1))
            
            ' Calculate the result of the denominator - numerator
            result = denominator - numerator
            
            ' Get the shape
            Set shapeToChange = inputSheet.Shapes("forwards")
            
            ' Check if the result is equal to 0
            If result = 0 Then
                ' Change the fill color of the shape to red
                shapeToChange.FILL.ForeColor.RGB = RGB(255, 0, 0)
            Else
                ' Change the fill color of the shape to dark green accent 3
                shapeToChange.FILL.ForeColor.RGB = darkGreenAccent3
            End If
        Else
            MsgBox "The fraction is not in the correct format 'X/Y' or is not numeric.", vbExclamation
        End If
    Else
        MsgBox "'Input No.' not found in cell B22.", vbExclamation
    End If
End Sub
Sub CheckFractionAndCallMacroBACK()
    Dim inputSheet As Worksheet
    Dim cellContent As String
    Dim fraction As String
    Dim splitFraction() As String
    Dim numerator As Long
    Dim denominator As Long
    Dim startPos As Long
    Dim shapeToChange As Shape
    Dim darkGreenAccent3 As Long

    ' Define the sheet
    Set inputSheet = ThisWorkbook.Sheets("Input")
    
    ' Define the RGB value for dark green accent 3
    darkGreenAccent3 = RGB(0, 97, 0)
    
    ' Get the content from cell B22
    cellContent = inputSheet.Range("B22").value
    
    ' Find the position of "Input No."
    startPos = InStr(cellContent, "Input No.")
    
    ' Check if "Input No." is found in the cell content
    If startPos > 0 Then
        ' Extract the fraction after "Input No."
        fraction = Trim(Mid(cellContent, startPos + Len("Input No.")))
        
        ' Split the fraction to get numerator and denominator
        splitFraction = Split(fraction, "/")
        
        ' Check if the fraction is in the correct format
        If UBound(splitFraction) = 1 And IsNumeric(splitFraction(0)) And IsNumeric(splitFraction(1)) Then
            numerator = CLng(splitFraction(0))
            denominator = CLng(splitFraction(1))
            
            ' Get the shape
            Set shapeToChange = inputSheet.Shapes("backwards")
            
            ' Check if the numerator is equal to 1
            If numerator = 1 Then
                ' Change the fill color of the shape to red
                shapeToChange.FILL.ForeColor.RGB = RGB(255, 0, 0)
                MsgBox "This is the first input available, cannot cycle through this way any further.", vbInformation
            Else
                ' Change the fill color of the shape to dark green accent 3
                shapeToChange.FILL.ForeColor.RGB = darkGreenAccent3
                ' Call your specific macro
                Call ARROWBACK
            End If
        Else
            MsgBox "The fraction is not in the correct format 'X/Y' or is not numeric.", vbExclamation
        End If
    Else
        MsgBox "'Input No.' not found in cell B22.", vbExclamation
    End If
    
    Call CheckNumeratorAndChangeShapeColor
End Sub
Sub CheckNumeratorAndChangeShapeColor()
    Dim inputSheet As Worksheet
    Dim cellContent As String
    Dim fraction As String
    Dim splitFraction() As String
    Dim numerator As Long
    Dim denominator As Long
    Dim shapeToChange As Shape
    Dim startPos As Long
    Dim darkGreenAccent3 As Long

    ' Define the sheet
    Set inputSheet = ThisWorkbook.Sheets("Input")
    
    ' Define the RGB value for dark green accent 3
    darkGreenAccent3 = RGB(0, 97, 0)
    
    ' Get the content from cell B22
    cellContent = inputSheet.Range("B22").value
    
    ' Find the position of "Input No."
    startPos = InStr(cellContent, "Input No.")
    
    ' Check if "Input No." is found in the cell content
    If startPos > 0 Then
        ' Extract the fraction after "Input No."
        fraction = Trim(Mid(cellContent, startPos + Len("Input No.")))
        
        ' Split the fraction to get numerator and denominator
        splitFraction = Split(fraction, "/")
        
        ' Check if the fraction is in the correct format
        If UBound(splitFraction) = 1 And IsNumeric(splitFraction(0)) And IsNumeric(splitFraction(1)) Then
            numerator = CLng(splitFraction(0))
            denominator = CLng(splitFraction(1))
            
            ' Get the shape
            Set shapeToChange = inputSheet.Shapes("backwards")
            
            ' Check if the numerator is equal to 1
            If numerator = 1 Then
                ' Change the fill color of the shape to red
                shapeToChange.FILL.ForeColor.RGB = RGB(255, 0, 0)
            Else
                ' Change the fill color of the shape to dark green accent 3
                shapeToChange.FILL.ForeColor.RGB = darkGreenAccent3
            End If
        Else
            MsgBox "The fraction is not in the correct format 'X/Y' or is not numeric.", vbExclamation
        End If
    Else
        MsgBox "'Input No.' not found in cell B22.", vbExclamation
    End If
End Sub


Sub UpdateNCRData()
    Dim wsInput As Worksheet
    Dim wsNCRData As Worksheet
    Dim fractionValue As String
    Dim numerator As Long
    Dim lastRow As Long
    Dim i As Long
    Dim parts() As String
    Dim fractionParts() As String
    
    ' Set references to the sheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsNCRData = ThisWorkbook.Sheets("NCR Data")
    
    ' Get the fraction value from cell B22 on the Input sheet
    fractionValue = wsInput.Range("B22").value
    
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
    Set wsInput = Nothing
    Set wsNCRData = Nothing
    
    Call UFillSequentialNumbersNCR
End Sub

Sub UFillSequentialNumbersNCR()
    Dim wsNCRData As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Set reference to the NCR Data sheet
    Set wsNCRData = ThisWorkbook.Sheets("NCR Data")

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
End Sub


Sub UpdateReworkData()
    Dim wsInput As Worksheet
    Dim wsNCRData As Worksheet
    Dim fractionValue As String
    Dim numerator As Long
    Dim lastRow As Long
    Dim i As Long
    Dim parts() As String
    Dim fractionParts() As String
    
    ' Set references to the sheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsNCRData = ThisWorkbook.Sheets("Rework Data")
    
    ' Get the fraction value from cell B22 on the Input sheet
    fractionValue = wsInput.Range("B22").value
    
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
    Set wsInput = Nothing
    Set wsNCRData = Nothing
    
    Call UFillSequentialNumbersRework
End Sub

Sub UFillSequentialNumbersRework()
    Dim wsNCRData As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Set reference to the NCR Data sheet
    Set wsNCRData = ThisWorkbook.Sheets("Rework Data")

    ' Get the last row in column A
    lastRow = wsNCRData.Cells(wsNCRData.Rows.Count, "A").End(xlUp).Row
    
    ' Fill column E with sequential numbers starting from 1 in row 2
    For i = 2 To lastRow
        If wsNCRData.Cells(i, "A").value <> "" Then
            wsNCRData.Cells(i, "D").value = i - 1
        Else
            Exit For
        End If
    Next i

    ' Clean up
    Set wsNCRData = Nothing
End Sub

Sub UpdateResponseData()
    Dim wsInput As Worksheet
    Dim wsNCRData As Worksheet
    Dim fractionValue As String
    Dim numerator As Long
    Dim lastRow As Long
    Dim i As Long
    Dim parts() As String
    Dim fractionParts() As String
    
    ' Set references to the sheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsNCRData = ThisWorkbook.Sheets("Response Data")
    
    ' Get the fraction value from cell B22 on the Input sheet
    fractionValue = wsInput.Range("B22").value
    
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
    Set wsInput = Nothing
    Set wsNCRData = Nothing
    
    Call UFillSequentialNumbersResponse
End Sub
Sub UFillSequentialNumbersResponse()
    Dim wsNCRData As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Set reference to the NCR Data sheet
    Set wsNCRData = ThisWorkbook.Sheets("Response Data")

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
End Sub
Sub FILL()
Call UFillSequentialNumbersNCR
Call UFillSequentialNumbersRework
Call UFillSequentialNumbersResponse
End Sub



Sub NEWDELETE()
Call ClearFiltersInput
Call UpdateNCRData
Call UpdateReworkData
Call UpdateResponseData



Call ARROWBACK
Call SubtractOneFromDenominator
Call CheckDenominatorMinusNumerator
End Sub

Sub DisplayRowCountAsFraction()
    Dim wsInput As Worksheet
    Dim wsNCRData As Worksheet
    Dim lastRow As Long
    Dim countRows As Long
    Dim i As Long
    
    ' Set references to the sheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsNCRData = ThisWorkbook.Sheets("NCR Data")
    
    ' Get the last row in column A of NCR Data sheet
    lastRow = wsNCRData.Cells(wsNCRData.Rows.Count, "A").End(xlUp).Row
    
    ' Count the number of rows with values in column A
    countRows = 0
    For i = 1 To lastRow
        If wsNCRData.Cells(i, "A").value <> "" Then
            countRows = countRows + 1
        End If
    Next i
    
    ' Display the row count as a fraction in cell B22 on the Input sheet
    wsInput.Range("B22").value = "Input No. " & countRows - 1 & "/" & countRows - 1
    
    ' Clean up
    Set wsInput = Nothing
    Set wsNCRData = Nothing
Call FILL
End Sub
Sub SubtractOneFromDenominator()
    Dim wsInput As Worksheet
    Dim fractionValue As String
    Dim numerator As Long
    Dim denominator As Long
    Dim parts() As String
    Dim fractionParts() As String
    
    ' Set reference to the Input sheet
    Set wsInput = ThisWorkbook.Sheets("Input")
    
    ' Get the fraction value from cell B22 on the Input sheet
    fractionValue = wsInput.Range("B22").value
    
    ' Extract the numerator and denominator from the fraction
    On Error GoTo ErrorHandler
    parts = Split(fractionValue, " ")
    fractionParts = Split(parts(2), "/")
    numerator = CLng(fractionParts(0))
    denominator = CLng(fractionParts(1))
    
    ' Subtract one from the denominator
    If denominator > 1 Then
        denominator = denominator - 1
    Else
        MsgBox "Denominator cannot be less than 1.", vbExclamation
        Exit Sub
    End If
    
    ' Update the fraction in cell B22
    wsInput.Range("B22").value = "Input No. " & numerator & "/" & denominator
    
    ' Clean up
    Set wsInput = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error parsing fraction in cell B22. Please check the format.", vbCritical
End Sub

