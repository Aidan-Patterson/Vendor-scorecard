Attribute VB_Name = "Module13"
Sub MatchAndTransferData()

    Dim wsPrintout As Worksheet
    Dim wsMaster As Worksheet
    Dim wsOutput As Worksheet
    Dim companyName As String
    Dim matchRow As Long
    Dim i As Integer
    
    ' Set the worksheets
    Set wsPrintout = Worksheets("Printout")
    Set wsMaster = Worksheets("Master Sheet")
    Set wsOutput = Worksheets("Output")
    
    ' Get the company name from cell A3 on the Printout sheet
    companyName = wsPrintout.Range("A3").value
    
    ' Find the matching company name in column A on the Master Sheet
    matchRow = 0
    On Error Resume Next
    matchRow = wsMaster.Columns("A").Find(What:=companyName, LookIn:=xlValues, LookAt:=xlWhole).Row
    On Error GoTo 0
    
    ' If a match is found, transfer the data
    If matchRow > 0 Then
        ' Clear the Output sheet first
        wsOutput.Cells.Clear
        
        ' Transfer the company name and corresponding values
        For i = 0 To 8 ' Columns A to I (0 to 8 as offset)
            wsOutput.Cells(1, i + 1).value = wsMaster.Cells(matchRow, i + 1).value
        Next i
        
        ' Set columns F and G to accounting number format
        
    Else
        MsgBox "Please select a company", vbExclamation
    End If
    Call CheckAndReplaceDiv0
    Call selectiongrade
    Call CalculateL4
    

    Call ColorN4BasedOnPercentageRange
End Sub

Sub CheckAndReplaceDiv0()
    Dim wsQuality As Worksheet
    Dim numerator As Double
    Dim denominator As Double
    Dim result As Double
    
    ' Set the worksheet
    Set wsQuality = ThisWorkbook.Sheets("Quality")
    
    ' Get the values from G3 and H3
    numerator = wsQuality.Range("G3").value
    denominator = wsQuality.Range("H3").value
    
    ' Check for division by zero
    If denominator = 0 Then
        wsQuality.Range("I3").value = 0
    Else
        ' Perform the division and multiply by 100
        result = (numerator / denominator) * 100
        wsQuality.Range("I3").value = result
    End If
End Sub


Sub SetPrintoutZoom()
    Dim wsPrintout As Worksheet
    
    ' Set the worksheet
    Set wsPrintout = ThisWorkbook.Sheets("Printout")
    
    ' Set the zoom level to 70%
    With wsPrintout
        .Activate
        ActiveWindow.Zoom = 70
    End With
End Sub

Sub SetVendorPromptInA4()
    Dim ws As Worksheet
    
    ' Set the worksheet (assuming you want to use the "Printout" sheet)
    Set ws = ThisWorkbook.Sheets("Printout")
    
    ' Set the value of cell A4
    ws.Range("A4").value = "Click to choose a month"
End Sub

Sub SetVendorPromptInA5()
    Dim ws As Worksheet
    
    ' Set the worksheet (assuming you want to use the "Printout" sheet)
    Set ws = ThisWorkbook.Sheets("Printout")
    
    ' Set the value of cell A4
    ws.Range("A5").value = "Click to choose a quarter"
End Sub

