VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    
    SetTextValues
    GoToPrintoutSheet
    
    
    SetTextValues2
    Call SetPrintoutSheetMessages
    
    Call ClearTroubleshootingCells
    Call ClearAndSetValues2
    Call clearcellsquality
    Call SetCellC16
    Call SetCellC7
    Call CalculateL4
End Sub
Sub InputFinderMessage()
    Dim wsInputFinder As Worksheet
    Set wsInputFinder = ThisWorkbook.Sheets("Input Finder")
    wsInputFinder.Range("G6").value = "Click here to pick a vendor"
    
End Sub




Sub SetPrintoutSheetMessages()
    Dim wsPrintout As Worksheet
    Set wsPrintout = ThisWorkbook.Sheets("Printout")
    
    ' Set the messages in the specified cells
    wsPrintout.Range("A3").value = "Click to choose a vendor"
    wsPrintout.Range("A4").value = "Click to choose a month"
    wsPrintout.Range("A5").value = "Click to choose a quarter"
    
    ' Select the Printout sheet
    wsPrintout.Select
End Sub


Sub ClearAndSetValues2()
    Dim wsInput As Worksheet
    Dim wsCost As Worksheet
    
    ' Set references to the worksheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsCost = ThisWorkbook.Sheets("Cost of Poor Quality")
    
    ' Clear cells A3, A4, and A5 on the "Input" sheet
    wsInput.Range("A3, A4, A5").ClearContents
    
    ' Clear cells A3, A4, and A5 on the "Cost of Poor Quality" sheet
    wsCost.Range("A3, A4, A5").ClearContents
    
    ' Set cell E7 on "Cost of Poor Quality" sheet
    wsCost.Range("E7").value = "Click here to choose a month"
    
    ' Set cell E16 on "Cost of Poor Quality" sheet
    wsCost.Range("E16").value = "Click here to choose a quarter"
End Sub


Sub CalculateL4()
    Dim ws As Worksheet
    Dim D5Value As Double
    Dim F5Value As Double
    Dim H5Value As Double
    Dim result As Double

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Printout")

    ' Get the values from D5, F5, and H5
    D5Value = ws.Range("D5").value
    F5Value = ws.Range("F5").value
    H5Value = ws.Range("H5").value

    ' Calculate the result
    result = 0.4 * D5Value + 0.4 * F5Value + 0.2 * H5Value

    ' Set the value of L4
    ws.Range("L4").value = result
End Sub

Private Sub Workbook_BeforePrint(Cancel As Boolean)
    Dim ws As Worksheet
    Dim hiddenSheet As Worksheet
    Dim finalizeShape As Shape
    
    Set ws = ThisWorkbook.Sheets("Printout")
    Set hiddenSheet = ThisWorkbook.Sheets("HiddenSheet")
    Set finalizeShape = ws.Shapes("finalize")
    
    ' Store original font colors
    hiddenSheet.Range("A1").value = ws.Range("A4").Font.Color
    hiddenSheet.Range("A2").value = ws.Range("A5").Font.Color
    hiddenSheet.Range("A3").value = ws.Range("A7").Font.Color
    hiddenSheet.Range("A4").value = ws.Range("A9").Font.Color
    
    ' Make specific cells invisible by setting font color to white
    ws.Range("A7").Font.Color = RGB(255, 255, 255)
    ws.Range("A9").Font.Color = RGB(255, 255, 255)
    
    ' Conditionally make A4 invisible
    If ws.Range("A4").value = "Click to choose a month" Then
        ws.Range("A4").Font.Color = RGB(255, 255, 255)
    End If
    
    ' Conditionally make A5 invisible
    If ws.Range("A5").value = "Click to choose a quarter" Then
        ws.Range("A5").Font.Color = RGB(255, 255, 255)
    End If
    
    ' Hide the shape named "finalize"
    finalizeShape.Visible = msoFalse
    
    ' Wait for the printing to complete
    Application.OnTime Now + TimeValue("00:00:01"), "RestoreColors"
End Sub





