Attribute VB_Name = "Module17"
Sub deletecellD7()
Attribute deletecellD7.VB_ProcData.VB_Invoke_Func = " \n14"
'
' deletecellD7 Macro
'

'
    Range("D7").Select
    Selection.ClearContents
    Range("C11").Select
End Sub
Sub clearcellsquality()
Attribute clearcellsquality.VB_ProcData.VB_Invoke_Func = " \n14"
'
' clearcellsquality Macro
'

'
    Selection.ClearContents
    Range("P11").Select
    Selection.ClearContents
    Range("P8").Select
    Selection.ClearContents
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

Sub RestoreColors()
    Dim ws As Worksheet
    Dim hiddenSheet As Worksheet
    Dim finalizeShape As Shape
    
    Set ws = ThisWorkbook.Sheets("Printout")
    Set hiddenSheet = ThisWorkbook.Sheets("HiddenSheet")
    Set finalizeShape = ws.Shapes("finalize")
    
    ' Restore original font colors unconditionally
    ws.Range("A4").Font.Color = hiddenSheet.Range("A1").value
    ws.Range("A5").Font.Color = hiddenSheet.Range("A2").value
    ws.Range("A7").Font.Color = hiddenSheet.Range("A3").value
    ws.Range("A9").Font.Color = hiddenSheet.Range("A4").value
    
    ' Show the shape named "finalize"
    finalizeShape.Visible = msoTrue
End Sub

