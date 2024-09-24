VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Cost of Poor Quality")
    
    ' Check if the selected cell is E7 or E16
    If Not Intersect(Target, ws.Range("E7")) Is Nothing Or Not Intersect(Target, ws.Range("E16")) Is Nothing Then
        ' Simulate a double-click to show the drop-down list
        Application.SendKeys "%{DOWN}"
    End If
End Sub

