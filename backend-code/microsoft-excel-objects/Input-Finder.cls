VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim rng As Range
    
    ' Define the range where you want the dropdown to appear
    Set rng = Me.Range("G6")  ' Updated to cell B6
    
    ' Check if the selected cell is within the defined range
    If Not Intersect(Target, rng) Is Nothing Then
        Application.SendKeys "%{DOWN}"
    End If
End Sub

