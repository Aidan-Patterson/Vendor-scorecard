VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim rngDate As Range, rngSel As Range
    Set rngDate = Range("A4")
    Set rngSel = Selection

    
    Dim rng As Range
    Dim cell As Range
    
    ' Define the range where you want the dropdown to appear
    Set rng = Me.Range("A3:A5")  ' Change this to your desired range
    
    ' Check if the selected cell is within the defined range
    If Not Intersect(Target, rng) Is Nothing Then
        Application.SendKeys ("%{DOWN}")
    End If
End Sub
    

