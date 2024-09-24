VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim rngDate As Range, rngSel As Range
    Set rngDate = Range("D7")
    Set rngSel = Selection
    If Intersect(rngDate, rngSel) Is Nothing Then
        Shapes("calendar1").Visible = False
    Else
        Shapes("calendar1").Visible = True
        Shapes("calendar1").Top = ActiveCell.Top + ActiveCell.Height
        Shapes("calendar1").Left = ActiveCell.Left + ActiveCell.Width
    End If
    
    Dim rng As Range
    Dim cell As Range
    
    ' Define the range where you want the dropdown to appear
    Set rng = Me.Range("B7")  ' Change this to your desired range
    
    ' Check if the selected cell is within the defined range
    If Not Intersect(Target, rng) Is Nothing Then
        Application.SendKeys ("%{DOWN}")
    End If
End Sub
