VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "UserForm5"
   ClientHeight    =   3012
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm5.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Activate()
    ' Display the current values from the "Input" sheet in the text boxes
    Me.txtReworkHours.value = ThisWorkbook.Sheets("Input").Range("J27").value
    Me.txtReworkWorkers.value = ThisWorkbook.Sheets("Input").Range("K27").value
End Sub

Private Sub cmdOK_Click()
    ' When OK is clicked, update the values on the "Input" sheet
    ThisWorkbook.Sheets("Input").Range("J27").value = Me.txtReworkHours.value
    ThisWorkbook.Sheets("Input").Range("K27").value = Me.txtReworkWorkers.value
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    ' When Cancel is clicked, just close the form without saving changes
    Me.Hide
End Sub

