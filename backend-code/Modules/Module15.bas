Attribute VB_Name = "Module15"
Sub ShowPleaseWaitMessageBox()
    MsgBox "Please wait; estimated time ~ 15 seconds", vbInformation, "Processing"
End Sub

Sub MainMacro1()
    ' Show the UserForm with the message
    frmMessage.Show vbModeless
    frmMessage.Repaint
    
    ' Call the other macro
    Call FilterByQuarterMaster
    Call SetCellA3
    
    ' Hide the UserForm after the other macro is done
    Unload frmMessage
End Sub

Sub MainMacro2()
    ' Show the UserForm with the message
    frmMessage.Show vbModeless
    frmMessage.Repaint
    
    ' Call the other macro
    Call FilterTablesByMonthAndYearMaster
    Call SetCellA3
    
    ' Hide the UserForm after the other macro is done
    Unload frmMessage
End Sub

Sub ClearTroubleshootingCells()
    Dim ws As Worksheet

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Troubleshooting")

    ' Clear cells A3, A4, and A5
    ws.Range("A3:A5").ClearContents
End Sub

Sub SetCellA3()
    ' Set the value of cell C16 to "Click here to pick a quarter"
    ThisWorkbook.Sheets("Printout").Range("A3").value = "Click here to pick a vendor"
End Sub

