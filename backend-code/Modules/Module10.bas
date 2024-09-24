Attribute VB_Name = "Module10"
Sub FillZerosInMasterSheet()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim cell As Range
    Dim col As Long

    ' Set the worksheet
    Set ws = Worksheets("Master Sheet")
    
    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through each row in column A
    For Each cell In ws.Range("A2:A" & lastRow)
        If cell.value <> "" Then
            ' Loop through columns B to I in the current row
            For col = 2 To 9 ' Columns B to I
                If IsEmpty(ws.Cells(cell.Row, col).value) Then
                    ws.Cells(cell.Row, col).value = 0
                    ' Ensure the accounting format is maintained if necessary
                    If ws.Cells(cell.Row, col).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)" Then
                        ws.Cells(cell.Row, col).NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
                    End If
                End If
            Next col
        End If
    Next cell
End Sub

Sub SetColumnsToGeneral()
    Dim ws As Worksheet
    Dim col As Long

    ' Set the worksheet
    Set ws = Worksheets("Master Sheet") ' Change this to your specific worksheet name if needed

    ' Loop through columns A to I and set the NumberFormat to "General"
    For col = 1 To 9 ' Columns A to I
        ws.Columns(col).NumberFormat = "General"
    Next col
End Sub

Sub FilterByQuarter2()
    Dim wsPrintout As Worksheet
    Dim wsPOData As Worksheet
    Dim quarter As String
    Dim lastRow As Long

    ' Set the worksheets
    Set wsPrintout = Worksheets("Printout")
    Set wsPOData = Worksheets("PO Data")

    ' Get the quarter value from cell A5 on the "Printout" sheet
    quarter = wsPrintout.Range("A5").value

    ' Determine the last row with data in column C on the "PO Data" sheet
    lastRow = wsPOData.Cells(wsPOData.Rows.Count, "C").End(xlUp).Row

    ' Remove any existing filters
    If wsPOData.AutoFilterMode Then
        wsPOData.AutoFilterMode = False
    End If

    ' Apply the filter to column C based on the quarter value
    With wsPOData.Range("A1:C" & lastRow)
        .AutoFilter Field:=3, Criteria1:=quarter
    End With
End Sub


