Attribute VB_Name = "Module22"
Sub UpdateYearTo2024()
    Dim wsDatar As Worksheet
    Dim wsDatap As Worksheet
    Dim lastRowDatar As Long
    Dim lastRowDatap As Long
    Dim i As Long
    Dim dateValue As Date

    ' Set worksheets
    Set wsDatar = ThisWorkbook.Sheets("datar")
    Set wsDatap = ThisWorkbook.Sheets("datap")

    ' Find last row in each sheet
    lastRowDatar = wsDatar.Cells(wsDatar.Rows.Count, "C").End(xlUp).Row
    lastRowDatap = wsDatap.Cells(wsDatap.Rows.Count, "E").End(xlUp).Row

    ' Update year to 2024 in datar sheet (columns C, D, I)
    For i = 2 To lastRowDatar ' Assuming headers in row 1
        If IsDate(wsDatar.Cells(i, "C").value) Then
            dateValue = wsDatar.Cells(i, "C").value
            wsDatar.Cells(i, "C").value = DateSerial(2024, Month(dateValue), Day(dateValue))
        End If
        
        If IsDate(wsDatar.Cells(i, "D").value) Then
            dateValue = wsDatar.Cells(i, "D").value
            wsDatar.Cells(i, "D").value = DateSerial(2024, Month(dateValue), Day(dateValue))
        End If
        
        If IsDate(wsDatar.Cells(i, "I").value) Then
            dateValue = wsDatar.Cells(i, "I").value
            wsDatar.Cells(i, "I").value = DateSerial(2024, Month(dateValue), Day(dateValue))
        End If
    Next i

    ' Update year to 2024 in datap sheet (columns E, F)
    For i = 2 To lastRowDatap ' Assuming headers in row 1
        If IsDate(wsDatap.Cells(i, "E").value) Then
            dateValue = wsDatap.Cells(i, "E").value
            wsDatap.Cells(i, "E").value = DateSerial(2024, Month(dateValue), Day(dateValue))
        End If
        
        If IsDate(wsDatap.Cells(i, "F").value) Then
            dateValue = wsDatap.Cells(i, "F").value
            wsDatap.Cells(i, "F").value = DateSerial(2024, Month(dateValue), Day(dateValue))
        End If
    Next i
End Sub

