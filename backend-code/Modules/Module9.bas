Attribute VB_Name = "Module9"
Sub TallyOnTimeStatusForCompanies1()
    Dim poDataSheet As Worksheet
    Dim poDataOutputSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim companyName As String
    Dim companyStatus As String
    Dim uniqueCompanies As Object
    Dim outputRow As Long
    Dim key As Variant ' Variable to hold each key in the dictionary
    Dim visibleCells As Range
    Dim cell As Range

    ' Set the worksheets
    Set poDataSheet = ThisWorkbook.Sheets("PO Data")
    Set poDataOutputSheet = ThisWorkbook.Sheets("PO DataOutput")
    
    ' Initialize the dictionary
    Set uniqueCompanies = CreateObject("Scripting.Dictionary")
    
    ' Find the last row of data in column B
    lastRow = poDataSheet.Cells(poDataSheet.Rows.Count, "B").End(xlUp).Row
    
    ' Clear existing data in columns A and B of PO DataOutput sheet
    poDataOutputSheet.Range("A:B").ClearContents
    
    ' Get the visible cells in column B
    On Error Resume Next
    Set visibleCells = poDataSheet.Range("B1:B" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    ' Check if there are any visible cells
    If Not visibleCells Is Nothing Then
        ' Loop through each visible cell to tally "On-Time" statuses
        For Each cell In visibleCells
            companyName = cell.value
            companyStatus = cell.Offset(0, 3).value ' Column E is 3 columns to the right of column B
            
            ' Ensure company name is not empty
            If Not IsEmpty(companyName) And companyName <> "" Then
                ' Check if company name is already in the dictionary
                If Not uniqueCompanies.Exists(companyName) Then
                    ' Add company name to dictionary with initial count
                    uniqueCompanies.Add companyName, 0
                End If
                
                ' Check if status is "On-Time" and increment count
                If companyStatus = "On-Time" Then
                    uniqueCompanies(companyName) = uniqueCompanies(companyName) + 1
                End If
            End If
        Next cell
        
        ' Output results to columns A and B in PO DataOutput sheet
        outputRow = 1
        For Each key In uniqueCompanies.Keys
            poDataOutputSheet.Cells(outputRow, "A").value = key ' Use key to access each company name
            poDataOutputSheet.Cells(outputRow, "B").value = uniqueCompanies(key) ' Use key to access the count
            outputRow = outputRow + 1
        Next key
    End If
    
    ' Autofit columns A and B in PO DataOutput sheet
    poDataOutputSheet.Columns("A:B").AutoFit
End Sub


Sub TallyCompanyOccurrences1()
    Dim poDataSheet As Worksheet
    Dim poDataOutputSheet As Worksheet
    Dim lastRow As Long
    Dim companyName As String
    Dim outputRow As Long
    Dim uniqueCompanies As Object
    Dim visibleCells As Range
    Dim cell As Range
    Dim companyStatus As String
    Dim key As Variant ' Variable to hold each key in the dictionary
    
    ' Set the worksheets
    Set poDataSheet = ThisWorkbook.Sheets("PO Data")
    Set poDataOutputSheet = ThisWorkbook.Sheets("PO DataOutput")
    
    ' Initialize the dictionary
    Set uniqueCompanies = CreateObject("Scripting.Dictionary")
    
    ' Find the last row of data in column B of the PO Data sheet
    lastRow = poDataSheet.Cells(poDataSheet.Rows.Count, "B").End(xlUp).Row
    
    ' Clear existing data in columns A and C of PO DataOutput sheet
    poDataOutputSheet.Range("C:C").ClearContents
    
    ' Get the visible cells in column B
    On Error Resume Next
    Set visibleCells = poDataSheet.Range("B1:B" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    ' Check if there are any visible cells
    If Not visibleCells Is Nothing Then
        ' Loop through each visible cell to tally company occurrences
        For Each cell In visibleCells
            companyName = cell.value
            companyStatus = cell.Offset(0, 3).value ' Column E is 3 columns to the right of column B
            
            ' Ensure company name is not empty and the status is not "Invalid Date"
            If Not IsEmpty(companyName) And companyName <> "" And companyStatus <> "Invalid Date" Then
                ' Check if company name is already in the dictionary
                If Not uniqueCompanies.Exists(companyName) Then
                    ' Add company name to dictionary with initial occurrences count
                    uniqueCompanies.Add companyName, 1
                Else
                    ' Increment occurrences count
                    uniqueCompanies(companyName) = uniqueCompanies(companyName) + 1
                End If
            End If
        Next cell
        
        ' Output results to columns A and C in PO DataOutput sheet
        outputRow = 1
        For Each key In uniqueCompanies.Keys
            poDataOutputSheet.Cells(outputRow, "A").value = key ' Use key to access each company name
            poDataOutputSheet.Cells(outputRow, "C").value = uniqueCompanies(key) ' Use key to access the occurrences count
            outputRow = outputRow + 1
        Next key
        
        ' Autofit columns A and C in "PO DataOutput" sheet
        poDataOutputSheet.Columns("A:C").AutoFit
    End If
End Sub


Sub SetVendorPrompt()
    Dim wsMaster As Worksheet
    
    ' Set the worksheet
    Set wsMaster = ThisWorkbook.Sheets("Master Sheet")
    
    ' Set the value of cell A1
    wsMaster.Range("A1").value = "Click here to pick a vendor"
End Sub

Sub SummarizeCompanyData()
    Dim poDataSheet As Worksheet
    Dim poDataOutputSheet As Worksheet
    Dim lastRow As Long
    Dim companyName As String
    Dim companyStatus As String
    Dim uniqueCompanies As Object
    Dim visibleCells As Range
    Dim cell As Range
    Dim outputRow As Long
    Dim key As Variant
    
    ' Set the worksheets
    Set poDataSheet = ThisWorkbook.Sheets("PO Data")
    Set poDataOutputSheet = ThisWorkbook.Sheets("PO DataOutput")
    
    ' Initialize the dictionary to hold unique companies and counts
    Set uniqueCompanies = CreateObject("Scripting.Dictionary")
    
    ' Clear existing data in columns A, B, and C of PO DataOutput sheet
    poDataOutputSheet.Range("A:C").ClearContents
    
    ' Process the visible cells in the "po" table
    With poDataSheet.ListObjects("po").DataBodyRange
        On Error Resume Next
        Set visibleCells = .Columns(2).SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        
        ' Check if there are any visible cells
        If Not visibleCells Is Nothing Then
            ' Loop through each visible cell to build the unique company list and count occurrences
            For Each cell In visibleCells
                companyName = cell.value
                companyStatus = cell.Offset(0, 3).value ' Column E is 4 columns to the right of column B (1-based index)
                
                ' Ensure company name is not empty
                If Not IsEmpty(companyName) And companyName <> "" Then
                    ' Check if company name is already in the dictionary
                    If Not uniqueCompanies.Exists(companyName) Then
                        ' Add company name to dictionary with initial counts
                        uniqueCompanies.Add companyName, Array(0, 0) ' Array(On-Time Count, Total Count)
                    End If
                    
                    ' Count the "On-Time" status
                    If companyStatus = "On-Time" Then
                        uniqueCompanies(companyName)(0) = uniqueCompanies(companyName)(0) + 1
                    End If
                    
                    ' Count the total of "On-Time" and "Late" statuses
                    If companyStatus = "On-Time" Or companyStatus = "Late" Then
                        uniqueCompanies(companyName)(1) = uniqueCompanies(companyName)(1) + 1
                    End If
                End If
            Next cell
            
            ' Output results to columns A, B, and C in PO DataOutput sheet
            outputRow = 1
            For Each key In uniqueCompanies.Keys
                poDataOutputSheet.Cells(outputRow + 1, "A").value = key ' Company name
                poDataOutputSheet.Cells(outputRow + 1, "B").value = uniqueCompanies(key)(0) ' On-Time count
                poDataOutputSheet.Cells(outputRow + 1, "C").value = uniqueCompanies(key)(1) ' Total count (On-Time + Late)
                outputRow = outputRow + 1
            Next key
            
            ' Autofit columns A, B, and C in PO DataOutput sheet
            poDataOutputSheet.Columns("A:C").AutoFit
        Else
            MsgBox "No visible data found to process.", vbExclamation
        End If
    End With
End Sub

Sub SumOnTimeForVisibleCompanies()
    Dim poDataSheet As Worksheet
    Dim poDataOutputSheet As Worksheet
    Dim lastRow As Long
    Dim lastOutputRow As Long
    Dim companyName As String
    Dim companyStatus As String
    Dim totalOnTime As Long
    Dim outputCell As Range
    Dim visibleCells As Range
    Dim cell As Range
    
    ' Set the worksheets
    Set poDataSheet = ThisWorkbook.Sheets("PO Data")
    Set poDataOutputSheet = ThisWorkbook.Sheets("PO DataOutput")
    
    ' Find the last row in the data sheet (column B)
    lastRow = poDataSheet.Cells(poDataSheet.Rows.Count, "B").End(xlUp).Row
    
    ' Find the last row in the output sheet (column A, starting from row 2)
    lastOutputRow = poDataOutputSheet.Cells(poDataOutputSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Get the visible cells in column B on the "PO Data" sheet
    On Error Resume Next
    Set visibleCells = poDataSheet.Range("B2:B" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    ' Check if there are any visible cells
    If Not visibleCells Is Nothing Then
        ' Loop through each company in the output sheet (starting from row 2)
        For Each outputCell In poDataOutputSheet.Range("A2:A" & lastOutputRow)
            companyName = outputCell.value
            totalOnTime = 0
            
            ' Loop through each visible corresponding value in the data sheet
            For Each cell In visibleCells
                If cell.value = companyName Then
                    companyStatus = cell.Offset(0, 3).value ' Column E is 3 columns to the right of column B
                    If companyStatus = "On-Time" Then
                        totalOnTime = totalOnTime + 1
                    End If
                End If
            Next cell
            
            ' Output the total "On-Time" count in column B next to the company name
            outputCell.Offset(0, 1).value = totalOnTime
        Next outputCell
    Else
        MsgBox "No visible data found to process.", vbExclamation
    End If
End Sub

Sub SumOnTimeAndLateForVisibleCompanies()
    Dim poDataSheet As Worksheet
    Dim poDataOutputSheet As Worksheet
    Dim lastRow As Long
    Dim lastOutputRow As Long
    Dim companyName As String
    Dim companyStatus As String
    Dim totalOnTimeLate As Long
    Dim outputCell As Range
    Dim visibleCells As Range
    Dim cell As Range
    
    ' Set the worksheets
    Set poDataSheet = ThisWorkbook.Sheets("PO Data")
    Set poDataOutputSheet = ThisWorkbook.Sheets("PO DataOutput")
    
    ' Find the last row in the data sheet (column B)
    lastRow = poDataSheet.Cells(poDataSheet.Rows.Count, "B").End(xlUp).Row
    
    ' Find the last row in the output sheet (column A, starting from row 2)
    lastOutputRow = poDataOutputSheet.Cells(poDataOutputSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Get the visible cells in column B on the "PO Data" sheet
    On Error Resume Next
    Set visibleCells = poDataSheet.Range("B2:B" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    ' Check if there are any visible cells
    If Not visibleCells Is Nothing Then
        ' Loop through each company in the output sheet (starting from row 2)
        For Each outputCell In poDataOutputSheet.Range("A2:A" & lastOutputRow)
            companyName = outputCell.value
            totalOnTimeLate = 0
            
            ' Loop through each visible corresponding value in the data sheet
            For Each cell In visibleCells
                If cell.value = companyName Then
                    companyStatus = cell.Offset(0, 3).value ' Column E is 3 columns to the right of column B
                    If companyStatus = "On-Time" Or companyStatus = "Late" Then
                        totalOnTimeLate = totalOnTimeLate + 1
                    End If
                End If
            Next cell
            
            ' Output the total "On-Time" + "Late" count in column C next to the company name
            outputCell.Offset(0, 2).value = totalOnTimeLate
        Next outputCell
    Else
        MsgBox "No visible data found to process.", vbExclamation
    End If
    Call RemoveZeroRows
End Sub

Sub RemoveZeroRows()
    Dim poDataOutputSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Set the worksheet
    Set poDataOutputSheet = ThisWorkbook.Sheets("PO DataOutput")
    
    ' Find the last row with data in column A
    lastRow = poDataOutputSheet.Cells(poDataOutputSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through each row, starting from the last row to the second row
    For i = lastRow To 2 Step -1
        If poDataOutputSheet.Cells(i, "B").value = 0 And poDataOutputSheet.Cells(i, "C").value = 0 Then
            poDataOutputSheet.Rows(i).Delete
        End If
    Next i
End Sub

