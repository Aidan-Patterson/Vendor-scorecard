Attribute VB_Name = "Module3"
Sub CopyVisibleIDsCompanyNamesAndAdditionalColumn()
    Dim datapSheet As Worksheet
    Dim poDataSheet As Worksheet
    Dim lastRow As Long
    Dim visibleRange As Range
    
    ' Set the worksheets
    Set datapSheet = ThisWorkbook.Sheets("datar")
    Set poDataSheet = ThisWorkbook.Sheets("PO Data")
    
    ' Find the last row of data in column A of "datap" sheet
    lastRow = datapSheet.Cells(datapSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Clear existing data in "PO Data" sheet
    poDataSheet.Cells.ClearContents
    
    ' Check if there are any visible cells in column A
    On Error Resume Next
    Set visibleRange = datapSheet.Range("B1:B" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not visibleRange Is Nothing Then
        ' Copy visible IDs from "datap" column A to "PO Data" column A as values
        visibleRange.Copy
        poDataSheet.Range("A1").PasteSpecial Paste:=xlPasteValues
        
        
        
        ' Copy visible additional data from "datap" column E to "PO Data" column C as values
        Set visibleRange = datapSheet.Range("D1:E" & lastRow).SpecialCells(xlCellTypeVisible)
        visibleRange.Copy
        poDataSheet.Range("D1").PasteSpecial Paste:=xlPasteValues
        
        ' Autofit columns A, B, and C in "PO Data" sheet
        poDataSheet.Columns("A:D").AutoFit
        
        ' Clear the clipboard
        Application.CutCopyMode = False
    End If
End Sub


Sub MatchIDsAndOutputNameDate()
    Dim poDataSheet As Worksheet
    Dim datapSheet As Worksheet
    Dim lastRowPOData As Long
    Dim lastRowDatap As Long
    Dim i As Long
    Dim id As Variant
    Dim dict As Object
    Dim name As Variant
    Dim dateValue As Variant
    
    ' Set the worksheets
    Set poDataSheet = ThisWorkbook.Sheets("PO Data")
    Set datapSheet = ThisWorkbook.Sheets("datap")
    
    ' Find the last rows of data in columns A for both sheets
    lastRowPOData = poDataSheet.Cells(poDataSheet.Rows.Count, "A").End(xlUp).Row
    lastRowDatap = datapSheet.Cells(datapSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Create a dictionary for ID to name and date mapping
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Populate the dictionary with ID, name, and date from "datap" sheet
    For i = 1 To lastRowDatap
        id = datapSheet.Cells(i, "A").value
        If Not dict.Exists(id) Then
            name = datapSheet.Cells(i, "B").value
            dateValue = datapSheet.Cells(i, "E").value
            dict.Add id, Array(name, dateValue)
        End If
    Next i
    
    ' Clear existing data in columns B and C on "PO Data" sheet
    poDataSheet.Range("B:C").ClearContents
    
    ' Loop through each ID in column A on "PO Data" sheet and output corresponding name and date
    For i = 1 To lastRowPOData
        id = poDataSheet.Cells(i, "A").value
        If dict.Exists(id) Then
            poDataSheet.Cells(i, "B").value = dict(id)(0) ' Output name from dictionary
            poDataSheet.Cells(i, "C").value = dict(id)(1) ' Output date from dictionary
        End If
    Next i
    
    ' Autofit columns B and C in "PO Data" sheet
    poDataSheet.Columns("B:C").AutoFit
End Sub

Sub CompareDatesAndOutputStatus()
    Dim poDataSheet As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Set the worksheet
    Set poDataSheet = ThisWorkbook.Sheets("PO Data")
    
    ' Set column D to short date format
    poDataSheet.Columns("D").NumberFormat = "m/d/yyyy"
    
    ' Find the last row of data in column C
    lastRow = poDataSheet.Cells(poDataSheet.Rows.Count, "C").End(xlUp).Row
    
    ' Loop through each row and compare the dates
    For i = 1 To lastRow
        If IsDate(poDataSheet.Cells(i, "C").value) And IsDate(poDataSheet.Cells(i, "D").value) Then
            If poDataSheet.Cells(i, "D").value <= poDataSheet.Cells(i, "C").value Then
                poDataSheet.Cells(i, "E").value = "On-Time"
            Else
                poDataSheet.Cells(i, "E").value = "Late"
            End If
        Else
            poDataSheet.Cells(i, "E").value = "Invalid Date"
        End If
    Next i
    
    ' Autofit column E in "PO Data" sheet
    poDataSheet.Columns("E").AutoFit
End Sub







