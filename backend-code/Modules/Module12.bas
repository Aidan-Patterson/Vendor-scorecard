Attribute VB_Name = "Module12"
Sub SummarizeCompanyValues()
    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    Dim dict As Object
    Dim lastRow As Long
    Dim dataArray As Variant
    Dim company As Variant
    Dim value As Double
    Dim visibleCells As Range
    Dim cell As Range
    Dim i As Long

    ' Set worksheets
    Set wsSource = Worksheets("Rework Data")
    Set wsOutput = Worksheets("Rework DataOutput")
    
    ' Clear previous output in columns A and B only
    wsOutput.Columns("A:B").ClearContents

    ' Create a dictionary to store company values
    Set dict = CreateObject("Scripting.Dictionary")

    ' Find the last row in the source sheet
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    ' Get the range of visible cells in columns A and C
    On Error Resume Next ' In case there are no visible cells
    Set visibleCells = wsSource.Range("A2:C" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    ' Check if there are any visible cells
    If Not visibleCells Is Nothing Then
        ' Loop through the visible cells in the array
        For Each cell In visibleCells.Columns(1).Cells
            company = cell.value
            value = cell.Offset(0, 2).value

            If dict.Exists(company) Then
                dict(company) = dict(company) + value
            Else
                dict.Add company, value
            End If
        Next cell

        ' Output the results to the output sheet
        i = 2 ' Start output in row 2
        For Each company In dict.Keys
            wsOutput.Cells(i, 1).value = company
            wsOutput.Cells(i, 2).value = dict(company)
            i = i + 1
        Next company

        ' Autofit the columns
        wsOutput.Columns("A:B").AutoFit
    End If

    ' Clean up
    Set dict = Nothing
    Set wsSource = Nothing
    Set wsOutput = Nothing
End Sub



Sub TotalCompanyValues()
    Dim ws As Worksheet
    Dim dict As Object
    Dim lastRowA As Long
    Dim lastRowF As Long
    Dim i As Long
    Dim company As String
    Dim value As Double
    Dim visibleCellsF As Range
    Dim cell As Range

    ' Set the worksheet
    Set ws = Worksheets("Rework DataOutput")
    
    ' Create a dictionary to store company values
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Find the last row in columns A and F
    lastRowA = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastRowF = ws.Cells(ws.Rows.Count, "F").End(xlUp).Row
    
    ' Clear column C before outputting results
    ws.Columns("C").ClearContents
    
    ' Get the visible cells in columns F and G
    On Error Resume Next ' In case there are no visible cells
    Set visibleCellsF = ws.Range("F2:F" & lastRowF).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    ' Loop through the visible cells in columns F and G
    If Not visibleCellsF Is Nothing Then
        For Each cell In visibleCellsF
            company = cell.value
            value = cell.Offset(0, 1).value ' Column G
            
            If dict.Exists(company) Then
                dict(company) = dict(company) + value
            Else
                dict.Add company, value
            End If
        Next cell
    End If
    
    ' Loop through the company names in column A and output the sums in column C
    For i = 2 To lastRowA
        company = ws.Cells(i, 1).value
        
        If dict.Exists(company) Then
            ws.Cells(i, 3).value = dict(company)
        Else
            ws.Cells(i, 3).value = 0
        End If
    Next i
    
    ' Clean up
    Set dict = Nothing
    Set ws = Nothing
End Sub


Sub CreateTablePOData()
    Dim ws As Worksheet
    Dim tblRange As Range
    Dim tbl As ListObject

    ' Set the worksheet
    Set ws = Worksheets("PO Data")

    ' Define the range for the table (columns A through E, starting from row 1)
    Set tblRange = ws.Range("A1:B" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)

    ' Check if the table already exists and delete it if it does
    On Error Resume Next
    Set tbl = ws.ListObjects("po")
    If Not tbl Is Nothing Then
        tbl.Delete
    End If
    On Error GoTo 0

    ' Create the table
    Set tbl = ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    tbl.name = "po"

    ' Format the table (optional)
    tbl.TableStyle = "TableStyleMedium2" ' You can change the table style as needed

    ' Autofit the columns
    ws.Columns("A:B").AutoFit
End Sub

