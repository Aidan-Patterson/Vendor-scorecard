Attribute VB_Name = "Module4"
Sub changetitle()
Attribute changetitle.VB_ProcData.VB_Invoke_Func = " \n14"
'
' changetitle Macro
'

'
    Range("C1").Select
    Selection.Copy
    Range("A1").Select
    ActiveSheet.Paste
End Sub
Sub Submit()
Call ClearFiltersInput

Call RecordNCRCompanyData
Call EnterReworkData
Call OutputResponseData
Call deletecellD7


Call PREVIOUS
Call DisplayRowCountAsFraction
Call CheckDenominatorMinusNumerator
End Sub

Sub CheckVendorScorecardTest()
    Dim wb As Workbook
    Dim isWorkbookOpen As Boolean
    
    ' Check if "Vendor Scorecard TEST" is open
    isWorkbookOpen = False
    For Each wb In Workbooks
        If wb.name = "Vendor Scorecard TEST.xlsm" Then
            isWorkbookOpen = True
            Exit For
        End If
    Next wb
    
    ' If the workbook is open, proceed and call your specific macro
    If isWorkbookOpen Then
        ' Place your specific macro call here
        ' Example: Call YourMacroName
        Call Submit ' Replace "YourMacroName" with the actual name of your macro
    Else
        ' If the workbook is not open, display a message
        MsgBox "Please open Excel workbook 'Vendor Scorecard TEST.xlsm' to proceed.", vbExclamation
    End If
End Sub

Sub RecordNCRCompanyData()
    Dim inputSheet As Worksheet
    Dim ncrSheet As Worksheet
    Dim otherWorkbook As Workbook
    Dim otherSheet As Worksheet
    Dim companyName As String
    Dim dateValue As Date
    Dim ncheckValue As String
    Dim ocheckValue As String
    Dim nextRow As Long
    Dim nextRowOther As Long
    Dim workbookName As String
    Dim otherWorkbookWasOpen As Boolean

    ' Set the worksheets
    Set inputSheet = ThisWorkbook.Sheets("Input")
    Set ncrSheet = ThisWorkbook.Sheets("NCR Data")
    

    ' Read data from Input sheet
    companyName = inputSheet.Range("B7").value
    dateValue = inputSheet.Range("D7").value
    ncheckValue = IIf(inputSheet.Shapes("ncheck").OLEFormat.Object.value = 1, "1", "0") ' Read ncheck checkbox value
    ocheckValue = IIf(inputSheet.Shapes("ocheck").OLEFormat.Object.value = 1, "1", "0") ' Read ocheck checkbox value

        ' Find the next available row in NCR Data sheet of current workbook
        nextRow = ncrSheet.Cells(ncrSheet.Rows.Count, "A").End(xlUp).Row + 1

        ' Output data to NCR Data sheet of current workbook
        ncrSheet.Cells(nextRow, "A").value = companyName
        ncrSheet.Cells(nextRow, "B").value = dateValue
        ncrSheet.Cells(nextRow, "C").value = ncheckValue
        ncrSheet.Cells(nextRow, "D").value = ocheckValue

        

    Call UFillSequentialNumbersNCR
End Sub




Sub ClearFilters3()
    Dim wsRework As Worksheet
    Dim wsNCR As Worksheet
    Dim qtRework As ListObject
    Dim qtNCR As ListObject
    
    ' Set worksheet references
    Set wsRework = Worksheets("Rework Data")
    Set wsNCR = Worksheets("NCR Data")
    
    ' Set table references
    Set qtRework = wsRework.ListObjects("rework")
    Set qtNCR = wsNCR.ListObjects("ncr")
    
    ' Clear filter on column B in "Rework Data" table
    On Error Resume Next
    qtRework.Range.AutoFilter Field:=2
    On Error GoTo 0
    
    ' Clear filter on column B in "NCR Data" table
    On Error Resume Next
    qtNCR.Range.AutoFilter Field:=2
    On Error GoTo 0
End Sub

Sub SumValuesByCompany()
    Dim ws As Worksheet
    Dim wsOutput As Worksheet
    Dim lastRow As Long
    Dim companyList As Object
    Dim company As Range
    Dim cell As Range
    Dim visibleCells As Range
    Dim companyName As String
    Dim totalC As Double, totalD As Double

    ' Set the worksheets
    Set ws = ThisWorkbook.Sheets("NCR Data") ' Replace "NCR Data" with your sheet name if different
    Set wsOutput = ThisWorkbook.Sheets("NCR DataOutput")
    
    ' Initialize a dictionary to store unique company names
    Set companyList = CreateObject("Scripting.Dictionary")
    
    ' Determine the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through visible cells in column A to get unique company names
    On Error Resume Next ' In case there are no visible cells
    Set visibleCells = ws.Range("A2:A" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    If Not visibleCells Is Nothing Then
        For Each company In visibleCells
            companyName = company.value
            If Not companyList.Exists(companyName) Then
                companyList.Add companyName, 0 ' Initialize company with total 0
            End If
        Next company
    End If

    ' Clear existing data in the output sheet columns A, B, and C
    wsOutput.Range("A:C").ClearContents
    
    ' Output unique company names in column A of the output sheet
    wsOutput.Range("A2").Resize(companyList.Count).value = WorksheetFunction.Transpose(companyList.Keys)
    
    ' Calculate totals for columns C and D and output in columns B and C of the output sheet
    If Not visibleCells Is Nothing Then
        For Each company In wsOutput.Range("A2:A" & companyList.Count + 1) ' Iterate through unique company names in column A of the output sheet
            totalC = 0
            totalD = 0
            
            ' Loop through visible cells in column A to find matching company names and sum values from columns C and D
            For Each cell In visibleCells
                If cell.value = company.value Then
                    totalC = totalC + cell.Offset(0, 2).value ' Column C
                    totalD = totalD + cell.Offset(0, 3).value ' Column D
                End If
            Next cell
            
            ' Output totals in columns B and C next to the corresponding company name in column A of the output sheet
            company.Offset(0, 1).value = totalC ' Column B
            company.Offset(0, 2).value = totalD ' Column C
        Next company
    End If
    
    ' Clean up
    Set companyList = Nothing
    Set ws = Nothing
    Set wsOutput = Nothing
End Sub

Sub ClearFiltersInput()
    Dim wsRework As Worksheet
    Dim wsNCR As Worksheet
    Dim wsResponse As Worksheet

    ' Set the worksheets
    Set wsRework = ThisWorkbook.Sheets("Rework Data")
    Set wsNCR = ThisWorkbook.Sheets("NCR Data")
    Set wsResponse = ThisWorkbook.Sheets("Response Data")
    
    ' Clear filter on column B for Rework Data table
    If wsRework.ListObjects("rework").AutoFilter.FilterMode Then
        wsRework.ListObjects("rework").AutoFilter.ShowAllData
    End If

    ' Clear filter on column B for NCR Data table
    If wsNCR.ListObjects("ncr").AutoFilter.FilterMode Then
        wsNCR.ListObjects("ncr").AutoFilter.ShowAllData
    End If

    ' Clear filter on column B for Response Data table
    If wsResponse.ListObjects("response").AutoFilter.FilterMode Then
        wsResponse.ListObjects("response").AutoFilter.ShowAllData
    End If

    ' Ensure column B is selected for each table
    wsRework.ListObjects("rework").Range.AutoFilter Field:=2
    wsNCR.ListObjects("ncr").Range.AutoFilter Field:=2
    wsResponse.ListObjects("response").Range.AutoFilter Field:=2
End Sub

Sub ClearFiltersInputTEST()
    Dim wsVendorRework As Worksheet
    Dim wsVendorNCR As Worksheet
    Dim wsVendorResponse As Worksheet
    Dim wbVendorScorecard As Workbook

    ' Assume the "Vendor Scorecard TEST.xlsm" workbook is already open
    Set wbVendorScorecard = Workbooks("Vendor Scorecard TEST.xlsm")
    
    ' Set the worksheets in "Vendor Scorecard TEST"
    Set wsVendorRework = wbVendorScorecard.Sheets("Rework Data")
    Set wsVendorNCR = wbVendorScorecard.Sheets("NCR Data")
    Set wsVendorResponse = wbVendorScorecard.Sheets("Response Data")

    ' Clear filter on column B for Rework Data table
    If wsVendorRework.ListObjects("rework").AutoFilter.FilterMode Then
        wsVendorRework.ListObjects("rework").AutoFilter.ShowAllData
    End If

    ' Clear filter on column B for NCR Data table
    If wsVendorNCR.ListObjects("ncr").AutoFilter.FilterMode Then
        wsVendorNCR.ListObjects("ncr").AutoFilter.ShowAllData
    End If

    ' Clear filter on column B for Response Data table
    If wsVendorResponse.ListObjects("response").AutoFilter.FilterMode Then
        wsVendorResponse.ListObjects("response").AutoFilter.ShowAllData
    End If

    ' Ensure column B is selected for each table
    wsVendorRework.ListObjects("rework").Range.AutoFilter Field:=2
    wsVendorNCR.ListObjects("ncr").Range.AutoFilter Field:=2
    wsVendorResponse.ListObjects("response").Range.AutoFilter Field:=2
End Sub


Sub CheckNCRDataAndSumValues()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim cell As Range
    Dim visibleDataExists As Boolean
    
    ' Set the worksheet and table
    Set ws = ThisWorkbook.Sheets("NCR Data")
    Set tbl = ws.ListObjects("ncr")
    
    visibleDataExists = False
    
    On Error GoTo NoVisibleCells
    ' Loop through the data body range of the table and check if there are visible cells with values
    For Each cell In tbl.DataBodyRange.SpecialCells(xlCellTypeVisible)
        If Not IsEmpty(cell.value) Then
            visibleDataExists = True
            Exit For
        End If
    Next cell
    On Error GoTo 0
    
    ' Check if there are visible data
    If visibleDataExists Then
        ' Call the macro to sum values by company
        Call SumValuesByCompany
    Else
        ' Output message if no data is found
        MsgBox "No data for NCRs in this time period"
    End If

    Exit Sub

NoVisibleCells:
    MsgBox "No data for NCRs in this time period"
    On Error GoTo 0
End Sub



