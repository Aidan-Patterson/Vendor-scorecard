Attribute VB_Name = "Module14"
Sub selectiongrade()
Attribute selectiongrade.VB_ProcData.VB_Invoke_Func = " \n14"
'
' selectiongrade Macro
'

'
    Range("L4").Select
End Sub

Sub FilterTablesByMonthAndYearMaster()
    Dim wsPrintout As Worksheet
    Dim wsPO As Worksheet
    Dim wsNCR As Worksheet
    Dim wsRework As Worksheet
    Dim wsResponse As Worksheet
    
    Dim monthName As String
    Dim monthNumber As String
    Dim currentYear As String
    Dim filterStartDate As Date
    Dim filterEndDate As Date
    
    Call CreateTablePOData

    
    ' Set the worksheets
    Set wsPrintout = ThisWorkbook.Sheets("Printout")
    Set wsPO = ThisWorkbook.Sheets("PO Data")
    Set wsNCR = ThisWorkbook.Sheets("NCR Data")
    Set wsRework = ThisWorkbook.Sheets("Rework Data")
    Set wsResponse = ThisWorkbook.Sheets("Response Data")
    
    ' Get the month and current year
    monthName = wsPrintout.Range("A4").value
    currentYear = Year(Now)
    
    ' Convert month name to month number
    monthNumber = Month(dateValue("1 " & monthName & " " & currentYear))
    
    ' Set the filter start and end dates
    filterStartDate = DateSerial(currentYear, monthNumber, 1)
    filterEndDate = DateSerial(currentYear, monthNumber + 1, 0)
    
    ' Clear previous filters and apply new filters
    ' PO Data sheet

    
    ' NCR Data sheet
    If wsNCR.ListObjects("ncr").AutoFilter.FilterMode Then
        wsNCR.ListObjects("ncr").AutoFilter.ShowAllData
    End If
    wsNCR.ListObjects("ncr").Range.AutoFilter Field:=2, Criteria1:=">=" & filterStartDate, Operator:=xlAnd, Criteria2:="<=" & filterEndDate
    
    ' Rework Data sheet
    If wsRework.ListObjects("rework").AutoFilter.FilterMode Then
        wsRework.ListObjects("rework").AutoFilter.ShowAllData
    End If
    wsRework.ListObjects("rework").Range.AutoFilter Field:=2, Criteria1:=">=" & filterStartDate, Operator:=xlAnd, Criteria2:="<=" & filterEndDate
    
    ' Response Data sheet
    If wsResponse.ListObjects("response").AutoFilter.FilterMode Then
        wsResponse.ListObjects("response").AutoFilter.ShowAllData
    End If
    wsResponse.ListObjects("response").Range.AutoFilter Field:=2, Criteria1:=">=" & filterStartDate, Operator:=xlAnd, Criteria2:="<=" & filterEndDate
    
    Call AFilterByMonthAndOutputDataMONTH
    Call AAProcessCompanyData
    
    Call CheckNCRDataAndSumValues
    Call SumValuesByCompany1
    Call FilterAndExtractDataMONTH
    Call MatchCompaniesAndUpdate
    Call TotalCompanyValues
    
    Call CheckResponseDataAndSumValues
    
    Call CopyPODataToMasterSheet
    Call CopyNCRDataToMasterSheet
    Call CopyReworkDataToMasterSheet
    Call CopyResponseDataToMasterSheet
    
    Call SetColumnsToGeneral
    Call FillZerosInMasterSheet
    
    Call MatchCompaniesAndUpdate
    
    Call SortCompaniesAlphabetically
    
    Call SetVendorPromptInA5
    
    Call SetPrintoutZoom
    
End Sub



Sub CheckResponseDataAndSumValues()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim cell As Range
    Dim visibleDataExists As Boolean
    
    ' Set the worksheet and table
    Set ws = ThisWorkbook.Sheets("Response Data")
    Set tbl = ws.ListObjects("response")
    
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
        Call AverageCompanyValuesUsingArrays
    Else
        ' Output message if no data is found
        MsgBox "No Response data in this time period"
    End If

    Exit Sub

NoVisibleCells:
    MsgBox "No Response data in this time period"
    On Error GoTo 0
End Sub


