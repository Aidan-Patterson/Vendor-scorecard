Attribute VB_Name = "Module7"
Sub FilterByQuarterMaster()
    Dim wsPrintout As Worksheet
    Dim quarter As String
    Dim startDate As Date, endDate As Date
    
    Call CreateTablePOData
    ' Set the worksheet
    Set wsPrintout = ThisWorkbook.Sheets("Printout")

    ' Get the quarter from cell A5
    quarter = wsPrintout.Range("A5").value

    ' Determine the start and end dates for the given quarter
    Select Case quarter
        Case "Quarter 1"
            startDate = DateSerial(Year(Date), 1, 1)
            endDate = DateSerial(Year(Date), 3, 31)
        Case "Quarter 2"
            startDate = DateSerial(Year(Date), 4, 1)
            endDate = DateSerial(Year(Date), 6, 30)
        Case "Quarter 3"
            startDate = DateSerial(Year(Date), 7, 1)
            endDate = DateSerial(Year(Date), 9, 30)
        Case "Quarter 4"
            startDate = DateSerial(Year(Date), 10, 1)
            endDate = DateSerial(Year(Date), 12, 31)
        Case Else
            MsgBox "Invalid quarter specified in A5"
            Unload frmLoading
            Exit Sub
    End Select

    ' Clear any existing filters before applying new filters
    Call ClearFilters("NCR Data")
    Call ClearFilters("Rework Data")
    Call ClearFilters("Response Data")


    ' Filter the tables based on the determined date range
    Call FilterTable("NCR Data", "ncr", startDate, endDate)
    Call FilterTable("Rework Data", "rework", startDate, endDate)
    Call FilterTable("Response Data", "response", startDate, endDate)
 

    Call AFilterAndOutputDataQUARTER
    Call AAProcessCompanyData
    

    
    Call CheckNCRDataAndSumValues
    Call SumValuesByCompany1
    Call FilterAndExtractDataByQUARTER
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
    
    Call SetVendorPromptInA4
    
    Call SetPrintoutZoom
    
    

End Sub

Sub ClearFilters(sheetName As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws.AutoFilterMode Then
        ws.AutoFilter.ShowAllData
    End If
End Sub

Sub FilterTable(sheetName As String, tableName As String, startDate As Date, endDate As Date, Optional isPOData As Boolean = False)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim filterField As Long
    
    Set ws = ThisWorkbook.Sheets(sheetName)
    Set tbl = ws.ListObjects(tableName)
    
    ' Clear any existing filters
    If tbl.AutoFilter.FilterMode Then
        tbl.AutoFilter.ShowAllData
    End If
    
    ' Determine the correct column to filter
    If isPOData Then
        filterField = 3 ' Column C
    Else
        filterField = 2 ' Column B
    End If
    
    ' Apply date filter
    tbl.Range.AutoFilter Field:=filterField, _
                         Criteria1:=">=" & startDate, Operator:=xlAnd, _
                         Criteria2:="<=" & endDate
End Sub




