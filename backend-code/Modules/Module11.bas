Attribute VB_Name = "Module11"
Sub FilterByWeekMaster()

    
    Call FilterByWeek
    
    Call CopyVisibleIDsCompanyNamesAndAdditionalColumn
    Call MatchIDsAndOutputNameDate
    Call CompareDatesAndOutputStatus
    

    
    Call TallyOnTimeStatusForCompanies1
    Call TallyCompanyOccurrences1
    
    
    Call SumValuesByCompany
    
    Call SumValuesByCompany1
    Call totalCost
    Call TotalCompanyValues
    
    Call AverageCompanyValuesUsingArrays
    
    Call CopyPODataToMasterSheet
    Call CopyNCRDataToMasterSheet
    Call CopyReworkDataToMasterSheet
    Call CopyResponseDataToMasterSheet
    
    Call SetColumnsToGeneral
    Call FillZerosInMasterSheet
    
End Sub

Sub SortCompaniesAlphabetically()
    Dim wsMaster As Worksheet
    Dim lastRow As Long

    ' Set the worksheet
    Set wsMaster = ThisWorkbook.Sheets("PO DataOutput")
    
    ' Find the last row with data in column A
    lastRow = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row
    
    ' Sort the data in columns A and B
    wsMaster.Sort.SortFields.Clear
    wsMaster.Sort.SortFields.Add key:=wsMaster.Range("A2:A" & lastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With wsMaster.Sort
        .SetRange wsMaster.Range("A1:I" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


