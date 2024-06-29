Attribute VB_Name = "AP_Supplier_Bal_Details_2"
'Written by King
'Created on 11/5/2024
'Updated on 11/5/2024 8.52 pm
'
'FUNCTIONS:
'1) Add total row for each company
'2) Sort By (All other currencies before SGD, SGD will be last in the order):
'       ~ USD
'       ~ AUD
'       ~ EUR
'       ~ JPY
'       ~ MYR
'       ~ CNH
'       ~ TW
'       ~ THB
'       ~ SGD
' -> Sort must be alphabetical order (note that when exported, it is already in alphabetical order)
'3) Add Grand Total for each currency at the end of the page
'4) Create AP Ageing Worksheet after tallying Grand Total (Sheet contains total for all organisations) [Update: 6/8/2022]
'5) Add Grand Total for AP Ageing Worksheet as well
'
'Updates:
' - Update Total cell formating (double line)
' - Remove bold and formating for all rows before calling sortCurrency) and pasting data over worksheet
' - Total and Grand Total changed to formulas and not tabulated figure (e.g. rng.Formula = "=SUM(A1:A200)")
' - Solved issue if company/supplier is the last company/supplier, unable to find and determine the last row of the company/supplier
'       -> Issue is resolved by referencing the timestamp row (which is usually the final row of the worksheet) and subtracting 3 from it to obtain the last company/supplier's last row.

Option Explicit

Private Enum supplierDetailsCols
    COL_COMPANY = 6
    COL_CURRENCY = 9
    COL_FOB = 10
    COL_ORIGINALAMOUINT = 8
    COL_START = 1
    COL_END = 10
End Enum
Private Enum supplierDetailsRows
    ROW_START = 2
End Enum
Private Enum supplierDetailArrayCols
    ARR_COMPANY = 0
    ARR_COMPANY_FROW = 1
    ARR_COMPANY_LROW = 2
    ARR_CURRENCY = 3
    ARR_ORDER = 4
End Enum
'2 dimensional Array for storing companyName, companyStartRow, companyEndRow, currency, order, companyRowsNo (companyStartRow - companyEndRow + 2 [add 2 to account for total row and empty row])
Private supplierDetailsArr() As Variant
Private currencyNamesArr() As Variant
Private sortOrderCurrency(0 To 8) As Variant
Private sortedCurrencyNames() As Variant
Private supplierDetailsArr_new() As Variant
Private grandTotalArr() As Variant

Sub Get_AP_Supplier_Bal_Details_2_MAIN()
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ActiveWorkbook
    Set ws = wb.Sheets(1)
    
    sortOrderCurrency(0) = "USD"
    sortOrderCurrency(1) = "AUD"
    sortOrderCurrency(2) = "EUR"
    sortOrderCurrency(3) = "JPY"
    sortOrderCurrency(4) = "MYR"
    sortOrderCurrency(5) = "CNH"
    sortOrderCurrency(6) = "TW"
    sortOrderCurrency(7) = "THB"
    sortOrderCurrency(8) = "SGD"
    
    Call addTotalRow(wb, ws)
    Call sortCurrency(wb, ws)
    Call calGrandTotal(wb, ws)
    Call createApAgeingSheet(wb, ws)
    Call formatDataSheet(wb, ws)
    
    MsgBox "Completed!"
End Sub

Sub addTotalRow(wb As Workbook, ws As Worksheet)
    Dim i, j, k, currencyCounter As Long
    Dim companyName As String
    Dim companyStartRW As Long, companyLastRW As Long
    Dim totalAmt As Double
    Dim totalLbl As String
    
    'Format currency from negative to positive and postive to negative values. (SAP default PU is negative, PC is positive)
    For i = ROW_START To ws.UsedRange.Rows.Count
        ws.Cells(i, COL_FOB) = -ws.Cells(i, COL_FOB)
        ws.Cells(i, COL_ORIGINALAMOUINT) = -ws.Cells(i, COL_ORIGINALAMOUINT)
    Next i
    
    totalLbl = "Total for "
    
    ReDim currencyNamesArr(0 To 0)
    i = ROW_START
    Do While i <= ws.UsedRange.Rows.Count
        If ws.Cells(i, COL_COMPANY) <> vbNullString Then
            companyName = ws.Cells(i, COL_COMPANY)
            ws.Cells(i + 1, COL_COMPANY).EntireRow.Delete
            companyStartRW = i + 1
            For j = companyStartRW To ws.UsedRange.Rows.Count
                'First condition is to determine if company is the last company (If yes, last row will be timestamp row - 3)
                If ws.Cells(j, COL_COMPANY) <> vbNullString Then
                    companyLastRW = j
                    Rows(companyLastRW).Insert
                    Rows(companyLastRW).Insert
                    Exit For
                ElseIf j = ws.UsedRange.Rows.Count Then
                    companyLastRW = j + 1   'Since its the last usedrange row, add one more row to account for total row for company
                End If
            Next j
            ws.Cells(companyLastRW, COL_COMPANY) = totalLbl & companyName
            ws.Cells(companyLastRW, COL_FOB).Formula = "=SUM(" & ws.Cells(companyStartRW, COL_FOB).Address & ":" & ws.Cells(companyLastRW - 1, COL_FOB).Address & ")"
            ws.Range(Cells(companyLastRW, COL_COMPANY), Cells(companyLastRW, COL_FOB)).Font.Bold = True
            i = j
            
            ReDim Preserve supplierDetailsArr(0 To 4, 0 To k)
            supplierDetailsArr(ARR_COMPANY, k) = companyName
            supplierDetailsArr(ARR_COMPANY_FROW, k) = companyStartRW
            supplierDetailsArr(ARR_COMPANY_LROW, k) = companyLastRW
            supplierDetailsArr(ARR_CURRENCY, k) = ws.Cells(companyStartRW, COL_CURRENCY)
            
            If Not IsInArray(ws.Cells(companyStartRW, COL_CURRENCY), currencyNamesArr) Then
                ReDim Preserve currencyNamesArr(0 To currencyCounter)
                currencyNamesArr(currencyCounter) = ws.Cells(companyStartRW, COL_CURRENCY)
                currencyCounter = currencyCounter + 1
            End If
            
            k = k + 1
        End If
        
        i = i + 1
    Loop
End Sub

Sub sortCurrency(wb As Workbook, ws As Worksheet)
    Dim i, j, arrPter, currencyPter As Long
    Dim str As String
    Dim lastrow As Long
    Dim rng As Range
    Dim currencyName As String
    Dim order As Long
    
    'Call clear formating function to remove random bolded or pre-formatted rows
    Call clearFormatingWS(ws)
    
    'Create a sorted array first before looping through supplierDetails array and populating the order element
    For i = LBound(sortOrderCurrency) To UBound(sortOrderCurrency)
        str = sortOrderCurrency(i)
        If IsInArray(str, currencyNamesArr) Then
            ReDim Preserve sortedCurrencyNames(0 To j)
            sortedCurrencyNames(j) = sortOrderCurrency(i)
            j = j + 1
        End If
    Next i
    
    'Populate order element in supplierDetailsArray
    order = 0
    For currencyPter = LBound(sortedCurrencyNames) To UBound(sortedCurrencyNames)
        For arrPter = 0 To UBound(supplierDetailsArr, 2)
            If supplierDetailsArr(ARR_ORDER, arrPter) = vbNullString And supplierDetailsArr(ARR_CURRENCY, arrPter) = sortedCurrencyNames(currencyPter) Then
                supplierDetailsArr(ARR_ORDER, arrPter) = order
                order = order + 1
            End If
        Next arrPter
    Next currencyPter
    
    'Loop through supplierDetailsArray by order to generate new sorted array (supplierDetailsArr_new)
    ReDim Preserve supplierDetailsArr_new(COL_START To COL_END, 1 To 1)
    For i = 0 To order - 1
        For arrPter = 0 To UBound(supplierDetailsArr, 2)
            If supplierDetailsArr(ARR_ORDER, arrPter) = i Then
                Dim tempArr As Variant
                With ws
                    tempArr = WorksheetFunction.Transpose(.Range(.Cells(supplierDetailsArr(ARR_COMPANY_FROW, arrPter) - 1, COL_START), .Cells(supplierDetailsArr(ARR_COMPANY_LROW, arrPter) + 1, COL_END)))
                End With
                supplierDetailsArr_new = Merge2DArray(supplierDetailsArr_new, tempArr)
                Exit For
            End If
        Next arrPter
    Next i
    
    'Clear Worksheet
    ws.Range(Cells(ROW_START, COL_START), Cells(ws.UsedRange.Rows.Count, COL_END)).Clear
    
    'Pasting Array on Worksheet and Formatting
    ws.Range(ws.Cells(ROW_START, COL_START), ws.Cells(UBound(supplierDetailsArr_new, 2) + 1, 10)) = WorksheetFunction.Transpose(supplierDetailsArr_new)
    ws.Range(ws.Cells(ROW_START, COL_START), ws.Cells(UBound(supplierDetailsArr_new, 2) + 1, 10)).Font.Size = 8
    ws.Range(ws.Cells(ROW_START, COL_START), ws.Cells(UBound(supplierDetailsArr_new, 2) + 1, 10)).Font.Name = "Arial"
    ws.Range(ws.Cells(ROW_START, COL_START), ws.Cells(UBound(supplierDetailsArr_new, 2) + 1, 10)).Columns(1).Font.Bold = True
    ws.Range(ws.Cells(ROW_START, COL_START), ws.Cells(UBound(supplierDetailsArr_new, 2) + 1, 10)).Columns(2).HorizontalAlignment = xlLeft
    
    For i = 1 To ws.UsedRange.Rows.Count
        On Error GoTo 0
        If Left(ws.Cells(i, COL_COMPANY), Len("Total for")) = "Total for" Then
            ws.Cells(i, COL_FOB).Font.Bold = True
            ws.Cells(i, COL_FOB).Borders(xlEdgeBottom).LineStyle = xlDouble
            ws.Cells(i, COL_FOB).Borders(xlEdgeBottom).Weight = xlThick
        End If
    Next i
End Sub

Sub calGrandTotal(wb As Workbook, ws As Worksheet)
    Dim i, j As Long
    Dim currencyName As String
    Dim currencyAmt As Double
    Dim outputStartRW As Long
    
    'Initialise grandTotalArr and populate currency label and currency amount to 0 first
    ReDim Preserve grandTotalArr(0 To UBound(sortedCurrencyNames), 0 To 1)
    For i = 0 To UBound(sortedCurrencyNames)
        grandTotalArr(i, 0) = sortedCurrencyNames(i)
        grandTotalArr(i, 1) = 0
    Next i
    
    'Look for "Total for" row to obtain the total for the supplier/company
    For i = ROW_START To ws.UsedRange.Rows.Count
        If Left(ws.Cells(i, COL_COMPANY), Len("Total for")) = "Total for" Then
            For j = 0 To UBound(grandTotalArr)
                If grandTotalArr(j, 0) = ws.Cells(i, COL_CURRENCY).Offset(-1, 0) Then
                    grandTotalArr(j, 1) = grandTotalArr(j, 1) + ws.Cells(i, COL_FOB)
                End If
            Next j
        End If
    Next i
    
    'Output Grand Total and format table
    outputStartRW = ws.UsedRange.Rows.Count + 3
    ws.Cells(outputStartRW, COL_START) = "Grand Total"
    ws.Range(ws.Cells(outputStartRW, COL_START), ws.Cells(outputStartRW, 2)).Merge
    ws.Range(ws.Cells(outputStartRW + 1, COL_START), ws.Cells(outputStartRW + 1 + UBound(grandTotalArr), 2)) = grandTotalArr
    ws.Range(ws.Cells(outputStartRW, COL_START), ws.Cells(outputStartRW + 1 + UBound(grandTotalArr), COL_START)).Font.Bold = True
    ws.Range(ws.Cells(outputStartRW, 2), ws.Cells(outputStartRW + 1 + UBound(grandTotalArr), 2)).NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    ws.Columns(2).EntireColumn.AutoFit
End Sub

Sub createApAgeingSheet(wb As Workbook, ws As Worksheet)
    Dim wsApAgeing As Worksheet
    Dim i As Long, j As Long
    
    Set wsApAgeing = Sheets.Add(After:=ws)
    On Error Resume Next
    wsApAgeing.Name = "AP AGEING"
    On Error GoTo 0
    
    'Headers for wsApAgeing
    With wsApAgeing
        .Range("A1") = "Supplier"
        .Range("B1") = "Currency"
        .Range("C1") = "Total"
    End With
    
    j = 2
    For i = ROW_START To ws.UsedRange.Rows.Count - 1
        If Left(ws.Cells(i, COL_COMPANY), Len("Total for")) = "Total for" Then
            wsApAgeing.Cells(j, 1) = Right(ws.Cells(i, COL_COMPANY), Len(ws.Cells(i, COL_COMPANY)) - Len("Total for "))
            wsApAgeing.Cells(j, 2) = ws.Cells(i, COL_CURRENCY).Offset(-1, 0)
            wsApAgeing.Cells(j, 3).Formula = "='" & ws.Name & "'!" & ws.Cells(i, COL_FOB).Address
            j = j + 1
        End If
    Next i
    
    'Multi sorting - sorting by currency then by supplier name
    Call multiSort(wsApAgeing)
    
    'Output Grand Total for AP AGEING Sheet
    With wsApAgeing
        .Range("E1") = "Grand Total"
        .Range(.Cells(2, 5), .Cells(UBound(grandTotalArr) + 1, 6)) = grandTotalArr
    End With
    
    'Formating
    wsApAgeing.Columns("A:C").EntireColumn.AutoFit
    wsApAgeing.Columns("C:C").EntireColumn.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    wsApAgeing.Columns("F:F").EntireColumn.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
    wsApAgeing.Rows(1).Font.Bold = True
End Sub

Sub formatDataSheet(wb As Workbook, ws As Worksheet)
    Dim i As Long
    
    'Delete index and Doc No (1st and 2nd columns) and "Type"
    ws.Columns("D:D").Delete
    ws.Columns("A:B").Delete
    
    'Bold Company name
    ws.Columns(3).Font.Bold = True
    
    'Delete every company's first row as it is repeated and not required
    For i = ROW_START To ws.UsedRange.Rows.Count
        If ws.Cells(i, 3) <> vbNullString And LCase(Left(ws.Cells(i, 3), 9)) <> "total for" Then
            ws.Cells(i, 3).EntireRow.Delete
            i = i - 1  'Minus 1 as a row has been deleted, row pointer should not increment
        End If
    Next i
    
    'Change data type of "Original Amount" & "Balance Due" to currency
    ws.Columns(5).Style = "Currency"
    ws.Columns(7).Style = "Currency"
    
    'Format negative values to RED
    For i = ROW_START To ws.UsedRange.Rows.Count
        If ws.Cells(i, 5) < 0 Then
            ws.Cells(i, 5).Font.Color = vbRed
        End If
        If ws.Cells(i, 7) < 0 Then
            ws.Cells(i, 7).Font.Color = vbRed
        End If
    Next i
End Sub

Sub multiSort(ws As Worksheet)
    Dim last_row As Long
    Dim custom_currency_order As String
    
    last_row = ws.UsedRange.Rows.Count
    custom_currency_order = Join(sortOrderCurrency, ",")
    
    'Add custom list
    Application.AddCustomList ListArray:=sortOrderCurrency
    ws.Sort.SortFields.Clear
    ' Sort by Currency first then by Supplier Name
    ws.Sort.SortFields.Add2 Key:=Range( _
        "B" & ROW_START & ":B" & last_row), SortOn:=xlSortOnValues, order:=xlAscending, CustomOrder:= _
        CVar(custom_currency_order), DataOption:=xlSortNormal
    ws.Sort.SortFields.Add2 Key:=Range( _
        "A" & ROW_START & ":A" & last_row), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:= _
        xlSortNormal
    'Apply sorting
    With ws.Sort
        .SetRange ws.Range("A1:C" & last_row)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Delete custom list
    Application.DeleteCustomList Application.CustomListCount
End Sub

Sub clearFormatingWS(ws As Worksheet)
    With ws
        .Range(.Cells(6, COL_START), .Cells(.UsedRange.Rows.Count - 1, COL_END)).ClearFormats
    End With
End Sub

Function Merge2DArray(arr As Variant, arr2 As Variant) As Variant
    Dim i, j, k As Long
    Dim oUBound_arr As Long, oLBound_arr As Long
    oUBound_arr = UBound(arr, 2)
    If UBound(arr, 2) = 1 Then
        oLBound_arr = 0
        oUBound_arr = 1
    Else:
        oLBound_arr = UBound(arr, 2)
        oUBound_arr = oUBound_arr + 1
    End If
    ReDim Preserve arr(COL_START To COL_END, 1 To oLBound_arr + UBound(arr2, 2))
    For i = COL_START To COL_END
        k = 1
        For j = oUBound_arr To UBound(arr, 2)
            arr(i, j) = arr2(i, k)
            k = k + 1
        Next j
    Next i
    Merge2DArray = arr
End Function

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False
End Function

Function RemoveElementFrom2DArray(index As Long, arr As Variant) As Variant
    Dim i As Long
    Dim newArr As Variant
    ReDim newArr(UBound(arr) - 1)
    For i = LBound(arr) To UBound(arr)
        If i < index Then
            newArr(i) = arr(i)
        ElseIf i > index Then
            newArr(i - 1) = arr(i)
        End If
    Next i
    RemoveElementFrom2DArray = newArr
End Function
