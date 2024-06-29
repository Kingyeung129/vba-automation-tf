Attribute VB_Name = "Get_AR_Summary_2"
'Written by King
'NOTE:
'11/05/2024 -

'Readme (Program Flow):
' 1) Check if 2nd sheetname is AR Ageing Summary
' 2) AR sheet pointer row starts at 6 and data sheet pointer row starts at 2
' 3) Get company start and end row with a loop first (Last row will be the last row that has "TOTAL" minus 1)
' 4) Get month no reference and tally with AR sheet month no, sum the value together. Loop through company start to end row to get total sum
' 5) output total sum for month in AR sheet then move on to next month column


Option Explicit
Private Enum sh_columns
    enum_sh_col_customer_name = 3
    enum_sh_col_date_of_transaction = 4
    enum_sh_col_amount = 8
End Enum

Sub Get_AR_Summary_2_MAIN()

    Dim i, j, k As Long
    Dim companyName As String
    Dim transaction_type As String
    Dim date_transaction As String
    Dim monthno As String
    Dim yearno As String
    Dim wb As Workbook
    Dim sh As Worksheet
    Dim ar_sh As Worksheet
    Dim sh_company_startrow As Long
    Dim sh_company_lastrow As Long
    Dim ar_sh_startcol, ar_sh_company_totalcolumn As Long
    Dim sh_startrow, sh_company_rowPt
    Dim ar_sh_rowPt, ar_sh_headerrow, ar_sh_startrow, ar_sh_lastrow As Long
    Dim sum_of_amount As Double
    Dim cell As Range
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Set wb = ActiveWorkbook
    Set ar_sh = wb.Sheets(2)
    Set sh = wb.Sheets(1)
    
    If UCase(ar_sh.Name) <> "AR AGEING SUMMARY" Then
        GoTo Err_Handler_SheetOrder
    End If
    
    '5 is header and 6 is starting row for ar_sh and sh (11/05/2024)
    ar_sh_headerrow = 5
    ar_sh_startrow = ar_sh_headerrow + 1
    ar_sh_rowPt = ar_sh_startrow
    '6 is starting row for sh (11/05/2024)
    sh_startrow = 2
    
    'Set date of transaction in datasheet as text type
    sh.Cells(1, enum_sh_col_date_of_transaction).EntireColumn.NumberFormat = "@"
    
    'Set AR sheet company name as text to prevent excel from converting string to date
    ar_sh.Columns("A:A").NumberFormat = "@"
      
    'Get Starting Column in AR Sheet. (Sometimes 2nd column is empty and not in the MMM'YY format, if that's the case, starting column will be 3 instead)
    If Trim(ar_sh.Cells(ar_sh_headerrow, 2)) <> vbNullString Then
        ar_sh_startcol = 2
    Else:
        ar_sh_startcol = 3
    End If
    
    'Get Total Column in AR Sheet
    If UCase(ar_sh.Cells(ar_sh_headerrow, ar_sh.UsedRange.Columns.Count)) = UCase("Total") Then
        ar_sh_company_totalcolumn = ar_sh.UsedRange.Columns.Count
    Else:
        MsgBox "A R Ageing Summary worksheet is not in the right format. Total should be the last column", vbCritical
        Exit Sub
    End If
    
    'First loop to get customer name, 2nd loop to get amount per month
    For i = sh_startrow To sh.UsedRange.Rows.Count
        If sh.Cells(i, enum_sh_col_customer_name) <> vbNullString Then
            companyName = sh.Cells(i, enum_sh_col_customer_name)
            'Debug.Print sh.Cells(i, enum_sh_col_customer_name)
            'Loop through each month header to get amount per month
            For j = ar_sh_startcol To ar_sh_company_totalcolumn - 1
                sum_of_amount = 0
                sh_company_rowPt = i + 2  'Add 2 as first row is tabulated total, followed by a blank row
                ar_sh.Cells(ar_sh_rowPt, 1) = companyName
                monthno = MonthNumber(Left(ar_sh.Cells(ar_sh_headerrow, j), 3))
                yearno = "20" & Right(ar_sh.Cells(ar_sh_headerrow, j), 2)
                If Not IsNumeric(monthno) Then
                    MsgBox "Check headers for Worksheet (A R Ageing Summary). Month Year headers should be in MMM'YY format", vbCritical
                    Exit Sub
                End If
                Do While (sh.Cells(sh_company_rowPt, enum_sh_col_customer_name) = vbNullString)
                    If sh_company_rowPt > sh.UsedRange.Rows.Count Then
                        Exit Do
                    End If
                    date_transaction = sh.Cells(sh_company_rowPt, enum_sh_col_date_of_transaction)
'                    Debug.Print "Month of Transaction: "; Format(Mid(date_transaction, 4, 2), "00"); "Year of Transaction: "; Format("20" & Mid(date_transaction, 7, 2), "0000")
                    If j = ar_sh_company_totalcolumn - 1 Then  'If is last AR sheet column before Totals column get all transaction both equivalent and earlier than month header
                        If CInt(Format("20" & Mid(date_transaction, 7, 2), "0000")) < CInt(yearno) Or (CInt(Format(Mid(date_transaction, 4, 2), "00")) <= CInt(monthno) And CInt(Format("20" & Mid(date_transaction, 7, 2), "0000")) <= CInt(yearno)) Then
                            sum_of_amount = sum_of_amount + sh.Cells(sh_company_rowPt, enum_sh_col_amount)
                        End If
                    ElseIf Format(Mid(date_transaction, 4, 2), "00") = monthno And CInt(Format("20" & Mid(date_transaction, 7, 2), "0000")) = CInt(yearno) Then
                        sum_of_amount = sum_of_amount + sh.Cells(sh_company_rowPt, enum_sh_col_amount)
                    End If
                    ar_sh.Cells(ar_sh_rowPt, j) = sum_of_amount
                    sh_company_rowPt = sh_company_rowPt + 1
                Loop
            Next j
            'Tabulate Company's Total Column
            ar_sh.Cells(ar_sh_rowPt, ar_sh_company_totalcolumn).Formula = "=sum(" & ar_sh.Cells(ar_sh_rowPt, ar_sh_startcol).Address & ":" & ar_sh.Cells(ar_sh_rowPt, ar_sh_company_totalcolumn - 1).Address & ")"
            ar_sh_rowPt = ar_sh_rowPt + 1
            i = sh_company_rowPt - 1
        End If
    Next i
    
    'Formulate Total for each column
    ar_sh_lastrow = ar_sh_rowPt
    ar_sh.Cells(ar_sh_lastrow, 1) = "TOTAL"
    For i = 2 To ar_sh_company_totalcolumn
        ar_sh.Cells(ar_sh_lastrow, i).Formula = "=sum(" & ar_sh.Cells(ar_sh_startrow, i).Address & ":" & ar_sh.Cells(ar_sh_lastrow - 1, i).Address & ")"
    Next i
    
    'Clear Formating AR Sheet for output
    ar_sh.Range(ar_sh.Cells(i, 1).Address & ":" & ar_sh.Cells(i, ar_sh_company_totalcolumn).Address).ClearFormats

    'Format Output (AR Summary)
    With ar_sh
        .Range(.Cells(ar_sh_startrow, 1).Address & ":" & Cells(ar_sh_lastrow, 1).Address).Font.Bold = True
        .Range(.Cells(ar_sh_startrow, 1).Address & ":" & Cells(ar_sh_lastrow, ar_sh_company_totalcolumn).Address).Font.Name = "Arial"
        .Range(.Cells(ar_sh_startrow, 1).Address & ":" & Cells(ar_sh_lastrow, ar_sh_company_totalcolumn).Address).Font.Size = 8
        .Range(.Cells(ar_sh_lastrow, 1).Address & ":" & Cells(ar_sh_lastrow, 1).Address).Font.Bold = True
        .Range(.Cells(ar_sh_startrow, 1).Address & ":" & Cells(ar_sh_lastrow, ar_sh_company_totalcolumn).Address).NumberFormat = "@"
        .Range(.Cells(ar_sh_startrow, 2).Address & ":" & Cells(ar_sh_lastrow, ar_sh_company_totalcolumn).Address).Style = "Currency"
'        .Range(.Cells(ar_sh_startrow, 2).Address & ":" & Cells(ar_sh_lastrow, ar_sh_company_totalcolumn).Address).Columns.AutoFit
    End With
    
    'Copy and paste valeus to avoid circular reference for formulas in "Total" column, then sort by Company Name Order
    With ar_sh
        .Range(.Cells(ar_sh_startrow, 1).Address & ":" & Cells(ar_sh_lastrow, .UsedRange.Columns.Count).Address).Copy
        .Cells(ar_sh_startrow, 1).PasteSpecial Paste:=xlPasteValues
        Application.CutCopyMode = False
        ar_sh.AutoFilterMode = False
        ar_sh.Range("A5").AutoFilter
        On Error Resume Next
        ar_sh.AutoFilter.Sort.SortFields.Clear
        On Error GoTo 0
        ar_sh.AutoFilter.Sort.SortFields.Add2 _
                Key:=Range(ar_sh.Cells(ar_sh_startrow - 1, 1).Address & ":" & ar_sh.Cells(ar_sh_lastrow, 1).Address), SortOn:=xlSortOnValues, order:=xlAscending, _
                DataOption:=xlSortNormal
    End With
    With ar_sh.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Format AR sheet negative values with red font
    For Each cell In ar_sh.Range(ar_sh.Cells(ar_sh_startrow, ar_sh_startcol), ar_sh.Cells(ar_sh.UsedRange.Rows.Count, ar_sh.UsedRange.Columns.Count))
        If cell.Value < 0 Then
            cell.Font.Color = vbRed
        End If
    Next
       
    'Completed, Notify user with message box
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.Calculate
    MsgBox "Completed!"
    Exit Sub
    
    
Err_Handler_SheetOrder:
    MsgBox "Wrong order for Worksheet (A R Ageing Summary) must be the second worksheet.", vbCritical
End Sub

Function MonthNumber(myMonthName As String)
    On Error GoTo Err_Handler
    MonthNumber = Month(DateValue("1 " & myMonthName & " 2020"))
    MonthNumber = Format(MonthNumber, "00")
    On Error GoTo 0
    Exit Function
    
Err_Handler:
    MsgBox "Please update month in this format (MMM'YY)", vbCritical
End Function

