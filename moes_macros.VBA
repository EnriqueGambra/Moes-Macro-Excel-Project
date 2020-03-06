Option Explicit

Dim row_counter As Integer
Dim last_row As Integer

Dim ez_sales_tax As Currency
Dim ez_cater As Currency
Dim mm_sales_tax As Currency
Dim mm_cater As Currency

Dim todays_date As Date

Dim dictionary_rows
Dim dictionary_values

Dim sh1 As Worksheet
Dim sh2 As Worksheet


Sub Moe_Macro_Main()
    'This program will compute all neccessary calculations for Moe's
    Dim got_date As Boolean
    Dim start_date_col As Integer
    Dim end_date_col As Integer
    Dim cur_date_col As Integer
    
    ReDim entries(1 To 4) As Boolean
    
    start_date_col = 2
    end_date_col = 8
    
    Set sh1 = ActiveWorkbook.Sheets("Table 1")
    
    got_date = has_date(sh1.Cells(1, 1).Value)
    If (got_date = False) Then
        Call create_date
    End If
    
    Call get_last_row
    Call fill_dictionary_rows
    
    If Not moe_macro_exists("Moe_Macro") Then
        Call create_worksheet
    End If
    
    Call convert_to_currency
    
    'Getting the date columns and populating appropriate data
    For cur_date_col = start_date_col To end_date_col
        'If values are present... proceed
        If (sh1.Cells(4, cur_date_col) <> 0) Then
            todays_date = sh1.Cells(1, cur_date_col)
            Call fill_dictionary_values(cur_date_col)
            
            'Start going through journal entries
            entries = get_present_entries
            
            If (entries(1) = True And entries(2) = True) Then
                MsgBox ("Both EZ Catering receivables and MM receivables are present. Adjusting entries as a result")
                Call ez_catering_entry(cur_date_col, 1)
                'Call mm_entry(cur_date_col, 1)
                'Call mj_entry(cur_date_col, 1)
            End If
            
        End If
    Next
    
End Sub

Sub ez_catering_entry(date_column As Integer, issue_num As Integer)
    'This method will fill the ez catering journal entry
    Dim ez_r As String
    Dim ez_t As String
    Dim ref_num As String
    
    Dim ez_r_v As Currency
    Dim ez_t_v As Currency
    
    ez_r_v = dictionary_values("- Alt Tend (EZ Cater)")
    ez_t_v = dictionary_values("EZ Cater Tips $")
    
    ref_num = get_reference_num("EZ")
    
    If (issue_num = 1) Then
        ez_sales_tax = format_currency("EZ Catering", "for total sales tax (under sales tax)")
        ez_cater = format_currency("EZ Catering", "for total catering in ez catering (subtotal + delivery fee)")
        '************* Start populating the rows **************
End Sub

Sub populate_sheet(moe_row As String, account_name As String, currency_val As Currency, ref_num As String, is_debit As Boolean)
    '**** This function will populate the 'Moe_Macro' sheet with appropriate values ****
    Sheets("Moe_Macro").Activate

    'This will increase the rows starting at 0 and going all the way up to know if we can populate the rows
    If (row_counter = 0) Then
        row_counter = 2
    Else
        row_counter = row_counter + 1
    End If
    
    'Inputting the date
    Worksheets("Moe_Macro").Range("A" & row_counter).Value = the_date
    'Inputting ref #
    Worksheets("Moe_Macro").Range("B" & row_counter).Value2 = ref_num
    'Inputting Account
    Worksheets("Moe_Macro").Range("C" & row_counter).Value2 = account_name
    'Inputting moe_row (memo)
    'Worksheets("Moe_Macro").Range("D" & row_counter).Value2 = moe_row
    'Not inputting class

    'Figuring out if it is a debit or credit and depending on that will input in appropriate column
    If (is_debit = True) Then
        Worksheets("Moe_Macro").Range("F" & row_counter).Value = currency_val
    Else
        Range("G" & row_counter).Value = currency_val
    End If
    
    If (account_name = "1201 Credit Card Receivable") Then
        If (InStr(ref_num, "DD") <> 0) Then
            Range("H" & row_counter).Value2 = "DoorDash"
        ElseIf (InStr(ref_num, "UE") <> 0) Then
            Range("H" & row_counter).Value2 = "UberEATS"
        End If
    ElseIf (account_name = "2200 Sales Tax Payable") Then
        Range("H" & row_counter).Value2 = "NYS Sales Tax"
    End If
    
End Sub

Function get_reference_num(journal_entry As String) As String
    get_reference_num = journal_entry & " " & todays_date
End Function

Function get_present_entries() As Variant
    'This function will check for what entries are present in the sheet and have a value
    Dim mm_r As Currency
    Dim ez_r As Currency
    Dim dd_r As Currency
    Dim ue_r As Currency
    
    'Getting the receivables values
    mm_r = dictionary_values("- Alt Tend (Onl Cater Credit)")
    ez_r = dictionary_values("- Alt Tend (EZ Cater)")
    dd_r = dictionary_values("- Delivery (DoorDash)")
    ue_r = dictionary_values("- Delivery (UberEATS)")
    
    ReDim temp_array(1 To 4) As Boolean
    
    If (mm_r > 0) Then
        temp_array(1) = True
    End If
    
    If (ez_r > 0) Then
        temp_array(2) = True
    End If
    
    If (dd_r > 0) Then
        temp_array(3) = True
    End If
    
    If (ue_r > 0) Then
        temp_array(4) = True
    End If
    
    get_present_entries = temp_array
End Function

Function formatted_currency(name_method As String, part_entering As String) As Currency
    'This function will ensure that the currency is formatted correctly and if it isn't will continue to call itself until it is entered correctly!
    
    On Error Resume Next
    Dim temp_value As Currency
    
    Do While True
        On Error Resume Next
        temp_value = InputBox("Enter in the total " & part_entering & " for " & name_method & " for the " & todays_date & " make sure it is in $#,#00.00 format!")
        
        If (Err.Number <> 0) Then
            MsgBox ("You did not enter the value in the correct format! Please enter again!")
        Else
            Exit Do
        End If
    Loop
    formatted_currency = temp_value
End Function

Sub convert_to_currency()
    'This method will convert all rows that include currency to the currency data type
    Dim i As Integer
    For i = 1 To last_row
        '****CHECK THE CURRENT VALUE STRING IF IT HAS A DECIMAL POINT IN IT, THEN IT IS A VALUE TO BE CONVERTED TO CURRENCY****
        'IN THE FUTURE I AM GOING TO HAVE TO ADD MORE IFS TO CHECK THE OTHER COLUMN CELLS
        If (sh1.Cells(i, 2) Like "*.*") Or (sh1.Cells(i, 2) Like "*.0*") Or (sh1.Cells(i, 2) Like "*0.*") Then
            Dim range_chars As String
            range_chars = "B" & i & ":I" & i
            'MsgBox (range_chars)
            Range(range_chars).Select
            Selection.NumberFormat = "$#,##0.00"
        End If
    Next
End Sub

Sub create_worksheet()
    'This method will create a new worksheet based from a template file we created
    
    'Gets a template of how we want the workbook to look from the below file location
    Dim str_template As String: str_template = "C:\Users\Owner\Documents\template_wb.xlsx"
    Dim wb As Workbook
    'Sets the active workbook as the workbook used currently
    Set wb = ActiveWorkbook
    'Creates a new sheet in the workbook
    Dim ws As Worksheet
    Set ws = wb.Sheets.Add(After:=wb.Worksheets(wb.Worksheets.Count), Type:=str_template)
End Sub

Function moe_macro_exists(WorksheetName As String) As Boolean
    '******** Function that will check to see if the sheet 'moe_macro' exists and will return a true or false depending on it ********
     Dim ws As Worksheets
     Dim i As Integer
     Dim ws_exists As Boolean
     ws_exists = False
     For i = 1 To Worksheets.Count
        If Worksheets(i).Name = WorksheetName Then
            ws_exists = True
        End If
    Next i
    moe_macro_exists = ws_exists
End Function

Sub fill_dictionary_rows()
    'This method will fill a dictionary with the row names and corresponding row values
    Dim row As Integer
    Dim row_name As String
    Set dictionary_rows = CreateObject("Scripting.Dictionary")
    
    For row = 1 To last_row
        row_name = sh1.Cells(row, 1)
        
        If dictionary_rows.Exists(row_name) Then
            dictionary_rows.Remove row_name
        End If
            
        dictionary_rows.Add row_name, row
    Next row
    
End Sub

Sub fill_dictionary_values(date_column As Integer)
    'This method will fill a dictionary with the row names and corresponding currency values
    Dim row As Integer
    Dim row_value As Variant
    Dim row_name As String
    Set dictionary_values = CreateObject("Scripting.Dictionary")
    
    For row = 1 To last_row
        row_name = sh1.Cells(row, 1)
        row_value = sh1.Cells(row, date_column)
        
        If dictionary_rows.Exists(row_name) Then
            dictionary_rows.Remove row_name
        End If
        
        dictionary_rows.Add row_name, row_value
   Next row
End Sub

Sub get_last_row()
    'Retrieves the last row in the table
    last_row = sh1.Cells(Rows.Count, 1).End(xlUp).row
End Sub

Sub create_date()
    'This method would be called if there is no date column when initially clicking the macros button
    
    Dim error_occured As Boolean
    error_occured = False
    
    Dim the_date As Date
    'starting_day gets the value from the first day of the week i.e. 'Monday' or 'Tuesday'
    Dim starting_day As Variant
    starting_day = sh1.Cells(1, 2).Value
    
    'Error handling, continues to resume regardless of an error
    On Error Resume Next
    the_date = InputBox("Enter in the date for " & starting_day & " in mm/dd/yyyy format")
    
    'If the error number is not 0... meaning there was an error it means that the date field was in the wrong format
    If Err.Number <> 0 Then
        error_occured = True
        MsgBox ("Did not enter in the date correctly... must enter in again!")
        
        'Recurisvely calls itself again to enter the correct date
        Call create_date
    End If
    
    'This is for the one time that the format was correct then we can set those values to the correct columns
    If error_occured = False Then
        Rows("1:1").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        Range("A1").Value = "Date"
        Range("B1").Value = the_date
        Range("C1").Value = the_date + 1
        Range("D1").Value = the_date + 2
        Range("E1").Value = the_date + 3
        Range("F1").Value = the_date + 4
        Range("G1").Value = the_date + 5
        Range("H1").Value = the_date + 6
    End If
End Sub

Function has_date(date_column As String) As Boolean
    'Function that will return true or false if there is a date row
    If (LCase(date_column) <> "date") Then
        has_date = False
    Else
        has_date = True
    End If
End Function