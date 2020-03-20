Attribute VB_Name = "FormatSheet"
'FormatSheet Module

Option Explicit

Dim row_counter As Integer
Dim last_row As Integer

Dim row_objects(1 To 300) As RowData

Dim ez_sales_tax As Currency
Dim ez_cater As Currency
Dim mm_sales_tax As Currency
Dim mm_cater As Currency

Dim todays_date As Date

Dim dictionary_rows

Dim sh1 As Worksheet
Dim sh2 As Worksheet

Sub set_sh1()
    Set sh1 = ActiveWorkbook.Sheets("Table 1")
End Sub

Sub create_worksheet()
    'This method will create a new worksheet based from a template file we created
    
    'Gets a template of how we want the workbook to look from the below file location
    Dim str_template As String: str_template = "C:\Users\egambra\Documents\template_workbook.xlsx"
    Dim wb As Workbook
    'Sets the active workbook as the workbook used currently
    Set wb = ActiveWorkbook
    'Creates a new sheet in the workbook
    Dim ws As Worksheet
    Set ws = wb.Sheets.Add(After:=wb.Worksheets(wb.Worksheets.Count), Type:=str_template)
End Sub

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

Function moe_macro_exists(WorksheetName As String) As Boolean
    '******** Function that will check to see if the sheet 'moe_macro' exists and will return a true or false depending on it ********
     Dim ws As Worksheets
     Dim i As Integer
     Dim ws_exists As Boolean
     ws_exists = False
     For i = 1 To Worksheets.Count
        If Worksheets(i).name = WorksheetName Then
            ws_exists = True
        End If
    Next i
    moe_macro_exists = ws_exists
End Function

Sub create_date()
    'This method would be called if there is no date column when initially clicking the macros button
    
    Dim error_occured As Boolean
    error_occured = False
    
    Dim the_date As Date
    'starting_day gets the value from the first day of the week i.e. 'Monday' or 'Tuesday'
    Dim starting_day As Variant
    starting_day = sh1.Cells(1, 2).value
    
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
        sh1.Range("A1").value = "Date"
        sh1.Range("B1").value = the_date
        sh1.Range("C1").value = the_date + 1
        sh1.Range("D1").value = the_date + 2
        sh1.Range("E1").value = the_date + 3
        sh1.Range("F1").value = the_date + 4
        sh1.Range("G1").value = the_date + 5
        sh1.Range("H1").value = the_date + 6
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



