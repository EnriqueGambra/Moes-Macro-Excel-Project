'Main module

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

Public entries(1 To 4) As Variant

Dim sh1 As Worksheet
Dim sh2 As Worksheet

Sub Moe_Macro_Main()
    'This program will compute all neccessary calculations for Moe's
    Dim got_date As Boolean
    Dim start_date_col As Integer
    Dim end_date_col As Integer
    Dim cur_date_col As Integer
    Dim issue_num As Integer
    
    Dim issue_2_present As Boolean
    Dim issue_4_present As Boolean
    
    row_counter = 0
    
    start_date_col = 2
    end_date_col = 8
    
    Set sh1 = ActiveWorkbook.Sheets("Table 1")
    
    Call FormatSheet.set_sh1
    
    got_date = FormatSheet.has_date(sh1.Cells(1, 1).value)
    If (got_date = False) Then
        Call FormatSheet.create_date
    End If
    
    Call get_last_row
    
    If Not FormatSheet.moe_macro_exists("Moe_Macro") Then
        Call FormatSheet.create_worksheet
    End If
    
    Call FormatSheet.convert_to_currency
    
    'Getting the date columns and populating appropriate data
    For cur_date_col = start_date_col To end_date_col
        'If values are present... proceed
        If (sh1.Cells(4, cur_date_col) <> 0) Then
            todays_date = sh1.Cells(1, cur_date_col)
            Call fill_dictionary_rows
            Call fill_dictionary_values(cur_date_col)
            Call reset_entries
            
            'Start going through journal entries
            Call get_present_entries
            
            issue_4_present = check_issue_4
            
            If (entries(1) = True And entries(2) = True) Then
                issue_num = 1
                MsgBox ("Both EZ Catering receivables and MM receivables are present for " & todays_date & ". Must adjust entries as a result")
                Call ez_catering_entry(cur_date_col, issue_num)
                Call mm_entry(cur_date_col, issue_num)
                Call mj_entry(cur_date_col, issue_num)
            ElseIf (entries(1) = True And entries(2) = False) Then
                issue_2_present = check_issue_2
                
                If (issue_2_present = True And issue_4_present = False) Then
                    MsgBox ("Must check Monkey Media website because MM a/r are less than the catering gross amount on " & todays_date)
                    issue_num = 2
                ElseIf (issue_2_present = False And issue_4_present = True) Then
                    MsgBox ("Must check website because Cash O/S (w/Tips) is suspiciously high (over 30 dollars) for the date of " & todays_date)
                    issue_num = 4
                ElseIf (issue_2_present = True And issue_4_present = True) Then
                    MsgBox ("BOTH ISSUES 2 AND 4 ARE PRESENT FOR " & todays_date)
                    issue_num = 2
                End If
                
                Call mm_entry(cur_date_col, issue_num)
                Call mj_entry(cur_date_col, issue_num)
                
            ElseIf (entries(1) = False And entries(2) = True) Then
                If (issue_4_present = True) Then
                    issue_num = 4
                    MsgBox ("Must check website because Cash O/S (w/Tips) is suspiciously high (over 30 dollars) for the date of " & todays_date)
                End If
                
                Call ez_catering_entry(cur_date_col, issue_num)
            ElseIf (entries(1) = False And entries(2) = False) Then
                
                If (issue_4_present = True) Then
                    issue_num = 4
                    MsgBox ("Must check website because Cash O/S (w/Tips) is suspiciously high (over 30 dollars) for the date of " & todays_date)
                End If
                
                Call mj_entry(cur_date_col, issue_num)
                End If
                
            If (entries(3) = True) Then
                Call dd_entry
            End If
            
            If (entries(4) = True) Then
                Call ue_entry
            End If
        End If
    Next
    
End Sub

Sub mj_entry(cur_date, issue_num)
    'This method will fill the main journal entry
    Dim ref_num As String
    
    ReDim rows_required(1 To 16) As Integer
    
    Dim temp_cater As Currency
    Dim total As Currency
    
    Dim i As Integer
    
    'Possible rows needed
    Dim mm_cater As Integer
    Dim alt_cater As Integer
    
    mm_cater = dictionary_rows("MM Catering")
    alt_cater = dictionary_rows("Alt Total Catering")
    
    'All rows required indeces
    Dim cd As Integer
    Dim amex As Integer
    Dim vmc As Integer
    Dim c_os As Integer
    Dim disc As Integer
    Dim p_o As Integer
    Dim g_c As Integer
    Dim olo_c As Integer
    Dim f_b As Integer
    Dim cater As Integer
    Dim s_t As Integer
    Dim gc_s As Integer
    Dim olo_t As Integer
    Dim olo_f As Integer
    Dim mm_tips As Integer
    Dim ez_tips As Integer
    
    'Calling food and beverage/sales tax to add appropriate values
    Call total_food_beverage
    Call total_sales_tax
    
    ref_num = get_reference_num("ME")
    
    cd = dictionary_rows("- Cash Deposits")
    amex = dictionary_rows("Total Amex $")
    vmc = dictionary_rows("Total V/MC/Discover $")
    c_os = dictionary_rows("= Cash O/S (w/Tips)")
    disc = dictionary_rows("- Discounts / Promos")
    p_o = dictionary_rows("- Paid Outs")
    g_c = dictionary_rows("- Gift Card Redeemed")
    olo_c = dictionary_rows("- Alt Tend (OLO)")
    f_b = dictionary_rows("+ Food And Beverage")
    
    'Must add in cater -- depending on the issue determines if debit credit and value amount...
    cater = dictionary_rows("+ Catering Sales (Gross)")
    'Must add in s_t it also depends on what the amount is and what issue number it is
    s_t = dictionary_rows("+ Sales Tax")
    
    gc_s = dictionary_rows("+ Gift Cards Sold")
    olo_t = dictionary_rows("+ OLO Dispatch Tip")
    olo_f = dictionary_rows("+ OLO Dispatch Fee $")
    mm_tips = dictionary_rows("Monkey Media Tips $")
    ez_tips = dictionary_rows("EZ Cater Tips $")
    
    'Filling the rows required array
    rows_required(1) = cd
    rows_required(2) = amex
    rows_required(3) = vmc
    rows_required(4) = c_os
    rows_required(5) = disc
    rows_required(6) = p_o
    rows_required(7) = g_c
    rows_required(8) = olo_c
    rows_required(9) = f_b
    
    If (issue_num = 2) Then
        rows_required(10) = alt_cater
    Else
        rows_required(10) = cater
    End If
    
    rows_required(11) = s_t
    rows_required(12) = gc_s
    rows_required(13) = olo_t
    rows_required(14) = olo_f
    rows_required(15) = mm_tips
    rows_required(16) = ez_tips
    
    'issue_num = 0... no issue just take catering straight from the sheet
    
    'Catering is not required because catering is put into the monkey media entries and ez catering entries
    If (issue_num = 1) Then
        row_objects(cater).currency_value = 0
    
    'Catering is subtracted from overall catering - mm_cater
    ElseIf (issue_num = 2) Then
        temp_cater = row_objects(cater).currency_value - row_objects(mm_cater).currency_value
        row_objects(alt_cater).currency_value = temp_cater
    End If
    
    'In this journal entry, ez tips and mm tips are debited and have account names
    row_objects(mm_tips).is_debit = True
    row_objects(ez_tips).is_debit = True
    row_objects(mm_tips).account_name = "6910 - CASH OVER"
    row_objects(ez_tips).account_name = "6910 - CASH OVER"
    
    For i = 1 To 16
        Call populate_sheet(row_objects(rows_required(i)), ref_num)
    Next i
    
    Call check_total(16, rows_required)
    
End Sub

Sub mm_entry(date_column As Integer, issue_num As Integer)
    'This method will fill the monkey media journal entry
    Dim ref_num As String
    
    Dim i As Integer
    Dim mm_t As Integer
    Dim mm_r As Integer
    Dim mm_c As Integer
    Dim mm_s_t As Integer
    Dim tot_cater As Integer
    Dim ez_c As Integer
    
    Dim temp_cater As Currency
    Dim temp_sales_tax As Currency
    Dim total As Currency
    
    ReDim rows_required(1 To 4) As Integer
    
    mm_t = dictionary_rows("Monkey Media Tips $")
    mm_r = dictionary_rows("- Alt Tend (Onl Cater Credit)")
    tot_cater = dictionary_rows("+ Catering Sales (Gross)")
    mm_c = dictionary_rows("MM Catering")
    mm_s_t = dictionary_rows("MM Sales Tax")
    ez_c = dictionary_rows("Catering (EZ)")
    
    rows_required(1) = mm_r
    rows_required(2) = mm_c
    rows_required(3) = mm_t
    rows_required(4) = mm_s_t
    
    ref_num = get_reference_num("MM")
    
    'For this entry you must make mm tips debiting false
    row_objects(mm_t).is_debit = False
    
    'If EZ Catering is present
    If (issue_num = 1) Then
        'Getting the monkey media catering total by subtracting ez catering from the total catering
        temp_cater = row_objects(tot_cater).currency_value - row_objects(ez_c).currency_value
        row_objects(mm_c).currency_value = temp_cater
        
        'Calculating the monkey media sales tax by subtracting the new mm cater amount + mm tips from mm a/r
        temp_sales_tax = row_objects(mm_r).currency_value - (row_objects(mm_t).currency_value + row_objects(mm_c).currency_value)
        row_objects(mm_s_t).currency_value = temp_sales_tax
    
    'If catering amount > monkey media
    ElseIf (issue_num = 2) Then
        temp_cater = formatted_currency("Monkey Media", "Monkey Media adjusted (add all subtotal + delivery fee from MM website where orders are NOT paid at the register)")
        row_objects(mm_c).currency_value = temp_cater
        
        If (row_objects(mm_t).currency_value > 0) Then
            MsgBox ("mm a/r = " & row_objects(mm_r).currency_value & "mm cater = " & row_objects(mm_c).currency_value)
            
            temp_sales_tax = row_objects(mm_r).currency_value - (row_objects(mm_t).currency_value + row_objects(mm_c).currency_value)
            row_objects(mm_s_t).currency_value = temp_sales_tax
        Else
            temp_sales_tax = row_objects(mm_r).currency_value - row_objects(mm_c).currency_value
            row_objects(mm_s_t).currency_value = temp_sales_tax
        End If
    End If
    
    For i = 1 To 4
        Call populate_sheet(row_objects(rows_required(i)), ref_num)
    Next i
    
    Call check_total(4, rows_required)
    
End Sub

Sub ez_catering_entry(date_column As Integer, issue_num As Integer)
    'This method will fill the ez catering journal entry
    Dim ez_r As String
    Dim ez_t As String
    Dim ref_num As String
    
    'Gets the indeces for these rows
    Dim ez_r_v As Integer
    Dim ez_t_v As Integer
    Dim ez_s_t As Integer
    Dim ez_c As Integer
    
    Dim i As Integer
    
    ReDim rows_required(1 To 4) As Integer
    
    ez_r_v = dictionary_rows("- Alt Tend (EZ Cater)")
    ez_t_v = dictionary_rows("EZ Cater Tips $")
    ez_s_t = dictionary_rows("Sales Tax (EZ)")
    ez_c = dictionary_rows("Catering (EZ)")
    
    rows_required(1) = ez_r_v
    rows_required(2) = ez_t_v
    rows_required(3) = ez_s_t
    rows_required(4) = ez_c
    
    ref_num = get_reference_num("EZ")
    
    If (issue_num = 1) Then
        MsgBox ("Please check the EZ Catering website for " & todays_date & " and input the following totals when asked.")
        ez_sales_tax = format_currency("EZ Catering", "for total sales tax (under sales tax)")
        ez_cater = format_currency("EZ Catering", "for total catering in ez catering (subtotal + delivery fee)")
        
        row_objects(ez_s_t).currency_value = ez_sales_tax
        row_objects(ez_c).currency_value = ez_cater
    End If
    
    For i = 1 To 4
        Call populate_sheet(row_objects(rows_required(i)), ref_num)
    Next i
        
    Call check_total(4, rows_required)
    
End Sub

Sub dd_entry()
    'This method will fill the door dash entry
    Dim dd As Integer
    Dim fb As Integer
    Dim gh As Integer
    Dim pm As Integer
    Dim one As Integer
    Dim two As Integer
    
    Dim total As Currency
    
    Dim rows_required(1 To 2) As Integer
    
    Dim ref_num As String
    ref_num = get_reference_num("DD")
    
    dd = dictionary_rows("- Delivery (DoorDash)")
    
    rows_required(1) = dd
    rows_required(2) = dd
    
    'If UberEATS is not a present entry... then get all the other delivery entries and add them in
    If (entries(4) = False) Then
        fb = dictionary_rows("- Delivery (Foodsby)")
        gh = dictionary_rows("- Delivery (GRUBHUB)")
        pm = dictionary_rows("- Delivery (POSTMATES)")
        one = dictionary_rows("- Delivery (Local 1)")
        two = dictionary_rows("- Delivery (Local 2)")
        
        'Add all values to the door dash entry
        total = row_objects(fb).currency_value + row_objects(gh).currency_value + row_objects(pm).currency_value + row_objects(one).currency_value + row_objects(two).currency_value + row_objects(dd).currency_value
        row_objects(dd).currency_value = total
    End If
        
    Call populate_sheet(row_objects(dd), ref_num)
    
    'Have to change the values for the door dash entry since it has multiple accounts
    row_objects(dd).is_debit = False
    row_objects(dd).account_name = "4000 Sales 4001 FOOD & Beverage"
    row_objects(dd).name = ""
    
    Call populate_sheet(row_objects(dd), ref_num)
    
    Call check_total(2, rows_required)
End Sub

Sub ue_entry()
    'This method will fill the UberEATS journal entry
    Dim ue As Integer
    Dim fb As Integer
    Dim gh As Integer
    Dim pm As Integer
    Dim one As Integer
    Dim two As Integer
    
    Dim total As Currency
    
    Dim rows_required(1 To 2) As Integer
    
    Dim ref_num As String
    ref_num = get_reference_num("UE")
    
    ue = dictionary_rows("- Delivery (UberEATS)")
    
    rows_required(1) = ue
    rows_required(2) = ue
    
    'UberEATS will always add up the other delivery entries
    fb = dictionary_rows("- Delivery (Foodsby)")
    gh = dictionary_rows("- Delivery (GRUBHUB)")
    pm = dictionary_rows("- Delivery (POSTMATES)")
    one = dictionary_rows("- Delivery (Local 1)")
    two = dictionary_rows("- Delivery (Local 2)")
        
    'Add all values to the door dash entry
    total = row_objects(fb).currency_value + row_objects(gh).currency_value + row_objects(pm).currency_value + row_objects(one).currency_value + row_objects(two).currency_value + row_objects(ue).currency_value
    row_objects(ue).currency_value = total
        
    Call populate_sheet(row_objects(ue), ref_num)
    
    'Have to change the values for the door dash entry since it has multiple accounts
    row_objects(ue).is_debit = False
    row_objects(ue).account_name = "4000 Sales 4001 FOOD & Beverage"
    row_objects(ue).name = ""
    
    Call populate_sheet(row_objects(ue), ref_num)
    
    Call check_total(2, rows_required)
    
End Sub

Sub total_sales_tax()
    'Calculating the total sales tax for the day and will store it in the sheets sales tax
    Dim s_t As Integer
    Dim mm_s_t As Integer
    Dim ez_s_t As Integer
    
    Dim t_sales_tax As Currency
    
    s_t = dictionary_rows("+ Sales Tax")
    mm_s_t = dictionary_rows("MM Sales Tax")
    ez_s_t = dictionary_rows("Sales Tax (EZ)")
    
    t_sales_tax = row_objects(s_t).currency_value - (row_objects(mm_s_t).currency_value + row_objects(ez_s_t).currency_value)
    row_objects(s_t).currency_value = t_sales_tax
    
End Sub

Sub total_food_beverage()
    'Calculating the total food and beverage
    Dim total As Currency
    
    Dim ue As Integer
    Dim dd As Integer
    Dim bev As Integer
    Dim food As Integer
    Dim t_f_b As Integer
    
    ue = dictionary_rows("- Delivery (UberEATS)")
    dd = dictionary_rows("- Delivery (DoorDash)")
    bev = dictionary_rows("+ Beverage Sales (Gross)")
    food = dictionary_rows("+ Food Sales (Gross)")
    t_f_b = dictionary_rows("+ Food And Beverage")
    
    total = (row_objects(bev).currency_value + row_objects(food).currency_value) - (row_objects(ue).currency_value + row_objects(dd).currency_value)
    row_objects(t_f_b).currency_value = total
    
End Sub

Sub check_total(arr_length As Integer, ByRef rows_required() As Integer)
    'This method will call calc_total and will correct entries if required
    Dim total As Currency
    
    Dim i As Integer
    
    total = calc_total(arr_length)
    
    'Problem occurs here... where the total for some reason is double the affected row... for now fixed it by multiplying the affected row
    If (total > 0) Then
        For i = 1 To 16
            If (row_objects(rows_required(i)).currency_value * 2 = total) Then
                If (row_objects(rows_required(i)).is_debit = False) Then
                    row_objects(rows_required(i)).is_debit = True
                Else
                    row_objects(rows_required(i)).is_debit = False
                End If
                
                Call readjust_sheet(row_objects(rows_required(i)))
            Exit For
            End If
        Next i
    End If
    
End Sub

Function calc_total(length_row_needed As Integer) As Currency
    'This method will calculate the total columns and tell the user if it is correct
    Sheets("Moe_Macro").Activate
    
    Dim last_row_inserted As Integer
    Dim first_row_inserted As Integer
    
    Dim temp_total As Currency
    
    last_row_inserted = row_counter
    first_row_inserted = last_row_inserted - length_row_needed + 1
    
    'MsgBox ("First row = " & first_row_inserted)
    'MsgBox ("Last row = " & last_row_inserted)
    
    Dim total_debit_column As Currency
    Dim total_credit_column As Currency
    
    total_debit_column = 0
    total_credit_column = 0
    
    Dim current_debit As Currency
    Dim current_credit As Currency
    
    Dim i As Integer
    For i = first_row_inserted To last_row_inserted
        current_debit = Range("F" & i).value
        current_credit = Range("G" & i).value
        
        If (current_debit > 0) And (current_credit = 0) Then
            'MsgBox ("In current debit = " & current_debit & " and current credit = " & current_credit & " row number = " & i)
            total_debit_column = total_debit_column + current_debit
        ElseIf (current_debit = 0) And (current_credit > 0) Then
            'MsgBox ("In current credit = " & current_debit & " and current credit = " & current_credit & " row number = " & i)
            total_credit_column = total_credit_column + current_credit
        End If
    Next i
    
    'MsgBox ("total debit = " & total_debit_column & " total credit = " & total_credit_column)
    If (total_debit_column > total_credit_column) Then
        temp_total = total_debit_column - total_credit_column
    ElseIf (total_debit_column < total_credit_column) Then
        temp_total = total_credit_column - total_debit_column
    ElseIf (total_debit_column = total_credit_column) Then
        temp_total = 0
    End If
    
    calc_total = temp_total
    
End Function

Function check_issue_2() As Boolean
    'This function will return a true/false value stating if issue 2 is present (mm a/r < catering gross amount)
    Dim mm_a_r As Integer
    Dim tot_cater As Integer
    
    Dim t_result As Boolean
    
    mm_a_r = dictionary_rows("- Alt Tend (Onl Cater Credit)")
    tot_cater = dictionary_rows("+ Catering Sales (Gross)")
    
    If (row_objects(mm_a_r).currency_value < row_objects(tot_cater).currency_value) Then
        t_result = True
    Else
        t_result = False
    End If
    
    check_issue_2 = t_result
    
End Function

Function check_issue_4() As Boolean
    'This function will return a true/false value stating if issue 4 is present (Cash O/S is over 30 dollars)
    
    '**** WHAT DO WE DO IN THIS CASE??? ****
    Dim c_os As Integer
    
    Dim t_result As Boolean
    
    c_os = dictionary_rows("= Cash O/S (w/Tips)")
    
    MsgBox (row_objects(c_os).currency_value)
    
    If (row_objects(c_os).currency_value > 30 Or row_objects(c_os).currency_value < -30) Then
        t_result = True
    Else
        t_result = False
    End If
    
    check_issue_4 = t_result
    
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

Sub readjust_sheet(row_obj As RowData)
    'This method will re-adjust the sheet
    Dim i As Integer
    
    Dim f_col As Currency
    Dim g_col As Currency
    
    For i = row_counter To 2 Step -1
        If (Worksheets("Moe_Macro").Range("F" & i).value = row_obj.currency_value) Then
            Worksheets("Moe_Macro").Range("F" & i).value = Empty
            Worksheets("Moe_Macro").Range("G" & i).value = row_obj.currency_value
            Exit For
        ElseIf (Worksheets("Moe_Macro").Range("G" & i).value = row_obj.currency_value) Then
            Worksheets("Moe_Macro").Range("G" & i).value = Empty
            Worksheets("Moe_Macro").Range("F" & i).value = row_obj.currency_value
            Exit For
        End If
    Next i
        
    
End Sub

Sub populate_sheet(row_obj As RowData, ref_num As String)
    '**** This function will populate the 'Moe_Macro' sheet with appropriate values ****
    Sheets("Moe_Macro").Activate

    'This will increase the rows starting at 0 and going all the way up to know if we can populate the rows
    If (row_counter = 0) Then
        row_counter = 2
    Else
        row_counter = row_counter + 1
    End If
    
    'Inputting the date
    Worksheets("Moe_Macro").Range("A" & row_counter).value = todays_date
    'Inputting ref #
    Worksheets("Moe_Macro").Range("B" & row_counter).Value2 = ref_num
    'Inputting Account
    Worksheets("Moe_Macro").Range("C" & row_counter).Value2 = row_obj.account_name
    'Inputting moe_row (memo)
    'Worksheets("Moe_Macro").Range("D" & row_counter).Value2 = moe_row
    'Not inputting class

    'Figuring out if it is a debit or credit and depending on that will input in appropriate column
    If (row_obj.is_debit = True) Then
        Worksheets("Moe_Macro").Range("F" & row_counter).value = row_obj.currency_value
    Else
        Range("G" & row_counter).value = row_obj.currency_value
    End If
    
    If (row_obj.name <> "") Then
        Range("H" & row_counter).Value2 = row_obj.name
    End If
    
End Sub

Function get_reference_num(journal_entry As String) As String
    get_reference_num = journal_entry & " " & todays_date
End Function

Sub get_present_entries()
    'This function will check for what entries are present in the sheet and have a value
    Dim mm_r As Integer
    Dim ez_r As Integer
    Dim dd_r As Integer
    Dim ue_r As Integer
    
    'Getting the receivables values
    mm_r = dictionary_rows("- Alt Tend (Onl Cater Credit)")
    ez_r = dictionary_rows("- Alt Tend (EZ Cater)")
    dd_r = dictionary_rows("- Delivery (DoorDash)")
    ue_r = dictionary_rows("- Delivery (UberEATS)")
    
    If (row_objects(mm_r).currency_value > 0) Then
        entries(1) = True
    End If
    
    If (row_objects(ez_r).currency_value > 0) Then
        entries(2) = True
    End If
    
    If (row_objects(dd_r).currency_value > 0) Then
        entries(3) = True
    End If
    
    If (row_objects(ue_r).currency_value > 0) Then
        entries(4) = True
    End If
    
End Sub

Sub fill_dictionary_rows()
    'This method will fill a dictionary with the row names and corresponding row values
    Dim row As Integer
    Dim counter As Integer
    
    Dim row_name As String
    
    Set dictionary_rows = CreateObject("Scripting.Dictionary")
    
    counter = 1
    
    ReDim rows_added(1 To 7) As String
    rows_added(1) = "Catering (EZ)"
    rows_added(2) = "Sales Tax (EZ)"
    rows_added(3) = "+ Food And Beverage"
    rows_added(4) = "EZ Catering Total"
    rows_added(5) = "MM Sales Tax"
    rows_added(6) = "MM Catering"
    rows_added(7) = "Alt Total Catering"
    
    dictionary_rows.Add "Catering (EZ)", last_row + 1
    dictionary_rows.Add "Sales Tax (EZ)", last_row + 2
    dictionary_rows.Add "+ Food And Beverage", last_row + 3
    dictionary_rows.Add "EZ Catering Total", last_row + 4
    dictionary_rows.Add "MM Sales Tax", last_row + 5
    dictionary_rows.Add "MM Catering", last_row + 6
    dictionary_rows.Add "Alt Total Catering", last_row + 7
    
    For row = 1 To last_row + 7
        If row <= last_row Then
            row_name = sh1.Cells(row, 1).value
        Else
            row_name = rows_added(counter)
            counter = counter + 1
        End If
        
        If dictionary_rows.Exists(row_name) Then
            dictionary_rows.Remove row_name
        End If
            
        dictionary_rows.Add row_name, row
        Set row_objects(row) = New RowData
        row_objects(row).row_name = row_name
        row_objects(row).index = row
        
        Call fill_row_objects(row_objects(row), row)
    Next row
End Sub

Sub fill_dictionary_values(date_column As Integer)
    'This method will fill a dictionary with the row names and corresponding currency values
    Dim row As Integer
    Dim row_value As Variant
    Dim key As Variant

    For Each key In dictionary_rows.Keys
        row = dictionary_rows(key)
        row_value = sh1.Cells(row, date_column)
        If (row_value <> Empty) Then
            row_objects(row).currency_value = row_value
        End If
   Next key
   
End Sub

Sub get_last_row()
    'Retrieves the last row in the table
    last_row = sh1.Cells(Rows.Count, 1).End(xlUp).row
End Sub

Sub reset_entries()
    'This method will reset the entries array
    entries(1) = False
    entries(2) = False
    entries(3) = False
    entries(4) = False
End Sub

Sub fill_row_objects(row_object As RowData, row_number As Integer)
    'Filling the row objects array with the appropriate account numbers
    
    If (row_object.row_name = "- Delivery (DoorDash)") Then
        row_objects(row_number).account_name = "1201 Credit Card Receivable"
        row_objects(row_number).name = "DoorDash"
    
    ElseIf (row_object.row_name = "- Delivery (UberEATS)") Then
        row_objects(row_number).account_name = "1201 Credit Card Receivable"
        row_objects(row_number).name = "UberEATS"
    
    '2 Account names
    ElseIf (row_object.row_name = "- Delivery (EZ Cater)") Then
        row_objects(row_number).account_name = "1200 Accounts Receivable"
        row_objects(row_number).name = "EZ Catering"
    
    'Created row
    ElseIf (row_object.row_name = "Catering (EZ)") Then
        row_objects(row_number).account_name = "4000 SALES 4010 CATERING"
        row_objects(row_number).is_debit = False
        
    'Created row
    ElseIf (row_object.row_name = "Sales Tax (EZ)") Then
        row_objects(row_number).account_name = "Sales Tax Payable"
        row_objects(row_number).is_debit = False
        row_objects(row_number).name = "NYS Sales Tax"
    
    'Account name changes based on which entry it is -- other is '6910 - CASH OVER' (Cr.) in ME
    ElseIf (row_object.row_name = "EZ Cater Tips $") Then
        row_objects(row_number).account_name = "6000 OPERATING EXPENSES 6300 - OTHER 6900"
        row_objects(row_number).is_debit = False
    
    ElseIf (row_object.row_name = "- Cash Deposits") Then
        row_objects(row_number).account_name = "1005 - FNBLI"
    
    ElseIf (row_object.row_name = "Total Amex $") Then
        row_objects(row_number).account_name = "1005 - FNBLI"
    
    ElseIf (row_object.row_name = "Total V/MC/Discover $") Then
        row_objects(row_number).account_name = "1005 - FNBLI"
    
    'Debit/Credit changes depending on the cash o/s
    ElseIf (row_object.row_name = "= Cash O/S (w/Tips)") Then
        row_objects(row_number).account_name = "6910 - CASH"
    
    ElseIf (row_object.row_name = "- Discounts / Promos") Then
        row_objects(row_number).account_name = "4050 - DISCOUNTS"
    
    ElseIf (row_object.row_name = "- Paid Outs") Then
        row_objects(row_number).account_name = "- Delivery - Catering"
    
    ElseIf (row_object.row_name = "- Gift Card Redeemed") Then
        row_objects(row_number).account_name = "2300 - Moe's Gift Card"
    
    ElseIf (row_object.row_name = "- Alt Tend (OLO)") Then
        row_objects(row_number).account_name = "1201 - Credit Cards"
    
    'Needs to be altered by food and beverage and an added row
    ElseIf (row_object.row_name = "+ Food And Beverage") Then
        row_objects(row_number).account_name = "4001 FOOD & BEVERAGE"
        row_objects(row_number).is_debit = False
        
    'We do use this amount in some cases... but in other special cases we subtract out of this amount and we credit... will set to debit for now
    ElseIf (row_object.row_name = "+ Catering Sales (Gross)") Then
        row_objects(row_number).account_name = "4010 - CATERING"
        row_objects(row_number).is_debit = True
        
    ElseIf (row_object.row_name = "+ Sales Tax") Then
        row_objects(row_number).account_name = "2200 - Sales Tax"
        row_objects(row_number).is_debit = False
        row_objects(row_number).name = "NYS Sales Tax"
        
    ElseIf (row_object.row_name = "+ Gift Cards Sold") Then
        row_objects(row_number).account_name = "2300 Moe's Gift Card"
    
    ElseIf (row_object.row_name = "+ OLO Dispatch Tip") Then
        row_objects(row_number).account_name = "2565 - Door Dash Payable"
        row_objects(row_number).is_debit = False
    
    ElseIf (row_object.row_name = "+ OLO Dispatch Fee $") Then
        row_objects(row_number).account_name = "4040 - Delivery Fee"
        row_objects(row_number).is_debit = False
    
    ElseIf (row_object.row_name = "Monkey Media Tips $") Then
        row_objects(row_number).account_name = "6000 OPERATING EXPENSES 6300 - OTHER 6900"
        row_objects(row_number).is_debit = False
    
    'Added in row
    ElseIf (row_object.row_name = "MM Sales Tax") Then
        row_objects(row_number).account_name = "2200 - Sales Tax Payable"
        row_objects(row_number).is_debit = False
        row_objects(row_number).name = "NYS Sales Tax"
    
    'Added in row
    ElseIf (row_object.row_name = "MM Catering") Then
        row_objects(row_number).account_name = "4000 - SALES 4010 CATERING"
        row_objects(row_number).is_debit = False
    
    ElseIf (row_object.row_name = "- Alt Tend (Onl Cater Credit)") Then
        row_objects(row_number).account_name = "1200 - Accounts Receivable"
        row_objects(row_number).name = "Monkey Media Receivables"
    
    'Added row... for total catering when altered in main journal entry
    ElseIf (row_object.row_name = "Alt Total Catering") Then
        row_objects(row_number).account_name = "4010 - CATERING"
        row_objects(row_number).is_debit = False
    End If
    
End Sub
