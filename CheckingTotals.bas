Attribute VB_Name = "CheckingTotals"
'This module will be used to check the totals individually after the Moe_Macro completes
Dim sh2 As Worksheet

Sub check_totals_main()
    '**** Main method to the check totals script ****
    If Not FormatSheet.moe_macro_exists("Moe_Macro") Then
        MsgBox ("Cannot use the check totals macro yet because you have not used the Moe Macro macro yet!!!!")
        Exit Sub
    End If
    
    Set sh2 = ActiveWorkbook.Sheets("Moe_Macro")
    
    Call loop_through_cells
    
    MsgBox ("Total Checking Complete!")
    
End Sub

Sub loop_through_cells()
    'This method will loop through the cells on the Moe_Macro sheet to check the journal entries located in cell of each row
    Sheets("Moe_Macro").Activate
    
    Dim beginning_row As Integer
    Dim previous_entry As String
    Dim entry As String
    Dim counter As Integer
    Dim row As Integer
    
    previous_entry = ""
    last_row = Macros.get_last_row
    
    For row = 2 To last_row
        entry = Range("B" & row).value
        
        If (entry <> previous_entry And previous_entry <> "") Then
            beginning_row = row - counter
            Call checking_total(row - 1, previous_entry, beginning_row)
            counter = 1
        Else
            counter = counter + 1
        End If
        previous_entry = entry
    Next row
End Sub

Sub checking_total(row As Integer, entry As String, beginning_row As Integer)
    'This method will check the total for the rows
    
    Sheets("Moe_Macro").Activate
    
    Dim temp_total As Currency
    Dim total_debit_column As Currency
    Dim total_credit_column As Currency
    Dim i As Integer
        
    total_debit_column = 0
    total_credit_column = 0
    
    Dim current_debit As Currency
    Dim current_credit As Currency

    For i = beginning_row To row
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
    
    If (temp_total <> 0) Then
        MsgBox ("The total is off by " & temp_total & " for the entry " & entry)
    End If
End Sub
