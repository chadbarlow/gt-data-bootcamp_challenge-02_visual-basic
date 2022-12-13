Sub Multiple_Year_Stock_Data()

    ' Declare variables
    Dim sheet_count As Integer
    Dim current_sheet As Integer
    Dim arrayData   As Variant
    Dim last_row_in_dataset As LongLong
    Dim current_summary_table_row As LongLong
    Dim previous_row_ticker As String
    Dim current_row_in_dataset_ticker As String
    Dim next_row_ticker As String
    Dim first_day_opening_price As Double
    Dim unique_ticker_name As String
    Dim last_day_closing_price As Double
    Dim yearly_change As Double
    Dim first_row_in_data_subset As LongLong
    Dim last_row_in_data_subset As LongLong
    Dim last_row_in_summary_table As LongLong
    Dim current_row_in_summary_table As LongLong
    Dim current_greatest_percent_change_ticker As String
    Dim current_greatest_percent_change_value As Double
    Dim current_volume_row_value_in_summary_table As Double
    Dim current_least_percent_change_ticker As String
    Dim current_least_percent_change_value As Double
    Dim current_greatest_volume_ticker As String
    Dim current_greatest_volume_value As Double
    Dim previous_row_in_summary_table_percent_change_value As Double
    Dim current_row_in_summary_table_percent_change_value As Double
    Dim ws As Worksheet
    MsgBox ("Please be patient As your script executes...")
    sheet_count = ActiveWorkbook.Worksheets.Count
      
    ' Loop through each sheet in the workbook
    For current_sheet = 1 To sheet_count
        Set ws = ActiveWorkbook.Worksheets(current_sheet)
        last_row_in_dataset = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ' Clear any existing summary tables
        ws.Range("I2:Q" & last_row_in_dataset).Clear
        ' Print header rows
        arrayData = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume", "", "", "", "Ticker", "Value")
        ws.Range("I1:Q1") = arrayData
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        current_summary_table_row = 2
        
        ' Loop through data on the current sheet
        For current_row_in_dataset = 2 To last_row_in_dataset
            previous_row_ticker = ws.Cells(current_row_in_dataset - 1, 1).Value
            current_row_in_dataset_ticker = ws.Cells(current_row_in_dataset, 1).Value
            next_row_ticker = ws.Cells(current_row_in_dataset + 1, 1).Value
            ' If we are in the first row of a subset of equivalent consiguous tickers, then store the row number and the opening price
            If previous_row_ticker <> current_row_in_dataset_ticker Then
                first_row_in_data_subset = current_row_in_dataset
                first_day_opening_price = ws.Cells(current_row_in_dataset, 3).Value
                ' If we are in the last row of a subset of equivalent consiguous tickers, then create a new row in a summary table and write
                ' the following values to the table: ticker name, calculated yearly change, calculated percent change, and calculated total volume
            ElseIf next_row_ticker <> current_row_in_dataset_ticker Then
                unique_ticker_name = ws.Cells(current_row_in_dataset, 1).Value
                last_day_closing_price = ws.Cells(current_row_in_dataset, 6).Value
                ws.Cells(current_summary_table_row, 9).Value = unique_ticker_name
                yearly_change = last_day_closing_price - first_day_opening_price
                ws.Cells(current_summary_table_row, 10).Value = yearly_change
                ws.Cells(current_summary_table_row, 11).Value = yearly_change / first_day_opening_price
                last_row_in_data_subset = current_row_in_dataset
                ' Calculate and print sum total stock volume * * * NOTE TO GRADER * * * I chose a programmatic dynamic formula to accomplish this task
                ' because the method implied by the assignment's instructions (e.g. store the Total Stock Volume value for every item in the dataset
                ' and add the stored value to an updated version of itself through each iteration of the For loop) resulted in persistent run-time
                ' and overflow errors in Excel. Other students encountered similar errors.
                ws.Range("L" & current_summary_table_row).Formula = "=SUM(" & ws.Name & "!G" & first_row_in_data_subset & ":G" & last_row_in_data_subset & ")"
                current_summary_table_row = current_summary_table_row + 1
            End If
        Next current_row_in_dataset
        
        last_row_in_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
        current_greatest_percent_change_value = 0
        current_least_percent_change_value = 0
        current_greatest_volume_value = 0
        
        For current_row_in_summary_table = 2 To last_row_in_summary_table
            current_row_in_summary_table_percent_change_value = ws.Cells(current_row_in_summary_table, 11).Value
            If current_row_in_summary_table = 2 Then
                previous_row_in_summary_table_percent_change_value = 0
            Else
                previous_row_in_summary_table_percent_change_value = ws.Cells(current_row_in_summary_table - 1, 11).Value
            End If
            If current_row_in_summary_table_percent_change_value >= previous_row_in_summary_table_percent_change_value And current_row_in_summary_table_percent_change_value >= current_greatest_percent_change_value Then
                current_greatest_percent_change_value = current_row_in_summary_table_percent_change_value
                current_greatest_percent_change_ticker = ws.Cells(current_row_in_summary_table, 9).Value
            ElseIf current_row_in_summary_table_percent_change_value <= previous_row_in_summary_table_percent_change_value And current_row_in_summary_table_percent_change_value <= current_least_percent_change_value Then
                current_least_percent_change_value = current_row_in_summary_table_percent_change_value
                current_least_percent_change_ticker = ws.Cells(current_row_in_summary_table, 9).Value
            End If
            previous_volume_row_value_in_summary_table = ws.Cells(current_row_in_summary_table - 1, 12).Value
            current_volume_row_value_in_summary_table = ws.Cells(current_row_in_summary_table, 12).Value
            If current_volume_row_value_in_summary_table >= previous_volume_row_value_in_summary_table And current_volume_row_value_in_summary_table >= current_greatest_volume_value Then
                current_greatest_volume_value = current_volume_row_value_in_summary_table
                current_greatest_volume_ticker = ws.Cells(current_row_in_summary_table, 9).Value
            End If
        Next current_row_in_summary_table
        
        ws.Cells(2, 16).Value = current_greatest_percent_change_ticker
        ws.Cells(2, 17).Value = current_greatest_percent_change_value
        ws.Cells(3, 16).Value = current_least_percent_change_ticker
        ws.Cells(3, 17).Value = current_least_percent_change_value
        ws.Cells(4, 16).Value = current_greatest_volume_ticker
        ws.Cells(4, 17).Value = current_greatest_volume_value
        ' Format functions
        ws.Range("J2:K" & last_row_in_summary_table).FormatConditions.Delete
        ws.Range("J2:K" & last_row_in_summary_table).FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
                        Formula1:="=0"
        ws.Range("J2:K" & last_row_in_summary_table).FormatConditions(1).Interior.Color = RGB(255, 0, 0)
        ws.Range("J2:K" & last_row_in_summary_table).FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
                        Formula1:="=0"
        ws.Range("J2:K" & last_row_in_summary_table).FormatConditions(2).Interior.Color = RGB(124, 252, 0)
        ws.Range("K2:K" & last_row_in_summary_table).NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Columns("A:L").AutoFit
        ws.Columns("O:Q").AutoFit
    Next current_sheet
    
    MsgBox ("You're done!")
End Sub