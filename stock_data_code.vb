Sub main_program()

'**************************************************************************
'********************************Outline ***********************************
'       1.  Variable Declarations
'       2.  Main Logic - collect row data
'       3.  Calculate ticker row values
'       4.  Display calculated ticker row values (create columns)
'       5.  Format Newly Created Columns
'       6.  Small Advanced Table - format and calc
'       7.  Functions
'           a.  number_sheets_in_wb()
'           b.  number_rows_in_sheet(i As Long)
'           c.  unique_ticker_labels_count(sh As Integer, cols As Integer)
'           d.  dict_max(sh As Integer, col_max As Long, col_label As Long, Optional pos_neg As Integer = 1)
'
'
'*************************************************************************
'*********************** Variable Declaration ****************************

Dim j As Integer                'sheet# iterator
Dim i As Integer                'unique ticker# iterator
Dim k As Long                   'row iterator
Dim last_k_value As Long        'saves the last index + 1 of the last k loop
                                'restarts next loop at this start point
Dim sum_daily_trade_volume As Double    'sum of k iteration values
Dim last_ticker_close As Double         'stores the last k iteration value
Dim first_ticker_open As Double         'stores the first k iteration value
Dim yearly_change As Double             'calculation using open/close values
Dim yearly_percent_change As Double     'calculation using yearly change and open value
Dim sheet_count As Integer              '#sheets from function


'*************************************************************************
'**************************** Main Logic *********************************

    sheet_count = number_sheets_in_wb()  'function determines #sheets in workbook

    For j = 1 To sheet_count  'j iterates over sheets
        last_k_value = 2  'start searching through row values at row 2
        
        For i = 2 To unique_ticker_labels_count((j), 1)   'i iterates to the last unique ticker label on the page
            
            carryover = 0  'this variable can be moved outside j iterator to create conintuous list
            sum_daily_trade_volume = 0  'initialize summing variable
            
            For k = last_k_value To number_rows_in_sheet((j))  'k iterates from last_k_value to the last row in sheet
                                                               'last_k_value is initialize as 2 but updates after each i iteration
                                                               'this allows looping to start from the end of the last loop
                first_ticker_open = Worksheets(j).Cells(last_k_value, 3)    'grab opening value of the year; aka last_k_value
                
                If (Worksheets(j).Cells(k, 1) = Worksheets(j).Cells(k + 1, 1)) Then
                    sum_daily_trade_volume = Worksheets(j).Cells(k, 7).Value + sum_daily_trade_volume
                Else
                    sum_daily_trade_volume = Worksheets(j).Cells(k, 7).Value + sum_daily_trade_volume
                    last_k_value = k + 1  'this stores the next iteration value that corresponds with the next unique ticker label
                    Exit For  'now that I have summed all values for one ticker label, i exit the loop and replace last_k_value as the new beginning iteration
                
                End If
                
                
            Next k 'iterates over sheet rows
            
        '**************************************************************
        '**************** Calculate ticker row values *****************
            
            'calculates ticker info summarized in k iteration
            'set ticker label
            last_ticker_close = Worksheets(j).Cells(last_k_value - 1, 6)
            
            'set change from open value of the year to the last close value of year
            yearly_change = last_ticker_close - first_ticker_open
            
            'set percent change calculation and filter out any divide by zero
            If first_ticker_open <> 0 Then
                yearly_percent_change = yearly_change / first_ticker_open
            Else
                yearly_percent_change = 0
            End If
                
         '****************************************************************
         '************ Display calculated ticker row values ***************
         
            'set row values to sheet
            Worksheets(j).Cells(i, 9).Value = Worksheets(j).Cells(last_k_value - 1, 1)
            Worksheets(j).Cells(i, 10).Value = yearly_change
            Worksheets(j).Cells(i, 11).Value = yearly_percent_change
            Worksheets(j).Cells(i, 12).Value = sum_daily_trade_volume
            Worksheets(j).Cells(i, 13).Value = Left(Worksheets(j).Cells(last_k_value - 1, 2), 4)
            
        Next i

    
        '**************************************************
        '************  Newly Created Columns  *************
        
        'Setup new column formats
        'set header values
        Worksheets(j).Cells(1, 9).Value = "Ticker"
        Worksheets(j).Cells(1, 10).Value = "Annual Change"
        Worksheets(j).Cells(1, 11).Value = "% Annual Change"
        Worksheets(j).Cells(1, 12).Value = "Stock Volume"
        Worksheets(j).Cells(1, 13).Value = "Year"
        
        'set color coding of Annual Change Column
        Worksheets(j).UsedRange.Columns("K").NumberFormat = "0.00%"
        Set positive_color = Worksheets(j).UsedRange.Columns("J").FormatConditions.Add(xlCellValue, xlGreater, 0)
        Set negative_color = Worksheets(j).UsedRange.Columns("J").FormatConditions.Add(xlCellValue, xlLess, 0)
            With positive_color
                .Interior.Color = vbGreen
            End With
            With negative_color
                .Interior.Color = vbRed
            End With
            
        'remove color coding of first row
        Worksheets(j).Range("J1:M1").Interior.ColorIndex = 0
        
        'set width of columns
        Worksheets(j).Columns("I:M").AutoFit
           
           
        '**************************************************
        '************  Small Advanced Table  **************
        
        'setup advanced table formats
        'set header values
        Worksheets(j).Cells(1, 17).Value = "Value"
        Worksheets(j).Cells(1, 16).Value = "Ticker"
        
        'set 1st row values
        dict_array1 = dict_max(j, 11, 9, 1)
        Worksheets(j).Cells(2, 17).Value = dict_array1(0)
        Worksheets(j).Cells(2, 16).Value = dict_array1(1)
        Worksheets(j).Cells(2, 15).Value = "Greatest % Increase"
        Worksheets(j).Range("Q2").NumberFormat = "0.00%"

        'set 2nd row values
        dict_array2 = dict_max(j, 11, 9, -1)
        Worksheets(j).Cells(3, 17).Value = dict_array2(0)
        Worksheets(j).Cells(3, 16).Value = dict_array2(1)
        Worksheets(j).Cells(3, 15).Value = "Greatest % Decrease"
        Worksheets(j).Range("Q3").NumberFormat = "0.00%"
        
        'set 3rd (final) row values
        dict_array3 = dict_max(j, 12, 9)
        Worksheets(j).Cells(4, 17).Value = dict_array3(0)
        Worksheets(j).Cells(4, 16).Value = dict_array3(1)
        Worksheets(j).Cells(4, 15).Value = "Greatest Total Volume"
        
        'set width of advancd table
        Worksheets(j).Columns("O:Q").AutoFit
        
        
    
    Next j  'sheet iterator


End Sub

'---------------------------------------------------------------------------------------------------------

Function number_sheets_in_wb()
'returns number of sheets in workbook
'Defn:  no inputs necessary

    Dim ws As Worksheet
    Dim i As Integer
    
    i = 0
    For Each ws In ThisWorkbook.Worksheets
        i = i + 1
    Next ws
    
    number_sheets_in_wb = i
     
End Function

'---------------------------------------------------------------------------------------------------------

Function number_rows_in_sheet(i As Long)
'returns number of rows on a specific sheet
'Defn:  i is the sheet#
'not necessary to be in a function
    
    number_rows_in_sheet = Worksheets(i).UsedRange.Rows.Count

End Function

'---------------------------------------------------------------------------------------------------------

Function unique_ticker_labels_count(sh As Integer, cols As Integer)
'calculates the number of ticker labels (companies) in a column in Workbook
'Defn:  sh is the sheet#
'Defn:  cols is the column#

    Dim i As Long
    Dim ticker_count As Long
    Dim string_value As String
    
    ticker_count = 1  'this could be zero but I corrected for this later
    string_value = Worksheets(sh).Cells(2, cols).Value
    
    For i = 2 To Worksheets(sh).UsedRange.Rows.Count
        If Worksheets(sh).Cells(i + 1, cols).Value = string_value Then
            'nothing happens
        Else
            string_value = Worksheets(sh).Cells(i + 1, cols).Value
            ticker_count = ticker_count + 1
            'MsgBox ("Ticker: " & string_value & "     Count:" & ticker_count)
        End If
    Next i
    
    unique_ticker_labels_count = ticker_count
    
    
End Function

'---------------------------------------------------------------------------------------------------------

Function dict_max(sh As Integer, col_max As Long, col_label As Long, Optional pos_neg As Integer = 1)
    'Defn:  sh is the sheet#
    'Defn:  col_max is the column# of the min/max value
    'Defn:  col_label is the column# of the ticker number
    'Defn:  pos_neg values are +1 or -1; Enter +1 for Maximum value capture; Enter -1 for Minimum value capture
    'After thought - I shoud have named this better
    
    Dim i As Long
    Dim max_Value As Double
    max_Value = Worksheets(sh).Cells(2, col_max).Value  'capture first value in column
    
    For i = 2 To Worksheets(sh).Cells(Rows.Count, col_max).End(xlUp).Row 'iterate through entire column values
        If pos_neg * Worksheets(sh).Cells(i + 1, col_max).Value > pos_neg * max_Value Then  'compare column values
            max_Value = Worksheets(sh).Cells(i + 1, col_max).Value  'store > value
            max_Label = Worksheets(sh).Cells(i + 1, col_label).Value 'store ticker name that corresponds with > value
        End If
    Next i
    dict_max = Array(max_Value, max_Label)  'output an array from the function
End Function



