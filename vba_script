Sub alphaTesting()

'Declaring variables
Dim ticker, next_ticker As String
Dim total_volume, print_counter, last_column, last_row2, i, last_row As Long
Dim first_open, last_close, yearly_change, total_vol As Double
Dim col_ticker, col_fo, col_lc, col_yc, col_pc, col_tv, col_ticker2, col_value, col_resume As Long
Dim maxPC, minPC, maxTV, j As Long
Dim maxPCt, minPCt, maxTVt As String
Dim size As String
Dim percentate_change As Double

'Variables for iterating between sheets
Dim mw As Workbook
Set mw = ActiveWorkbook
Dim ws As Worksheet

'Start sheet iterations
For Each ws In Worksheets
    ws.Activate
    
    'Calculating the last row and column
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    last_column = Cells(1, Columns.Count).End(xlToLeft).Column
    
    'Assingning the columns
    col_ticker = last_column + 4
    col_fo = last_column + 5
    col_lc = last_column + 6
    col_yc = last_column + 7
    col_pc = last_column + 8
    col_tv = last_column + 9
    col_ticker2 = last_column + 13
    col_value = last_column + 14
    col_resume = last_column + 12
    
    
    'Creating titles
    Cells(1, col_ticker) = "Ticker"
    Cells(1, col_fo) = "First Open"
    Cells(1, col_lc) = "Last Close"
    Cells(1, col_yc) = "Yearly Change"
    Cells(1, col_pc) = "Percentage Change"
    Cells(1, col_tv) = "Total Volume"
    
    Cells(1, col_ticker2) = "Ticker"
    Cells(1, col_value) = "Value"
    Cells(2, col_resume) = "Greatest % of Increase"
    Cells(3, col_resume) = "Greatest % of Decrease"
    Cells(4, col_resume) = "Greatest Total Volume"
    
    'Intializing variables
    ticker = Cells(2, 1)
    total_volume = 0
    print_counter = 2
    first_open = Cells(2, 3)
    
    'Start of the main procedure for one sheet
    For i = 2 To last_row
        'Calculate total volume
        If Cells(i, 1) = ticker Then
            total_volume = total_volume + Cells(i, 7)
        End If
        
        If Cells(i, 1) <> ticker Then
            'Calculate the other variables
            last_close = Cells(i - 1, 6)
            yearly_change = last_close - first_open
            
            If first_open = 0 And last_close = 0 Then
                percentate_change = 0
                
            ElseIf first_open = 0 Then
                percentate_change = 1
                
            Else
                percentate_change = (last_close / first_open) - 1
                
            End If
            
            'Print variables
            Cells(print_counter, col_ticker) = ticker
            Cells(print_counter, col_tv) = total_volume
            Cells(print_counter, col_fo) = first_open
            Cells(print_counter, col_lc) = last_close
            Cells(print_counter, col_yc) = yearly_change
            Cells(print_counter, col_pc) = percentate_change
            Cells(print_counter, col_pc).NumberFormat = ".00%"
            
            'Change the interior color of the yearly_change column
            If yearly_change <= 0 Then
                Cells(print_counter, col_yc).Interior.ColorIndex = 3
            ElseIf yearly_change > 0 Then
                Cells(print_counter, col_yc).Interior.ColorIndex = 4
            End If
            
            'Re-initialize variables
            ticker = Cells(i, 1)
            first_open = Cells(i, 3)
            total_volume = Cells(i, 7)
            print_counter = print_counter + 1
        End If
    Next i
    
    ' Looking for the Greatest % increase", "Greatest % Decrease" and "Greatest total volume
    last_row2 = Cells(Rows.Count, col_ticker).End(xlUp).Row
    maxPC = Cells(2, col_pc)
    minPC = Cells(2, col_pc)
    maxTV = Cells(2, col_tv)
    maxPCt = Cells(2, col_ticker)
    minPCt = Cells(2, col_ticker)
    maxTVt = Cells(2, col_ticker)
    
    For j = 3 To last_row2
        If Cells(j, col_pc) > maxPC Then
            maxPC = Cells(j, col_pc)
            maxPCt = Cells(j, col_ticker)
        End If
        
        If Cells(j, col_pc) < minPC Then
            minPC = Cells(j, col_pc)
            minPCt = Cells(j, col_ticker)
        End If
            
        If Cells(j, col_tv) > maxTV Then
            maxTV = Cells(j, col_tv)
            maxTVt = Cells(j, col_ticker)
        End If
    
    Next j
    
    'Print the results in the worksheet
    Cells(2, col_ticker2) = maxPCt
    Cells(3, col_ticker2) = minPCt
    Cells(4, col_ticker2) = maxTVt
    
    Cells(2, col_value) = maxPC
    Cells(3, col_value) = minPC
    Cells(4, col_value) = maxTV
    
    'Format percentage change
    Cells(2, col_value).NumberFormat = "0.00%"
    Cells(3, col_value).NumberFormat = "0.00%"
    
    'Autofit of all data
    ActiveSheet.Range(Cells(1, 1), Cells(1, col_value + 1)).EntireColumn.AutoFit
Next ws

End Sub



Sub clear()

'Variables for iterating between sheets
Dim mw As Workbook
Set mw = ActiveWorkbook
Dim ws As Worksheet


For Each ws In Worksheets
    ws.Activate
    ws.Range("K1:V1").EntireColumn.clear
Next ws
