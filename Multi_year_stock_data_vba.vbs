Attribute VB_Name = "Module1"
Sub TickerCount():

'----BASE HOMEWORK----


'declare vars
'current worksheet
Dim Current As Worksheet
'stock ticker name
Dim stock_name As String
'opening price
Dim open_price As Double
'closing price
Dim close_price As Double
'count of unique tickers
Dim ticker_count As Integer
'yearly change
Dim yearly_change As Double
'percent change
Dim percent_change As Double
'stock volume
Dim stock_vol As Double

'for loop through worksheets
For Each Current In Worksheets

    'create headers for summary table
    Current.Range("I1").Value = "Ticker"
    Current.Range("J1").Value = "Yearly Change"
    Current.Range("K1").Value = "Percent Change"
    Current.Range("L1").Value = "Total Stock Volume"
    
    'determine the last row
     LastRow = Current.Cells(Rows.Count, 1).End(xlUp).Row
     
    'set ticker_count to 0
    ticker_counter = 0
    
    'set first opening price
    open_price = Current.Range("C2").Value
    
    'loop through stock tickers
    For i = 2 To LastRow
        'check if stock name changed
        If Current.Cells(i + 1, 1).Value <> Current.Cells(i, 1).Value Then
            'if changed, do the following
            'set stock name as last ticker name
            stock_name = Current.Cells(i, 1).Value
            'record closing value
            close_price = Current.Cells(i, 6).Value
            'add one to the ticker count
            ticker_counter = ticker_counter + 1
            'add last day to stock vol
            stock_vol = stock_vol + Current.Cells(i, 7).Value
            
            'record outputs to summary table
            'print stock name
            Current.Cells(1 + ticker_counter, 9).Value = stock_name
            'calc and print the yearly change
            yearly_change = close_price - open_price
            Current.Cells(1 + ticker_counter, 10).Value = yearly_change
            'calc and print the percent change
            percent_change = yearly_change / open_price
            Current.Cells(1 + ticker_counter, 11).Value = percent_change
            'print stock vol
            Current.Cells(1 + ticker_counter, 12).Value = stock_vol
            
            'take note of new opening stock value
            open_price = Current.Cells(i + 1, 3).Value
            
            'reset stock_vol
            stock_vol = 0
            
        'if stock name is unchanged
        Else
            'add to stock vol
            stock_vol = stock_vol + Current.Cells(i, 7).Value
        
        'end if
        End If
        
    'next row of stock data
    Next i

    'format summary table
    
    'find last row of summary table
    LastRow2 = Current.Cells(Rows.Count, 9).End(xlUp).Row
    
    'autofit summary table
    Current.Columns("I:L").AutoFit
    
    'format percent change col to percent
    Current.Range("K2:K" & LastRow2).NumberFormat = "0.00%"
    
    'format yearly change col
    For j = 2 To LastRow2
        'check if postive change
        If Current.Cells(j, 10) > 0 Then
            'if positive, color green
            Current.Cells(j, 10).Interior.ColorIndex = 4
        'if not positive
        ElseIf Current.Cells(j, 10) < 0 Then
            'color red
            Current.Cells(j, 10).Interior.ColorIndex = 3
        'end if
        End If
    'next row of summary table
    Next j
    
    
    
 '-----BONUS-----
 
 
    'create headers and row labels
    Current.Range("O2").Value = "Greatest % Increase"
    Current.Range("O3").Value = "Greatest % Decrease"
    Current.Range("O4").Value = "Greatest Total Volume"
    Current.Range("P1").Value = "Ticker"
    Current.Range("Q1").Value = "Value"
    
    'pull values from summary table and print
    'greatest percentage increase
    PerMax = WorksheetFunction.Max(Current.Range("K2:K" & LastRow2))
    Current.Range("Q2").Value = PerMax
    'greatest percentage decrease
    PerMin = WorksheetFunction.Min(Current.Range("K2:K" & LastRow2))
    Current.Range("Q3").Value = PerMin
    'Greatest Volume
    VolMax = WorksheetFunction.Max(Current.Range("L2:L" & LastRow2))
    Current.Range("Q4").Value = VolMax
    
    'loop through stock names to find permax and permin names
    For k = 2 To LastRow2
        'checks if stock's percent change matches the max
        If Current.Cells(k, 11).Value = PerMax Then
            'if so, pulls name into table
            Current.Range("P2").Value = Current.Cells(k, 9)
        'checks if stock's percent change matches the min
        ElseIf Current.Cells(k, 11).Value = PerMin Then
            'if so, pulls name into table
            Current.Range("P3").Value = Current.Cells(k, 9)
        End If
    Next k
    
    'separate loop for vol max, just in case it happens to be one of the permax or permin names
    For m = 2 To LastRow2
        'checks if stock's total vol matches the max
        If Current.Cells(m, 12).Value = VolMax Then
            'if so, pulls name into table
            Current.Range("P4").Value = Current.Cells(m, 9)
        End If
    Next m
    
    'autofit bonus table
    Current.Columns("O:Q").AutoFit
    
    'format permax and permin to percent
    Current.Range("Q2:Q3").NumberFormat = "0.00%"
    
'next sheet
Next
            
End Sub

