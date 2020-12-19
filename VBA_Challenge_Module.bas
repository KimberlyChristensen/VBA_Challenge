Attribute VB_Name = "Module1"
Sub VBAChallenge()
    'declare variables
    Dim ticker As String
    Dim yearly_change As Double
    Dim pct_change As Double
    
    Dim total_volume As LongLong
    
    Dim high_pct_change As Double
    Dim low_pct_change As Double
    Dim high_volume As LongLong
    
    Dim next_row As Long
    
    'these strings will hold the ticker for each of the highest percentage mover, lowest percentage mover, and high volume securities
    Dim high_pct_ticker As String
    Dim low_pct_ticker As String
    Dim high_volume_ticker As String
    
    Dim Summary_Table_Row As Integer
    
    'set initial values to zero
        yearly_change = 0
        pct_change = 0
        volume = 0
    
        low_pct_change = 0
        high_pct_change = 0
        high_volume = 0
        
        'Can set this low_pct_change value at zero to start because
        'there are negative values in the dataset, otherwise would set to first value
    
        
    Dim current_ws As Worksheet
    
    'keep track of each row in data from first to last row
    Dim row_count As Long
    
    Dim open_price As Double
    Dim close_price As Double
   
    
    ' Cycle through all sheets
'
    For Each current_ws In Worksheets
    last_row = current_ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'counts the last row of each worksheet
        
    row_count = 2
    Summary_Table_Row = 2
    
    'Add in column headers for summary table
    
        current_ws.Range("I1") = "Ticker"
        current_ws.Range("J1") = "Yearly Change"
        current_ws.Range("K1") = "Percentage Change"
        current_ws.Range("L1") = "Total Stock Volume"
        
        current_ws.Range("P1") = "Ticker"
        current_ws.Range("Q1") = "Value"
        
        current_ws.Range("O2") = "Greatest % Increase"
        current_ws.Range("O3") = "Greatest % Decrease"
        current_ws.Range("O4") = "Greatest Total Volume"
   
    open_price = current_ws.Cells(row_count, 3).Value
    
        'in this case, the data is set to where the first date of the year is populated as the first observation
        'and the observations are earliest to latest

    For row_count = 2 To last_row
    If current_ws.Cells(row_count + 1, 1).Value <> current_ws.Cells(row_count, 1).Value Then
        
        ticker = current_ws.Cells(row_count, 1).Value
        
        close_price = current_ws.Cells(row_count, 6).Value
    
    'this is looking for the cell at which the ticker changes and taking the last observation
    'it takes the close price of the last observation
    
    'Adding in a method to move to the next row when the initial observations start at zero, until the first is non-zero
                
                Do Until open_price <> 0
            
                    open_price = current_ws.Cells(row_count + next_row, 3).Value
                               
                    next_row = next_row + 1
                    
                Loop
                                                    
    
    'calculate percentage change as yearly change/open price
                
        yearly_change = close_price - open_price
       
                pct_change = yearly_change / open_price
           
        
        current_ws.Range("J" & Summary_Table_Row).Value = yearly_change
        
         'yearly change -- conditional formatting
         'conditional formatting for yearly change - green if positive, red if negative
            'can either use Interior.Color = vbRed or Interior.Color.Index = 3
            'can either use Interior.Color = vbGreen or Interior.Color.Index = 4
         
            If current_ws.Range("J" & Summary_Table_Row).Value < 0 Then
        
                current_ws.Range("J" & Summary_Table_Row).Interior.Color = vbRed
                
                Else
                
                current_ws.Range("J" & Summary_Table_Row).Interior.Color = vbGreen
                
            End If
            
    'populate percentage change for ticker in summary table column K
        current_ws.Range("K" & Summary_Table_Row).Value = pct_change
        
            If pct_change < low_pct_change Then
                low_pct_change = pct_change
                    low_pct_ticker = ticker

                ElseIf pct_change > high_pct_change Then

                high_pct_change = pct_change
                    high_pct_ticker = ticker


            End If

      
        total_volume = total_volume + current_ws.Cells(row_count, 7).Value
        
        current_ws.Range("I" & Summary_Table_Row).Value = ticker
        current_ws.Range("L" & Summary_Table_Row).Value = total_volume
        
        If total_volume > high_volume Then

                high_volume = total_volume
                
                high_volume_ticker = ticker

            End If
        
       
        Summary_Table_Row = Summary_Table_Row + 1
        
        total_volume = 0
        
        open_price = current_ws.Cells(row_count + 1, 3).Value
        
        
        Else
        
            total_volume = total_volume + current_ws.Cells(row_count, 7).Value
            
            End If
        

                
        Next row_count
        
        'format columns to reflect data type
       
        
        'percentage change -- percentage formatting
        current_ws.Range("J2:J" & last_row).NumberFormat = "#0.00"
        current_ws.Range("K2:K" & last_row).NumberFormat = "#0.00%"
        
        'volume with commas
        current_ws.Range("L2:L" & last_row).NumberFormat = "#,##0"
        
        'format cells in Greatest % Change(+/-) and Volume Table
        current_ws.Range("Q2").NumberFormat = "#,##0.00%"
        current_ws.Range("Q3").NumberFormat = "0.00%"
        current_ws.Range("Q4").NumberFormat = "#,##0"
        current_ws.Range("O:O").ColumnWidth = 20
        current_ws.Range("Q:Q").ColumnWidth = 13
        
        
        'Populate statistics for highest and lowest percentage change and highest volume, with associated ticker
        current_ws.Range("Q2").Value = high_pct_change
        current_ws.Range("Q3").Value = low_pct_change
        current_ws.Range("Q4").Value = high_volume
             
        current_ws.Range("P2").Value = high_pct_ticker
        current_ws.Range("P3").Value = low_pct_ticker
        current_ws.Range("P4").Value = high_volume_ticker
        
        'reset highest values
        high_pct_change = 0
        low_pct_change = 0
        high_volume = 0
        
        Next current_ws

End Sub

