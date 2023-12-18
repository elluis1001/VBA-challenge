Attribute VB_Name = "Module2"
Sub Stocks()
    
    'Declarations/Variables:
    Dim ws As Worksheet
    Dim last_row As Long
    Dim tick_symbol As String
    Dim open_price As Double
    Dim close_price As Double
    Dim year_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    Dim x As Long
    Dim tick_counter As Long
    
    For Each ws In Worksheets
        
        'Column Names:
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Stock counter to start on row 2:
        tick_counter = 2
        last_row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        'Loop through all rows to fill in columns:
        For x = 2 To last_row
            
            'Loop through to see stock ticker name changes:
            If ws.Cells(x + 1, 1).Value <> ws.Cells(x, 1).Value Then
                'Stock symbol to fill in column #9:
                tick_symbol = ws.Cells(x, 1).Value
                ws.Cells(tick_counter, 9).Value = tick_symbol
                
                ' Calculate yearly change for the current ticker symbol
                open_price = ws.Cells(x - 1, 3).Value
                close_price = ws.Cells(x, 6).Value
                
                'the open and close price are the same:
                If open_price = close_price Then
                    year_change = 0
                Else
                    'Calculate yearly change
                    year_change = close_price - open_price
                End If
                
                'Fill in column #10:
                ws.Cells(tick_counter, 10).Value = year_change
                
                'Increment the ticker counter for the next row:
                tick_counter = tick_counter + 1
            
            End If
        
        Next x
    
    Next ws

End Sub

