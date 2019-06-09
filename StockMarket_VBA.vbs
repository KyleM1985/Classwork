Sub Total_Volume_Script()

    'Loop through all sheets
    For Each ws In Worksheets

    'Set variables for ticker
    Dim ticker As String
    
    'Set variable for volume
    Dim vol As Double
    vol = 0

    'Set variable to track location of ticker in summary table
    Dim table_row As Integer
    table_row = 2

    'Set variable that stores the number of rows on the sheet
    Dim Row_Count As Double
    Row_Count = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1
    
    'Set variable for yearly change
    Dim yearly As Double
    
    'Set variable for percent change
    Dim percent As Double

    'Set variable to store number of trading days to calculate change
    Dim trading_days As Double
    trading_days = 0

    'Add headers for summary table
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Volume"

    'Loop through all stocks
    For i = 2 To Row_Count
    
        'Set variables for open and close price
        Dim open_price As Double
        Dim close_price As Double
        open_price = ws.Cells(i - trading_days, 3).Value
        close_price = ws.Cells(i, 6).Value

            'Check if the ticker has changed, if not then:
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                'Set the ticker
                ticker = ws.Cells(i, 1).Value
                
                'Add to the tracking volume
                vol = vol + ws.Cells(i, 7).Value
                
                'Print the ticker in the summary table
                ws.Range("I" & table_row).Value = ticker
                        
                'Calculate the yearly change in stock price
                yearly = close_price - open_price
                
                'Print the yearly change in stock price
                ws.Range("J" & table_row).Value = yearly
            
                'Conditionally format the change based on gains or losses
                If yearly > 0 Then
                    ws.Range("J" & table_row).Interior.ColorIndex = 4
                ElseIf yearly < 0 Then
                    ws.Range("J" & table_row).Interior.ColorIndex = 3
                ElseIf yearly = 0 Then
                    ws.Range("J" & table_row).Interior.ColorIndex = 2
                End If
                
                'Calculate the percent change in stock price
                If open_price = 0 Then
                    percent = 0
                Else
                    percent = (close_price - open_price) / open_price
                End If
                
                'Print the percent change in stock price
                ws.Range("K" & table_row).Value = percent
                
                'Format percent change as percent
                ws.Range("K" & table_row).Style = "Percent"
                            
                'Print the total volume in the summary table
                ws.Range("l" & table_row).Value = vol
                
                'Add one to the summary table row
                table_row = table_row + 1
                
                'Reset the total trading volume
                vol = 0
                
                'Reset the number of trading days
                trading_days = 0
                
            'If we are still on the same ticker
            Else
            
                'Add to the total trading volume
                vol = vol + ws.Cells(i, 7).Value
                
                'Add one to the number of trading days
                trading_days = trading_days + 1
                            
        End If
        
    Next i

Next ws

End Sub

