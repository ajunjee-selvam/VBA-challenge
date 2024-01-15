Sub stockinfo()
    'The above function is used to return the ticker symbol, yearly price change, % price change, and stock volume for each individual stock symbol
    ' in each of the worksheets. Additionally, the great % increase and decrease, as well as the greatest total volume are extracted from each worksheet
    
    
    'to loop through each worksheet in the Excel
    For Each ws In Worksheets
        'declaration statements
        Dim j As Integer
        Dim startingprice As Double
        Dim endingprice As Double
        
        'adding headings to each of the worksheets
        ws.Cells(1, 10).Value = "Ticker Symbol"
        ws.Cells(1, 11).Value = "Starting Opening Price"
        ws.Cells(1, 12).Value = "Ending Closing Price"
        ws.Cells(1, 13).Value = "Yearly Change ($)"
        ws.Cells(1, 14).Value = "Yearly Change (%)"
        ws.Cells(1, 15).Value = "Total Volume"
        ws.Cells(1, 17).Value = "Metric"
        ws.Cells(2, 17).Value = "Greatest % Increase"
        ws.Cells(3, 17).Value = "Greatest % Decrease"
        ws.Cells(4, 17).Value = "Greatest Total Volume"
        ws.Cells(1, 18).Value = "Ticker Symbol"
        ws.Cells(1, 19).Value = "Value"
        
        'for determing the last row in the stock data
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'for listing out the individual stocks from the data
        j = 2
        
        'looks through the data and returns the list of stocks to be summarized
        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ws.Cells(j, 10).Value = ws.Cells(i, 1).Value
                j = j + 1
            End If
        Next i
        
        'for determing the row number of the last stock to be summarized based on the extracted stock list
        lastRowList = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
        'looking through the extracted stock list
        For x = 2 To lastRowList
            'initializing the total stock volume as 0
            totalvolume = 0
            'looping through the stock data
            For i = 2 To lastRow
                
                'if the stock name in the data matches the stock in the list, then add up the volume
                If ws.Cells(i, 1).Value = ws.Cells(x, 10).Value Then
                    totalvolume = totalvolume + ws.Cells(i, 7).Value
                End If
                
                'if the stock name in the data matches the stock in the list and the previous row contains a different stock, set the starting price using this row
                If ws.Cells(i, 1).Value = ws.Cells(x, 10).Value And ws.Cells(i - 1, 1).Value <> ws.Cells(x, 10).Value Then
                    startingprice = ws.Cells(i, 3).Value
                    ws.Cells(x, 11).Value = startingprice
                End If
                
                'if the stock name in the data matches the stock in the list and the next row contains a different stock, set the ending price using this row
                If ws.Cells(i, 1).Value = ws.Cells(x, 10).Value And ws.Cells(i + 1, 1).Value <> ws.Cells(x, 10).Value Then
                    endingprice = ws.Cells(i, 6).Value
                    ws.Cells(x, 12).Value = endingprice
                End If
    
            Next i
            'return the total volume calculated from the previous loop
            ws.Cells(x, 15).Value = totalvolume
            'calculate the yearly difference and return the value in a cell
            ws.Cells(x, 13).Value = ws.Cells(x, 12).Value - ws.Cells(x, 11).Value
            'calculate the % yearly difference, format as a percentage, and return the value in a cell
            ws.Cells(x, 14).Value = Format((ws.Cells(x, 13).Value / ws.Cells(x, 11).Value), "Percent")
        
        Next x
        
        
        'initializing the values of the 3 metrics below
        maxvolume = 0
        maxinc = 0
        maxdec = 0
        'interating throw the stock list summary again
        For k = 2 To lastRowList
            'if the max volume of the current cell is greater than the existing maxvolume value, update the maxvolume to the current cell
            If ws.Cells(k, 15).Value > maxvolume Then
                maxvolume = ws.Cells(k, 15).Value
                ws.Cells(4, 19).Value = maxvolume
                ws.Cells(4, 18).Value = ws.Cells(k, 10).Value
            End If
            
            'if the max increase of the current cell is greater than the existing maxinc value, update the maxinc to the current cell
            If ws.Cells(k, 14).Value > maxinc Then
                maxinc = ws.Cells(k, 14).Value
                ws.Cells(2, 19).Value = Format(maxinc, "Percent")
                ws.Cells(2, 18).Value = ws.Cells(k, 10).Value
            End If
            
            'if the max decrease of the current cell is lower than the existing maxdec value, update the maxdec to the current cell
            If ws.Cells(k, 14).Value < maxdec Then
                maxdec = ws.Cells(k, 14).Value
                ws.Cells(3, 19).Value = Format(maxdec, "Percent")
                ws.Cells(3, 18).Value = ws.Cells(k, 10).Value
            End If
        Next k
        
    Next ws
End Sub
