Sub stockinfo()
    
    'To loop through each worksheet in the Excel
    Dim ws As Worksheet
        For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
            
            'For looping through each row of stock data
            Dim i As Long
            'For looping through the summary list to find the required metrics
            Dim j As Long
            'For identifying the row number and cell number to be pulled or populated
            Dim row As Integer
            row = 2
            Dim column As Integer
            column = 1
            
            'For determing the last row in the stock data
            Dim lastRow As Long
            lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
            'For determing the last row of the summary list
            Dim lastRowList As Long
            
            'Remaining declaration statements for the stock data metrics
            Dim ticker As String
            Dim startingprice As Double
            Dim endingprice As Double
            Dim yearlychange As Double
            Dim percentchange As Double
            Dim totalvolume As Double
            totalvolume = 0
            Dim maxvolume As Double
            maxvolume = 0
            Dim maxinc As Double
            maxinc = 0
            Dim maxdec As Double
            maxdec = 0
            
            'Adding headings to the worksheet
            ws.Cells(1, 9).Value = "Ticker Symbol"
            ws.Cells(1, 10).Value = "Yearly Change ($)"
            ws.Cells(1, 11).Value = "Yearly Change (%)"
            ws.Cells(1, 12).Value = "Total Volume"
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            ws.Cells(1, 16).Value = "Ticker Symbol"
            ws.Cells(1, 17).Value = "Value"
            
            'Set the initial ticker symbol and its starting price
            ticker = Cells(row, column).Value
            startingprice = Cells(row, column + 2).Value
             
            'For looping through the entire stock data
            For i = 2 To lastRow
                'If the ticker symbol in the next row is different, pull the end-of-year values and calculate the yearly changes
                If Cells(i + 1, column).Value <> Cells(i, column).Value Then
                    'Set the ticker symbol for the summary row
                    ticker = Cells(i, column).Value
                    Cells(row, column + 8).Value = ticker
                    'Set the ending price for the ticker
                    endingprice = Cells(i, column + 5).Value
                    'Calculate yearly change and populate the summary row
                    yearlychange = endingprice - startingprice
                    Cells(row, column + 9).Value = yearlychange
                    'Calculate percentage change based on yearly change and starting price and format as percentage
                    percentchange = yearlychange / startingprice
                    Cells(row, column + 10).Value = percentchange
                    Cells(row, column + 10).NumberFormat = "0.00%"
                    'Determine total stock volume and convert from scientific notation
                    totalvolume = totalvolume + Cells(i, column + 6).Value
                    Cells(row, column + 11).Value = totalvolume
                    Cells(row, column + 11).NumberFormat = "0"
                    'Change row count to prepare for next summary row
                    row = row + 1
                    'Reset starting price for the next ticker
                    startingprice = Cells(i + 1, column + 2)
                    'Reset total volume for the next ticker
                    totalvolume = 0
                'If the cell is still the same ticker, add up the stock volume
                Else
                    totalvolume = totalvolume + Cells(i, column + 6).Value
                End If
            Next i
            
            
            'Determine the last row of the summary list
            lastRowList = ws.Cells(Rows.Count, column + 8).End(xlUp).row
            'Loop through the summary list to find the max and min change, max volume, and their associated ticker symbols
            For j = 2 To lastRowList
                'Max change and its ticker
                If Cells(j, column + 10).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastRowList)) Then
                    Cells(2, column + 15).Value = Cells(j, column + 8).Value
                    Cells(2, column + 16).Value = Cells(j, column + 10).Value
                    Cells(2, column + 16).NumberFormat = "0.00%"
                'Min change and its ticker
                ElseIf Cells(j, column + 10).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastRowList)) Then
                    Cells(3, column + 15).Value = Cells(j, column + 8).Value
                    Cells(3, column + 16).Value = Cells(j, column + 10).Value
                    Cells(3, column + 16).NumberFormat = "0.00%"
                'Max volume and its ticker
                ElseIf Cells(j, column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastRowList)) Then
                    Cells(4, column + 15).Value = Cells(j, column + 8).Value
                    Cells(4, column + 16).Value = Cells(j, column + 11).Value
                    Cells(4, column + 16).NumberFormat = "0"
                End If
            Next j
    'Move to next worksheet to run the above code again
    Next ws
            
End Sub