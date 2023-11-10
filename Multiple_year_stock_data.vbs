Attribute VB_Name = "Module1"
Sub stock_analysis()
    
For Each ws In Worksheets

    Dim stock_rows As Long 'placehold for row in stock market data
    Dim ticker_row As Long 'place hold for row in summary data
    stock_rows = ws.Cells(Rows.Count, 1).End(xlUp).Row 'find last row in worksheet
    ticker_row = 2 'row of first ticker entry in summary
    
    Dim year_open As Double
    Dim year_close As Double
    Dim year_change As Double
    Dim percent_change As Double
    Dim tvolume As Double 'total volume
    Dim greatestInc As Double
    Dim GPIticker As String
    Dim greatestDec As Double
    Dim GPDticker As String
    Dim greatestTV As Double
    Dim GTVticker As String
        
    ws.Range("I1").Value = "Ticker" 'record each Ticker symbol from yearly stock data in column I
    ws.Range("J1").Value = "Yearly Change" 'record yearly change for each Ticker symbol in column J
    ws.Range("K1").Value = "Percent Change" 'record percent yearly change for each Ticker symbol in column K
    ws.Range("L1").Value = "Total Volume" 'record yearly total volum for each Ticker symbol in column L
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    For I = 2 To stock_rows
        If ws.Cells(I, 1).Value <> ws.Cells(I - 1, 1).Value Then 'check if current row symbol is different from preceding row, i.e. new ticker symbol
            ws.Cells(ticker_row, 9).Value = ws.Cells(I, 1).Value 'display new symbol in Ticker summary
            year_open = ws.Cells(I, 3).Value 'record opening price of each unique stock
            tvolume = 0 'reset total volume of stock with new ticker symbol; will record volume from first row outside IF
        ElseIf ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value Then 'check to see if next row is different stock, i.e. final row of current ticker symbol
            year_close = ws.Cells(I, 6).Value 'record closing price of each unique stock
            year_change = year_close - year_open
            ws.Cells(ticker_row, 10).Value = year_change 'record change from year open to close
            percent_change = year_change / year_open 'calculate percent change
            ws.Cells(ticker_row, 11).Value = percent_change 'FormatPercent(year_change / year_open, 2) 'calculate percent change
            tvolume = tvolume + ws.Cells(I, 7).Value 'add volume for current row to total volume counter
            ws.Cells(ticker_row, 12).Value = tvolume 'record total volume from final row of current ticker symbol
            ticker_row = ticker_row + 1 'increment row in ticker summary
        End If
        'remaining code in for loop will execute for all rows for each ticker symbol
        tvolume = tvolume + ws.Cells(I, 7).Value 'add volume for current row to tvolume; total already recorded for final row of ticker symbol
    Next I
    
    'reset metrics for each sheet
    greatestInc = 0
    greatestDec = 0
    tvolume = 0
    
    'find greatest % inc, % dec, and total volume
    For k = 2 To ticker_row - 1
        If ws.Cells(k, 11).Value > greatestInc Then
            greatestInc = ws.Cells(k, 11).Value
            GPIticker = ws.Cells(k, 9).Value
            'MsgBox ("new high!" + Str(greatestInc)) 'for testing
        ElseIf ws.Cells(k, 11).Value < greatestDec Then
            greatestDec = ws.Cells(k, 11).Value
            GPDticker = ws.Cells(k, 9).Value
            'MsgBox ("new low!" + Str(greatestDec)) 'for testing
        End If
        
        If ws.Cells(k, 12).Value > greatestTV Then
            greatestTV = ws.Cells(k, 12).Value
            GTVticker = ws.Cells(k, 9).Value
        End If
    Next k
    
    'output greatest % inc, % dec, and total volume
    ws.Range("P2").Value = GPIticker
    ws.Range("Q2").Value = FormatPercent(greatestInc, 2)
    ws.Range("P3").Value = GPDticker
    ws.Range("Q3").Value = FormatPercent(greatestDec, 2)
    ws.Range("P4").Value = GTVticker
    ws.Range("Q4").Value = greatestTV
    
    'format Yearly and % Change: green if positive, red if negative; if yearly change pos/neg, % change is pos/neg
    For j = 2 To ticker_row - 1
        If ws.Cells(j, 10).Value > 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 10 'set interior color Green, if positive change
            ws.Cells(j, 11).Interior.ColorIndex = 10 'set interior color Green, if positive change
        ElseIf ws.Cells(j, 10).Value < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3 'set interior color Red, if negative change
            ws.Cells(j, 11).Interior.ColorIndex = 3 'set interior color Red, if negative change
        End If
        ws.Cells(j, 11).Value = FormatPercent(ws.Cells(j, 11).Value, 2) 'format %'s
    Next j
    
    ws.Columns("A:Q").AutoFit
Next ws
    
End Sub
