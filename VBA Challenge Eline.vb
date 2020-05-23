Sub VBA Challenge Eline()

' Make sure to loop through each worksheet in the workbook


For Each ws In Worksheets


'Then set an initial variable for holding the ticker symbol name, the opening price, the closing price, the price change (I set a new variable to compare the first opening price value in a ticker category against the final closing price value in that category), the percentage change, the total stock volume

    Dim TickerSymbol As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim PriceChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    
'I Keep track of the location for each ticker category in the summary table

    Dim SummaryTableRow As Integer
    SummaryTableRow = 2
    
    ' First Opening Price
    
    OpeningPrice = ws.Cells(2, 3).Value
    
    ' Find the last row in the dataset
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    ' Set the header as "Ticker", "Yearly Change", "Percent Change", "Total Stock Volume" in our first analysis table and then as "Ticker", "Value" in a second table, and "Greatest % Increase", "Greatest % Decrease" and "Greatest Total Volume" in the rows
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    ' Make sure to loop through all stocks
    
    For i = 2 To LastRow
    
    'If the value for Opening Price is zero
    
    If OpeningPrice = 0 Then
    
    GoTo skipthisiteration
    
    End If
    
    ' Check if we are still within the same ticker name, if not
    
       If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
       
            'Set the Ticker Symbol name
            
            TickerName = ws.Cells(i, 1).Value
            
            ' Set the Closing Price
            
            ClosingPrice = ws.Cells(i, 6).Value
            
            ' Calculate the Price Change
            
            PriceChange = ClosingPrice - OpeningPrice
            
            ' Calculate the Percent Change
            
            PercentChange = (PriceChange / OpeningPrice)
            
            ' Format into Percent
            
            Percent = FormatPercent(PercentChange, 2)
            
            'Update the Opening Price
            
            OpeningPrice = ws.Cells(i + 1, 3).Value
            
            'Add to the Total Stock Volume
            
            TotalStockVol = TotalStockVol + ws.Cells(i, 7).Value
            
            ' Print the Ticker Symbol name in the Summary Table
            
            ws.Range("I" & SummaryTableRow).Value = TickerName
            ws.Range("J" & SummaryTableRow).Value = PriceChange
            ws.Range("K" & SummaryTableRow).Value = Percent
            ws.Range("L" & SummaryTableRow).Value = TotalStockVol
        
        ' Conditional format Percent_Change to green and red
        
        If PriceChange > 0 Then ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
        If PriceChange < 0 Then ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
    
        ' Add 1 to the summary table row
        
        SummaryTableRow = SummaryTableRow + 1
        
        'Reset the stock volume
        
        TotalStockVol = 0
        
        'If the cell following a row is the same ticker
        Else
    
        'Add to the Volume Total
        
        TotalStockVol = TotalStockVol + ws.Cells(i, 7).Value
        
        End If
        
skipthisiteration:
    
        
        Next i
        
        
'CHALLENGE

    ' Set extra variables for challenge
    
    Dim FinalRow As Integer
    
    Dim TickerSymbolBis As String
    Max = 0
    Min = 0
    GreatestVol = 0
    
    FinalRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
    For i = 2 To FinalRow
        
    'Check percentage increase in ticker category
        
        If ws.Cells(i, 11).Value > Max Then
            Max = ws.Cells(i, 11).Value
            TickerSymbolBis = ws.Cells(i, 9).Value
            PercentBis = FormatPercent(Max, 2)
            ws.Range("Q2").Value = PercentBis
            ws.Range("P2").Value = TickerSymbolBis
        End If
        
        'Check percentage decrease in ticker category
        
        If ws.Cells(i, 11).Value < Min Then
            Min = ws.Cells(i, 11).Value
            TickerSymbolBis = ws.Cells(i, 9).Value
            PercentBis = FormatPercent(Min, 2)
            ws.Range("Q3").Value = PercentBis
            ws.Range("P3").Value = TickerSymbolBis
        End If
        
        'Check greatest total volume in ticker category
        
        If ws.Cells(i, 12).Value > GreatestVol Then
            GreatestVol = ws.Cells(i, 12).Value
            TickerSymbolBis = ws.Cells(i, 9).Value
            ws.Range("Q4").Value = GreatestVol
            ws.Range("P4").Value = TickerSymbolBis
            
            
         'Format values to fit column width and number format
         
        ws.Cells(4, 17).NumberFormat = "General"
        ws.Range("N:P").Columns.AutoFit
        ws.Range("I:L").Columns.AutoFit
        
        
        
        End If
        
Next i

Next ws

End Sub