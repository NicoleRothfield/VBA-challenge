Sub Stocks()

    ' Set up so script runs on all worksheets
    For Each ws In Worksheets
        ' Define Variables
        Dim TickerCode As String
        
        Dim YearlyChange As Double
        
        Dim PercentChange As Double
        
        Dim TotalStockVolume As Double
        
        Dim StockOpen As Double
        
        Dim StockClose As Double
        
        Dim NextEmptyRowNumber As Integer
        NextEmptyRowNumber = 2
        
        ' Define Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Insert columns and headers
        ws.Range("J1").EntireColumn.Insert
        
        ws.Range("J1").Value = "Ticker Code"
        
        ws.Range("K1").EntireColumn.Insert
        
        ws.Range("K1").Value = "Yearly Change"
        
        ws.Range("L1").EntireColumn.Insert
        
        ws.Range("M1").EntireColumn.Insert
        
        ws.Range("M1").Value = "Stock Volume"
        
        ' Format Percent Change column to percent
        ws.Range("L1").Value = "Percent Change"
        ws.Range("L1").EntireColumn.NumberFormat = "0.00%"
        
        ' Begin with the first ticker symbol in row 2. We must track when our loop has moved
        ' to a new ticker symbol in order to reset the accumulators
        TickerCode = ws.Cells(2, 1)
        TotalStockVolume = 0
        StockOpen = ws.Cells(2, 3)
        StockClose = 0
        YearlyChange = 0
        PercentChange = 0
        
         ' loop through all stock daily data
        For i = 2 To LastRow

                ' check if we are within the same Ticker Code
                If ws.Cells(i, 1).Value <> TickerCode Then
                    ' We've moved to a new ticker symbol - write out the current accumulators and
                    ' re-initialize for the new ticker
                    
                    ' Calculate Yearly Change
                    YearlyChange = StockClose - StockOpen
                    
                    'Calculate Percent Change
                    If StockOpen <> 0 Then
                        PercentChange = (YearlyChange / StockOpen)
                    Else
                        PercentChange = 100 ' Fix divide by zero issue
                    End If
                
                    ' Print Ticker Code to NextEmptyRowNumber
                    ws.Range("J" & NextEmptyRowNumber).Value = TickerCode
                    
                    ' Print the Total Stock amount to NextEmptyRowNumber
                    ws.Range("M" & NextEmptyRowNumber).Value = TotalStockVolume
                    
                    ' Print the Yearly Change amount to NextEmptyRowNumber
                    ws.Range("K" & NextEmptyRowNumber).Value = YearlyChange
                    
                    ' Format the color of yearly change: if negative red, if positive, green
                    If YearlyChange >= 0 Then
                            ws.Range("K" & NextEmptyRowNumber).Interior.ColorIndex = 43
                    Else
                            ws.Range("K" & NextEmptyRowNumber).Interior.ColorIndex = 22
                    End If
                    
                    ' Print the Percent Change amount to NextEmptyRowNumber
                    ws.Range("L" & NextEmptyRowNumber).Value = PercentChange
                    
                    ' Add one to NextEmptyRow
                    NextEmptyRowNumber = NextEmptyRowNumber + 1
                    
                    TickerCode = ws.Cells(i, 1)
                    TotalStockVolume = 0
                    StockOpen = ws.Cells(i, 3)
                    StockClose = 0
                    YearlyChange = 0
                    PercentChange = 0
                End If
                
                ' Always sum up Total Stock Volume
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                
                ' Always update stock close to whatever the latest close price is - we keep overwriting
                ' this until the last daily close for the ticker symbol, which is the yearly close
                StockClose = ws.Cells(i, 6).Value
    
        Next i

    Next ws

End Sub


