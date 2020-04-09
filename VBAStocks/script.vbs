Sub script():

    Dim ws As Worksheet
    Dim LastRow As LongLong
    Dim LastColumn As LongLong
    'helper variables
    Dim i As LongLong
    Dim OpeningPrice As Double
    OpeningPrice = 0
    Dim ClosingPrice As Double
    Dim SolutionRow As Integer
    SolutionRow = 2
    Dim OpeningPriceIsSet As Boolean
    OpeningPriceIsSet = False

    'output/solution variables
    Dim Ticker As String
    Dim PriceChange As Double
    Dim PercentChange As Double
    Dim StockVolumeTotal As Double
    StockVolumeTotal = 0
    
    'Hard solution variables
    Dim GreatestPercentIncrease As Double
    Dim GreatestPercentDecrease As Double
    Dim GreatestTotalVolume As Double
    GreatestPercentIncrease = 0
    GreatestPercentDecrease = 0
    GreatestTotalVolume = 0
    
    'Hard solution helper variables
    Dim GPITicker As String
    Dim GPDTicker As String
    Dim GTVTicker As String
    
    
    For Each ws In Worksheets
    
        'Retrieve last row and last column as integers (Citation: Slack Channel # 01-class-activities)
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        'moderate solution headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'hard solution headers
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        'For all rows
        For i = 2 To LastRow
        
            'Grab Ticker String
            Ticker = ws.Cells(i, 1).Value
            
            'Add stock volume to running total
            StockVolumeTotal = StockVolumeTotal + ws.Cells(i, 7).Value
            
            
            'When turing is at first date for a given ticker
            If Ticker <> ws.Cells(i - 1, 1).Value Then
                
                'reset open price flag
                OpeningPriceIsSet = False
                
            End If
            
            'This next block accounts for opening prices that occur as zero in the data set until a later date, see README for details
            '
            'if open price flag is not set
            If OpeningPriceIsSet = False Then
            
                'grab the opening price at the current date
                OpeningPrice = ws.Cells(i, 3).Value
                
                'if the price is not 0
                If OpeningPrice <> 0 Then
                
                    'set the open price flag
                    OpeningPriceIsSet = True
                    
                End If
                
            End If
                
                
            
            
            'When turing reaches the last date for a given ticker
            If Ticker <> ws.Cells(i + 1, 1).Value Then
                
                'output ticker to solution
                ws.Cells(SolutionRow, 9).Value = Ticker
                
                'Grab the closing price at the end of the year
                ClosingPrice = ws.Cells(i, 6).Value
                'Calculate the price change and output it to solution
                PriceChange = ClosingPrice - OpeningPrice
                ws.Cells(SolutionRow, 10).Value = PriceChange
                
                'Format solution cell green if positive, red if negative
                If PriceChange >= 0 Then
                    ws.Cells(SolutionRow, 10).Interior.ColorIndex = 4
                    ws.Cells(SolutionRow, 11).Interior.ColorIndex = 4
                Else
                    ws.Cells(SolutionRow, 10).Interior.ColorIndex = 3
                    ws.Cells(SolutionRow, 11).Interior.ColorIndex = 3
                End If
                
                'Calculate PercentChange and output it to solution (only if non-zero divisor)
                If OpeningPrice <> 0 Then
                    PercentChange = (PriceChange / OpeningPrice)
                'if non-zero divisor, set to percentchange to zero
                Else
                    PercentChange = 0
                End If
                ws.Cells(SolutionRow, 11).Value = PercentChange
                
                'Output total stock volume to solution
                ws.Cells(SolutionRow, 12).Value = StockVolumeTotal
                
                
                'Update hard solutions if greater/smaller, and keep track of associated ticker
                If PercentChange > GreatestPercentIncrease Then
                    GreatestPercentIncrease = PercentChange
                    GPITicker = Ticker
                End If
                If PercentChange < GreatestPercentDecrease Then
                    GreatestPercentDecrease = PercentChange
                    GPDTicker = Ticker
                End If
                If StockVolumeTotal > GreatestTotalVolume Then
                    GreatestTotalVolume = StockVolumeTotal
                    GTVTicker = Ticker
                End If
                               
                'reset total stock volume
                StockVolumeTotal = 0
                
                'Increase solutionrow count for next solution
                SolutionRow = SolutionRow + 1
                
            End If
        
        Next i
        
        'Output hard solution values
        ws.Cells(2, 15).Value = GPITicker
        ws.Cells(3, 15).Value = GPDTicker
        ws.Cells(4, 15).Value = GTVTicker
        ws.Cells(2, 16).Value = GreatestPercentIncrease
        ws.Cells(3, 16).Value = GreatestPercentDecrease
        ws.Cells(4, 16).Value = GreatestTotalVolume
        
        
        'Format solution percentage column
        ws.Range("K2:K" & (SolutionRow - 1)).NumberFormat = "0.00%"
        
        'Format hard solution percentages
        ws.Range("P2:P3").NumberFormat = "0.00%"
        
        
        'Format column width
        ws.Columns("A:P").AutoFit
        
        'Reset solution row
        SolutionRow = 2
        
        'Reset hard solution variables
        GreatestPercentIncrease = 0
        GreatestPercentDecrease = 0
        GreatestTotalVolume = 0
        
    Next ws
    
End Sub

