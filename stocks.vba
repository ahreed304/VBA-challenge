Sub stocks()
     Cells(1, "J") = "Ticker"
     Cells(1, "K") = "Yearly Change"
     Cells(1, "L") = "Percent Change"
     Cells(1, "M") = "Total Stock Volume"
     
     Dim ticker, greatestPercentDecreaseTicker, greatestPercentIncreaseTicker, greatestTotalVolumeTicker As String
     Dim yearOpenPrice, yearEndPrice, volumeTotal, greatestPercentIncrease, greatestPercentDecrease, greatestTotalVolume As Double
     Dim stockCounter, i As Long
     stockCounter = 2
     volumeTotal = 0
     greatestPercentIncrease = 0
     greatestPercentDecrease = 0
     greatestTotalVolume = 0
     ticker = ""
     lastRow = Cells(Rows.Count, "A").End(xlUp).Row
     
     For i = 2 To lastRow
        lastTicker = ticker
        ticker = Cells(i, "A").value
        nextTicker = Cells(i + 1, "A").value
        
        volumeTotal = volumeTotal + Cells(i, "G")
        
        If ticker <> lastTicker Then
            yearOpenPrice = Cells(i, "C")
        End If
        
        If ticker <> nextTicker Then
            yearClosePrice = Cells(i, "F").value
            
            Cells(stockCounter, "J") = ticker
            yearlyChange = yearClosePrice - yearOpenPrice
            If yearOpenPrice <> 0 Then
                percentageChange = yearlyChange / yearOpenPrice
            Else
                percentageChange = "undefined"
            End If
            Dim color As Integer
            If yearlyChange > 0 Then
                color = 4
            ElseIf yearlyChange < 0 Then
                color = 3
            Else
                color = 0
            End If
            
            Cells(stockCounter, "K") = yearlyChange
            Cells(stockCounter, "K").Interior.ColorIndex = color
            Cells(stockCounter, "L") = percentageChange
            Cells(stockCounter, "L").Interior.ColorIndex = color
            Cells(stockCounter, "L").NumberFormat = "0.00%"
            Cells(stockCounter, "M") = volumeTotal
            
            If percentageChange > greatestPercentIncrease And percentageChange <> "undefined" Then
                greatestPercentIncrease = percentageChange
                greatestPercentIncreaseTicker = ticker
            End If
            If percentageChange < greatestPercentDecrease And percentageChange <> "undefined" Then
                greatestPercentDecrease = percentageChange
                greatestPercentDecreaseTicker = ticker
            End If
            If volumeTotal > greatestTotalVolume Then
                greatestTotalVolume = volumeTotal
                greatestTotalVolumeTicker = ticker
            End If
            volumeTotal = 0
            stockCounter = stockCounter + 1
        End If
    Next
    
    Cells(1, "P") = "Ticker"
    Cells(1, "Q") = "Value"
    
    Cells(2, "O") = "Greatest % Increase"
    Cells(2, "P") = greatestPercentIncreaseTicker
    Cells(2, "Q") = greatestPercentIncrease
    Cells(2, "Q").NumberFormat = "0.00%"
    
    Cells(3, "O") = "Greatest % Decrease"
    Cells(3, "P") = greatestPercentDecreaseTicker
    Cells(3, "Q") = greatestPercentDecrease
    Cells(3, "Q").NumberFormat = "0.00%"
    
    Cells(4, "O") = "Greatest Total Volume"
    Cells(4, "P") = greatestTotalVolumeTicker
    Cells(4, "Q") = greatestTotalVolume
    
    Range("J:Q").Columns.AutoFit
End Sub

