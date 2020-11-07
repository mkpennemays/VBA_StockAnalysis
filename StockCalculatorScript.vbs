Sub StockCalculator()
Dim currStock As String
Dim yearOpenPrice As Double
Dim yearEndPrice As Double
Dim yearPriceChange As Double
Dim yearPercentChange As Double


Dim stockYearTotalVolume As Double

Dim reportRow As Double
Dim tickerCol, yrChangeCol, yrPercentCol, yrTotalVolCol As Integer
tickerCol = 9
yrChangeCol = 10
yrPercentCol = 11
yrTotalVolCol = 12

Dim highestTotalVolume, highestIncrease, highestDecrease As Double
Dim highTotalVolumeStock, highIncreaseStock, highDecreaseStock As String



Dim lRow As Double
lRow = Cells(Rows.Count, 1).End(xlUp).Row
Dim i As Double
Dim ws As Worksheet

For Each ws In Worksheets

    ws.Activate

    'create headers
    Cells(1, tickerCol).Value = "Stock"
    Cells(1, yrChangeCol).Value = "Year Price Change"
    Cells(1, yrPercentCol).Value = "Year Percent Change"
    Cells(1, yrTotalVolCol).Value = "Year Total Volume"

    Cells(1, yrTotalVolCol + 3).Value = "Ticker"
    Cells(1, yrTotalVolCol + 4).Value = "Value"
    
    Cells(2, yrTotalVolCol + 2).Value = "Greatest Percent Increase"
    Cells(3, yrTotalVolCol + 2).Value = "Greatest Percent Decrease"
    Cells(4, yrTotalVolCol + 2).Value = "Greatest Total Volume"

    'initialize beginning values
    currStock = Cells(2, 1).Value
    yearOpenPrice = Cells(2, 3).Value
    yearPriceChange = 0
    stockYearTotalVolume = 0
    reportRow = 2
    highestTotalIncrease = 0
    highestIncrease = 0
    highestDecrease = 0

    For i = 2 To lRow
        stockYearTotalVolume = stockYearTotalVolume + Cells(i, 7).Value
        If Cells(i + 1, 1).Value <> currStock Then
            
            'we have arrived at a new stock, calculate and post values
            Cells(reportRow, tickerCol).Value = currStock
            Cells(reportRow, yrTotalVolCol).Value = stockYearTotalVolume
            Cells(reportRow, yrPercentCol).NumberFormat = "0.00%"
            If (Cells(i, 6).Value > yearOpenPrice) Then
                yearPriceChange = Cells(i, 6).Value - yearOpenPrice
                Cells(reportRow, yrChangeCol).Value = yearPriceChange
                Cells(reportRow, yrChangeCol).Interior.ColorIndex = 4
                yearPercentChange = yearPriceChange / yearOpenPrice
                Cells(reportRow, yrPercentCol).Value = yearPercentChange
            ElseIf (Cells(i, 6).Value = yearOpenPrice) Then
                Cells(reportRow, yrChangeCol).Value = 0
                Cells(reportRow, yrPercentCol).Value = 0
                yearPercentChange = 0
            Else
                yearPriceChange = (yearOpenPrice - Cells(i, 6).Value) * -1
                Cells(reportRow, yrChangeCol).Value = yearPriceChange
                Cells(reportRow, yrChangeCol).Interior.ColorIndex = 3
                yearPercentChange = (yearPriceChange / yearOpenPrice)
                Cells(reportRow, yrPercentCol).Value = yearPercentChange

            End If
            
            If stockYearTotalVolume > highestTotalVolume Then
                highestTotalVolume = stockYearTotalVolume
                highTotalVolumeStock = currStock
            End If
            If yearPercentChange > highestIncrease Then
                highestIncrease = yearPercentChange
                highIncreaseStock = currStock
            End If
            If yearPercentChange < highestDecrease Then
                highestDecrease = yearPercentChange
                highDecreaseStock = currStock
            End If
            
            'reinitialize
            reportRow = reportRow + 1
            currStock = Cells(i + 1, 1).Value
            stockYearTotalVolume = 0
            yearOpenPrice = Cells(i + 1, 3).Value
            
        Else
        'same stock - check for the case where opening price is zero, but have found a start price
            If yearOpenPrice = 0 And Cells(i, 3).Value > 0 Then
                yearOpenPrice = Cells(i, 3).Value
            End If
           
         End If
    
    
    Next i
    'display additional statistics
    Cells(2, yrTotalVolCol + 3).Value = highIncreaseStock
    Cells(2, yrTotVolColal + 4).NumberFormat = "0.00%"
    Cells(2, yrTotalVolCol + 4).Value = highestIncrease
            
    Cells(3, yrTotalVolCol + 3).Value = highDecreaseStock
    Cells(3, yrTotalVolCol + 4).NumberFormat = "0.00%"
    Cells(3, yrTotalVolCol + 4).Value = highestDecrease
            
    Cells(4, yrTotalVolCol + 3).Value = highTotalVolumeStock
    Cells(4, yrTotalVolCol + 4).Value = highestTotalVolume
    

Next 'worksheet

End Sub
