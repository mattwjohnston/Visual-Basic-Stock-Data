Sub VolumeLoop()

    Dim Total As Double
    SymbolRow = 2
    ResultRow = 2
    StockSymbol = Cells(SymbolRow, 1).Value
    NextStock = Cells(SymbolRow + 1, 1).Value
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Total Stock Volume"

    Do While StockSymbol <> ""

        ' adds the volume of the day to the total for that stock
        Total = Total + Cells(SymbolRow, 7).Value

        ' checks if the next symbol is different from this one
        If StockSymbol <> NextStock Then
            ' if its different, it inserts the stock value, then the volume total,
            ' resets total, then move onto the next line in the results columns
            Cells(ResultRow, 9).Value = Cells(SymbolRow, 1).Value
            Cells(ResultRow, 10).Value = Total
            Total = 0
            ResultRow = ResultRow + 1

        End If

        SymbolRow = SymbolRow + 1
        StockSymbol = Cells(SymbolRow, 1).Value
        NextStock = Cells(SymbolRow + 1, 1).Value

    Loop

End Sub
