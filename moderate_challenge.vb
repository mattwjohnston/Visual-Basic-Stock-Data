Sub VolumeLoop()

    Dim Total As Double
    Dim YearOpen As Double
    Dim YearClose As Double
    SymbolRow = 2
    ResultRow = 2
    StockSymbol = Cells(SymbolRow, 1).Value
    NextStock = Cells(SymbolRow + 1, 1).Value
    YearOpen = Cells(SymbolRow, 3).Value
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"

    Do While StockSymbol <> ""

        ' adds the volume of the day to the total for that stock
        Total = Total + Cells(SymbolRow, 7).Value

        ' checks if the next symbol is different from this one
        If StockSymbol <> NextStock Then
            ' if its different, it inserts the stock value, then the volume total,
            ' resets total, then move onto the next line in the results columns
            Cells(ResultRow, 9).Value = Cells(SymbolRow, 1).Value
            Cells(ResultRow, 12).Value = Total

            YearClose = Cells(SymbolRow, 3).Value
            Cells(ResultRow, 10).Value = YearClose - YearOpen

            If Cells(ResultRow, 10).Value > 0 Then
                Cells(ResultRow, 10).Interior.Color = vbGreen
            Else
                Cells(ResultRow, 10).Interior.Color = vbRed
            End If


            If YearOpen = 0 Then
                YearOpen = 1E-09
            End If

            Cells(ResultRow, 11).Value = Cells(ResultRow, 10).Value / YearOpen

            YearOpen = Cells(SymbolRow + 1, 3).Value

            Total = 0
            ResultRow = ResultRow + 1

        End If

        SymbolRow = SymbolRow + 1
        StockSymbol = Cells(SymbolRow, 1).Value
        NextStock = Cells(SymbolRow + 1, 1).Value

    Loop

End Sub
