Sub VolumeLoop()

    Dim Total As Double
    Dim YearOpen As Double
    Dim YearClose As Double
    Dim BestIncrease As Double
    Dim BestIncreaseSym As String
    Dim BestDecrease As Double
    Dim BestDecreaseSym As String
    Dim BestVolume As Double
    Dim BestVolumeSym As String

    SymbolRow = 2
    ResultRow = 2
    StockSymbol = Cells(SymbolRow, 1).Value
    NextStock = Cells(SymbolRow + 1, 1).Value
    YearOpen = Cells(SymbolRow, 3).Value
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Range("K:K").NumberFormat = "0.00%"

    Do While StockSymbol <> ""

        ' adds the volume of the day to the total for that stock
        Total = Total + Cells(SymbolRow, 7).Value

        ' checks if the next symbol is different from this one
        If StockSymbol <> NextStock Then


            ' if its different, it inserts the stock value, then the volume total,
            ' resets total, then move onto the next line in the results columns
            Cells(ResultRow, 9).Value = Cells(SymbolRow, 1).Value
            Cells(ResultRow, 12).Value = Total

            YearClose = Cells(SymbolRow, 6).Value

            If YearClose = 0 Then
                i = SymbolRow
                Do While YearClose = 0
                    YearClose = Cells(i, 6).Value
                    i = i - 1
                Loop
            End If

            Cells(ResultRow, 10).Value = YearClose - YearOpen

            If Cells(ResultRow, 10).Value > 0 Then
                Cells(ResultRow, 10).Interior.Color = vbGreen
            Else
                Cells(ResultRow, 10).Interior.Color = vbRed
            End If





            Cells(ResultRow, 11).Value = Cells(ResultRow, 10).Value / YearOpen

            If Cells(ResultRow, 11).Value > BestIncrease Then
                BestIncrease = Cells(ResultRow, 11).Value
                BestIncreaseSym = Cells(ResultRow, 9).Value
            ElseIf Cells(ResultRow, 11).Value < BestDecrease Then
                BestDecrease = Cells(ResultRow, 11).Value
                BestDecreaseSym = Cells(ResultRow, 9).Value
            End If

            If Cells(ResultRow, 12).Value > BestVolume Then
                BestVolume = Cells(ResultRow, 12).Value
                BestVolumeSym = Cells(ResultRow, 9).Value
            End If

            If NextStock = "" Then
                Exit Do
            End If

            YearOpen = Cells(SymbolRow + 1, 3).Value
            If YearOpen = 0 Then
                i = SymbolRow + 2
                Do While YearOpen = 0
                    YearOpen = Cells(i, 3).Value
                    i = i + 1
                Loop
            End If
            Total = 0
            ResultRow = ResultRow + 1

        End If

        SymbolRow = SymbolRow + 1
        StockSymbol = Cells(SymbolRow, 1).Value
        NextStock = Cells(SymbolRow + 1, 1).Value

    Loop

    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(2, 15).Value = BestIncreaseSym
    Cells(2, 16).Value = BestIncrease
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(3, 15).Value = BestDecreaseSym
    Cells(3, 16).Value = BestDecrease
    Cells(4, 14).Value = "Greatest Total Volume"
    Cells(4, 15).Value = BestVolumeSym
    Cells(4, 16).Value = BestVolume


End Sub
