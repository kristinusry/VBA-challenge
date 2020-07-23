Sub AnalyzeStockData()

Dim Stocks As Worksheet
    For Each Stocks In ActiveWorkbook.Worksheets
    Stocks.Activate

    'Add column headings
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
        
    'Create variables
    Dim tickerName As String
    Dim openPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim i As Long

    'Set initial values
    totalVolume = 0
    outputRow = 2
    openPrice = Cells(2, 3).Value
    
    'Set the last row of data
    lastRow = Stocks.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastRow
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

               'Set closing price
                closingPrice = Stocks.Cells(i, 6).Value

                'Calculate yearly change
                yearlyChange = closingPrice - openPrice
                Cells(outputRow, 10).Value = yearlyChange

                'Set ticker name
                tickerName = Cells(i, 1).Value
                Cells(outputRow, 9).Value = tickerName

                'Calculate percent change
                If (openPrice = 0 And closingPrice = 0) Then
                    percentChange = 0
                ElseIf (openPrice = 0 And closingPrice <> 0) Then
                    percentChange = 1
                Else
                    percentChange = yearlyChange / openPrice
                    Cells(outputRow, 11).Value = percentChange
                End If
                
                'Calculate total volume
                totalVolume = totalVolume + Cells(i, 7).Value
                Cells(outputRow, 12).Value = totalVolume

                'Add row and reset values: open price and total volume
                outputRow = outputRow + 1
                openPrice = Cells(i + 1, 3)
                totalVolume = 0

            Else
                totalVolume = totalVolume + Cells(i, 7).Value
            End If
        Next i
        
        'Set the last row of yearly change
        yearlyChangelastRow = Stocks.Cells(Rows.Count, 9).End(xlUp).Row
        'Set the cell colors
        For j = 2 To yearlyChangelastRow
            If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 4
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
        
        'Add Greatest increase, decrease and total volume headers
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        'Find the greatest value and its ticker name
        For k = 2 To yearlyChangelastRow
            If Cells(k, 11).Value = Application.WorksheetFunction.Max(Stocks.Range("K2:K" & yearlyChangelastRow)) Then
                Cells(2, 16).Value = Cells(k, 9).Value
                Cells(2, 17).Value = Cells(k, 11).Value
            ElseIf Cells(k, 11).Value = Application.WorksheetFunction.Min(Stocks.Range("K2:K" & yearlyChangelastRow)) Then
                Cells(3, 16).Value = Cells(k, 9).Value
                Cells(3, 17).Value = Cells(k, 11).Value
            ElseIf Cells(k, 12).Value = Application.WorksheetFunction.Max(Stocks.Range("L2:L" & yearlyChangelastRow)) Then
                Cells(4, 16).Value = Cells(k, 9).Value
                Cells(4, 17).Value = Cells(k, 12).Value
            End If
        Next k
        
    Next Stocks
        
End Sub
