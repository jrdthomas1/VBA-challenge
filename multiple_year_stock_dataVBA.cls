Sub StockAnalysisCompleted()
    Dim ws As Worksheet
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim increaseTicker As String
    Dim decreaseTicker As String
    Dim volumeTicker As String
    Dim lastRow As Long
    Dim i As Long
    Dim outputRow As Long
    Dim startRow As Long

    For Each ws In Worksheets
        ws.Activate

        ' Set variables
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        startRow = 2
        outputRow = 2
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0

        ' Output headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        ' Loop through all rows
        For i = 2 To lastRow
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Or i = lastRow Then
                ' Capture the ticker, open and close prices, and calculate volume
                ticker = ws.Cells(i, 1).Value
                openPrice = ws.Cells(startRow, 3).Value
                closePrice = ws.Cells(i, 6).Value
                totalVolume = Application.WorksheetFunction.Sum(ws.Range("G" & startRow & ":G" & i))
                
                ' Calculate quarterly change and percentage change
                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = (quarterlyChange / openPrice) * 100
                Else
                    percentChange = 0
                End If
                
                ' Output results for the ticker
                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = quarterlyChange
                ws.Cells(outputRow, 11).Value = percentChange
                ws.Cells(outputRow, 12).Value = totalVolume
                
                ' Move to the next output row
                outputRow = outputRow + 1

                ' Track the greatest increase, decrease, and volume
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    increaseTicker = ticker
                End If
                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    decreaseTicker = ticker
                End If
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    volumeTicker = ticker
                End If

                ' Reset for next ticker
                startRow = i + 1
                totalVolume = 0
            End If
        Next i

        ' Output summary results
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = increaseTicker
        ws.Cells(2, 17).Value = greatestIncrease & "%"
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = decreaseTicker
        ws.Cells(3, 17).Value = greatestDecrease & "%"
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = volumeTicker
        ws.Cells(4, 17).Value = greatestVolume
        
    Next ws
End Sub
