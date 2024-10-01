# VBA-challenge
Module 2- Challenge

In this challenge, I was tasked with creating a VBA script that loops through all the stocks for each quarter in the Multiple_year_stock_data.xlsx spreadsheets and outputs the following:
-Ticker symbol
-Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
-The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.
-The total stock volume of the stock.

Additionally, I added functionality to return the stock with the Greatest % increase, Greatest % decrease, and Greatest Volume.
* Conditional formatting was also used to highlight positive change in green and negative change in red for "Quartely Change" and "Percent Change" columns

' Sub StockAnalysisCompleted()
-Defines the beginning of VBA subroutine. Contains the entire script that will execute when running macro.

' Variables
- ws As Worksheet: Declares a variable ws to represent each worksheet in the workbook. Allowing the script to loop through all the sheets.
- ticker As String: Stores the stock ticker symbol for each stock.
- openPrice As Double, closePrice As Double: Store the opening and closing prices for a stock in each quarter. These are named Double because they are numerical values.
- totalVolume As Double: Stores the sum of all stock volumes for the ticker over a quarter.
- quarterlyChange As Double: Holds the calculated change in price from the beginning to end of a quarter
- percentChange As Double: Holds the percentage change based on the opening and closing prices of the stock
- greatestIncrease, greatestDecrease, greatestVolume As Double: Store values of the greatest percentage increase/decrease and volume
- Store the ticker symbols for the stocks with the greatest increase/decrease and volume
- lastRow As Long, i As Long: lastRow stores the last row in the worksheet and i is used as the row iterator within loops
- startRow As Long, outputRow As Long: startRow tracks the row where a new ticker starts and outputRow tracks the current row for outputting results

  'For Each ws In Worksheets
    ws.Activate - Loop goes through each worksheet in the workbook one by one

   ' Set variables
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row: Decides the last non-empty row in the worksheet by counting from bottom of column 1(A) upwards. This gets the range of data to loop
        startRow = 2 : Being that the row starts at 2
        outputRow = 2 : Output will start from row 2 in columns I to L

  ' Output headers : Makes sure the column headers are placed in the correct column starting at column I (9)
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

' Loop through all rows
  For i = 2 To lastRow
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Or i = lastRow
  - This loop iterates over each row of the worksheet to process the stock data.

    ' Capture the ticker, open and close prices, and calculate volume
                ticker = ws.Cells(i, 1).Value : Retrieves the stock ticker from the first column
                openPrice = ws.Cells(startRow, 3).Value : Gets the opening price from the start of the quarter
                closePrice = ws.Cells(i, 6).Value : Gets closing price at the end of the quarter
                totalVolume = Application.WorksheetFunction.Sum(ws.Range("G" & startRow & ":G" & i)) : Sums up the total volume for the ticker from the first to current row

      ' Calculate quarterly change and percentage change
                quarterlyChange = closePrice - openPrice : Calculates the difference between the closing & opening prices
                If openPrice <> 0 Then
                    percentChange = (quarterlyChange / openPrice) * 100 : Calculate the % change in price.
                Else
                    percentChange = 0
                End If
    
  ' Output results for the ticker : Calculated values are output to columns I to L
                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = quarterlyChange
                ws.Cells(outputRow, 11).Value = percentChange
                ws.Cells(outputRow, 12).Value = totalVolume
                
    
  ' Move to the next output row : outputRow tracks the current row for output
                 outputRow = outputRow + 1

   ' Track the greatest increase, decrease, and volume : As the loop processes each stock, it checks if the current & change or volume is greater than the previous value. If so, it updates the greatest increase, decrease, or volume values with their correct tickers
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
    
   ' Reset for next ticker : After processing each stock, the startRow is updated to the row of the next ticker
                startRow = i + 1

   ' Output summary results : The script outputs the tickers with the greatest % increase, decrease, and greatest total volume to columns O to Q
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
- Loop and script ends
