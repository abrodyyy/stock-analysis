# stock-analysis
## Project Overview
The purpose of this project is to assist a financial advisor in stock market analysis for green energy stocks. We've written and refactored code in Visual Basic Applications(VBA) to efficiently analyze the stock data. The original code (green_stocks.xlsm) is sufficient for analyzing a small number of stocks, however, in order to analyze thousands, we refactored it to use less memory and improve efficiency (VBA_challenge.xlsm). 

## Resources
Data Source: green_stocks.xlsm
- Software: Microsoft Excel 16.66.1, Visual Studio Code, 1.70.2

## Results: Refactor VBA code and measure performance
> 1. The tickerIndex is set equal to zero before looping over the rows.

    1b) Create a ticker Index
    tickerIndex = 0
    
    1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

> 2. Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

    2a) Create a for loop to initialize the tickerVolumes to zero.
    For j = 0 To 11
        tickerVolumes(j) = 0
    Next j

    2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount

> 3. The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays.

        3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        3c) check if the current row is the last row with the selected ticker. If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

            3d) Increase the tickerIndex.
            tickerIndex = tickerIndex + 1

        End If

    Next i
    
> 4. The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices.

    4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    Next i

> 5. Code for formatting the cells in the spreadsheet is working.

    5) Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
        Else
            Cells(i, 3).Interior.Color = vbRed
        End If
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    End Sub

2017 

![VBA_green_stocks_2017](https://user-images.githubusercontent.com/111623064/197681265-a885b73b-5c03-4e2d-989c-d137386abc11.png)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/111623064/197681235-1b6b67f9-a176-4bf6-81b7-6c20a3dc147d.png)

While both versions of code produce the same results, we can determine that the refactored code ran about 0.2 seconds faster than the original. 

2018

![VBA_green_stocks_2018](https://user-images.githubusercontent.com/111623064/197681322-d9596cc5-b428-4381-b5f5-f93c3494cd73.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/111623064/197681306-a4652268-e0a8-4dd8-99ac-86abb9fd1ba1.png)

While both versions of code produce the same results, we can determine that the refactored code ran about 0.2 seconds faster than the original. 

## Challenge Summary
Advantages of refactoring code: 

Disadvantages of refactoring code: 

There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).
