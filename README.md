# stock-analysis
## Project Overview
The purpose of this project is to assist a financial advisor in stock market analysis for green energy stocks. We've written and refactored code in Visual Basic Applications(VBA) to efficiently analyze stock data. The original code (green_stocks.xlsm) is sufficient for analyzing a small number of stocks, however, in order to analyze thousands, we refactored it to use less memory and improve efficiency (VBA_challenge.xlsm). 

## Resources
Data Source: green_stocks.xlsm
- Software: Microsoft Excel 16.66.1, Visual Studio Code, 1.70.2

## Results: Refactor VBA code and measure performance
> 1. The tickerIndex is set equal to zero before looping over the rows. (5 pt).

    1b) 1a) Create a ticker Index
    tickerIndex = 0
    1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

> 2. Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices (15 pt).

    2a) Create a for loop to initialize the tickerVolumes to zero.
    For j = 0 To 11
        tickerVolumes(j) = 0
    Next j

    2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount

> 3. The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays (15 pt).

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
    
> 4. The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices (25 pt).

> 5. Code for formatting the cells in the spreadsheet is working (5 pt).

> 6. There are comments to explain the purpose of the code (5 pt).

> 7. The outputs for the 2018 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module (5 pt).

> 8. The pop-up messages showing the elapsed run time for the script are saved as VBA_Challenge_2018.png and VBA_Challenge_2018.png (5 pt).

## Challenge Summary
Advantages of refactoring code: 

Disadvantages of refactoring code: 

There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).
