# stock-analysis
## Project Overview
The purpose of this project is to assist a financial advisor in stock market analysis for green energy stocks. We've written and refactored code in Visual Basic Applications(VBA) to efficiently analyze stock data. The original code (green_stocks.xlsm) is sufficient for analyzing a small number of stocks, however, in order to analyze thousands, we refactored it to use less memory and improve efficiency (VBA_challenge.xlsm). 

## Resources
Data Source: green_stocks.xlsm
- Software: Microsoft Excel 16.66.1, Visual Studio Code, 1.70.2

## Results
1. The tickerIndex is set equal to zero before looping over the rows. (5 pt).

2. Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices (15 pt).

3. The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays (15 pt).

4. The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices (25 pt).

5. Code for formatting the cells in the spreadsheet is working (5 pt).

6. There are comments to explain the purpose of the code (5 pt).

7. The outputs for the 2018 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module (5 pt).

8. The pop-up messages showing the elapsed run time for the script are saved as VBA_Challenge_2018.png and VBA_Challenge_2018.png (5 pt).

## Challenge Summary
Advantages of refactoring code: 

Disadvantages of refactoring code: 

There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).
