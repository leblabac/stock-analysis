Click here to view the file: [VBA_Challenge.xlsm](https://github.com/leblabac/stock-analysis/blob/main/VBA_Challenge.xlsm)

# Overview of Project
The purpose of this project was to refactor a Microsoft Excel VBA code to collect stock information for the years 2017 and 2018, and to determine whether or not the stocks are worth investing in. This process was originally completed within the green-stocks.xlsm file using multiple subroutines, however, the the task for this challenge was to increase the efficiency of the original code.

## Data
The data provided included two tables with stock information on 12 different stocks. Each stock had the following data: 
- a unique ticker value
- the date the stock was issued
- the opening price for a given date
- the high price for a given date
- the low price for a given date
- the closing price for a given date
- the adjusted close for a given date
- and the volume of the stock traded

The goal of the assignment was to refactor created code designed to locate/retrieve the ticker, to retrieve it's total volume traded for a given year, and to return its rate of return.

## Results
In order to analyze the stock performance for 2017 and 2018, using VBA, a worksheet was created and a title and table headers created, after which the series of stock tickers were initialized.  An InputBox was used to retrieve the year for which stock analysis was preferred.

```
yearValue = InputBox("What year would you like to run the analysis on?")
```

After creating a ticker index and output arrays for the tickerVolumes and ticker starting and ending prices, a nested loop was created to retrieve and increase the tickerVolumes, based on an If-Then condition. These were designed to look for the initial ticker row, one to look for the ending row of the same ticker in order to include these in the scope of the calculation, and then to move on to the next ticker index.

```
'3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
    '3c) check if the current row is the last row with the selected ticker
        'If the next row's ticker not match, increase the tickerIndex.
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
         
    '3d Increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
```        
Lastly, the values gathered by the code were outputted into a table, and then formatted to provide easy-glance interpretation of stocks whose returns were performing well versus those which were not.













as well as the execution times of the original script and the refactored script.






## Summary






### Pros and Cons of Refactoring Code
In terms of the pros and cons of refactoring code, the pros are that: refactored code has the ability to provide good debugging of code, increase the performance of the code, and, if instruction comments are written well, the ability to understand the code in a better way.  The cons of refactoring code are that it may take time to do - and its a con if the "cost" of refactoring the code is more "expensive" than actually keeping the original code.  Also, refactoring later in the code developemnt process can lead to less testing if a deadline is approaching, so may not be desirable if that is the case.











