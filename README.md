# Stocks Analysis with VBA

## Overview of Project

The scenario here is we have a friend, Steve, who recently graduated from college and now works as a financial planner.  Steve’s parents are his first customer and want to invest fully in green energy.  More specifically, they want to invest fully in one stock named DAQO New Energy Corp.  The reason is the stock’s ticker symbol is DQ, which reminds Steve’s parents of when they first met at Dairy Queen.  Steve, however, is concerned about the lack of diversification.

### Purpose

The purpose of this analysis was to help Steve do his due diligence by evaluating DAQO’s performance and comparing it against other green energy stocks.  We helped make Steve’s task easier by writing macros to automate some of his work.  We also refactored the code to make it run more efficiently.  The resulting refactored macro is named AllStocksAnalysis and is in the file VBA_Challenge.xlsm located [here](https://github.com/mshideler/Stocks-Analysis.git).

## Results

In most cases, performance for the green energy stocks evaluated here was great in 2017 but bad in 2018, which shows how volatile these stocks are.  Below is the output from the macro we wrote.

![Stock Performance](https://github.com/mshideler/Stocks-Analysis/blob/main/Resources/Stock_Performance.png)

The VBA code used in the analysis macro initially focused on using loops and conditionals to save values to variables to output at the end (see code below).
```
    For i = 0 To 11
        
        ticker = tickers(i)
        totalVolume = 0
        
        '5) Loop through rows in the data.
        Sheets(yearValue).Activate
        
        For j = 2 To RowCount
        
            '5a) Find the total volume for the current ticker.
            If Cells(j, 1).Value = ticker Then
            
               totalVolume = totalVolume + Cells(j, 8).Value
               
            End If
        
            '5b) Find the starting price for the current ticker.
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                
                startingPrice = Cells(j, 6).Value
                
            End If
                           
            '5c) Find the ending price for the current ticker.
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
                endingPrice = Cells(j, 6).Value
                
            End If
            
        Next j
        
    '6) Output the data for the current ticker.Worksheets("All Stocks Analysis").Activate
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
    Next i
```

After refactoring, the VBA code used no nested loops and more arrays (see code below).
```
    For i = 2 To RowCount
    
        
        '3a) Increase ticker volumes
        If Cells(i, 1).Value = tickers(tickerIndex) Then
        
            'MsgBox (tickers(tickerIndex))
            
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            
            'MsgBox (tickerVolumes(tickerIndex))
        
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
'        If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

        End If

        '3c) check if the current row is the last row with the selected ticker
        'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

        End If
        
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1

        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11

        Worksheets("All Stocks Analysis").Activate

        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1

    Next i

```

Also included in both versions of the analysis code was a timer to show how long it took to execute the code.  

### Green Stocks (Original Code)

![Green Stocks 2017](https://github.com/mshideler/Stocks-Analysis/blob/main/Resources/Green_Stock_2017.png)	![Green Stocks 2018](https://github.com/mshideler/Stocks-Analysis/blob/main/Resources/Green_Stock_2018.png)

### VBA Challenge (Refactored Code)

![VBA Challenge 2017](https://github.com/mshideler/Stocks-Analysis/blob/main/Resources/VBA_Challenge_2017.png)	![VBA Challenge 2018](https://github.com/mshideler/Stocks-Analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary

1.  What are the advantages or disadvantages of refactoring code?

    - One advantage is refactored code takes less time to execute.
  
    - Another advantage of refactoring code is making it easier to read so another person reviewing the code is able to understand it more easily than code that hasn’t been refactored.
    
    - A couple of disadvantages I stumbled across include the possibility of breaking the code and the amount of time to troubleshoot.  If taken too far, these disadvantages may defeat the purpose of refactoring.

2.  How do these pros and cons apply to refactoring the original VBA script?

    - One pro that resulted from refactoring the original VBA script can be seen in the table with screenshots. Prior to refactoring, the analysis code took longer to run and didn’t include any formatting code.  The refactored code took less time and included formatting code.  Also, I felt the use of arrays instead of variables made the code more readable and easier to manipulate the data.
    
    - The cons inadvertently caused a speed bump in my progress.  Making several changes to refactor the code temporarily broke my for loop, which then took time to troubleshoot.  I ended up reverting to the code I originally used because I felt to continue troubleshooting would not have been productive.


