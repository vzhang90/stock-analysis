# **Stock Analysis Refactor VBA**
[VBA_Challenge.xlsm](https://github.com/vzhang90/stock-analysis/blob/main/VBA_Challenge.xlsm)

## Overview of VBA Challenge
Having recently graduated with a finance degree, Steve is taking on his parents as his first clients. Diversification is of utmost importance in his investment outlook - so, he wants to analyze data of several green energy stocks in the [uploaded excel file.](https://github.com/vzhang90/stock-analysis/blob/main/VBA_Challenge.xlsm) 

**Purpose of this module:** refactor different VBA macros to loop through all data one time in a single subroutine, in order to compile analysis on the stocks.
 
## Results
The refactored code now analyzes *Volume* and *Return* for each stock ticker depending on year in a single macro.

As datasets become larger, it is important to know how fast the VBA code will compile results. To help Steve perform analysis on larger datasets in the future, a script to capture start & end time of the executed code is added to measure this VBA code performance. 
1) Initialize variables `startTime` and `endTime` as `Single` data types underneath the `sub AllStocksAnalysisRefactored()` subroutine. 
2) To start the timer after entering the year in the `InputBox90`, set the `startTime` variable equal to the `Timer` function underneath/after the `yearValue`variable set  
```
Sub AllStocksAnalysis()

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

       startTime = Timer
```             
3) At the end of 'AllStocksAnalysisRefactored' script before the `End Sub` command and after the last `Next i`, set the `endTime` variable equal to the `Timer` function. 
4) Finally, create a message box statement `MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)` to calculate elapsed time of the code by subtracting `startTime` from `endTime` to measure code performance.  
---
**Stock Performance & Elapsed Time for macro VBA Challenge 2017**
![VBA_Challenge_2017](https://github.com/vzhang90/stock-analysis/blob/main/VBA_Challenge_2017.png)  
  
**Stock Performance & Elapsed Time for macro VBA Challenge 2018**
![VBA_Challenge_2018](https://github.com/vzhang90/stock-analysis/blob/main/VBA_Challenge_2018.png)
  
2018 saw horrible overall returns for a majority of green energy stocks compared to those of 2017 (which saw overall positive returns). The only 2 stocks to post consecutive positive returns in 2017 to 2018 are stock tickers ***ENPH*** and ***RUN***. On the other hand, ***TERP*** is the only stock to have negative returns in both 2017 & 2018.


## Summary
The general advantage to refactoring is improving a body of code's internal structure into one efficient subroutine. While aggregating various macros into one, refactoring makes code more efficient by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. The biggest disadvantage to refactoring code is the syntax error from combining multiple macros into one. 

The original VBA scripts being compartmentalized into separate subroutines is cumbersome to sort through. Besides increased execution speed compiling data, another big advantage to refactored VBA scripts vs the original VBA scripts is organized code. 
