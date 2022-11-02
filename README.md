# **Stock Analysis Refactor VBA**
[VBA_Challenge.xlsm](https://github.com/vzhang90/stock-analysis/blob/main/VBA_Challenge.xlsm)

## Overview of VBA Challenge
Having recently graduated with a finance degree, Steve is taking on his parents as his first clients. Steve thinks his parents' investments should be more diversified - so, he wants to analyze several green energy stocks in the uploaded excel file [VBA_Challenge.xlsm](https://github.com/vzhang90/stock-analysis/blob/main/VBA_Challenge.xlsm) 

The purpose of this module is to refactor (or editing) different subroutines to loop through all the data one time in a single macro in order to analyze the Volume and Return for each stock ticker depending on year.
 
## Results
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
3) At the end of 'AllStocksAnalysisRefactored' script before the `End Sub` command and after the last `Next i`, set the `endTime` variable equal to the `Timer` function. Finally, create a message box statement `MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)` to calculate elapsed time of the code by subtracting `startTime` from `endTime` to measure code performance.  
---
**Stock Performance & Elapsed Time for macro VBA Challenge 2017**
![VBA_Challenge_2017](https://github.com/vzhang90/stock-analysis/blob/main/VBA_Challenge_2017.png)  
  
**Stock Performance & Elapsed Time for macro VBA Challenge 2018**
![VBA_Challenge_2018](https://github.com/vzhang90/stock-analysis/blob/main/VBA_Challenge_2018.png)
  
2018 saw horrible returns compared to the overal positive returns in 2017.


## Summary
When refactoring code, there is not necessarily newly added functionality. The biggest benefit to refactoring code is making it more efficient by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. The biggest disadvantage to refactoring code is the syntax error from combining multiple macros into one. 

advantages and disadvantages of original and refactored VBA script
While originally separated into multiple macros across different modules, the original VBA scripts being compartmentalized into separate macros 
