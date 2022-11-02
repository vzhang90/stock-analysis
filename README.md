# **Stock Analysis Refactor VBA**
[VBA_Challenge](https://github.com/vzhang90/stock-analysis/blob/main/VBA_Challenge.xlsm)

## Overview of VBA Challenge
Steve has recently graduated with his finance degree and taking on his parents as his first clients. While his parents want to invest all their money into DAQO New Energy Corporation, a company that makes silicon wafers for solar panels, Steven this his parents' investments should be more diversified. So, he wants to analyze several green energy stocks in the excel file [VBA_Challenge](https://github.com/vzhang90/stock-analysis/blob/main/VBA_Challenge.xlsm) using VBA. 

The purpose of this module is to better understand the process of refactoring (or editing) different codes to loop through all the data one time in order to collect the same information across multiple macros. 
 
## Results

As datasets become larger, it is important to know how fast the VBA code will compile results.  

To help Steve perform analysis on larger datasets in the future, a script to catpure start & end time of the executed code is added to measure this VBA code performance. To caculate how long the code takes to execute and output the elapsed time in a message box:
>Initialize variables `startTime` and `endTime` as `Single` data types underneath the `sub AllStocksAnalysisRefactored()` subroutine. To start the timer after entering the year in the `InputBox90`, set the `startTime` variable equal to the `Timer` function underneath/after the `yearValue`variable set. At the end of 'AllStocksAnalysisRefactored' script before the `End Sub` command and after the last `Next i`, set the `endTime` variable equal to the `Timer` function. Finally, create a message box statement `MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)` to calculate elapsed time of the code by subtracting `startTime` from `endTime` to measure code performance.  

**Elapsed Time for macro VBA Challenge 2017
![VBA_Challenge_2017](https://github.com/vzhang90/stock-analysis/blob/main/VBA_Challenge_2017.png)

**Elapsed Time for macro VBA Challenge 2018
![VBA_Challenge_2018](https://github.com/vzhang90/stock-analysis/blob/main/VBA_Challenge_2018.png)

## Summary
advantages and disadavntages of refactoring code in general
When refactoring code, there is not necessarily newly added functionality. The biggest benefit to refactoring code is making it more efficient by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.

advantages and disadvantages of original and refactored VBA script
While originally separated into multiple macros across different modules, the original VBA scripts being compartmentalized into separate macros 
