# VBA of Wall Street - Challenge 2

## Overview of VBA Project
Steve recently began his career in the finance industry. 
His first job is to analyze 12 green energy stocks for his parents including DQ which is their current stock choice. 

### Purpose
Steve needs to be able to quickly analyze the entire dataset for stocks in 2017 and 2018.
Utilizing readable VBA code will enable him to automate tasks in the dataset while increasing the speed and the accuracy of results.
The VBA macro will return the "Total Daily Volume" and "Return" by ticker symbol for both 2017 and 2018.
The original code used to generate the output of "Total Daily Volume" and "Return" will be refactored to determine if the VBA script can run faster.

## Results

### Comparison of Stock Performance Between 2017 and 2018
Overall stock returns in 2017 where significantly higher than in 2018.
92% of the stocks yielded a positive return in 2017 while only 17% of stocks in 2018 yielded a positive return.
In 2018 the total daily volume for all the stocks was higher than in 2017, but I do not consider a 4% increase to be a material difference.
[VBA_Challenge.xlsm](VBA_Challenge.xlsm)
The "All Stocks Analysis" worksheet includes three buttons with macros assigned. 
The "Run Analysis for All Stocks" button is the original code, after selecting that button a message box will prompt the user to enter the year to run analysis on.
The "Challenge 2 - Refactored Stock Analysis" button is the refactored code, after selecting that button a message box will prompt the user to enter the year to run analysis on.
Both the original and refactored code utilize the code below which is the start of a for loop to loop over all the rows in the dataset. 
'For i = 2 To RowCount'
Within the for loop is an if-then statement which allows the program to calculate the starting and ending prices that are used in the "Results" calculation.

### Execution Times of Original Script and Refactored Script
The refactored code has faster execution times for both 2017 and 2018.
The average execution time for the refactored code was .115234 seconds.
The average execution time for the refactored code was .748047 seconds.
The refactored code was on average .632813 seconds faster than the original code.
The code used to calculate and display the codes run time is below.
'endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)'
Output messages with the execution time for the refactored code are included for 2017 and 2018.
[VBA_Challenge_2017.png](Resources/VBA_Challenge_2017.png)
[VBA_Challenge_2018.png](Resources/VBA_Challenge_2018.png)

## Summary

### Advantages and Disadvantages of Refactoring Code in General
#### Advantages
Refactored code should improve the design to help the program execute faster.
Making more efficient code can be accomplished by improving logic to remove bad code smell. 
It can decrease the number of steps, use less memory and make the code more clear.
#### Disadvantages
Refactoring code can be time consuming to complete.
There are risks that while refactoring new bugs can be introduced which could negatively impact the original code.

### Pros and Cons of Refactoring the Original Stock Analysis VBA script
#### Pros
For this project refactoring the original VBA script decreased the execution time.
The refactored code utilizes arrays for tickers, tickerVolumes, tickerStartingPrices and tickerEndingPrices.
The variable tickerIndex is used to access the stocker ticker index for the arrays.
#### Cons
Refactoring the original code was time consuming.
During the process I encountered multiple error messages which had to be corrected.
For the current dataset the increased execution speed does not create a noticeable benefit.