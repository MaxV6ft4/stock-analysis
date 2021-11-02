# Analysis of Energy Stock Performance

## Overview
Steve supplied me with alternative energy stock data for the years 2017 and 2018, as his parents expressed the desire to invest in clean energy.  At first they were only interested in Daqo (ticker:DQ) but the stock did not perform well in 2018 (dropped 63% from 2017), so they want to see how other alternative energy stocks fared both in 2017 and in 2018.  

### Purpose
For this project I wrote a macro in VBA called AllStocksAnalysis to display the total daily volumes and yearly returns for each stock during the two years provided.  I also created a run button in the output sheet for Steve to run the analysis on his own.  The code worked, but it involved looping through each individual row of stock.  Therefore I have refactored the code and have created a second run button specifically for this code.  The output displays the same results as those in the original code, but does so in a shorter amount of time.  This report will strictly cover the refactored code.

## Analysis

[Energy Stock Analysis](https://github.com/MaxV6ft4/stock-analysis/blob/main/VBA_Challenge.xlsm)

### Index and Arrays
I created a new macro in VBA called AllStocksAnalysisRefactored, maintaining both the same output formatting and the same array of tickers as was written in the original code (the tickers array's data type is string, since it simply displays text).  However, in this code I looped over the entire data *at once* instead of row by row.  To do this I created a ticker index, a variable that holds all the tickers inside itself, and initialized it to zero since in VBA the first number in an index is always zero.  Next, I created three new output arrays, one for ticker volumes (data type: long, since the numbers will be in the hundreds of millions), one for starting prices and one for ending prices (data types: single, meaning only one number to the right of the decimal point).

### Loops
I created three separate loops: the first to strictly loop over the volumes array, the second to loop over all the rows in the two data sheets (2017 and 2018), and the third to loop over the output arrays.

#### First loop
The first loop initalized the ticker volumes array to zero, allowing it to reset once the loop reached the next ticker (represented by the letter i here).

    For i = 0 To 11
    
          tickerVolumes(i) = 0
          
    Next i
        

#### Second loop
The second loop is key.  For each ticker, I increased the volume by adding the inital volume (zero) to the value of each cell in colummn H of the two data sheets.  The difference is that in the original code I added the volume to the value of the cells in H individually by checking each row to make sure it contained the proper ticker.  Here, the volumes are looped over all at once because, as an array they *immediately have access to the entire ticker index*, allowing them to instantly locate the correct ticker.  Using the letter i to represent the number of rows, I began the loop as follows:

    For i = 2 To RowCount
    
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

All I had to do afterwards was make sure that the row selected in the data sheet was the first row containing the desired ticker by checking to see if the next row up did not contain that ticker.  If so, the selected row is where the starting price (in Column F of the data sheets) for the ticker began.  

        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                

Similarly, I made sure that the row selected in the sheet was the last row containing the desired ticker by checking to see if the next row down did not contain that ticker.  If so, the selected row would contain the ending price for the ticker.  

        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

Remember, the starting and ending prices are arrays as well, so just like the volumes array they can instantly locate the correct ticker.  To allow the loop to continue from ticker to ticker without interruption I simply added 1 to the ticker index.  Once all tickers had been looped over, I closed the loop.

        tickerIndex = tickerIndex + 1
        
        End If
        
    Next i

#### Third loop
The third loop outputted the list of tickers plus total daily volumes and yearly returns for each ticker to the All Stocks Analysis sheet, just like in the original code.  In this case however, the four arrays had access to the ticker index so they have to contain a variable (i) representing the index.

    For i = 0 To 11
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
        
    Next i

### Formatting
I did not have to change any output formatting from the original code, but I did add a second run button on the output sheet to run only the refactored code.

### Results
It is clearly evident that in 2017 just about every stock performed better than in 2018. Even DAQO had a whopping 199% increase in return!  Only one stock (ticker:TERP) had a dip in return (a mere 7%).  In 2018, returns were awful.  Only two stocks (tickers:ENPH and RUN) performed well (total daily volume of 607,000,000; 82% increase in return, and total daily volume of 503,000,000; 84% increase in return, respectively).

#### Run times
The times needed to run my original code to output the results for 2017 and 2018 were 0.78 seconds and 0.77 seconds, respectively.  However, the time needed to run the refactored code was *significantly* less: 0.14 seconds for both 2017 and 2018!

![2017 run time with refactored code](https://github.com/MaxV6ft4/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![2018 run time with refactored code](https://github.com/MaxV6ft4/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary

### Benefits and Drawbacks of Refactoring Code
Refactoring code can be a huge benefit to both coder and viewer for mulitple reasons:

-By modifying the conditionals, the coder can create more efficient loops and logic statements that can go through thousands of rows of data at a blistering rate.  This will result in less lines of code to create as well.
-In the future, the coder will be able to view and/or edit any part of the code in a quicker amount of time.
-The viewer should be able to understand the code more easily as a result.

However, there are a couple drawbacks to refactoring that require attention:

-It can be very easy to get lost in the code while refactoring.  This could lead to multiple bugs, resulting in the code not properly running.
-In addition, the coder could also be unaware of the amount of time it could take to refactor the code.  If not prepared, the coder could be working for an exhorbitant amount of time.  This could be compounded by the first drawback taking place too.

### Application to this Particular Code

By adding a ticker index to the stock analysis code, I was able to create faster loops that went over the entire dataset at once.  In addition, the code no longer required individual row checking, resulting in fewer if-then statements.  Should I have to return to this code in the future, it will be very easy for me to run it with a differnt amount of tickers (the code involved in looping over all the rows would not have to change).

Knowing what to loop over in this code was easy but creating the new loops was tricky.  There were multiple issues with the code upon adding the ticker index at first.  I had to display the data type of and assign values inside each array before the loops began.  The code would not run if the wrong data type was assigned to an array.  Also, I had to modify the if-then statements to make sure they only applied to checking the first and last row of each ticker and not to checking each row individually.  It was easy to get lost in the second for loop, but using text to explain each line of code proved to be a beneficial guide in getting unstuck.  Refactoring this code as a whole did take a decent amount of time to complete.  Luckily, the formatting remained the same as in the original code.
