# Analysis of Energy Stock Performance

## Overview
For this project I wrote a macro in VBA called AllStocksAnalysis to display the total daily volume and yearly return for each stock during the years 2017 and 2018.  I also created a run button in the output sheet for Steve to run the analysis on his own.  The code worked, but it involved looping through each individual row of stock.  Therefore I have refactored the code and have created a second run button specifically for this code.  The output displays the same results as those in the original code, but does so in a shorter amount of time.

## Results

[Energy Stock Analysis](https://github.com/MaxV6ft4/stock-analysis/blob/main/VBA_Challenge.xlsm)

### Running The Refactored Code
I looped over the entire data *at once* instead of row by row in the refactored code.  To do so, I created:
- a ticker index (held all twelve tickers inside) and initialized it to zero since in VBA the first number in an index is always zero.  
- three new output arrays:
    - ticker volumes 
    - starting prices
    - ending prices
- three separate loops:
    - the first loop initalized the ticker volumes array to zero, allowing it to reset once the next ticker (represented by the letter i here), was reached.

            For i = 0 To 11
    
                tickerVolumes(i) = 0
        
    - the second loop increased the volume by adding the inital volume (zero) to the value of each cell in colummn H of the two data sheets.  Here, the volumes are looped over all at once because they *immediately have access to the entire ticker index*, allowing them to instantly locate the correct ticker.  Using the letter i to represent the number of rows, I began the loop as follows:

            For i = 2 To RowCount
    
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

        - all I had to do afterwards was make sure that the row selected in the data sheet was the first row containing the desired ticker by checking to see if the next row up did not contain that ticker.  If so, the selected row is where the starting price (in Column F of the data sheets) for the ticker began.  

                 If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            
                     tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                

        - similarly, I made sure that the row selected in the sheet was the last row containing the desired ticker by checking to see if the next row down did not contain that ticker.  If so, the selected row would contain the ending price for the ticker.  

                 If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            
                    tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

        - to allow the loop to continue from ticker to ticker without interruption I simply added 1 to the ticker index.  Once all tickers had been looped over, I closed the loop.

                tickerIndex = tickerIndex + 1
        
    - the third loop outputted the list of tickers plus total daily volume and yearly return for each ticker to the All Stocks Analysis sheet.  Here, the four arrays had access to the ticker index so they have to contain a variable (i) representing the index.

            For i = 0 To 11
        
                    Cells(4 + i, 1).Value = tickers(i)
                    Cells(4 + i, 2).Value = tickerVolumes(i)
                    Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1

### Results
For the year 2017:
- just about every stock performed better than it did in 2018.
- only one stock (ticker:TERP) had a dip in return (a mere 7%).  

For 2018:
- only two stocks (tickers:ENPH and RUN) performed well.
    - ticker:ENPH had a total daily volume of 607,000,000 and an 82% increase in return.
    - ticker:RUN had a total daily volume of 503,000,000 and an 84% increase in return.

Run times:
 - original code:
    - 0.78 seconds for 2017 results
    - 0.77 seconds for 2018 results
 - refactored code:
    - 0.14 seconds for 2017 results
    - 0.14 seconds for 2018 results
 
![2017 run time with refactored code](https://github.com/MaxV6ft4/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![2018 run time with refactored code](https://github.com/MaxV6ft4/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary

### Benefits and Drawbacks of Refactoring Code
Refactoring code can be a huge benefit to both coder and viewer for mulitple reasons:

1. By modifying the conditionals, the coder can create more efficient loops and logic statements that can go through thousands of rows of data at a blistering rate.  This will result in less lines of code to create as well.
2. In the future, the coder will be able to view and/or edit any part of the code in a quicker amount of time.
3. The viewer should be able to understand the code more easily as a result.

However, there are a couple drawbacks to refactoring that require attention:

1. It can be very easy to get lost in the code while refactoring.  This could lead to multiple bugs, resulting in the code not properly running.
2. In addition, the coder could also be unaware of the amount of time it could take to refactor the code.  If not prepared, the coder could be working for an exhorbitant amount of time.  This could be compounded by the first drawback taking place too.

### Application to this Particular Code

By adding a ticker index to the stock analysis code, I was able to create faster loops that went over the entire dataset at once.  In addition, the code no longer required individual row checking, resulting in fewer if-then statements.  Should I have to return to this code in the future, it will be very easy for me to run it with a differnt amount of tickers (the code involved in looping over all the rows would not have to change).

Knowing what to loop over in this code was easy but creating the new loops was tricky.  There were multiple issues with the code upon adding the ticker index at first.  I had to display the data type of and assign values inside each array before the loops began.  The code would not run if the wrong data type was assigned to an array.  Also, I had to modify the if-then statements to make sure they only applied to checking the first and last row of each ticker and not to checking each row individually.  It was easy to get lost in the second for loop, but using text to explain each line of code proved to be a beneficial guide in getting unstuck.  Refactoring this code as a whole did take a decent amount of time to complete.  Luckily, the formatting remained the same as in the original code.
