# Analysis of Energy Stock Performance

## Overview
Steve supplied me with alternative energy stock data for the years 2017 and 2018, as his parents desire to invest in clean energy.  At first they were only interested in Daqo (ticker:DQ) but the stock did not perform well in 2018 (dropped 63% from 2017), so they want to see how other alternative energy stocks fared both in 2017 and in 2018.  

### Purpose
For this project we want to specifically know the difference between total daily volumes and yearly returns for the two years provided.  Initially I wrote a macro in VBA that performed analysis on all twelve stocks, and I created a run button for Steve to run the analysis on his own.  The code worked, but it involved looping through each individual row of stock.  Therefore I have refactored the code to make it run faster and have created a second run button specifically for this code.  The output displays the same results as those in the original code, but at a blistering rate.  This report will strictly cover the refactored code.

## Results

[Energy Stock Analysis](https://github.com/MaxV6ft4/stock-analysis/blob/main/VBA_Challenge.xlsm)

I created a new macro in VBA called AllStocksAnalysisRefactored, maintaining both the same output formatting and the same array of twelve tickers as was done in the original code.  However, in this code I looped over the entire data *at once* instead of row by row.  To do this I created a ticker index, a variable that holds all the tickers inside itself, called it tickerIndex and initialized it to zero since in VBA the first number in an index is always 0.  Next, I created three new output arrays, one for ticker volumes, one for starting prices and one for ending prices.  Then I created three separate loops: two to loop over arrays and one to loop over all the rows in the two data sheets (2017 and 2018).

### Loops

#### First loop
The first loop initalized the ticker volumes array to zero, allowing it to reset once the loop reached the next ticker (represented by the letter i here).

SHOW CODE.

#### Second loop
The second loop is key.  For each ticker, I increased the volume by adding the inital volume (zero) to the value of each cell in colummn H of the two data sheets.  The difference is that in the original code I added the volume to the value of the cells in H individually by checking each row to make sure it contained the proper ticker.  Here, the volume is looped over all at once because, as an array the volume *immediately has access to the entire ticker index*, allowing it to instantly locate the correct ticker.  

SHOW CODE.

All I had to do afterwards was make sure that the row selected in the data sheet was the first row containing the desired ticker by checking to see if the next row up did not contain that ticker.  If so, the selected row is where the starting price (in Column F of the data sheets) for the ticker began.  

SHOW CODE.  

Similarly, I made sure that the row selected in the sheet was the last row containing the desired ticker by checking to see if the next row down did not contain that ticker.  If so, the selected row would contain the ending price for the ticker.  

SHOW CODE.  

Remember, the starting and ending prices are arrays as well here, so just like the volume array they can instantly locate the correct ticker.  To allow the loop to continue from ticker to ticker without interruption I simply added 1 to the tickerIndex variable.

SHOW CODE.

#### Third loop
The third loop outputted the list of tickers plus total daily volumes and yearly returns for each ticker to the All Stocks Analysis sheet, just like in the original code.  Here, however, the four arrays had access to the ticker index so they must be represented by a variable (i) containing the tickers.  

SHOW CODE.

### Formatting
I did not have to change any output formatting from the original code, but I did add a second run button to run only the refactored code.

### Outcome
It is clearly evident that in 2017, just about every stock performed better than in 2018. Even DAQO had a whopping 199% increase in return!  Only one stock (ticker:TERP) had a dip in total daily volume and return (a mere 7%).  In 2018, total daily volumes and returns were awful.  Only two stocks (tickers:ENPH and RUN) had increases in both categories.  Hopefully more clean energy stocks performed better in later years...

#### Run times
The times needed to run my original code to output the results for 2017 and 2018 were 0.78 seconds and 0.77 seconds, respectively.  However, the time needed to run the refactored code was *significantly* less: 0.14 seconds for both 2017 and 2018!

![2017 run time with refactored code](https://github.com/MaxV6ft4/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![2018 run time with refactored code](https://github.com/MaxV6ft4/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

## Summary
As shown above, refactoring code can help loop over thousands of rows of data in a split second.  Using an index as a variable inside arrays can speed up the looping process.  However, there is more room for error in writing the code.  It is important to make sure that new variables are correctly placed inside arrays and that each loop is closed.

In the refactored code for the stock analysis, the run times are more than 0.6 seconds faster than they were in the original code.
The ticker index expedited the looping process.  I did have to assign new variables for each loop.
