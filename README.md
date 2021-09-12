# Stock-analysis

## Overview

I developed a code using VBA for a friend Steve to help his parents with their investment decision. As a result, he can analyze stock performances of an entire dataset with the click of a button. In order to see if the code that was developed can run faster and more efficiently than its previous iteration, I used refactoring to loop through all the data one time to collect the same information. 

The basic purpose of code refactoring is to make the code more efficient and maintainable. It is intended to improve the design, structure, and/or implementation while preserving its functionality.

## Refactoring ##

In order to make the code more efficient, I switched the nesting order of the 'for' loops. I created 4 different arrays - tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices. The tickers array was used to identify the ticker symbol of a stock. The rest of the arrays were matched with the tickers array using a variable called  tickerIndex. This variable allowed me to assign the tickerVolumes, tickerStartingPrices, and tickerEndingPrices to each ticker symbol before looping through the data set. This would complete the analysis much faster than using the nested 'for' loop as in the previous version of the code. 

The refactored code (extract) looks something like this:
   
    '1a) Create a ticker Index
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
    
    ''2b) Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount
        
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
    
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        'End If
        End If
        
        '3c) check if the current row is the last row with the selected ticker
        'If the next rows ticker doesnt match, increase the tickerIndex.
        
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
        
        End If
        
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        Cells(i + 4, 1).Value = tickers(i)
        Cells(i + 4, 2).Value = tickerVolumes(i)
        Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
    Next i
    
## Results

### Run-time ###

Using the original code, the run times were as follows for each of the two years, 2017 and 2018.

![image](https://github.com/amberwnaushahi/stock-analysis/blob/main/Resources/original_code_2017.png)

![image](https://github.com/amberwnaushahi/stock-analysis/blob/main/Resources/original_code_2018.png)

Once the code was refactored, run times were as follows:

![image](https://github.com/amberwnaushahi/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![image](https://github.com/amberwnaushahi/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

**Based on the run-times, the refactored code runs over 0.5 seconds faster than the original code making it more efficient.**

### Stock Performance - 2017 and 2018

In 2017, the majority of stocks that were analyzed performed well. The return for every stock except for 'TERP' was a net positive, with 'DQ' and 'SEDG' generating almost a 200% return. Based on this metric, they appear to be stand out options for Steve's parents to invest in. Based on total daily volume traded for each stock, which could reflect on a stock's popularity, potential growth and market sentiment, 'SPWR' and 'FSLR' were the clear leaders. Other stocks that drew large amounts of public interest were 'CSIQ', 'ENPH', 'RUN', and 'SEDG' each drawing over 200,000,000 in trading volume.

![image](https://github.com/amberwnaushahi/stock-analysis/blob/main/Resources/performance_2017.png)

2018 seemed to take a complete turn in terms of performance. Only two stocks 'ENPH' and 'RUN' showed positive returns based on opening and closing prices and also demonstrated some of the largest increases in total daily volume traded between 2017 and 2018. This appears to be a good indicator of the public sentiment and its potential growth in the future.

![image](https://github.com/amberwnaushahi/stock-analysis/blob/main/Resources/performance_2018.png)

## Summary

### Advantages of Refactoring Code
In general, some of advantages of refactoring code are:
* It can make the code more efficient, with fewer steps and quicker to execute
* Could possibly use less computer memory
* It can improve the code's logic for current and future users when code is well structured


### Disadvantages of Refactoring Code
Some of the cons of refactoring include:
* It can be time-consuming, making it inefficient.
* There are chances of errors, especially if original code writer is not the one refactoring. Some logical steps could be missed.

### How do these advantages and disadvantages apply to refactoring the original VBA Code

In our case, refactoring the original VBA script resulted in a more efficient program that resulted in faster run times. However, the time-saving was minimal compared to the time it took to actually refactor. This exercise would be very useful when data sets are very large and executed multiple times.

The code flow was more logical after refactoring. It can incorporate additional data as and when it is added, for example subsequent years' performance. 
