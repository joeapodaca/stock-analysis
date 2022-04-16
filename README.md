# Stock-Analysis
Steve’s parents want to invest in DQ stocks, but don't know much about the stock market.  Steve would like to help and needs a quick way to analyze the stock market.

## Overview of Project
Steve had a limited number of stocks that he was interested in, but now needs a way to analyze many stocks at the same time.  The first code was good and did the job, but when looking at many stocks it is too slow.  Code must be refactored to analyze many different stocks in a short time.

### Purpose
The purpose of this project is to refactor the code to use less memory and return the analyzed results on many stocks faster.  This will give Steve a tool to quickly analyze what stocks his parents should be investing in.

## Results
Beginning with the original code I began adding variables and created an array.  I created for loops and nested loops to go through all the tickers for ticker volumes, ticker starting prices and ticker ending prices.  

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
        'If  Then
        
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
          
                 
        '3c) check if the current row is the last row with the selected ticker
         'If the next rows ticker doesn't match, increase the tickerIndex.
        'If  Then
            
             If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
          
            '3d Increase the tickerIndex.
            
             If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then tickerIndex = tickerIndex + 1
            
        'End if
            
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
    Next i



### Comparison of 2018 original code to refactored code run times
When looking at 2018 stocks the time to analyze was .738 seconds.
![2018 before](https://github.com/joeapodaca/stock-analysis/blob/main/2018%20Analyze.PNG)

After refactoring the code 2018 stocks can now run in .113 seconds.
![2018 refactured](https://github.com/joeapodaca/stock-analysis/blob/main/2018%20Refactored%20Analyze.PNG)

### Comparison of 2017 original code to refactored code run times
When looking at 2017 stocks the time to analyze was .718 seconds.
![2017 before](https://github.com/joeapodaca/stock-analysis/blob/main/2018%20Analyze.PNG)

After refactoring the code 2017 stocks can now run in .105 seconds.
![2017 refactured](https://github.com/joeapodaca/stock-analysis/blob/main/2018%20Refactored%20Analyze.PNG)

## Summary

### Advantages and Disadvantages of Refactored Code
The advantage of the refactored code is that it now runs 7X faster than the original code and uses less memory to accomplish the same task.  It is makes it more organized, improves the logic and makes it easier for the next person that may need to work on the code. The disadvantage is that it is a little harder to write the code and may not work with all cases or may not work at all. 

### Advantages and Disadvantages of Original and Refactored Code
The advantage of the original code is that it is easier to write and it works.  It also gives you the foundation needed to refactor the code as you can reuse much of the original code to refactor.  The disadvantage is that you make break the original code. Make sure to save the original code incase refactoring breaks the code.


