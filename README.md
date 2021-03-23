# Stock Analysis using VBA

## Overview of project
This project uses a VBA script to analyze a set of 12 green stocks to idnetify those with high trade volume and a positive yearly return, so client can make reccommendations on investment choices.

### Purpose and Background
The purpose of this project is to analyze a set of predetermined stock tickers and find their total trading volume and return over a given year to determine if they are worthy investments. We refactored a VBA script designed to analyze a spreadsheet specific to a given year and/or stock to take input from the user on the year to analyze and output trading volume and return results for all of the identified stock tickers in that given year. The goal is to make the VBA script more scalable and efficient.

## Analysis and Challenges

### Analysis
The stock analysis for the years 2017 and 2018 shows only two stocks with positive returns for the years 2017 and 2018. The two stocks that performed positively for both 2017 and 2018 were [ENPH (Enphase Energy, Inc)](https://www.morningstar.com/stocks/XNAS/ENPH/quote) and [SUN (Sunrun, Inc.)](https://www.morningstar.com/stocks/xnas/run/quote). Both stocks were frequently traded throughout both years of analysis totaling over 829m and 770m trades, respectively. Only [TERP (Terreis)](https://www.morningstar.com/stocks/chix/terp/quote) performed negatively over the two year timeframe. The most traded stocks were solar companies including [SPWR (SunPower Corp)](https://www.morningstar.com/stocks/xnas/spwr/quote) and [FSLR (First Solar)](https://www.morningstar.com/stocks/xnas/fslr/quote) both of which were traded over 1b times in the two-year timeframe. 

In the tables below, we can see overall stock return performance for all 12 stocks in 2017 was better than 2018 with 11 out of the 12 stocks posting positive returns as opposed to 2 out of 12 stocks in 2018. 

### Challenges and Difficulties
One of the challenges I ran into refactoring the VBA script was setting the tickerIndex variable to increase only when all of that tickers info had been looped through. Without the conditional if/then statement the counter was going too high within the loop. The error was resulting from not checking to ensure the previous "Ticker" row no longer matched the next row. This was resulting in the index counter increasing too fast within the for loop and would end up with an "out of range error." 

My initial fix to run a nested for loop to search for each ticker in ticker array was ineffcient. 
'''
For tickerIndex = 0 To 11
    'for loop through all rows
Next tickerIndex
'''

The final refactoring elimnated the nested for loop need by adding a conditional statement to only increase the tickerIndex variable once it identified a new ticker name was identified.
'''
If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
End If
'''
An additional issue I had was generating a populated pre-populated array using a for loop. The fix is using a simple short for loop to loop through the array's index and populate it with zeros.
'''
'Populates tickerVolumes array with zeros
For k = 0 To 11
    tickerVolumes(k) = 0
Next k
'''
Limitations of this script are limited to the sheets they are tied to as we manually identifed the various different tickers to be used. Additionally, the data worksheets have to be arranged in such a way that each ticker is in a "block' of cells organized by thte name of the ticker to identify the starting and ending prices for the ticker.
The limitations of this data set can limit our ability to make recommendations on any specific stock as the dataset is not normalized to overall market performance in the sector or overall. 

## Results
*****Insert color chart
Based on the stock performance analysis, we would recommend looking deeper into ENPH, SUN, and [SEDG (SolarEdge Technologies)](https://www.morningstar.com/stocks/xnas/sedg/quote) based on their two year return performance. They performed positively with a large trade volume over the two year timeframe. We would recommend avoiding investing in [JKS (JinkoSolar)], TERP, and SPWR as they all had the lowest 2 year yield.

## Summary

### Advantages of refactoring code in general
Refactoring code is important as you learn to code more efficiently. The more you code, the more you learn methods and concept to use memory more effciently and improve the logic. Coding is an iterative excercise to continually make improvements in effciency, so the code can be scaled and understood by other developers more easily.

### Advantages of refactored VBA script over the original
The refactored code removes the original nested for loop that iterates through the various tickers searching through every cell in every row. Instead, the refactored code iterates through the tickerIndex variable based on the same elimnating the need for a nested for loop, therefore simplfying the logic. The refactored code also waits to print the output arrays until they've all been populated by the for loop instead of populating each indexed part of the array as it loops through looking for it's specific ticker.

### Areas to improve current VBA script
The current VBA script can be modified to more efficiently identify the different tickers to populate the tickers array. We can do this by searching the ticker column and for every "different" ticker add an additional variable to the index of the array. This would make the script scalable to a spreadsheet where we may not know how many different stock tickers are in the spreasheet, so the script can count the distinct ones.