# Stock analysis using VBA

## Overview of project
This project uses a VBA script to analyze a set of 12 green stocks to identify green stocks with high trade volume and a positive yearly return, so client can make proper recommendations on investments.

### Purpose and background
The purpose of this project is to analyze a set of predetermined "green" stock tickers and find their total trading volume and return over a given year to determine if they are worthy investments. Client requested to refactor a VBA script originally designed to analyze a spreadsheet specific to a given year and/or stock to take input from the user on the year to analyze, output trading volume and return results for all identified stock tickers in that given year. The goal is to make the VBA script more scalable and efficient to discover the best green stocks for investment.

## Analysis and challenges

### Green stock analysis
Tables 1.1 and 1.2 below show stock return performance for all 12 stocks. 2017 performance was better than 2018 as 11 out of the 12 stocks posted positive returns as opposed to 2 out of 12 stocks. 

**Table 1.1 Green Stock Performance (2017)**
![2017 Green Stock Performance](https://github.com/joshuanallen/stock-analysis/blob/c0af656872fe23da68a4d3c580af77d1664fb68a/Resources/VBA_Challenge_2017_results.png)

**Table 1.2 Green Stock Performance (2018)**
![2018 Green Stock Performance](https://github.com/joshuanallen/stock-analysis/blob/c0af656872fe23da68a4d3c580af77d1664fb68a/Resources/VBA_Challenge_2018_results.png)


Table 1.3 below shows the return performance for the 2 year timeframe (2017-2018). Of the 12 stocks present in the data, all had significant trade volume, 8 had a positive return, 2 had a neutral return, and 3 had a negative return over the 2 year timeframe.

**Table 1.3 2-Year Green Stock Overall Performance (2017-2018)**
![2-Year Overall Green Stock Performance (2017-2018)](https://github.com/joshuanallen/stock-analysis/blob/c0af656872fe23da68a4d3c580af77d1664fb68a/Resources/All_stocks_2yr_performance.png)

The positive performing stocks are:
- [ENPH (Enphase Energy, Inc)](https://www.morningstar.com/stocks/XNAS/ENPH/quote)
- [SEDG (Solar Edge Technologies)](https://www.morningstar.com/stocks/xnas/sedg/quote)
- [RUN (Sunrun, Inc)](https://www.morningstar.com/stocks/xnas/run/quote)
- [VSLR (Vivint Solar, Inc)](https://www.marketbeat.com/stocks/NYSE/VSLR/)
- [FSLR (First Solar, Inc)](https://www.morningstar.com/stocks/xnas/fslr/quote)
- [DQ (Daqo New Energy Corp ADR)](https://www.morningstar.com/stocks/xnys/dq/quote)
- [CSIQ (Canadian Solar, Inc)](https://www.morningstar.com/stocks/xnas/csiq/quote)

The neutral performing stocks are:
- [AY (Atlantica Sustainable Infrastructure, PLC)](https://www.morningstar.com/stocks/xnas/ay/quote)
- [HASI (Hannon Armstrong Sustainable Infrastructure Capital, Inc)](https://www.morningstar.com/stocks/xnys/hasi/quote)

The negative performing stocks are:
- [TERP (Terreis)](https://www.morningstar.com/stocks/chix/terp/quote)
- [SPWR (SunPower Corp)](https://www.morningstar.com/stocks/xnas/spwr/quote)
- [JKS (JinkSolar Holding Co Ltd ADR)](https://www.morningstar.com/stocks/xnys/jks/quote)

### Challenges with refactoring VBA script
One of the challenges I ran into refactoring the VBA script was setting the tickerIndex variable to increase only when all of that tickers info had been looped through. Without the conditional if/then statement the counter was going too high within the loop. The error was resulting from not checking to ensure the previous "Ticker" row no longer matched the next row. This was resulting in the index counter increasing too fast within the for loop and would end up with an "out of range error." 

My initial fix to run a nested for loop to search for each ticker in ticker array was ineffcient. 

    For tickerIndex = 0 To 11
        'for loop through all rows
    Next tickerIndex


The final refactoring elimnated the nested for loop need by adding a conditional statement to only increase the tickerIndex variable once it identified a new ticker name was identified.

    If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If

An additional issue I had was generating a populated pre-populated array using a for loop. The fix is using a simple short for loop to loop through the array's index and populate it with zeros.

    Populates tickerVolumes array with zeros
    For k = 0 To 11
        tickerVolumes(k) = 0
    Next k

Limitations of this script are limited to the sheets they are tied to as we manually identifed the various different tickers to be used. Additionally, the data worksheets have to be arranged in such a way that each ticker is in a "block' of cells organized by thte name of the ticker to identify the starting and ending prices for the ticker.

### Limitations of data set
The limitations of this data set can limit our ability to make recommendations on any specific stock as the dataset is not normalized to overall market performance in the sector or overall. Data is limited to 2 year timeframe, thus limiting the scope of the research. Additional years and more recent data would be helpful in building out a more complete picture.

## Results

### Green stock recommendations
*****Insert color chart 2 year results
Based on the stock performance analysis above, our recommendations for the following stocks are:

Recommendations for investment (Top performers):
1. [ENPH](https://www.morningstar.com/stocks/XNAS/ENPH/quote)
2. [RUN](https://www.morningstar.com/stocks/xnas/run/quote)
3. [SEDG](https://www.morningstar.com/stocks/xnas/sedg/quote)
4. [VSLR](https://www.marketbeat.com/stocks/NYSE/VSLR/)

Recommendations for further research:
5. [FSLR](https://www.morningstar.com/stocks/xnas/fslr/quote)
6. [DQ](https://www.morningstar.com/stocks/xnys/dq/quote)
7. [CSIQ](https://www.morningstar.com/stocks/xnas/csiq/quote)
8. [AY](https://www.morningstar.com/stocks/xnas/ay/quote)
9. [HASI](https://www.morningstar.com/stocks/xnys/hasi/quote)

Recommendations against investments (Worst performers):
10. [TERP](https://www.morningstar.com/stocks/chix/terp/quote)
11. [SPWR](https://www.morningstar.com/stocks/xnas/spwr/quote)
12. [JKS](https://www.morningstar.com/stocks/xnys/jks/quote)

Stock recommendations are limited to 2017-2018 performance, therefore it is advised to pair with research with more contemporary data and industry outlook for future investment.

### Refactoring results
The refactored VBA script signifcantly increases the efficiency of the analysis. The results for the refactored code are shown in the pictures below:

The 2017 analysis runtime improved from 0.5 seconds(s) to 0.125s after script refactoring.
**Original VBA script timer for 2017 data:**
![Original VBA script timer for 2017 data](https://github.com/joshuanallen/stock-analysis/blob/c0af656872fe23da68a4d3c580af77d1664fb68a/Resources/Original_VBA_script_2017_timer.png)
**Refactored VBA script timer for 2017 data:**
![Refactored VBA script timer for 2017 data](https://github.com/joshuanallen/stock-analysis/blob/c0af656872fe23da68a4d3c580af77d1664fb68a/Resources/VBA_Challenge_2017_timer.png)

The 2018 analysis runtime improved from 0.496 seconds(s) to 0.132s after script refactoring.
**Original VBA script timer for 2018 data:**
![Original VBA script timer for 2018 data](https://github.com/joshuanallen/stock-analysis/blob/c0af656872fe23da68a4d3c580af77d1664fb68a/Resources/Original_VBA_script_2018_timer.png)
**Refactored VBA script timer for 2018 data:**
![Refactored VBA script timer for 2018 data](https://github.com/joshuanallen/stock-analysis/blob/c0af656872fe23da68a4d3c580af77d1664fb68a/Resources/VBA_Challenge_2018_timer.png)

## Summary of VBA script refactoring

### Advantages of refactoring code in general
Refactoring code is important as you learn to code more efficiently. The more you code, the more you learn methods and concept to use memory more effciently and improve the logic. Coding is an iterative excercise to continually make improvements in effciency, so the code can be scaled and understood by other developers more easily.

### Advantages of refactored VBA script over the original
The refactored code removes the original nested for loop that iterates through the various tickers searching through every cell in every row. Instead, the refactored code iterates through the tickerIndex variable based on the same logic that identifies the end of the ticker rows. This elimnates the need for a nested for loop, therefore simplfying the logic. The refactored code also waits to print the output arrays until they've all been populated by the for loop instead of populating each indexed part of the array as it loops through looking for it's specific ticker. Additionally, the refactored code includes formatting the table to identify positive and negative performers eliminating the need for another script to run to apply formatting on the output table.

### Areas to improve current VBA script
The current VBA script can be modified to be more efficient by identifying the different tickers to populate the tickers array. We can do this by searching the ticker column and for every "different" ticker add an additional variable to the index of the array. This would make the script scalable to a spreadsheet where we may not know how many different stock tickers are in the spreasheet.
