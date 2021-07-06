# Stock Analysis

## Overview of Project
Analysis of various green energy company stocks from years 2017 and 2018

### Purpose
The purpose of this project was to assist a young finance professional (hereby refered to as "client") in deciding which company's stock their parents should invest in. The client mentioned their parents had gone ahead to invest in a green company named DQ and would like to better visualize how well DQ has done in comparison to other green companies so they make the best investement decision.

The stock performance of a total of 12 green energy companies were analyzed. These companies include:
1. AY
2. CSIQ
3. DQ
4. ENPH
5. FSLR
6. HASI
7. JKS
8. RUN
9. SEDG
10. SPWR
11. TERP
12. VSLR

## Results
All stocks were analyzed as tickers and their total daily volume, starting price and ending price were identified for both 2017 and 2018. Please view the `VBA_Challenge.xlsm` file for all data and analysis. All corresponding screenshots are included in the resources folder.

### Stock Performance
From the information gathered on the stocks, which is also shown in the images below, it can be seen that 2017 was a better year for green energy companies in comparison to 2018. The only companies that did well both years were ENPH and RUN.

<img width="340" alt="All Stocks (2017)" src="https://user-images.githubusercontent.com/86085601/124535424-13d5e900-dde4-11eb-8011-9a134478efe2.png">
<img width="340" alt="All Stocks (2018)" src="https://user-images.githubusercontent.com/86085601/124535315-e6893b00-dde3-11eb-888c-5cb5ab652e6e.png">


As mentioned previously, the client's family had an interest in DQ and although in the year 2017 it was not a bad idea, based on the return in 2018, it is not adviced. The analysis shows that the client's family has a higher chance of getting good returns when investing in either of this companies. It is a better idea to invest in companies that have consistent good returns in stock.
 
### VBA Code Run Time
The VBA code used to condense the data for analysis was written in two different ways. Please view the `VBA_Challenge.xlsm` file macro for full detailed code.

#### Original Code
The "Original" code was where the tota volumes, starting price and ending price were stored as individual values. The nested `for` loop code initially used to create the tables is shown below.

```
'user to determine year of interest
yearValue = InputBox("What year would you like to run the analysis on?") 

startTime = Timer

For i = 0 To 11
    ticker = tickers(i) 'ticker used as stock name identifier
    totalVolume = 0

'Loop through rows in the data.
    
    Worksheets(yearValue).Activate
      For j = 2 To RowCount
    
    'Find the total volume for the current ticker.
    
        If Cells(j, 1).Value = ticker Then
            totalVolume = totalVolume + Cells(j, 8).Value
        End If
                  
    'Find the starting price for the current ticker.
    
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            startingPrice = Cells(j, 6).Value
        End If
    
    'Find the ending price for the current ticker.
    
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            endingPrice = Cells(j, 6).Value
        End If
        
      Next j
    
'Output the data for the current ticker.
    Worksheets("All Stocks Analysis").Activate
    
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
Next i

endTime = Timer
MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
```

As seen in the code, the time taken to run the code was also calculated and outputed through `endTime - startTime`. The resulting code run time for both years were ~ 0.66s and ~0.61s for 2017 and 2018 respectively, as shown in the images below.

<img width="414" alt="2017_Original" src="https://user-images.githubusercontent.com/86085601/124535584-5b5c7500-dde4-11eb-9669-02fa9727ff1f.png">   <img width="415" alt="2018_Original" src="https://user-images.githubusercontent.com/86085601/124535593-5f889280-dde4-11eb-8440-26b367bd3e0b.png">

#### Refactored code

A "refactored" code was also performed to determine if the "original" code can be rewritten to run faster. A refactored version was where the desired outputs were stored as arrays as seen in the following code block.

```
'user to determine year of interest
yearValue = InputBox("What year would you like to run the analysis on?") 

startTime = Timer

For stock = 0 To 11
    ticker = tickers(tickerIndex)
    tickerVolumes(tickerIndex) = 0
        
    'Loop over all the rows in the spreadsheet.
    Worksheets(yearValue).Activate
      For i = 2 To RowCount
    
      'Increase volume for current ticker
        If Cells(i, 1).Value = ticker Then
          tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If
        
      'Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
          tickerStartingPrice(tickerIndex) = Cells(i, 6).Value
        End If
        
      'check if the current row is the last row with the selected ticker
        If Cells(i + 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
          tickerEndingPrice(tickerIndex) = Cells(i, 6).Value
        End If
      Next i
    
    'Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        
    Worksheets("All Stocks Analysis").Activate
    
    Cells(4 + stock, 1).Value = ticker
    Cells(4 + stock, 2).Value = tickerVolumes(tickerIndex)
    Cells(4 + stock, 3).Value = (tickerEndingPrice(tickerIndex) / tickerStartingPrice(tickerIndex)) - 1
    tickerIndex = tickerIndex + 1
    
Next stock

endTime = Timer
MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
```
The resulting run times for 2017 and 2018 were ~0.59s and ~0.6s respectively as seen in the screenshots below.
<img width="415" alt="2017_Refactored" src="https://user-images.githubusercontent.com/86085601/124535700-9494e500-dde4-11eb-9094-41c617ddd4e0.png">   <img width="412" alt="2018_Refactored" src="https://user-images.githubusercontent.com/86085601/124535702-965ea880-dde4-11eb-9b70-fb895c9826fe.png">


#### Code time Comparison
The conclusion from the resulting screenshots show that there is a minute difference in code times. Excel is able to run VBA codes faster after running through the same code multiple times. This means although it seems the refactored code is running slightly faster, it is possible that the original code could catch up after multiple runs. The use of arrays instead of single data might not be record breaking.

## Summary
### Advantages/Disadavantages of Refactored Code


### Refactored Code Effect on Original VBA Script


