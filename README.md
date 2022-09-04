# stock-analysis

## Overview of Project
### Purpose: 
The project aims to improve query execution time so that a client can analyze multiple stocks of green energy more quickly as he wants to expand the dataset to include the entire stock market in recent years.

## Results
### Code:
To refactor the code, we created arrays to store the volume, starting price and ending price data and separated them into several For blocks instead of calculating them all as variables within the same loop, which made the application slower.  
Arrays created:  
 Dim tickerVolumes(12) As Long  
 Dim tickerStartingPrices(12) As Single  
 Dim tickerEndingPrices(12) As Single  
We created a variable as index(tickerIndex) to identify which array should be calculated each time.  
Dim tickerIndex As Variant  
     tickerIndex = 0  
We create a For to set the ticketVolume array at zero.  
 For i = 0 To 11  
          tickerVolumes(i) = 0  
     Next i  
Another For to populate the ticketVolume, tickerStartingPrices and tickerEndingPrices.  
 tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value  
 tickerStartingPrices(tickerIndex) = Cells(i, 6).Value  
 tickerEndingPrices(tickerIndex) = Cells(i, 6).Value  
And finally a For to present the results of the arrays according to their index.  
Cells(4 + i, 1).Value = tickers(i)  
Cells(4 + i, 2).Value = tickerVolumes(i)  
Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1  
With the creation of these arrays we were able to reduce the processing time as we can see in the following images:  
Previous code processing time  
![VBA_Challenge2_2018](https://user-images.githubusercontent.com/111664141/188336404-a1def02a-de02-4c15-86fd-620f9f91c7f7.png)

Processing time after refactored code  
![VBA_Challenge_2018](https://user-images.githubusercontent.com/111664141/188336365-974ba9af-308f-4ca4-b2e8-1b50f96fbbe9.png)

### Analysis:
As the client's first intention was to verify the DQ company for a possible investment, after performing the analyses, we can see that the company in question would not be suitable to make an investment since its return during the last year available for consultation was negative (-62.60%).
Analyzing the other companies as a possible investment, it would be interesting to study the ENPH and RUN companies more deeply, as they were the only ones that presented a positive return in these analyzes as well as a high volume traded.  
Below is a comparison of the companies for the year 2018:  

![AllStockAnalysis](https://user-images.githubusercontent.com/111664141/188336465-056c84ab-f271-4dad-b298-0ca49f8e9e22.png)


## Summary
### Advantages or Disadvantages:  
For code refactoring, the main advantage is the agility of the search, making the query much more dynamic and being able to cover a much wider range of data. The downside of refactoring is the time spent thinking about the solution and coding in addition to the need to create more lines of code and the possible errors during this refactoring.    
In refactoring this specific VBA, we had to create more variables, arrays and For, breaking a single loop into three. Although it was a laborious process, the objective of making the search much more agile was successful.


