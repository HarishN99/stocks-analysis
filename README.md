# Stocks-Analysis

## Overview of Project

### Purpose

This analysis was intended to refactor and compare the refactored code with the code originally designed to help Steve analyze a handful of green energy stocks to see which are best for his parents to invest into. The original code was refactored to run through the data faster in hopes of being able to analyze thousands of stocks in the future. 

## Results

Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

### Stock Performance by Year

![VBA_Challenge_2017_Raw Table](https://user-images.githubusercontent.com/94864663/149045447-3ea4dfcb-d68a-4122-83e2-37d5cb74c220.png) 

**Table 1. All Stock Analysis for 2017** showing stock ticker, total daily volume and return wiht positive return's highlighted in green, and negative returns highlighted in red.  


![VBA_Challenge_2018_Raw Table](https://user-images.githubusercontent.com/94864663/149045462-592f55fa-229a-4b8c-af28-6f65bfba208a.png)

**Table 2. All Stock Analysis for 2018** showing stock ticker, total daily volume and return wiht positive return's highlighted in green, and negative returns highlighted in red.  

Nearly all the stocks in the dataset for 2017 provided a positive return on investment which were highlighted green in table 1. The return for TERP was -7.2% which means the stock dropped in value during this year, therefore it is highlighted in red on the table. The best green stocks to invest in during 2017 would be DQ, SEDG, and ENPH. In 2018, most stocks experienced a drop in value as exhibited by the red highlighted cells in table 2. The best stock of 2017, DQ, dropped by nearly 63% in 2018 which is why, based on the data, this stock is not a good investment. SEDG, the second highest return of 2017, experienced a drop of about 8% during 2018. The highest stock return of 2018 was RUN which saw an increase of 84% while it had a return of 5.5 % in year 2017. However, based on these results, the best stock to invest in appears to be ENPH which had the third highest return in 2017 and experienced an increase of about 82% in 2018. 


### Comparision of Analysis Execution Times (Original vs. Refactored)


**Refactored Script**

![VBA_Challenge_2017](https://user-images.githubusercontent.com/94864663/149047164-e77c4bfc-acde-4698-8ba9-36b4243b5428.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/94864663/149047200-c906f109-e567-4fcd-aadd-0ec258f84543.png)


**Original Script**

![Old Code 2017](https://user-images.githubusercontent.com/94864663/149047220-e65b5768-4fa4-4e0f-b4bc-1c198e634bce.png)
![Old Code 2018](https://user-images.githubusercontent.com/94864663/149047234-76a25d2c-2a61-4421-855a-02b9cd6aaaf3.png)

The refactored code ran for 0.176 and 0.199 seconds for the years 2017 and 2018 respectively. These are evidently lower runtimes when compared to the original code which ran in 0.852 and 0.887 seconds for 2017 and 2018 respectively. Therefore, refactoring the script allowed the code to run about 20% faster than the original code. 

The main changes to the original code include adding a tickerindex variable and removing the need for another iterator in the script. The original code for the loop changes from this:

```
For j = 2 to RowCount
    
    If Cells(j,1.Value = ticker Then
        
        totalVolume = totalVolume + Cells(j,8).Value
     
     End if 
```

to this in the refactored code:

```
For i = 2 To RowCount
      
      tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
```

Removing the need for another iterator allows the code to go through the data and collect the information once for each iteration rather than going through the “i” iterator for every iteration and then continuing to the “j” iterator for every iteration once again. As shown before, the refactored code allows the analysis to be run faster but it also streamlines the code and in turn makes it more readable as shown above. The refactored code does include a new line of code (shown below) which checks if the current row in the iteration is the first row with the selected tickerindex. This is a safety check to make sure no new issues arise with the refactored code above. 

```
If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
     
     tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
 End If
 ```

## Summary

In the case of the original VBA script, the original code was a lot less readable compared to the refactored code and as mentioned before ran a lot slower having to go through both the i and the j iterators. For this reason, the amount of memory used by the refactored code can be assumed to be lower seeing as the code loops through all the data in one go. Although, this does streamline the code, it does introduce potential for new bugs which is why a safe net was included to minimize this risk. However, the original code did not include the safety check, so this line also introduces some potential for error. With that being said, the refactored code ran smoothly with no errors in a fraction of the time the original code did and due to its reduced complexity, this code will be easier to maintain in the future. 

In the case of the original VBA script, the original code was a lot less readable compared to the refactored code and as mentioned before ran a lot slower having to go through both the i and the j iterators. For this reason, the amount of memory used by the refactored code can be assumed to be lower seeing as the code loops through all the data in one go. Although, this does streamline the code, it does introduce potential for new bugs which is why a safe net was included to minimize this risk. However, the original code did not include the safety check, so this line also introduces some potential for error. With that being said, the refactored code ran smoothly with no errors in a fraction of the time the original code did and due to its reduced complexity, this code will be easier to maintain in the future. 

