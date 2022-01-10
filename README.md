# VBA stock market trends

### By: Stacey Booysen

For this project, we were tasked with helping Steve analyze and present data to his retiring parents who wish to invest in stocks. They don’t know much about any of the stocks presented and have tasked their son, as his first customers, with helping them find the right fit.
The results found through analysis, showed that the original choice for investment, DQ, was likely not a very safe one. While DQ had a positive return in 2017, it went down drastically in 2018, as did the return for many of the options presented to Steve’s parents.

![Results 2017]( https://github.com/sbooysen/stock-analysis/blob/main/2017_Stocks_Return.png)
![Results 2018]( https://github.com/sbooysen/stock-analysis/blob/main/2018_Stocks_Return.png)

In the most recent year, 2018, ENPH and RUN seem to have the best return. By running the codes to calculate the outcomes for the year, we were able to help Steve and his parents see the big picture more easily.
We also managed to improve the runtime of the analysis by shaving off a few seconds of processing time. In this way, Steve should be able to run data for more years in a quicker amount of time if he needs to. In particular, the nested loops helped to run multiple analysis at once to provide Steve with a concise result and a readable format.


![2017 Original]( https://github.com/sbooysen/stock-analysis/blob/main/VBA_Challenge%202017-First%20Run.png)
![2018 Original]( https://github.com/sbooysen/stock-analysis/blob/main/VBA_Challenge%202018-First%20Run.png)

Vs.

![2017 Refactored]( https://github.com/sbooysen/stock-analysis/blob/main/VBA_Challenge_2017.png)
![2018 Refactored]( https://github.com/sbooysen/stock-analysis/blob/main/VBA_Challenge_2018.png)


Refactoring the code helped in the sense that we had a base to work with already and were not starting from scratch. However, there can be some confusion when working with pre-established code, especially if a lot of it is somewhat new to you. In this case, there were several times in which there needed to be debugging because there was a missed “End If” or “Next i”. In one case I had missed an ‘End If’ on the code below and caused an error: 

```
  If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
```

Luckily, the debugging highlights the issues most of the time, but if the code is even longer than what we used in this project, then it’s liable to get even messier. Still, it was a great help to have a base to work off of and in this project, it helped to retain the information from the spreadsheets that we had already worked on.

Overall, the results provided an easy-to-use format for Steve, and one that he can hopefully reuse with more data. Thanks to the formatting, the time it takes to process large data should also be slightly less than before. With this, Steve should be able to present his findings clearly for his parents in order to help them make a good decision.


