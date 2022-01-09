# Stock Analysis
This project contains the stock analysis Excel workbook that Steve requested. It also includes the various VBA functions that are capable of processing daily stock results and combining them into a summarized analysis by year. Screenshots are also included to show refactored code performance.

## Overview of Project
We have been helping Steve by building various functions in VBA to perform stock analysis research for his parents. Throughout the module we have written code that is functional but that does not scale well when expanding the list of stocks significantly. The primary purpose of this project is to refactor the code to perform faster, describe the reasons for the performance gains, and dive into the code differences deeper and highlight additional pros and cons of refactoring.

## Results
Our refactored algorithm performed quite a bit better. Let's break it out by the two years of data we were given:
  1. For 2017, the refactored code ran in about 113 milliseconds:  
  ![2017 Refactored](https://github.com/haldud/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)  
  The original code typically took about 900 milliseconds to complete when processing the 2017 data.
  
  2. For 2018, the results were very similar in that we had about a 700% increase in performance with the recorded time for the new algorithm being about 129 milliseconds.  
  ![2018 Refactored](https://github.com/haldud/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)  
  The original code typically took about 1 second to complete the 2018 analysis.
  
If we look at the flow of the original algorithm that performed analysis on all stocks, we can perhaps follow the logic a little easier in that it is going stock by stock and looping through all of the daily stock rows. The primary reason for the performance improvement is that we are only looping through the long daily list of stock values once whereas the original algorithm looped through them multiple times, once for each stock. The below code in the original algorithm is the primary reason for the performance difference.
```
For i = 0 To 11  
  ..  
  ..  
  For j = 2 To RowCount  
    ..  
    ..  
  End If  
  ..  
End If
```  
  
The new code only contains one for loop:  
```
For i = 2 To RowCount
  ..
  ..
End If
```
  
The new code also contains some additional loops, but those loops are on the smaller list of stocks which is less impactful on overall running time than looping through the daily values.  

## Summary
Refactoring code can be quite beneficial in several ways. First, like we saw in our scenario, we saw a performance improvement in the amount of time it took to generate the analysis. Performance issues are one of the main reasons for refactoring code. Another reason for code refactoring might be to make the code more readable and supportable by you or others in the future. Originally, we may write the code in such a way as to just get something done as we try different approaches to get to our result. If you come back some time later with a fresh mindset you might identify places where it can be improved in terms of readability without necessarily changing any algorithms.

On the other hand, refactoring code can also be problematic in some ways. In the quest to improve performance, the code can become more complicated in general. This means that we've used more sophisticated coding techniques which can at times cause the code to be more difficult to understand. That is why it is important to document your code well and explain your thought process as you are refactoring it. Another challenge with refactoring can be the introduction of bugs and/or unexpected behavior. At times, it can seem like starting from scratch and you are left with solving the same problem in a different way. It is important to have the ability to compare the original algorithm's output to the refactored one.

As the results showed, the primary advantage of the refactored algorithm is the performance. Throughout my testing, the new algorithm's performance was approximately seven times better than the original. This is a significant difference with the fairly small set of stocks that we are processing. One can conclude after examining the nested loop in the old algorithm further that this time difference would continue to grow larger as more stocks are added to the data set. With more stocks, there are also more daily rows to process so we can see that the old algorithm would suffer from serious performance degradation quite quickly. Another advantage of the new algorithm is that all of the formatting is done in the same function, instead of calling multiple functions separately.

One could argue that the original algorithm was easier to understand. It iterates the data for each stock and outputs the result for each without too much complication. I found it helpful to document each line of the new algorithm's code to better understand it. The new algorithm has multiple loops - one loop to initialize the results to zero, another to iterate the daily values, and then another to write out the results. The new algorithm certainly has added complexity that may not be needed for a small set of data.

Both algorithms suffer from some assumptions made on the data. For example, many of the cell references are hard-coded and relying on a specific structure. Both algorithms also have a list of specific stocks they are processing and not reading the daily values to identify unique tickers. This could become problematic as new stocks are added. In addition, both algorithms assume the data is sorted correctly by ticker and date.


