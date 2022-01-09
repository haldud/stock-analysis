# Stock Analysis
This project contains the stock analysis Excel workbook that Steve requested. It also includes the various VBA functions that are capable of processing daily stock results and combining them into a summarized analysis by year. Screenshots are also included to show refactored code performance.

## Overview of Project
We have been helping Steve by building various functions in VBA to perform stock analysis research for his parents. Throughout the module we have written code that is functional but that does not scale well when expanding the list of stocks significantly. The primary purpose of this project is to refactor to code to perform faster, describe the reasons for the performance gains, and dive into the code differences deeper and highlight additional pros and cons of the refactoring.

## Results
Our refactored algorithm performed quite a bit better. Lets's break it out by the two years of data we were given:
  1. For 2017, the refactored code ran in about 113 milliseconds:
 
 ![2017 Refactored](https://github.com/haldud/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

## Summary
If we look at the flow of the original algorithm that performed analysis on all stocks, we can perhaps follow the logic a little easier in that it is going stock by stock and looping through all of the daily stock rows.
