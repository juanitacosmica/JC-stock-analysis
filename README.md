# Stock Analysis using VBA

## 1. Overview of Project

The purpose of this analysis was to refactor the code studied during Module 2 (Excel VBA); when refactoring we intend to improve the code in terms of design, structure and its implementation (how long we take to run it). Refactoring is not a change in function, the code was still aiming to loop through the stocks data to collect the Total Daily Volume and the Return per ticker upon user entered the year to analyze.

## 2. Results

The analysis was performed mainly through the use of:
- Index, 
- Loop, and,
- Nested Loops.

The results of the stock analysis for 2017 and 2018 remained the same before and after refactoring, when refactoring the code, its function does not change (tehrefore the results remain the same) but what improves is the time it takes to be run. With that being said, let's see how the performance was for 2017 and 2018 accordingly before jumping to see the difference between the time it took for the code to run prior and after refactoring.

![2017 All Stocks Analysis Results](/Resources/All_Stocks_Analysis_2017.png)

![2018 All Stocks Analysis Results](/Resources/All_Stocks_Analysis_2018.png)

From the current data and results, the Daily volume from 2017 to 2018 was not too distant across all tickers; however, overall the return dicreased average 75%, and only in two tickers the return was above 0 (81.9% for ENPH and 84% for RUN, in both cases the volume increased almost by 200%). 

Before giving Steve an advice on whether DQ is worth to invest or not, recommendation is to gather more historic data for a more detailed and fair comparison.

Refactoring the code allowed an improvement of 82% for the 2017 analysis, and, 79% for the 2018 analysis. As shown in the chart below and in detail in the images below.

![Challenge_2017vs2018_comparison](/Resources/VBA_Challenge_2017vs2018_comparison.png)

* *Before refactoring* *

![2018_Result_before_Refactoring](/Resources/2018_Result_before_Refactoring)

![2017_Result_before_Refactoring](/Resources/2017_Result_before_Refactoring)

* *After refactoring* *

![VBA_Challenge_2018](/Resources/VBA_Challenge_2018.png)

![VBA_Challenge_2017](/Resources/VBA_Challenge_2017.png)

Therefore, whenever the data allows us to refactor (improve) the code, it is worth trying. See below the refactored code:

![Refactored Code](/Resources/Refactored_Code.png)

## 3. Summary

Improving the time by 80% is definitely an advantage when refactoring the code in VBA. The commands used seemed to flow smoother and it will be easier to be transcribed for other analysis, in other words, it can be easily applied accross different scenarios.
