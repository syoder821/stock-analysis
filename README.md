# Stock-Analysis

## Overview of Project
Extract the total daily volume and yearly return for each stock in the dataset provided by the client.  Create a VBA macro to provide the information based on year with a simply button click in the excel sheet. Output the information in a formatted table in order to better visualize the stocks with the highest and lowest yearly return.  Additional request to refactor the code to make sure it is scalable for much larger data sets of stocks.      

### Purpose
To refactor the initial code generated to analyze the total daily volume and yearly return of the stocks provided in the data set to be scalable for much larger data sets.  Demonstrate and document the reduced runtime with the refactored code.  

## Results
- Comparision of the runtime of the intial code vs the refactored code:
  - 2017 Stock analysis runtime (intial script): 1.03 seconds 
  - 2017 Stock analysis runtime (refactored script): 0.15 seconds
  - 2018 Stock analysis runtime (intial script): 1 second
  - 2017 Stock analysis runtime (refactored script): 0.17 seconds
  - Results show a significant decrease in runtime for the refactored script of almost a factor of 10. 
- All stocks with the exception of TERP had positive returns in 2017.
- All stocks with the exception of ENPH and RUN had negative returns in 2018.  Only ENPH and RUN had positive returns in both years.

### 2017 Stock Analysis
![Alt Text](https://github.com/syoder821/stock-analysis/blob/main/VBA_Challenge_2017.png)


### 2018 Stock Analysis
![Alt Text](https://github.com/syoder821/stock-analysis/blob/main/VBA_Challenge_2018.png)
 
## Summary
- The purposes of refactoring according to Martin Fowler (Father of Code Smell) are stated in the following:
  - Refactoring Improves the Design of Software
  - Refactoring Makes Software Easier to Understand
  - Refactoring Helps Finding Bugs
  - Refactoring Helps Programming Faster

- Some general disadvantages of refactoring code are it can be risky when the application is big, when the existing code doesn't have proper test cases, and the developers do not fully understand the code. 
 
- In this case the main advantage of the refactored VBA script for the Stock-Analysis is the decrease in run time by almost a factor of 10. This reduction in runtime greatly improves the scalibility of the code for analysis of larger number of stocks.  
