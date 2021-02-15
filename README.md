# stock-analysis
Bootcamp Module Two VBA Portfolio Project 

# Overview of Project
The purpose of this project was to write a VBA script for Steve, which could be run with the push of a button, in order to provide him with an analysis that could assist in advising his parents in their stock investment decisions. The VBA script analyzes both the total daily volume and return from the datasets "2017" and "2018" from the VBA_Challenge excel workbook. The datasets contain trading performance data for a handful of stocks for the years 2017 and 2018. 

Steve has also expressed a desire to expand the datasets to include the *entire* stockmarket over the last few years, and so the original Module 2 Solution script needed refactoring for more effective performance in order to allow for potential future scalability. 

# Results
The final refactored script, titled "AllStocksAnalysisRefactored()," can be run with the push a button from the "All Stocks Analysis" worksheet of the VBA_Challenge excel workbook. The original Module 2 Solution script, titled "AllStocksAnalysis()", can also be run with the push of a button from the same "All Stocks Analysis" worksheet.

## Runtime Results 

#### Runtime for the original Module 2 Solution script (left) and refactored script (right) analyzing the 2017 dataset: 
![All_Stocks_Runtime_2017](Resources/All_Stocks_Runtime_2017.png)   ![VBA_Challenge_2017](Resources/VBA_Challenge_2017.png)

#### Runtime for the original Module 2 Solution script (left) and refactored script (right) analyzing the 2018 dataset: 
![All Stocks Runtime 2018](Resources/All_Stocks_Runtime_2018.png)  ![VBA_Challenge_2018](Resources/VBA_Challenge_2018.png)

As is clear from the results above, the refactored script had a much more effective runtime than the original Module 2 Solution script for analyzing both the 2017 and 2018 datasets. The refactored script will be the better choice going forward for scaling to analyze a dataset including the entire stockmarket for the past few years. 

## Stock Performance Results 
![Refactored_Output_2017](Resources/Refactored_Output_2017.png)
![All_Stocks_Output_2017](Resources/All_Stocks_Output_2017.png)
![Refactored_Output_2018](Resources/Refactored_Output_2018.png)
![All_Stocks_Output_2018](Resources/All_Stocks_Output_2018.png)
# Summary

There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script
