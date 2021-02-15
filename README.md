# stock-analysis
Bootcamp Module Two VBA Portfolio Project 

# Overview of Project
The purpose of this project was to write a VBA script for Steve, which could be run with the push of a button, in order to provide him with an analysis that could assist in advising his parents in their stock investment decisions. The VBA script analyzes both the total daily volume and return from the datasets "2017" and "2018" from the VBA_Challenge excel workbook. The datasets contain trading performance data for a handful of stocks for the years 2017 and 2018. 

Steve has also expressed a desire to expand the datasets to include the *entire* stockmarket over the last few years, and so the original Module 2 Solution script has been refactored for more effective performance in order to allow for potential future scalability. 

# Results
The final refactored script, titled "AllStocksAnalysisRefactored()," can be run with the push a button from the "All Stocks Analysis" worksheet of the VBA_Challenge excel workbook. The original Module 2 Solution script, titled "AllStocksAnalysis()", can also be run with the push of a button from the same "All Stocks Analysis" worksheet.

## Runtime Results 

The images below display the side-by-side runtimes for the original script and refactored script on both the 2017 and 2018 datasets. 

#### 2017 dataset: 

Runtime for the original Module 2 Solution script:  |  Runtime for the refactored script: 
:-------------------------:|:-------------------------:
![All_Stocks_Runtime_2017](Resources/All_Stocks_Runtime_2017.png)  |  ![VBA_Challenge_2017](Resources/VBA_Challenge_2017.png)

#### 2018 dataset: 

Runtime for the original Module 2 Solution script:  |  Runtime for the refactored script: 
:-------------------------:|:-------------------------:
![All Stocks Runtime 2018](Resources/All_Stocks_Runtime_2018.png)  |  ![VBA_Challenge_2018](Resources/VBA_Challenge_2018.png)


As clear from the results above, the refactored script had a much more effective runtime than the original Module 2 Solution script for analyzing both the 2017 and 2018 datasets. The refactored script will be the better choice going forward for scaling to analyze a dataset including the entire stockmarket for the past few years. 

## Stock Performance Results 

#### 2017 dataset: 

The images below display side-by-side analyses by the original Module 2 Solution script and refactored script on both the 2017 and 2018 data. The output of both scripts on each datset is identical, outside of rounding differences. There are clear differences, however, in the performance data for the analyzed stocks between the years 2017 and 2018. 

The list of stocks in the dataset seemed to perform much better overall in 2017 than 2018, with eleven out of twelve stocks having positive returns for the year in 2017, versus only two out of twelve stocks in 2018. 

Output for the original Module 2 Solution script:  |  Output for the refactored script: 
:-------------------------:|:-------------------------:
![All_Stocks_Output_2017](Resources/All_Stocks_Output_2017.png)  |  ![Refactored_Output_2017](Resources/Refactored_Output_2017.png)

#### 2018 dataset: 

Output for the original Module 2 Solution script:  |  Output for the refactored script: 
:-------------------------:|:-------------------------:
![All_Stocks_Output_2018](Resources/All_Stocks_Output_2018.png)  |  ![Refactored_Output_2018](Resources/Refactored_Output_2018.png)

Based on the data available, Steve may want to advise his parents to invest in ENPH, a stock which had strong positive returns for both 2017 and 2018. 

# Summary

There are always certain pros and cons to be weighted in refactoring code for a project. Refactoring can make a script more maintainable or scalable, and it can elimate "code smell," or general bad coding practices like the use of duplicate code or unnecessarily complex code. However, refactoring also takes time, and carries the risk of introducing new bugs to the code. 

In this instance, refactoring the original code was beneficial in that it significantly decreased the runtime, making the script more scalable for use on a larger dataset. The script introduced three additional "output" arrays, eliminating the need to structure a nested loop to generate output, which reduced some of the script's complexity. While this change made the runtime faster and the script more readable, it should be noted that the additional arrays will take up more memory. 

One issue with the original code that the refactored code did not address is the need to initialize an array of all tickers being analyzed at the start of the script. While this is not so much of an issue with the current use-case of analyzing twelve tickers, if Steve wishes to expand the script to include the entire stockmarket, expanding the initialized array of tickers could be a time-consuming adaptation with lots of potential for human error, if no form of automation is introduced to this part of the process. 

