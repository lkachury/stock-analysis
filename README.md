# Stock Analysis with VBA

## Overview of Project
Refactoring VBA code to make the stock analysis run more efficiently. 

### Purpose
In this module, we have been utilizing VBA to analyze stock performance in a dataset for a single and then a dozen stocks. With an expanded dataset to include thousands of stocks, the current VBA code may not work as well or run as quickly. The purpose of this analysis is to first refactor the current VBA code as it loops through the expanded dataset to collect the stock performance infomation and then to determine if this edit successfully ran the analysis faster and more efficiently. Refactoring code allows the code to run in fewer steps, use less memory, and improve on the current logic of the code.

## Results

### Original
The images below compare the stock performance between 2017 and 2018 and displays their execution times with the original VBA script:

![Initial_2017_RunTime](https://user-images.githubusercontent.com/108038989/178812431-1655e76b-d69b-4b90-8edf-11c238eefa3f.png)

![Initial_2018_RunTime](https://user-images.githubusercontent.com/108038989/178812447-b0d60dd4-86b9-4fc2-bd68-c7aab6d41da2.png)

### Refactored 

Using the starter code provided, the code was refactored to loop through the data one time and collect all of the information. 



Your refactored code should run faster than it did in this module.



The tickerIndex is set equal to zero before looping over the rows. (5 pt).

Arrays are created for tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices (15 pt).

The tickerIndex is used to access the stock ticker index for the tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices arrays (15 pt).

The script loops through stock data, reading and storing all of the following values from each row: tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices (25 pt).

Code for formatting the cells in the spreadsheet is working (5 pt).

There are comments to explain the purpose of the code (5 pt).

The outputs for the 2017 and 2018 stock analyses in the VBA_Challenge.xlsm workbook match the outputs from the AllStockAnalysis in the module (5 pt).

The pop-up messages showing the elapsed run time for the script are saved as VBA_Challenge_2017.png and VBA_Challenge_2018.png (5 pt).



## Summary
In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?

There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).

Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. 


