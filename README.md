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

Using the code provided, the code was refactored to loop through the data one time and collect all of the information with the following steps: 

#### Step 1a:
> Create a tickerIndex variable and set it equal to zero before iterating over all the rows. You will use this tickerIndex to access the correct index across the four different arrays you’ll be using: the tickers array and the three output arrays you’ll create in Step 1b.

    '1a) Create a ticker Index
    'The tickerIndex is set to equal to zero before looping over the rows
    tickerIndex = 0

#### Step 1b:
> Create three output arrays: tickerVolumes, tickerStartingPrices, and tickerEndingPrices. The tickerVolumes array should be a Long data type. The tickerStartingPrices and tickerEndingPrices arrays should be a Single data type.

    '1b) Create three output arrays
    'Arrays are created for tickers and all three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

#### Step 2a:
> Create a for loop to initialize the tickerVolumes to zero.

    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i

#### Step 2b:
> Create a for loop that will loop over all the rows in the spreadsheet.

#### Step 3a:
> Inside the for loop in Step 2b, write a script that increases the current tickerVolumes (stock ticker volume) variable and adds the ticker volume for the current stock ticker. Use the tickerIndex variable as the index.

#### Step 3b:
> Write an if-then statement to check if the current row is the first row with the selected tickerIndex. If it is, then assign the current starting price to the tickerStartingPrices variable.

#### Step 3c:
> Write an if-then statement to check if the current row is the last row with the selected tickerIndex. If it is, then assign the current closing price to the tickerEndingPrices variable. 

#### Step 3d:
> Write a script that increases the tickerIndex if the next row’s ticker doesn’t match the previous row’s ticker.

    ''2b) Loop over all the rows in the spreadsheet.
    'The tickerIndex is used to access the stock ticker index for the tickers array and all three output arrays
    'The script loops through stock data, reading and storing all of the following values from each row: tickers, volumes, starting prices, ending prices
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
           
            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
            End If
            
        'End If
    
    Next i

#### Step 4:
> Use a for loop to loop through your arrays (tickers, tickerVolumes, tickerStartingPrices, and tickerEndingPrices) to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet.

    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    'The output for the 2017 and 2018 stock analyses match the outputs from the AllStockAnalysis
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i

After this stock analysis ran, we confirmed the outputs for 2017 and 2018 were the same as in the original analysis. The images below compare the stock performance between 2017 and 2018 and displays their execution times with the new refactored VBA script:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/108038989/178817941-3da911f7-a9a1-4249-b34b-5dc9b08c86e4.png)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/108038989/178817950-88b41b91-1456-4041-834e-53689d77072e.png)

## Summary
In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?

There is a detailed statement on the advantages and disadvantages of refactoring code in general (3 pt).
There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script (3 pt).

Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. 

Your refactored code should run faster than it did in this module.

