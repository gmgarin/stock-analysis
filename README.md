# Stock Analysis Using Excel VBA


## Overview of the Project

### Purpose

#### The purpose of this analysis is to study the stock market over the last few years to help Steve with his financial decisions. This was done using VBA codes to help him simplify gigantic amount of stock data. The years 2017 and 2018 are the focus of this stock market analysis.  The goal was to compare the returns of the 12 different identified in this analyis. 

## Results

### 2017 vs 2018 Stock Perfomance
![VBA_Challenge_2017](https://user-images.githubusercontent.com/96095956/149698348-3550eeda-e4ee-4ae6-b9c9-8e262b8a32b3.png) ![VBA_Challenge_2018](https://user-images.githubusercontent.com/96095956/149698390-0274e89d-7f6c-4fde-bd55-0fbb97e45ece.png)
##### 2017 Stock Performance                                                2018 Stock Performance



#### Analysis shows that there is a general significant decline in returns in the year 2018 from the year 2017. The data also shows substantial decline in total daily volume. 

### VBA Codes 
#### The codes used to make analysis is shown below. Since the file is run from top to bottom, the code started by declaring variables. Then, the relavant sheets were activated. Next, the codes were used to create header row for the input box. Then, the for loops, if-then, and nested loops codes were written to inpect the data in its entirety. Lastly, the conditional formatting codes as well as code performance were included.



    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
   
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
   
    Range("A1").Value = "All Stocks (" + yearValue + ")"
   
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
   
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
   
    'Activate data worksheet
    Worksheets(yearValue).Activate
   
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
   
    '1a) Create a ticker Index
        tickerIndex = 0

    '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single

    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
            tickerVolumes(i) = 0
            
        Next i

    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount

    '3a) Increase volume for current ticker
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
   
    '3b) Check if the current row is the first row with the selected tickerIndex.
    'If  Then
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
    End If
   
    '3c) check if the current row is the last row with the selected ticker
    'If  Then
     If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
     End If

        '3d Increase the tickerIndex.
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If

    Next i

    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
   
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
       
    Next i

    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
       
        If Cells(i, 3) > 0 Then
           
            Cells(i, 3).Interior.Color = vbGreen
           
        Else
       
            Cells(i, 3).Interior.Color = vbRed
           
        End If
       
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)


### Execution Times
#### The run times were drastically reduced after refactoring. The macro run time for both years is at approximately 0.14 seconds compared to a bit longer run times before refactoring. 

![2017_run_time](https://user-images.githubusercontent.com/96095956/149701396-c13ae844-051b-4b76-a97e-c38665333a53.png)
![2018_run_time](https://user-images.githubusercontent.com/96095956/149701382-2e94df32-184a-4d2f-a4f1-b07c155486fe.png)


## Summary

### Advantages and Disadvantages of Refactoring Codes
#### Refactoring codes is great way to optimize codes by restructuring existing ones without affecting or changing its output. It also helps in fixing bugs and generally decreases execution times. On the other hand, refactoring codes can be time-consuming and downright confusing.

### Advantages and Disadvantages of Original and Refactored VBA Script
#### One of the clear benefits of refactored VBA script is a faster run times. The original code took a bit longer in terms of excution time, whereas refactored codes run smoother and faster. Another benefit is a simpler and easier to read script. However, refactoring could be very confusing.

