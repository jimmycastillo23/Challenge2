# VBA with Excel

## Overview of Project

​VBA code editing from the M​​odule2's challenge exercise.

## Overview of Project

​To experience how editing, or refactoring codes make the VBA script run faster.  
 

### Purpose

​To edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information to determine if investing in the stocks available is really good investment. And to compare and evaluate the performance of running codes for 2017 data and 2018, and observe whether refactoring the code successfully made the VBA script run faster. Finally, explain the findings.

## Results

​The measure code performance was organized in two parts of the entire code. The first part was placed at the beginning right after the first line subroutine "Sub AllStocksAnalysisChalleng()" and included as follow:

Sub AllStocksAnalysisChalleng()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

       startTime = Timer

The second part was positioned at the end of the subroutine Sub AllStocksAnalysisChalleng() script, right after the Next i loop script. And it is stated as follow: 

 endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

​As for the analysis of the performance comparing both 2017 and 2018, it can be observed that the performance of the 2017 image shows "This code ran 0.203125 seconds which is faster than the 2018 which shows 0.2070313. Still both performances are very fast. A​s an​ experiment as ​I ​also compared the performance of the GreenStock Challenge's file performance and showed similar execution performance.  
---
[This image shows the performance when running the codes for 2017](https://github.com/jimmycastillo23/Challenge2/blob/main/VBA_Challenge_2017.png)
---
[This image shows the performance when running the codes for 2017](https://github.com/jimmycastillo23/Challenge2/blob/main/VBA_Challenge_2018.png)

### Summary: Two pieces of analysis are shown below:    

## There is a detailed statement on the advantages and disadvantages of refactoring code in general:

The refactoring scripts can possibly be considered as advantage in terms of quality, but it could be a disadvantage to test again all functionality and fix potential errors. However, dealing with functionality and fixing errors is a daily challenge all developer face when writing codes. 

## There is a detailed statement on the advantages and disadvantages of the original and refactored VBA script

when comparing performance of both files gree_stock's file from the module2's materials and the assignment VBA challenge, the results ended taking longer than the refactored file.  

Below the organization of the refactored scripts: 

Sub Sub AllStocksAnalysisperformance()
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
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
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

End Sub

### Conclusion and analysis of findings: 

​Refactoring can lead to making scripts more cleaned, organized, readable and easier to understand. Based on the findings in 2017 there are two tickers that can be more feasible worthy stocks to trade to, these are DQ with 199.4% and SEDG 184.5% respectively. And In 2018 occurred the opposite of 2017. In 2018 There are two tickers that show positive numbers and those are ENPH and RUN with 81.9% and 84%. The higher the number is, the more challenging it can be to trade these ticke​rs​ to because on one hand it can be considered that the result of 2017 can mean high risk for people interested in these tickers, however the same can happen with the result of the 2018. It is recommended for those interested in purchasing stocks or trading these ticke​rs​ to explore other resources of analysis such as historic performance, local and global economical and marketing impact over the years, and consider possible predictions when trading these tickers.

      
