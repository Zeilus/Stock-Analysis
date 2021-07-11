# Stock-Analysis

### **Overview of Project**

The aim of this project was to automate the process of checking various company stocks and observe if the company had a positive return. Data used for this project has information on twelve company's stock opening price, closing price, highest gains, lowest loses, and total volume for the years of 2017 and 2018.

A script was made in VBA to ask which year the user wanted results for and variables were made to hold the values for company names, total daily volume, and yearly returns. 

    Sub AllStocksAnalysis()
      Dim startTime As Single
      Dim endTime As Single
    
      yearValue = InputBox("What year would you like to run the analysis on?")
    
         startTime = Timer
       
    '1) Format the output sheet on All Stocks Analysis worksheet
       Worksheets("All Stocks Analysis").Activate
       Range("A1").Value = "All Stocks (2018)"
       'Create header row'
       Cells(3, 1).Value = "Year"
       Cells(3, 2).Value = "Total Daily Volume"
       Cells(3, 3).Value = "Return"
    
    

Afterwards an array holding each companies names was made along with a for loop to check each companies names, total volumes, and yearly returns.

      '3a) Initialize variables for starting price and ending price
         Dim startingPrice As Single
         Dim engingPrice As Single

      '3b) Activate data worksheet
         Worksheets("2018").Activate
   
      '3c) Get the number of rows to loop over
         RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
      '4) Loop through tickers
         For i = 0 To 11
             ticker = tickers(i)
      '5) loop through rows in the data
         Worksheets("2018").Activate
         For j = 2 To RowCount
            '5a) Get total volume for current ticker
             If Cells(j, 1).Value = ticker Then
             totalVolume = totalVolume + Cells(j, 8).Value
             
         End If
         
            '5b) get starting price for current ticker
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            startingPrice = Cells(j, 6).Value
            
            End If
            
           '5c) get ending price for current ticker
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            endingPrice = Cells(j, 6).Value
            
            End If
            

       Next j
       
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
     
Next data was formatted to fit each column to the length of the values for each catagory, bold the headers, make the yearly return values into percentages, and color code return results as red for losses and green for gains.
 
       Next i
       
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C25").NumberFormat = "0.0%"
    Columns("B").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd
        If Cells(i, 3) > 0 Then
            'Change cell color green
            Cells(i, 3).Interior.Color = vbGreen
        ElseIf Cells(i, 3) < 0 Then
            'Change cell color red
            Cells(i, 3).Interior.Color = vbRed
        Else
            Cells(i, 3).Interior.Color = xlNone
            
        
        End If
        
        
    
    Next i

Lastly a message showing the total runtime of the script and an end to the subroutine.

    endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
   
    End Sub

The script was then refactored to loop through each ticker all at once.

    Sub AllStocksAnalysisRefactored()
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
    For n = 0 To 11
        tickerVolumes(n) = 0
        tickerStartingPrices(n) = 0
        tickerEndingPrices(n) = 0
        
    Next n
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
        If Cells(i, 1).Value = tickers(tickerIndex) Then
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
       

            '3d Increase the tickerIndex.
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

    End Sub


## Results

By using the script, the results showed that only ENPH and RUN company made any postitive results.


![image](https://user-images.githubusercontent.com/85597961/125180686-b827a900-e1b1-11eb-88e7-24b27ad94419.png)


## Summary

The first script was able to run through all the data using variables for company name, total volume, and yearly returns by looking through 12 different "ticker" variables to output a formatted numerical value for total volume and returns. Refactoring the scritp to run through all the calculations through the variable "tickerInex" made the script run faster and had a run time of around .79 seconds for 2017 and .77 seconds for 2018.

2017

![image](https://github.com/Zeilus/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png.PNG)

2018

![image](https://github.com/Zeilus/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.png.PNG)
