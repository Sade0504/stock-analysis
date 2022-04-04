# stock-analysis
## Overview of Project
### Purpose
A client was interested in investing in a perticular green energy stock "DQ". The client had no idea of how the stock was performing and wanted to invest solely on the fascination of the name. The original question was to access the total daily volume and the yearly return for the "DQ" stock. With data from 2017 and 2018, we were able to compile an effective analysis.
## Results
It was determined that "DQ's" 2018 yearly return was a lost of 63%. So we expanded our research to 11 other green stocks for analysis. The objective was to find the ticker, the total daily volume, and the return on each stock. See the steps below for an illustration of the steps taken to provide such information.
    Dim startTime As Single
    Dim endTime As Single
  
  yearValue = InputBox("What year would you like to run the analysis on?")
  
  startTime = Timer

'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"

    'Create a header row
    
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume "
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
    
    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
    End If
    
    '3c) check if the current row is the last row with the selected ticker
    
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

