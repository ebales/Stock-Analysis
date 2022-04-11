# Stock-Analysis
Week 2

Insights:

I wasn't able to successfully run the refactor code to reduce the amount of time required to analyze the stocks data.

It would be useful to reduce time when working with larger data sets, but something in the reduced loops got off track for me and I wasn't able to complete the program.

Code I had:

Sub StockAnalysis()



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
        Dim tickerVolume(11) As Long
        Dim tickerStartingPrice(11) As Long
        Dim tickerEndingPrice(11) As Long
        Dim volume As Long
        
    Worksheets(yearValue).Activate
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        tickerIndex = 0
            tickerVolume(tickerIndex) = 0
        
        
      ''2b) Loop over all the rows in the spreadsheet.
        For j = 2 To 3013
    
        '3a) Increase volume for current ticker
        If Cells(j, 1).Value = tickers(tickerIndex) Then
            volume = Cells(j, 8).Value
            tickerVolume(tickerIndex) = tickerVolume(tickerIndex) + volume
        End If
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
         If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrice(tickerIndex) = Cells(j, 6).Value
            
        End If
                
        
        '3c) check if the current row is the last row with the selected ticker
        
        If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrice(tickerIndex) = Cells(j, 6).Value
            
            
        End If
            
     
     '3d Increase the tickerIndex.
        If Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        End If
    
      Next j
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For j = 1 To 11
        
        Worksheets("All Stocks Analysis").Activate
            Cells(3 + j, 1).Value = tickers(j)
            Cells(3 + j, 2).Value = tickerVolume(j)
            Cells(3 + j, 3).Value = (tickerEndingPrice(j) - tickerStartingPrice(j)) / tickerStartingPrice(j)
            
            
                    
        
    Next j

End Sub

