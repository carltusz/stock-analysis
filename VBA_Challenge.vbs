Attribute VB_Name = "RefactoredStockAnalysis"
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
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
    'initialize ticker index to 0
    Dim tickerIndex As Integer
    tickerIndex = 0
    

    '1b) Create three output arrays
    
    'Output arrays created for ticker volume, start prices and end prices
    
    Dim tickerVolumes(12) As Double
    Dim tickerStartingPrices(12) As Double
    Dim tickerEndingPrices(12) As Double
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    'initialize volumes to 0
    
    For i = 0 To UBound(tickerVolumes)
        tickerVolumes(i) = 0
    Next i
    
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    
    'loop through all stonk data
    For i = 2 To rowCount
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If the previous row is not the same, then the current row is the first with the selected index
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            'set starting price
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
                
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i + 1, 1) <> tickers(tickerIndex) Then
            'set ending price
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
            
        End If
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        'ensure output goes to "All Stocks Analysis"
        Worksheets("All Stocks Analysis").Activate
        'Print ticker, volume, % change.
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

    datarowstart = 4
    dataRowEnd = 15

    For i = datarowstart To dataRowEnd
        'if %change increases, indicate gain with green fill. Else indicate loss with red fill
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
    'close the timer
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
