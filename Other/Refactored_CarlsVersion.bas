Attribute VB_Name = "Refactored"
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single
    Dim i, j As Integer
    Dim yearValue As String
    Dim rowCount As Integer
    Dim output() As Variant
    Dim dailyVolume As Double
    Dim stockReturn As Double
    Dim openValue, closeValue As Double
    Dim ticker As String
    Dim tbl As ListObject

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
    

    '1b) Create three output arrays
    ReDim output(1 To 5, 1 To UBound(tickers))
    
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For j = 1 To UBound(output, 2)
        output(1, j) = tickers(j - 1)
        output(2, j) = 0
        output(3, j) = 0
        
    Next j
    
    j = 1
    
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To rowCount
    
        'get the variables
        ticker = Cells(i, 1)
        dailyVolume = Cells(i, 8)
        closeValue = Cells(i, 6)
    
        'loop through each row of the tickers to increase volume
        For j = 1 To UBound(output, 2)
        '3a) Increase volume for current ticker
            If ticker = output(1, j) Then
                'check if year open or close value
                If ticker <> Cells(i - 1, 1) Then output(3, j) = closeValue
                If ticker <> Cells(i + 1, 1) Then output(4, j) = closeValue
                
                'update volume
                output(2, j) = output(2, j) + dailyVolume
            End If
        Next j
    Next i
    
    'activate worksheet
    Worksheets("All Stocks Analysis").Activate
    
    'need to calculate return
    For j = 1 To UBound(output, 2)
        output(5, j) = output(4, j) / output(3, j) - 1
        
        Cells(j + 3, 1) = output(1, j)
        Cells(j + 3, 2) = output(2, j)
        Cells(j + 3, 3) = output(5, j)
    Next j
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        
        
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
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
