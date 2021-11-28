Attribute VB_Name = "Classwork"
Sub MacroCheck()

    Dim testMessage As String
    testMessage = "Hello World!"
    
    'MsgBox (testMessage)
    Dim r As Range
    Set r = Application.InputBox("Select Range", "Table Range", "$A$1", Type:=8)
    
    Dim t As ListObject
    Set t = ThisWorkbook.ActiveSheet.ListObjects.Add(Source:=r)
    

End Sub

Sub DQAnalysis()
    
    'activate the DQ Analysis worksheet
    Worksheets("DQ Analysis").Activate
    Range("A1").Value = "DAQO (Ticker: DQ)"
    
    'create a worksheet object to reference the active worksheet (DQ Analysis)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("DQ Analysis")
    
    'create error handling in case of a pre-existing listobject
    'On Error Resume Next
    
    'clear pre-existing tables where applicable (cover everything in the macro)
    
    
    
    'this is the original lesson method
    'creating table header row
    ws.Cells(3, 1).Value = "Year"
    ws.Cells(3, 2).Value = "Total Daily Volume"
    ws.Cells(3, 3).Value = "Return"
    
    
'    'This is carl's special way
'    'creating a content-less table with only table headers defined
'    Dim t As ListObject
'    Set t = ws.ListObjects.Add(SourceType:=xlSrcRange, Source:=ws.Range(ws.Cells(3, 5), ws.Cells(3, 7)), Destination:=ws.Cells(1, 1))
'    t.HeaderRowRange(1, 1).Value = "Year"
'    t.HeaderRowRange(1, 2).Value = "Total Daily Volume"
'    t.HeaderRowRange(1, 3).Value = "Return"
'
    'back to the class method
    'creating a loop to sum the 2018 values
    Dim ws2018 As Worksheet
    Set ws2018 = ThisWorkbook.Worksheets("2018")
    
    'activate 2018
    Worksheets("2018").Activate
    
    'explicitly dimensionalized variables out of habit
    Dim rowStart As Integer
    rowStart = 2
    Dim rowEnd As Integer
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row ' from stackoverflow ref
    Dim totalVolume As Long
    totalVolume = 0
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    startingPrice = 0
    endingPrice = 0
    
    'set up loop
    For i = rowStart To rowEnd
        'increase only if ticker is DQ
        If ws2018.Cells(i, 1) = "DQ" Then
            
            'increase total volume
            totalVolume = totalVolume + ws2018.Cells(i, 8).Value
            
            'check if starting price
            If ws2018.Cells(i - 1, 1) <> "DQ" Then
                startingPrice = ws2018.Cells(i, 6)
            End If
            
            'check if ending price
            If ws2018.Cells(i + 1, 1) <> "DQ" Then
                endingPrice = ws2018.Cells(i, 6)
            End If
            
            
        End If
    Next i
    
    'MsgBox (totalVolume)
    
    'provide output
    Worksheets("DQ Analysis").Activate
    ws.Cells(4, 1).Value = 2018
    ws.Cells(4, 2).Value = totalVolume
    ws.Cells(4, 3).Value = endingPrice / startingPrice - 1
    
    
End Sub


Public Sub AllStocksAnalysis()

'dimensionalize variables
Dim allAnalysis As Worksheet
Dim dataSource As Worksheet
Dim i, j As Integer
Dim ticker As String
Dim tickers(11) As String
Dim rowEnd As Integer
Dim rowStart As Integer
Dim totalVolume As Double
Dim startValue As Double
Dim endValue As Double
Dim year As Integer
Dim startTime, endTime As Single

'get user input for target year
year = InputBox("Enter the year for analysis", "Stocks Analysis", Default:=2018)


'start timer
startTime = Timer


'set objects
Set allAnalysis = ThisWorkbook.Worksheets("All Stocks Analysis")
Set dataSource = ThisWorkbook.Worksheets(CStr(year))


'set noRows
rowStart = 2
rowEnd = dataSource.Cells(Rows.Count, "A").End(xlUp).Row ' from stackoverflow ref


'set title for worksheet
allAnalysis.Cells(1, 1).Value = "All Stocks (2018)"

'set headers for worksheet
allAnalysis.Cells(3, 1).Value = "Ticker"
allAnalysis.Cells(3, 2).Value = "Total Daily Volume"
allAnalysis.Cells(3, 3).Value = "Return"

'assign tickers
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

'WERE AT THE NESTED LOOPS SECTION OF 2.3.2
For i = 0 To 11
    
    ticker = tickers(i)
    
    'clear old data
    startValue = 0
    endValue = 0
    totalVolume = 0
    
    For j = rowStart To rowEnd
    
        'check for if the ticker matches
        If dataSource.Cells(j, 1) = ticker Then
            'increase total volume
            totalVolume = totalVolume + dataSource.Cells(j, 8)
            
            'check if opening value
            If dataSource.Cells(j - 1, 1) <> ticker Then
                startValue = dataSource.Cells(j, 6)
            End If
            
            'check if closing value
            If dataSource.Cells(j + 1, 1) <> ticker Then
                endValue = dataSource.Cells(j, 6)
            End If
        End If
    
    Next j
    
    'print out values
    allAnalysis.Cells(i + 4, 1).Value = ticker
    allAnalysis.Cells(i + 4, 2).Value = totalVolume
    allAnalysis.Cells(i + 4, 3).Value = endValue / startValue - 1
    
Next i


'Formatting
allAnalysis.Range("A3:C3").Font.FontStyle = "Bold"
allAnalysis.Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
allAnalysis.Range("B4:B15").NumberFormat = "#,##0"
allAnalysis.Range("C4:C15").NumberFormat = "0.0%"
allAnalysis.Columns("B").AutoFit

'apply conditional formatting
'set range limits
rowStart = 4
rowEnd = 15
'loop through range adjusting color based on returns
For i = rowStart To rowEnd
    Select Case allAnalysis.Cells(i, 3).Value
        Case Is > 0
            allAnalysis.Cells(i, 3).Interior.Color = vbGreen
        Case Is < 0
            allAnalysis.Cells(i, 3).Interior.Color = vbRed
        Case Else
            allAnalysis.Cells(i, 3).Interior.Color = xlNone
    End Select
Next i
        
endTime = Timer

MsgBox "Elapsed time " & (endTime - startTime)


End Sub

Sub ClearCells()

    ActiveSheet.Cells.Clear
    

End Sub
