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
    dim tickerIndex as single
    tickerIndex = 0
    'tickerIndex starts at 0 in the for loop...

    '1b) Create three output arrays   
    dim tickerVolumes(12) as long
    dim tickerStartingPrices(12) as Single
    dim tickerEndingPrices(12) as Single
    'tickers array already initialized

    ''2a) Create a for loop to initialize the tickerVolumes to zero. 
    For tickerIndex = 0 to 11
    tickerVolumes(tickerIndex) = 0

        ''2b) Loop over all the rows in the spreadsheet.
        sheets(yearValue).activate
        For i = 2 To RowCount
    
            '3a) Increase volume for current ticker
            if Cells(i, 1).Value = tickers(tickerIndex) then
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + cells(i, 8).value
            end if

            '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = cells(i, 6).value
            End If
        
            '3c) check if the current row is the last row with the selected ticker
            'If the next row’s ticker doesn’t match, increase the tickerIndex.
            If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = cells(i, 6).value            

                '3d) Increase the tickerIndex.
                tickerIndex = tickerIndex + 1
                'Reset tickerVolumes    
                tickerVolumes(tickerIndex) = 0
            
            End If
        next i
    Next tickerIndex
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For j = 0 To 11
        Worksheets("All Stocks Analysis").Activate
        cells(j + 4, 1).value = tickers(j)
        cells(j + 4, 2).value = tickerVolumes(j)
        cells(j + 4, 3).value = tickerEndingPrices(j)/tickerStartingPrices(j) - 1
    Next j
    
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