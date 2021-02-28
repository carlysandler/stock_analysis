Attribute VB_Name = "Module1"
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    Range("A1").Font.Bold = True
    
    
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
    RowCount = Cells(Rows.Count, "A").End(xlUp).row
    
    '1a) Create a ticker Index
    For i = 0 To 11
        tickerIndex = tickers(i)
    
    '1b) Create three output arrays
    Dim tickerVolumes As Long
    Dim tickerStartingPrices As Single
    Dim tickerEndingPrices As Single
    

    ''2a) Create a for loop to initialize the tickerVolumes to zero.
        Worksheets(yearValue).Activate
        tickerVolumes = 0

  
        ''2b) Loop over all the rows in the spreadsheet.
        For j = 2 To RowCount
        
            'for every row in spreadsheet, if the value of cell(j,1) = "ticker(i)" then
            If Cells(j, 1).Value = tickerIndex Then
        
                '3a) Increase volume for current ticker
                 tickerVolumes = tickerVolumes + Cells(j, 8).Value
            
            End If
            
            '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
            
                tickerStartingPrices = Cells(j, 6).Value
      
            'End If
            End If
        
            '3c) check if the current row is the last row with the selected ticker
            If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
                
                'If the next row’s ticker doesn’t match, increase the tickerIndex.
                tickerEndingPrices = Cells(j, 6).Value
            
            'End If
            End If
            
        Next j
        
    'Code to output data for current ticker(i), "Total Daily Volume", and "Return"
    Worksheets("All Stocks Analysis").Activate
        
    Cells(4 + i, 1).Value = tickerIndex
    Cells(4 + i, 2).Value = tickerVolumes
    Cells(4 + i, 3).Value = tickerEndingPrices / tickerStartingPrices - 1
        
    Next i
    
    'Formatting cells in spreadsheet for visualization and correcting value format of each varaible
    
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("A3:C3").Font.Background = xlBackgroundAutomatic
    Range("A3:C3").Font.Italic = True
    Range("A3:C3").Interior.Color = RGB(200, 255, 255)
    Range("A3:C3").HorizontalAlignment = xlCenter
    
    'MsgBox Selection.Interior.ColorIndex
    'apply center alignment of header range
    'format range of output values for "All Stocks Analysis" worksheet taken from: https://software-solutions-online.com/excel-vba-set-horizontal-alignment-sample-code/
    
    Range("A4:C15").HorizontalAlignment = xlCenter
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.00%"
    Columns("B").AutoFit
    
    'Conditional formatting "Return" column output values for visualization
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd
    
        If Cells(i, 3) > 0 Then
        
            'Change cell color to green
            Cells(i, 3).Interior.Color = vbGreen
            
        ElseIf Cells(i, 3) < 0 Then
        
            'Change cell color to red
            Cells(i, 3).Interior.Color = vbRed
            
        Else
            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone
            
        End If
    
    Next i
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

