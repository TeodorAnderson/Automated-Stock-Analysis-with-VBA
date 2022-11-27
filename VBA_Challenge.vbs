Attribute VB_Name = "Module1"
Sub AllStocksAnalysis()
    '1)Ask for year input and start timer
   Dim startTime As Single
    Dim endTime  As Single
    '
    yearValue = InputBox("What year would you like to run the analysis on?")
    startTime = Timer
    '2)Format the output sheet on All Stocks Analysis worksheet
Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (" + yearValue + ")"""
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    '3) Initialize array of all tickers
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
   '3a) Initialize variables and arrays for starting price, ending price and volume
   Dim tickerStartingPrice(12) As Single
   Dim tickerEndingPrice(12) As Single
   Dim tickerVolume(12) As Long
   For TV = 0 To 11
   tickerVolume(TV) = 0
   Next TV
        '3b) Activate data worksheet
    Worksheets(yearValue).Activate
   '3c) Get the number of rows to loop over
   RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    '4) Loop through tickers and create variable for ticker index
   tickerIndex = 0
    ticker = tickers(tickerIndex)
    '5) loop through rows in the data
       Worksheets(yearValue).Activate
       For rowloop = 2 To RowCount
           '5a) Get total volume for current ticker
                tickerVolume(tickerIndex) = tickerVolume(tickerIndex) + Cells(rowloop, 8).Value
           '5b) get starting price for current ticker
           If Cells(rowloop - 1, 1).Value <> tickers(tickerIndex) And Cells(rowloop, 1).Value = tickers(tickerIndex) Then
                tickerStartingPrice(tickerIndex) = Cells(rowloop, 6).Value
            End If
           '5c) get ending price for current ticker
           If Cells(rowloop + 1, 1).Value <> tickers(tickerIndex) And Cells(rowloop, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrice(tickerIndex) = Cells(rowloop, 6).Value
           End If
           '5c)Increase ticker
           If Cells(rowloop + 1, 1).Value <> tickers(tickerIndex) And Cells(rowloop, 1).Value = tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
        Next rowloop
       '6)Loop through all arrays and output values
    For arrayloop = 0 To 11
    Worksheets("All Stocks Analysis").Activate
        Cells(4 + arrayloop, 1).Value = tickers(arrayloop)
        Cells(4 + arrayloop, 2).Value = tickerVolume(arrayloop)
        Cells(4 + arrayloop, 3).Value = tickerEndingPrice(arrayloop) / tickerStartingPrice(arrayloop) - 1
    Next arrayloop
     '7)Formatting output
      Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit
    dataRowStart = 4
    dataRowEnd = 15
    '7a)Portray green or red color for postive and negative values, respectively
    For tickerIndex = dataRowStart To dataRowEnd
    If Cells(tickerIndex, 3) > 0 Then
            Cells(tickerIndex, 3).Interior.Color = vbGreen
            Else
            Cells(tickerIndex, 3).Interior.Color = vbRed
    End If
    Next tickerIndex
    '8)Output runtime
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub

Sub ClearWorksheet()

    Cells.Clear

End Sub
