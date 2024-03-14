Attribute VB_Name = "Module1"
Sub stockAnalysis()

    'This sections establishes all of our needed variables
    Dim ws As Worksheet
    Dim CurrentTicker As String
    Dim TotalVolume As Double
    Dim i As Long
    Dim j As Long
    Dim LastRow As Long
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim GreatestPerInc As Double
    Dim GreatestPerDec As Double
    Dim GreatestTotalVol As Double
    Dim GreatestPerIncTicker As String
    Dim GreatestPerDecTicker As String
    Dim GreatestTotalVolTicker As String
    
    'This loops moves us through each sheet in this workbook
    For Each ws In ThisWorkbook.Worksheets
    
        Worksheets(ws.Name).Select 'selects the current sheet
        LastRow = Cells(Rows.Count, "A").End(xlUp).row 'determines the last row in the main table of data
        CurrentTicker = Cells(2, 1)  'sets the first ticker
        j = 2 'sets my initial value for the row number in the new data
        Cells(2, 9) = CurrentTicker  'stores the first ticker
        OpenPrice = Cells(2, 3) 'sets the first opening price
        GreatestTotalVol = 0 'sets the baseline for greastest total volume
        GreatestPerInc = 0 'sets the baseline for greastest percentage increase
        GreatestPerDec = 0 'sets the baseline for greastest percentage decrease
        
        'This is the main loop for each sheets calculations
        For i = 2 To LastRow
        
            'This conditional reads for a change in the ticker symbol in the first column
            'and compared to the current stored ticker symbol
            If CurrentTicker = Cells(i, 1) Then
                TotalVolume = TotalVolume + Cells(i, 7)  'totals the current tickers volume
            Else
                Cells(j, 12) = TotalVolume 'places the total volume for the current ticker in the sheet
                ClosePrice = Cells(i - 1, 6) 'locks in the closing price for the current ticker
                YearlyChange = ClosePrice - OpenPrice 'Calculates the yearly change for the current ticker
                PercentChange = YearlyChange / OpenPrice 'Calculates the percentage change for the current ticker
                Cells(j, 10) = YearlyChange 'places the yearly change for the current ticker in the sheet
                Cells(j, 11) = PercentChange 'places the percentage change for the current ticker in the sheet
                
                'This loop applies color formatting to the yearly change data
                If Cells(j, 10) >= 0 Then
                    Cells(j, 10).Interior.Color = RGB(0, 200, 0)
                Else
                    Cells(j, 10).Interior.Color = RGB(200, 0, 0)
                End If
                
                'This loop stores the ticker values and symbols for the stocks with the
                'greatest percentage increase and decrease
                If PercentChange > GreatestPerInc Then
                    GreatestPerInc = PercentChange
                    GreatestPerIncTicker = CurrentTicker
                ElseIf PercentChange < GreatestPerDec Then
                    GreatestPerDec = PercentChange
                    GreatestPerDecTicker = CurrentTicker
                End If
                
                'This loop stores the ticker value and symbol for the stock with the greatest total volume
                If TotalVolume > GreatestTotalVol Then
                    GreatestTotalVol = TotalVolume
                    GreatestTotalVolTicker = CurrentTicker
                End If
                
                OpenPrice = Cells(i, 3) ' resets open price to the current tickers open price
                TotalVolume = Cells(i, 7) ' resets and begins new volume ticker total
                CurrentTicker = Cells(i, 1)  ' resets the ticker to the new current ticker
                j = j + 1 'sets the next value for the row number in the new data
                Cells(j, 9) = CurrentTicker  'stores the current ticker into the next cell
            End If
            
        Next i
        
        'This section addresses the last row of data that gets missed by the previous loop
        ClosePrice = Cells(LastRow, 6)
        YearlyChange = ClosePrice - OpenPrice
        Cells(j, 10) = YearlyChange
        If Cells(j, 10) >= 0 Then
            Cells(j, 10).Interior.Color = RGB(0, 200, 0)
        Else
            Cells(j, 10).Interior.Color = RGB(200, 0, 0)
        End If
        PercentChange = YearlyChange / OpenPrice
        Cells(j, 11) = PercentChange
        Cells(j, 12) = TotalVolume
        
        'This next section populates the singular values in the sheet
        Cells(2, 17) = GreatestPerInc
        Cells(3, 17) = GreatestPerDec
        Cells(4, 17) = GreatestTotalVol
        Cells(2, 16) = GreatestPerIncTicker
        Cells(3, 16) = GreatestPerDecTicker
        Cells(4, 16) = GreatestTotalVolTicker
        
        'This next section applies headers/labels for all data to be collected
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greastest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        'This next section applies our percentage formatting and autofits the columns to an ideal width
        Columns("K:K").NumberFormat = "0.00%"
        Range("Q2:Q3").NumberFormat = "0.00%"
        Columns("A:Q").EntireColumn.AutoFit
        
    Next ws
    
End Sub
