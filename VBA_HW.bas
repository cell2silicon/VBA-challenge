Sub VBA_HW():

' Variable for ticker symbol
Dim Ticker As String

' Variable for number of tickers
Dim TickerNumber As Integer

' Variable for last row of worksheet
Dim LastRow As Long

' Variable for the oprning price of the year
Dim OpeningPrice As Double

' Variable for the closing price of the year
Dim ClosingPrice As Double

' Variable for change in yearly price
Dim YearlyChange As Double

' Variable for percent of chnage of yearly value
Dim PercentChange As Double

' Variable for volume of stock
Dim StockVolume As Double

' Variable for greatest percent increase in value
Dim GreatestPercentIncrease As Double

' Variable for greatest percent increase ticker
Dim GreatestPercentIncreaseTicker As String

' Variable for greatest percent decrease in value
Dim GreatestPercentDecrease As Double

' Variable for greatest percent decrease ticker
Dim GreatestPercentDecreaseTicker As String

' Variable for greatest stock volume
Dim GreatestStockVolume As Double

' Variable for ticker of greatest stock volume
Dim GreatestStockVolumeTicker As String

' For each worksheet to go in loop
For Each ws In Worksheets

    ' To activate current worksheet
    ws.Activate

    ' Find the last row of each worksheet
    
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' To put headers in each worksheet
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Initialize variables for each worksheet.
    TickerNumber = 0
    Ticker = ""
    YearlyChange = 0
    OpeningPrice = 0
    PercentChange = 0
    StockVolume = 0
    
    ' Looping through the ticker list by skipping the header row
    For i = 2 To LastRow

        ' Value of current ticker symbol in column
        Ticker = Cells(i, 1).Value
        
        ' Starting price of the ticker for the year
        
        If OpeningPrice = 0 Then
            OpeningPrice = Cells(i, 3).Value
        End If
        
        ' Total stock volume for the given ticker
        StockVolume = StockVolume + Cells(i, 7).Value
        
        ' To calculate the above for next ticker
        
        If Cells(i + 1, 1).Value <> Ticker Then
        
            ' Add 1 when we get to next ticker in column
            
            TickerNumber = TickerNumber + 1
            Cells(TickerNumber + 1, 9) = Ticker
            
            ' Year end closing price for the ticker
            ClosingPrice = Cells(i, 6)
            
            ' Change in value for a year
            YearlyChange = ClosingPrice - OpeningPrice
            
            ' Adding yearly change value to its cell
            Cells(TickerNumber + 1, 10).Value = YearlyChange
            
            ' # Moderate_Solution

            ' Conditional formatting of cells
            
            ' If yearly change is positive, highlight cell with green
            If YearlyChange > 0 Then
                Cells(TickerNumber + 1, 10).Interior.ColorIndex = 4
                
            ' If yearly change is negative, highlight cell with red
            ElseIf YearlyChange < 0 Then
                Cells(TickerNumber + 1, 10).Interior.ColorIndex = 3

            End If
            
            
            ' Ticker change value percentage
            If OpeningPrice = 0 Then
                PercentChange = 0
            Else
                PercentChange = (YearlyChange / OpeningPrice)
            End If
            
            
            ' Formatting percent change value to percentage
            Cells(TickerNumber + 1, 11).Value = Format(PercentChange, "Percent")
            
            
            
            ' When we get to different ticker, set value to 0 again to start above calculations for new ticker
            OpeningPrice = 0
            
            ' Add stock volume to the relevant cell of each worksheet.
            Cells(TickerNumber + 1, 12).Value = StockVolume
            
            ' Setting stock volume to 0, when we get different ticker
            StockVolume = 0
        End If
        
    Next i
    
    '# Bonus Part

    ' Adding following headers and titles to relevant cells on each worksheet
    
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    ' Getting the value of the last row of ticker column("I")
    
    LastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    ' Set variable value for the follwing variables
    
    GreatestPercentIncrease = Cells(2, 11).Value
    GreatestPercentIncreaseTicker = Cells(2, 9).Value
    GreatestPercentDecrease = Cells(2, 11).Value
    GreatestPercentDecreaseTicker = Cells(2, 9).Value
    GreatestStockVolume = Cells(2, 12).Value
    GreatestStockVolumeTicker = Cells(2, 9).Value
    
    ' # Hard Solution

    ' Looping through the list by skipping the header row
    For i = 2 To LastRow
    
        ' Ticker with the greatest percent increase
        If Cells(i, 11).Value > GreatestPercentIncrease Then
            GreatestPercentIncrease = Cells(i, 11).Value
            GreatestPercentIncreaseTicker = Cells(i, 9).Value
        End If
        
        ' Ticker with the greatest percent decrease
        If Cells(i, 11).Value < GreatestPercentDecrease Then
            GreatestPercentDecrease = Cells(i, 11).Value
            GreatestPercentDecreaseTicker = Cells(i, 9).Value
        End If
        
        ' Ticker with the greatest stock volume
        If Cells(i, 12).Value > GreatestStockVolume Then
            GreatestStockVolume = Cells(i, 12).Value
            GreatestStockVolumeTicker = Cells(i, 9).Value
        End If
        
    Next i
    
    ' Formatting and adding above values to their relevant cells
    
    Range("P2").Value = Format(GreatestPercentIncreaseTicker, "Percent")
    Range("Q2").Value = Format(GreatestPercentIncrease, "Percent")
    Range("P3").Value = Format(GreatestPercentDecreaseTicker, "Percent")
    Range("Q3").Value = Format(GreatestPercentDecrease, "Percent")
    Range("P4").Value = GreatestStockVolumeTicker
    Range("Q4").Value = GreatestStockVolume
    
Next ws


End Sub

