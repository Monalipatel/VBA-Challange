' The VBA of Wall Street
Sub MultipleYearStockData(): 

    ' Loop All Worksheets
    For Each ws In Worksheets

        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("N2").Value = "Greatest % increase"
        ws.Range("N3").Value = "Greatest % decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Range("O1").Value = "Ticker"
        ws.Range("P1").Value = "Value"

    
        Dim TickerName As String
        Dim LastRow As Long
        Dim LastRowValue As Long
        
        Dim TotalTickerVolume As Double
        TotalTickerVolume = 0
        
        Dim SummaryTableRow As Long
        SummaryTableRow = 2
        
        Dim YearlyOpen As Double
        Dim YearlyClose As Double
        Dim YearlyChange As Double
        
        Dim PreviousAmount As Long
        PreviousAmount = 2
        
        Dim PercentChange As Double
        
        Dim GreatestIncrease As Double
        GreatestIncrease = 0
        
        Dim GreatestDecrease As Double
        GreatestDecrease = 0
        
        
        
        Dim GreatestTotalVolume As Double
        GreatestTotalVolume = 0

        ' Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow

            ' Add To Ticker Total Volume
            TotalTickerVolume = TotalTickerVolume + ws.Cells(i, 7).Value
            
            ' Check If We Are Still in the The Same Ticker Name or not
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then


                ' Set Ticker Name
                TickerName = ws.Cells(i, 1).Value
                
                ' Print Name of the ticker In The Summary Table
                ws.Range("I" & SummaryTableRow).Value = TickerName
                
                ' Print Total Amount of ticker To The Summary Table
                ws.Range("L" & SummaryTableRow).Value = TotalTickerVolume
                
                ' Reset
                TotalTickerVolume = 0

                ' Set Yearly Open, Yearly Close and Yearly Change Name
                YearlyOpen = ws.Range("C" & PreviousAmount)
                YearlyClose = ws.Range("F" & i)
                YearlyChange = YearlyClose - YearlyOpen
                ws.Range("J" & SummaryTableRow).Value = YearlyChange

               
                If YearlyOpen = 0 Then
                    PercentChange = 0
                Else
                    YearlyOpen = ws.Range("C" & PreviousAmount)
                    PercentChange = YearlyChange / YearlyOpen
                End If
                '
                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                ws.Range("K" & SummaryTableRow).Value = PercentChange

                '
                If ws.Range("J" & SummaryTableRow).Value >= 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                End If
            
                
                SummaryTableRow = SummaryTableRow + 1
                PreviousAmount = i + 1
                End If
            Next i

'

            ' Greatest % Increase, Greatest % Decrease and Greatest Total Volume
            LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
            ' Start Loop For Final Results
            For i = 2 To LastRow
                If ws.Range("K" & i).Value > ws.Range("P2").Value Then
                    ws.Range("P2").Value = ws.Range("K" & i).Value
                    ws.Range("O2").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("K" & i).Value < ws.Range("P3").Value Then
                    ws.Range("P3").Value = ws.Range("K" & i).Value
                    ws.Range("O3").Value = ws.Range("I" & i).Value
                End If

                If ws.Range("L" & i).Value > ws.Range("P4").Value Then
                    ws.Range("P4").Value = ws.Range("L" & i).Value
                    ws.Range("O4").Value = ws.Range("I" & i).Value
                End If

            Next i
        
            ws.Range("P2").NumberFormat = "0.00%"
            ws.Range("P3").NumberFormat = "0.00%"
       
        ws.Columns("I:P").AutoFit

    Next ws

End Sub
