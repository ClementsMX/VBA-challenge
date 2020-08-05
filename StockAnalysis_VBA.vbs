
' Steps:
' ----------------------------------------------------------------------------

' Create a script that will loop through all the stocks for one year and output the following information.

' The ticker symbol.

' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

' The total stock volume of the stock.

' You should also have conditional formatting that will highlight positive change in green and negative change in red.


Sub StockAnalysis()
    
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS (In test file there is one year of information)
    ' --------------------------------------------
    For Each wks In Worksheets

        ' --------------------------------------------
        ' Variables
        ' --------------------------------------------

        Dim RowDisplay As Integer
        Dim InitialRow As Variant
        
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        Dim PercentChange As Double
        Dim YearlyChange As Double
                        
        Dim TotalStockVol As Variant
        
        Dim GreatestIncrease As Double
        Dim TickerIncrease As String
        Dim GreatestDecrease As Double
        Dim TickerDecrease As String
        Dim GreatestVolume As Variant
        Dim TickerVolume As String
        
                       
        RowDisplay = 1
        InitialRow = 2
        TotalStockVol = 0
        
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestVolume = 0
        
        ' Headers
        wks.Cells(RowDisplay, 9).Value = "Ticker"
        wks.Cells(RowDisplay, 10).Value = "Yearly Change"
        wks.Cells(RowDisplay, 11).Value = "Percent Change"
        wks.Cells(RowDisplay, 12).Value = "Total Stock Volume"
        
        ' Determine the Last Row
        LastRow = wks.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = InitialRow To LastRow
        
            ' Saving the opening price at the beginning of a given year
            If InitialRow = i Then
                OpeningPrice = wks.Cells(i, 3)
            End If
            
            ' Sum the Vol
            TotalStockVol = TotalStockVol + wks.Cells(i, 7).Value
            
            ' Checking if the next element in the column is the same
            If wks.Cells(i + 1, 1).Value <> wks.Cells(i, 1).Value Then
            
                RowDisplay = RowDisplay + 1
                ClosingPrice = wks.Cells(i, 6)
                
                YearlyChange = ClosingPrice - OpeningPrice
                
                    If OpeningPrice > 0 Then
                        PercentChange = ((YearlyChange / OpeningPrice))
                    Else
                        PercentChange = ((YearlyChange / 1))
                    End If
                    
                     
                        If PercentChange > GreatestIncrease Then
                            GreatestIncrease = PercentChange
                            TickerIncrease = wks.Cells(i, 1)
                        End If
                        
                        If PercentChange < GreatestDecrease Then
                            GreatestDecrease = PercentChange
                            TickerDecrease = wks.Cells(i, 1)
                        End If
                        
                        If TotalStockVol > GreatestVolume Then
                            GreatestVolume = TotalStockVol
                            TickerVolume = wks.Cells(i, 1)
                        End If
                        
                                
                ' Display calculations
                wks.Cells(RowDisplay, 9).Value = wks.Cells(i, 1)
                wks.Cells(RowDisplay, 10).Value = YearlyChange
                
                wks.Cells(RowDisplay, 11).Value = Format(PercentChange, "0.00%")
                
                
                wks.Cells(RowDisplay, 12).Value = TotalStockVol
                
                    If YearlyChange < 1 Then
                        wks.Cells(RowDisplay, 10).Interior.ColorIndex = 3
                    Else
                        wks.Cells(RowDisplay, 10).Interior.ColorIndex = 4
                    End If
                
                InitialRow = i + 1
                TotalStockVol = 0
                OpeningPrice = 0
                ClosingPrice = 0
                YearlyChange = 0
                PercentChange = 0
                
            End If
        Next i
                ' Display challenge
                wks.Cells(2, 15).Value = "Greatest % Increase"
                wks.Cells(3, 15).Value = "Greatest % Decrease"
                wks.Cells(4, 15).Value = "Greatest Total Volume"
                
                wks.Cells(1, 16).Value = "Ticker"
                wks.Cells(2, 16).Value = TickerIncrease
                wks.Cells(3, 16).Value = TickerDecrease
                wks.Cells(4, 16).Value = TickerVolume
                
                
                wks.Cells(1, 17).Value = "Value"
                wks.Cells(2, 17).Value = Format(GreatestIncrease, "0.00%")
                wks.Cells(3, 17).Value = Format(GreatestDecrease, "0.00%")
                
                wks.Cells(4, 17).Value = GreatestVolume
                
    
    Next wks
End Sub ' End StockAnalysis
