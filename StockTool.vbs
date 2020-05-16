Sub StockTool():

    ' Set a variable for workbook
    Dim wb As Workbook
    
    ' Set a variable for worksheet
    Dim ws As Worksheet
    
    ' Set a variable for holding the stock name
    Dim StockName As String
    
    ' Set a variable for holding the number of days being analyzed
    Dim DayCounter As Integer
    
    ' Set initial value of Days
    DayCounter = 0
    
    ' Set a variable for holding the total volume of stock trades
    Dim StockVolumeTotal As Double
    
    ' Set initial value of StockVolumeTotal
    StockVolumeTotal = 0
    
    ' Set a variable for holding the yearly change of a stock's price
    Dim YearlyChange As Double
    
    ' Set a variable for holding the percent change of a stock's price
    Dim PercentChange As Double

    ' Set a variable for tracking location on final summary table
    Dim SummaryRow As Integer

    ' Set initial value of SummaryRow
    SummaryRow = 2
    
        For Each ws In ActiveWorkbook.Worksheets
    
            ' Set initial value for the last row on the stock list
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
            ' Create column headers for final summary table
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
    
            ' Loop through stock data
            For i = 2 To LastRow

                ' If StockTool is still within the same StockName, then:
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                    ' Set StockName
                    StockName = ws.Cells(i, 1).Value
                
                    ' Set YearlyChange of StockName
                    YearlyChange = ws.Cells(i, 6).Value - ws.Cells(i - DayCounter, 3).Value
                
                        ' Set PercentChange of StockName
                        If ws.Cells(i - DayCounter, 3).Value = 0 Then
                        PercentChange = 0
                    
                        Else
                        PercentChange = (ws.Cells(i, 6).Value - ws.Cells(i - DayCounter, 3).Value) / ws.Cells(i - DayCounter, 3).Value
                    
                        End If
                
                        ' Set StockVolumeTotal of StockName
                        StockVolumeTotal = StockVolumeTotal + ws.Cells(i, 7).Value

                    ' Print StockName in final summary table
                    ws.Range("I" & SummaryRow).Value = StockName
                
                    ' Print YearlyChange in final summary table
                    ws.Range("J" & SummaryRow).Value = YearlyChange
                
                        ' If YearlyChange is positive color cell green, if not color cell red
                        If ws.Range("J" & SummaryRow).Value > 0 Then
                        ws.Range("J" & SummaryRow).Interior.Color = vbGreen
                    
                        Else
                        ws.Range("J" & SummaryRow).Interior.Color = vbRed
                    
                        End If
                
                    ' Print PercentChange in final summary table
                    ws.Range("K" & SummaryRow).Value = PercentChange

                    ' Print StockVolumeTotal in final summary table
                    ws.Range("L" & SummaryRow).Value = StockVolumeTotal

                    ' Add to SummaryRow to move down one row within final summary table
                    SummaryRow = SummaryRow + 1
      
                    ' Reset StockVolumeTotal for the next StockName
                    StockVolumeTotal = 0
                
                    ' Reset the number of days being analyzed for the next StockName
                    DayCounter = 0

                ' If StockTool is still within the same StockName
                Else

                    ' Add to StockVolumeTotal of StockName
                    StockVolumeTotal = StockVolumeTotal + ws.Cells(i, 7).Value
                
                    ' Add to the number of days being analyzed
                    DayCounter = DayCounter + 1

                End If
            
            Next i
            
        ' Insert labels for increase/decrease table
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ' Set a variable for the maxmium PercentChange increase
        Dim MaxPercentIncrease As Double
        
        ' Find the greatest PercentChange increase and insert on sheet
        MaxPercentIncrease = Application.WorksheetFunction.Max(ws.Range("K:K"))
        ws.Cells(2, 17).Value = MaxPercentIncrease
        
        ' Set a variable for the greatest PercentChange decrease
        Dim MaxPercentDecrease As Double
        
        ' Find the greatest PercentChange decrease and insert on sheet
        MaxPercentDecrease = Application.WorksheetFunction.Min(ws.Range("K:K"))
        ws.Cells(3, 17).Value = MaxPercentDecrease
        
        ' Find the greatest total volume and insert on sheet
        MaxTotalVolume = Application.WorksheetFunction.Min(ws.Range("L:L"))
        ws.Cells(4, 17).Value = MaxTotalVolume
     
            ' Loop through data again
            For j = 2 To LastRow
            
                If ws.Cells(j, 11).Value = MaxPercentIncrease Then
                ws.Cells(2, 16).Value = ws.Cells(j, 1).Value
                
                ElseIf ws.Cells(j, 11).Value = MaxPercentDecrease Then
                    ws.Cells(3, 16).Value = ws.Cells(j, 1).Value
                    
                    ElseIf ws.Cells(j, 12).Value = MaxTotalVolume Then
                        ws.Cells(4, 16).Value = ws.Cells(j, 1).Value
                    
                End If
                
            Next j
            
        ' Reset SummaryRow for the next worksheet
        SummaryRow = 2
        
    Next ws
            
End Sub



