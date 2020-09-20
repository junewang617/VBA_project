Sub Stocks()
    Dim Ticker As String
    Dim Yearly_Change As Double, Percent_Change As Double
    Dim Stock_Volume As Double
    Dim Opening_Price As Double, Closing_Price As Double
   
    For Each ws In Worksheets
        Stock_Volume = 0
        Opening_Price = ws.Cells(2, 3).Value
        Dim Stock_Summary_Row As Integer
        Stock_Summary_Row = 2
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Account for division by zero 
                if Opening_Price = 0 Then
                    goto skipthisiteration
                End if

                Ticker = ws.Cells(i, 1).Value
                Closing_Price = ws.Cells(i, 6).Value
                Yearly_Change = Closing_Price - Opening_Price
                Percent_Change = ((Closing_Price - Opening_Price) / Opening_Price) * 100
                Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
                
                ws.Range("I" & Stock_Summary_Row).Value = Ticker
                ws.Range("J" & Stock_Summary_Row).Value = Yearly_Change
                ws.Range("K" & Stock_Summary_Row).Value = Percent_Change
                ws.Range("L" & Stock_Summary_Row).Value = Stock_Volume
                
                If Yearly_Change > 0 Then
                    ws.Range("J" & Stock_Summary_Row).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & Stock_Summary_Row).Interior.ColorIndex = 3
                End If
                
                Stock_Summary_Row = Stock_Summary_Row + 1
                Stock_Volume = 0
                Opening_Price = ws.Cells(i+1, 3).Value

            Else
                Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
            End If
            
            skipthisiteration:
        Next i
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Range("K2:K" & Stock_Summary_Row).NumberFormat = "0.00\%"

        'Next summary table
        Dim GreatestPercentIncrease as Double
        Dim GreatestPercentDecrease as Double
        Dim GreatestTotalVolume as Double
        Dim MaxStock as String, LowStock as String, MaxVolume as String
            NextLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
            Ticker = ws.Cells(i, 9).Value
            
            GreatestPercentIncrease = WorksheetFunction.Max(ws.Range("K2:K" & NextLastRow))
            MaxStock = WorksheetFunction.Match(GreatestPercentIncrease, ws.Range("K2:K" & NextLastRow), 0)
            ws.Cells(2, 15).Value = ws.Cells(MaxStock + 1, 9)
            ws.Cells(2, 16).Value = GreatestPercentIncrease
            
            GreatestPercentDecrease = WorksheetFunction.Min(ws.Range("K2:K" & NextLastRow))
            LowStock = WorksheetFunction.Match(GreatestPercentDecrease, ws.Range("K2:K" & NextLastRow), 0)
            ws.Cells(3, 15).Value = ws.Cells(LowStock + 1, 9)
            ws.Cells(3, 16).Value = GreatestPercentDecrease
            
            GreatestTotalVolume = WorksheetFunction.Max(ws.Range("L2:L" & NextLastRow))
            MaxVolume = WorksheetFunction.Match(GreatestTotalVolume, ws.Range("L2:L" & NextLastRow), 0)
            ws.Cells(4, 15).Value = ws.Cells(MaxVolume + 1, 9)
            ws.Cells(4, 16).Value = GreatestTotalVolume
            
            ws.Cells(1, 15).Value = "Ticker"
            ws.Cells(1, 16).Value = "Value"
            ws.Cells(2, 14).Value = "Greatest % Increase"
            ws.Cells(3, 14).Value = "Greatest % Decrease"
            ws.Cells(4, 14).Value = "Greatest Total Volume"
            
            ws.Range("P2:P3").NumberFormat = "0.00\%"
            ws.Cells.EntireColumn.Autofit
    Next ws
End Sub