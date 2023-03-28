Sub StockData()


For Each ws In Worksheets
    
    Dim yearlyChange As Double
    
    Dim percentChange As Double
    
    Dim i As Long
    
    ws.Cells(1, 10) = "Ticker"
    
    ws.Cells(1, 11) = "Yearly Change"
    
    ws.Cells(1, 12) = "Percent Change"
    
    ws.Cells(1, 13) = "Total Stock Volume"
    
    ws.Cells.EntireColumn.AutoFit
    
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    yearlyChange = 0
    
    totalStockVol = 0
        
    Rowcounter = 2

    a = 2
    
For i = 2 To lastRow
    
    
    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        
        totalStockVol = totalStockVol + ws.Cells(i, 7).Value
            
            ws.Cells(Rowcounter, 10).Value = ws.Cells(i, 1).Value
            ws.Cells(Rowcounter, 13).Value = totalStockVol
        
            
        yearlyChange = ws.Cells(i, 6) - ws.Cells(a, 3)
            
            ws.Cells(Rowcounter, 11).Value = yearlyChange
        
        percentChange = (yearlyChange / (ws.Cells(a, 3)))
            
            ws.Cells(Rowcounter, 12).Value = percentChange
            ws.Cells(Rowcounter, 12).Value = FormatPercent(Cells(Rowcounter, 12).Value)
            
        Rowcounter = Rowcounter + 1
        
        a = i + 1
        
        totalStockVol = 0
    
        
    Else
        totalStockVol = totalStockVol + ws.Cells(i, 7).Value
            
            End If
            
        
        
         If ws.Cells(Rowcounter, 11).Value > 0 Then
            ws.Cells(Rowcounter, 11).Interior.ColorIndex = 4
            
            Else
            ws.Cells(Rowcounter, 11).Interior.ColorIndex = 3
            
            End If
            
          If ws.Cells(Rowcounter, 12).Value > 0 Then
            ws.Cells(Rowcounter, 12).Interior.ColorIndex = 4
            
            Else
            ws.Cells(Rowcounter, 12).Interior.ColorIndex = 3
            
            End If
        
        Next i
        
lastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
    Dim increase_greatest_ticker As String
    
    Dim decrease_greatest_ticker As String
    
    Dim greatest_total_vol_ticker As String
    
    Dim increase_greatest As Double

    Dim decrease_greatest As Double

    Dim greatest_total_vol As Double

For i = 2 To lastRow

        If ws.Cells(i, 12).Value > increase_greatest Then
            
            increase_greatest = ws.Cells(i, 12).Value
            
            increase_greatest_ticker = ws.Cells(i, 10).Value
            
            
        End If
          
        If ws.Cells(i, 12).Value < decrease_greatest Then
            
            decrease_greatest = ws.Cells(i, 12).Value
            
            decrease_greatest_ticker = ws.Cells(i, 10).Value
              
        End If
        
         If ws.Cells(i, 13).Value > greatest_total_vol Then
            
            greatest_total_vol = ws.Cells(i, 13).Value
            
            greatest_total_vol_ticker = ws.Cells(i, 10).Value
        
        End If
        
        Next i
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = increase_greatest_ticker
        ws.Cells(2, 17).Value = increase_greatest
        ws.Cells(2, 17).Value = FormatPercent(ws.Cells(2, 17).Value)
        
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = decrease_greatest_ticker
        ws.Cells(3, 17).Value = decrease_greatest
        ws.Cells(3, 17).Value = FormatPercent(ws.Cells(3, 17).Value)
        
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = greatest_total_vol_ticker
        ws.Cells(4, 17).Value = greatest_total_vol
        
    Next ws
End Sub
