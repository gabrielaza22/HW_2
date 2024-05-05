Sub stock_activity()
    
    Dim i As Long 'row number
    Dim vol_stock As Long 'column G value
    Dim Total_stock As Long 'total of the stocks
    Dim ticker As String 'ticker
    
    
    Dim close_price As Double
    Dim open_price As Double
    Dim qua_change As Double 'quaterly change column
    Dim percent_change As Double 'percent column K
    
    Dim great_increase As Double
    Dim great_drecrease As Double
    Dim great_vol As LongLong
    
    Dim K As Long
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
    
    
        Dim lastrow As Long
        
        ws.Range("J1").Value = "Quaterly change"
        ws.Range("I1").Value = "ticker"
        ws.Range("k1").Value = "percent change"
        ws.Range("L1").Value = "Total stock Volume"
        
       
       
        K = 2 'new ticker column
        Total_stock = 0
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
        open_price = Cells(2, 3).Value
    
        
        
        For i = 2 To lastrow:
        vol_stocks = ws.Cells(i, 7).Value
        ticker = ws.Cells(i, 1).Value
        
        
        
        'loop rows
        'check if next row ticker is different
        'if the same we only need to add to the Total_stock
        'if different, we add the last row and jump to the next ticker
        'reset the total stock to 0
        
        If (ws.Cells(i + 1, 1).Value <> ticker) Then
            total_stocks = total_stocks + vol_stocks
            
        'closing price
        close_price = ws.Cells(i, 6).Value
        qua_change = close_price - open_price
        percent_change = qua_change / open_price
        
        
        
        
            ws.Cells(K, 9).Value = ticker
            ws.Cells(K, 12).Value = total_stocks
            ws.Cells(K, 10).Value = qua_change
            ws.Cells(K, 11).Value = percent_change
            
            'color formatting
            If (qua_change > 0) Then
                ws.Cells(K, 10).Interior.ColorIndex = 4
                ws.Cells(K, 11).Interior.ColorIndex = 4
            ElseIf (qua_change < 0) Then
                ws.Cells(K, 10).Interior.ColorIndex = 3
                ws.Cells(K, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(K, 10).Interior.ColorIndex = 2
                ws.Cells(K, 11).Interior.ColorIndex = 2
            End If
            
           'Second leaderboard
                If ticker = ws.Cells(2, 1).Value Then
                 great_vol = Total_stock
                 great_increase = percent_change
                 great_decrease = percent_change
                 
                 Else
                 'compare
                    If Total_stock > great_vol Then
                      great_vol = Total_stock
                    End If
                    
                    
                    If percent_change > great_increase Then
                        great_increase = percent_change
                      End If
                    
                    
                    If percent_change < great_decrease Then
                         great_decrease = percent_change
                         End If
                    
            End If
            
             
                
                'percentage
                ws.Columns("k:k").NumberFormat = "0.00%"
                
                
                
                ws.Cells(3, 11).Value = great_decrease
                ws.Cells(4, 11).Value = great_vol
                ws.Cells(2, 11).Value = great_increase
            
                
                'reset
                total_stocks = 0
                K = K + 1
                open_price = ws.Cells(i + 1, 3).Value
                
            
       
     Else
         'we just add to the total
         total_stocks = total_stocks + vol_stocks
         
         End If
       
     
      
      
     Next
     Next ws
     
     
End Sub