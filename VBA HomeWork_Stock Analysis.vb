Sub StockData_Analysis()

    For Each ws In Worksheets
    
        MsgBox ("Looping for Year " & ws.Name)
    
        Dim lastrow As Double
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim ticker As String
        
        Dim total_volume As Double
        total_volume = 0
        
        Dim yearly_change As Double
        
        Dim percentage_change As Double
        
        Dim summary_row As Integer
        summary_row = 2
        
        Dim closing_price As Double
        
        Dim opening_price As Double
        
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 9).Font.Bold = True
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 10).Font.Bold = True
        ws.Cells(1, 11) = "Percentage Change"
        ws.Cells(1, 11).Font.Bold = True
        ws.Cells(1, 12) = "Total Volume"
        ws.Cells(1, 12).Font.Bold = True
         
        For i = 2 To lastrow
        
            total_volume = total_volume + ws.Cells(i, 7)                    ' Part (i) Calculating the total volume
            
            If ws.Cells(i, 1) <> ws.Cells(i - 1, 1) Then
            opening_price = ws.Cells(i, 3)
            
            ElseIf ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
            ticker = ws.Cells(i, 1)
            closing_price = ws.Cells(i, 6)
            yearly_change = (closing_price - opening_price)
                        
                If opening_price = 0 Then
                percentage_change = 0
                
                Else
                percentage_change = yearly_change / opening_price
                
                End If
                    
            ws.Cells(summary_row, 9) = ticker                               ' Returning all the unique tickers along with total_volume
            ws.Cells(summary_row, 10) = yearly_change                       ' Part (ii) Calculating the yearly change and percentage change
            ws.Cells(summary_row, 11) = percentage_change                   ' in the above loop 
            ws.Cells(summary_row, 12) = total_volume
                    
            total_volume = 0
            summary_row = summary_row + 1
            opening_price = 0
            closing_price = 0
                        
            End If
        
        Next i
        
        ws.Range("K:K").NumberFormat = "0.00%"

    
        lastrow_new = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    
            For i = 2 To lastrow_new
    
                If ws.Cells(i, 11) > 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 4                 ' Conditional Formatting, if yearly change > 0 then green and                
                                                                        ' if less than zero then red
                Else
                ws.Cells(i, 11).Interior.ColorIndex = 3
        
            End If
    
            Next i

        Dim ticker1, ticker2, ticker3 As String                        ' Part (iii) creating 3 variables of ticker to return 3 different tickers 
        
        Dim greatest_value As Double
        greatest_value = 0
        
        Dim lowest_value As Double
        lowest_value = 0
        
        Dim value As Double
        value = 0
        
        Dim summary_row2 As Integer
        summary_row2 = 2
        
        ws.Cells(1, 15) = "Ticker"
        ws.Cells(1, 15).Font.Bold = True
        ws.Cells(1, 16) = "Value"
        ws.Cells(1, 16).Font.Bold = True
        ws.Cells(2, 14) = "Greatest % Increase"
        ws.Cells(2, 14).Font.Bold = True
        ws.Cells(3, 14) = "Greatest % Decrease"
        ws.Cells(3, 14).Font.Bold = True
        ws.Cells(4, 14) = "Greatest Total Volume"
        ws.Cells(4, 14).Font.Bold = True
        
                For i = 2 To lastrow_new
    
                    If ws.Cells(i, 11) > ws.Cells(i + 1, 11) And ws.Cells(i, 11) > greatest_value Then      ' Calculating the greatest_value, lowest_value from the values calculated from Part (ii)
                    greatest_value = ws.Cells(i, 11)
                    ticker1 = ws.Cells(i, 9)
                
                    ElseIf ws.Cells(i, 11) < ws.Cells(i + 1, 11) And ws.Cells(i, 11) < lowest_value Then
                    lowest_value = ws.Cells(i, 11)
                    ticker2 = ws.Cells(i, 9)
                
                    ElseIf ws.Cells(i, 12) > ws.Cells(i + 1, 12) And ws.Cells(i, 12) > value Then
                    value = ws.Cells(i, 12)
                    ticker3 = ws.Cells(i, 9)
                
                    End If
            
                    Next i
        
        ws.Cells(2, 15) = ticker1
        ws.Cells(3, 15) = ticker2
        ws.Cells(4, 15) = ticker3
        ws.Cells(2, 16) = greatest_value
        ws.Cells(2, 16).NumberFormat = "0.00%"
        ws.Cells(3, 16) = lowest_value
        ws.Cells(3, 16).NumberFormat = "0.00%"
        ws.Cells(4, 16) = value
            
    Next ws
        
End Sub