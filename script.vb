Sub ticker():

    For Each ws In Worksheets

        Dim i As Long
        Dim rowcount As Long
        Dim lastrow As Long
        Dim sum As Double
        Dim percentage_change As Double
        Dim opening As Double
        Dim closing As Double
        
        lastrow = ws.Cells(Rows.count, 1).End(xlUp).Row
        rowcount = 2
        sum = 0
        opening = ws.Cells(2, 3).Value

        ws.Cells(1, 9) = "ticker"
        ws.Cells(1, 10) = "yearly change"
        ws.Cells(1, 11) = "percentage change"
        ws.Cells(1, 12) = "total stock volumn"
        
        Cells(2, 15).Value = "Greatest % increase"
        Cells(3, 15).Value = "Greatest % decrease"
        Cells(4, 15).Value = "Greatest total volumn"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        Dim ws_count As Integer
        Dim ticker_positive As String
        Dim ticker_negative As String
        
        ws_count = ActiveWorkbook.Worksheets.count
     
        Cells(2, 17).Value = 0
        Cells(3, 17).Value = 0
        Cells(4, 17).Value = 0
                  
        
            
        For i = 2 To lastrow
        
            sum = sum + ws.Cells(i, 7).Value
            
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
                ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(rowcount, 12).Value = sum
                
                closing = ws.Cells(i, 6).Value
                percentage_change = (closing - opening) / opening
                ws.Cells(rowcount, 10).Value = closing - opening
                
                If (ws.Cells(rowcount, 10).Value > 0) Then
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 4
                ElseIf (ws.Cells(rowcount, 10).Value < 0) Then
                    ws.Cells(rowcount, 10).Interior.ColorIndex = 3
                End If
                
                ws.Cells(rowcount, 11).Value = percentage_change
                ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                
                opening = ws.Cells(i + 1, 3).Value
 
                rowcount = rowcount + 1
                sum = 0
                
             End If
             
               'seperate table
             If (ws.Cells(i, 11).Value > 0 And ws.Cells(i, 11).Value > Cells(2, 17).Value) Then
                Cells(2, 17).Value = ws.Cells(i, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"
                ticker_positive = ws.Cells(i, 1).Value
                
            ElseIf (ws.Cells(i, 11).Value < 0 And ws.Cells(i, 11).Value < Cells(3, 17).Value) Then
                Cells(3, 17).Value = ws.Cells(i, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
                ticker_negative = ws.Cells(i, 1).Value
            End If
            
        'greatest total
            If (ws.Cells(i, 12).Value > Cells(4, 17).Value) Then
                Cells(4, 17).Value = ws.Cells(i, 12).Value
                Cells(4, 16).Value = ws.Cells(i, 1).Value
            End If
            
        Next i
        Cells(2, 16).Value = ticker_positive
        Cells(3, 16).Value = ticker_negative

            
    Next ws
    

End Sub
