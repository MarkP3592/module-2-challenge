Sub Multiple_Year()
    Dim i, j, k As Long
    Dim ws As Worksheet
    
    
    For Each ws In ThisWorkbook.Worksheets ' loop through worksheets
    Dim ticker As String
    Dim open_price, close_price, qtly_change, percent_change, volume As Double
    Dim lastrow, qtly_change_lastrow As Long
    Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        volume = 0
        lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        open_price = ws.Cells(2, 3).Value
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Volume"
    
        ' begin loop
        For i = 2 To lastrow
            
            ' set ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
            ticker = ws.Cells(i, 1).Value
            ws.Range("I" & Summary_Table_Row).Value = ticker
            
            ' add to volume and print
            volume = volume + ws.Cells(i, 7).Value
            ws.Range("L" & Summary_Table_Row).Value = volume
            
            ' grab close price
            close_price = ws.Cells(i, 6).Value
            
            ' calculate quarterly change and print
            qtly_change = close_price - open_price
            ws.Range("J" & Summary_Table_Row).Value = qtly_change
            
            ' calculate percent change and print
            percent_change = (qtly_change / open_price)
            ws.Range("K" & Summary_Table_Row).Value = percent_change
            ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

            Summary_Table_Row = Summary_Table_Row + 1
            
            open_price = ws.Cells(i + 1, 3).Value
            
            ' reset volume and opening price
            volume = 0

            Else
                volume = volume + ws.Cells(i, 7).Value
                

            End If
            
        Next i
        
        qtly_change_lastrow = ws.Cells(Rows.Count, "J").End(xlUp).Row
        
        ' format cells
        For j = 2 To qtly_change_lastrow
            If ws.Cells(j, 10).Value >= 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 10
            
            ElseIf ws.Cells(j, 10).Value < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3
            
            End If
            
        Next j
    
    ' add headers for greatest increase, decrease, and volume
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
            
    ' Find last row of summary table
    Dim lastrow2 As Long
    lastrow2 = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
    
    ' begin loop
    For k = 2 To lastrow2
    ' find greatest percent increase
        If ws.Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow2)) Then
            ws.Cells(2, 16).Value = ws.Cells(k, 9).Value
            ws.Cells(2, 17).Value = ws.Cells(k, 11).Value
            ws.Cells(2, 17).NumberFormat = "0.00%"
        ' find greatest percent decrease
        ElseIf ws.Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow2)) Then
            ws.Cells(3, 16).Value = ws.Cells(k, 9).Value
            ws.Cells(3, 17).Value = ws.Cells(k, 11).Value
            ws.Cells(3, 17).NumberFormat = "0.00%"
        ' find greatest total volume
        ElseIf ws.Cells(k, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow2)) Then
            ws.Cells(4, 16).Value = ws.Cells(k, 9).Value
            ws.Cells(4, 17).Value = ws.Cells(k, 12).Value
        End If
    Next k
    
    
    Next ws
    
    
End Sub
