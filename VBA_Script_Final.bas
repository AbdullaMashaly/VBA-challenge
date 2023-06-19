Attribute VB_Name = "Module1"
Sub WorksheetLoop()
        
    'To loop through worksheets
    
    Dim WS_Count As Integer
    Dim I As Integer
    
    'Set WS_count as the number of sheets
    WS_Count = ActiveWorkbook.Worksheets.Count

    ' Begin the loop.
    For I = 1 To WS_Count
        Dim ws As Worksheet
        
        'Set ws to actice worksheet
        Set ws = ActiveWorkbook.Worksheets(I)
        
        ' Name the Headers for the outputs
        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Yearly Change"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Total Stock Volume"
            
        ' To loop through all rows to the last row
        Dim x As Long
        Dim lastrow As Long
        
        'Set of variables that we want to look for in the data
        Dim ticker_name As String
        Dim open_price As Double
        Dim close_price As Double
        Dim stock_volume As Double
        Dim Summary_Table_Row As Integer
        
        'Set lastrow as the number of rows
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        'Set open price for the first ticker
        open_price = ws.Cells(2, 3).Value
        'Set output row for the first ticker
        Summary_Table_Row = 2
        'Set stock volume to zero
        stock_volume = 0
        
        ' To return outputs
        For x = 2 To lastrow
            If ws.Cells(x + 1, 1).Value = ws.Cells(x, 1).Value Then
                'to addup stock volume for the same ticker
                stock_volume = stock_volume + ws.Cells(x, 7).Value
            ElseIf ws.Cells(x + 1, 1).Value <> ws.Cells(x, 1).Value Then
                'set ticker name
                ticker_name = ws.Cells(x, 1).Value
                'set close price for ticker
                close_price = ws.Cells(x, 6).Value
                'to set total stock volume for this ticker
                stock_volume = stock_volume + ws.Cells(x, 7).Value
                'outputs to our new table
                ws.Range("I" & Summary_Table_Row).Value = ticker_name
                ws.Range("J" & Summary_Table_Row).Value = close_price - open_price
                ws.Range("K" & Summary_Table_Row).Value = ws.Range("J" & Summary_Table_Row).Value / open_price
                ws.Range("L" & Summary_Table_Row).Value = stock_volume
                
                'reset output row
                Summary_Table_Row = Summary_Table_Row + 1
                'reset open price for the next ticker
                open_price = ws.Cells(x + 1, 3).Value
                'reset stock volume for the next ticker
                stock_volume = 0
            End If
        Next x
        
       'Conditional formatting
        Dim z As Double
        For z = 2 To lastrow
        If ws.Cells(z, "J").Value > 0 Then
        'set positive cells interior as green
        ws.Cells(z, "J").Interior.ColorIndex = 4
        ElseIf ws.Cells(z, "J").Value < 0 Then
        'set negative cells interior as red
        ws.Cells(z, "J").Interior.ColorIndex = 3
        End If
        Next z
        
        'Format Column K as percentage
        ws.Range("K:K").NumberFormat = "0.00%"
        
        'to make column L autofit cells content
        ws.Range("L:L").EntireColumn.AutoFit
        
        'To return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
        ws.Cells(1, "P").Value = "ticker"
        ws.Cells(1, "Q").Value = "Value"
        ws.Cells(2, "O").Value = "Greatest % Increase"
        ws.Cells(3, "O").Value = "Greatest % Decrease"
        ws.Cells(4, "O").Value = "Greatest Total Volume"
        'to get maximum value in percent change
        ws.Cells(2, "Q").Value = WorksheetFunction.Max(ws.Range("K:K"))
        'to get minimum value in percent change
        ws.Cells(3, "Q").Value = WorksheetFunction.Min(ws.Range("K:K"))
        'to get maximum stock volume
        ws.Cells(4, "Q").Value = WorksheetFunction.Max(ws.Range("L:L"))
        
        'to match the name of ticker for each value mentioned above
        Dim y As Double
        For y = 2 To lastrow
            If ws.Cells(y, "K").Value = ws.Cells(2, "Q").Value Then
                ws.Cells(2, "P").Value = ws.Cells(y, "I").Value
            ElseIf ws.Cells(y, "K").Value = ws.Cells(3, "Q").Value Then
                ws.Cells(3, "P").Value = ws.Cells(y, "I").Value
            ElseIf ws.Cells(y, "L").Value = ws.Cells(4, "Q").Value Then
                ws.Cells(4, "P").Value = ws.Cells(y, "I").Value
            End If
        Next y
        
        'format cells in the new table
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("O:O").EntireColumn.AutoFit
        ws.Range("Q:Q").EntireColumn.AutoFit
    Next I
End Sub


