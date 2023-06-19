Attribute VB_Name = "Module1"
Sub WorksheetLoop()
        
    Dim WS_Count As Integer
    Dim I As Integer
    WS_Count = ActiveWorkbook.Worksheets.Count

    ' Begin the loop.
    For I = 1 To WS_Count
        Dim ws As Worksheet
        Set ws = ActiveWorkbook.Worksheets(I)
        
        ' Name the Headers for the outputs
        ws.Cells(1, "I").Value = "Ticker"
        ws.Cells(1, "J").Value = "Yearly Change"
        ws.Cells(1, "K").Value = "Percent Change"
        ws.Cells(1, "L").Value = "Total Stock Volume"
            
        ' To loop through all rows to the last row
        Dim x As Long
        Dim lastrow As Long
        Dim ticker_name As String
        Dim open_price As Double
        Dim close_price As Double
        Dim stock_volume As Double
        Dim Summary_Table_Row As Integer
        
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        open_price = ws.Cells(2, 3).Value
        Summary_Table_Row = 2
        stock_volume = 0
        
        ' To return outputs
        For x = 2 To lastrow
            If ws.Cells(x + 1, 1).Value = ws.Cells(x, 1).Value Then
                stock_volume = stock_volume + ws.Cells(x, 7).Value
            ElseIf ws.Cells(x + 1, 1).Value <> ws.Cells(x, 1).Value Then
                ticker_name = ws.Cells(x, 1).Value
                close_price = ws.Cells(x, 6).Value
                stock_volume = stock_volume + ws.Cells(x, 7).Value
                ws.Range("I" & Summary_Table_Row).Value = ticker_name
                ws.Range("J" & Summary_Table_Row).Value = close_price - open_price
                ws.Range("K" & Summary_Table_Row).Value = ws.Range("J" & Summary_Table_Row).Value / open_price
                ws.Range("L" & Summary_Table_Row).Value = stock_volume
                
                Summary_Table_Row = Summary_Table_Row + 1
                open_price = ws.Cells(x + 1, 3).Value
                stock_volume = 0
            End If
        Next x
        
        Dim z As Double
        For z = 2 To lastrow
        If ws.Cells(z, "J").Value > 0 Then
        ws.Cells(z, "J").Interior.ColorIndex = 4
        ElseIf ws.Cells(z, "J").Value < 0 Then
        ws.Cells(z, "J").Interior.ColorIndex = 3
        End If
        Next z
        
        'Format Column K as percentage
        ws.Range("K:K").NumberFormat = "0.00%"
        
        'To return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume".
        ws.Cells(1, "P").Value = "ticker"
        ws.Cells(1, "Q").Value = "Value"
        ws.Cells(2, "O").Value = "Greatest % Increase"
        ws.Cells(3, "O").Value = "Greatest % Decrease"
        ws.Cells(4, "O").Value = "Greatest Total Volume"
        ws.Cells(2, "Q").Value = WorksheetFunction.Max(ws.Range("K:K"))
        ws.Cells(3, "Q").Value = WorksheetFunction.Min(ws.Range("K:K"))
        ws.Cells(4, "Q").Value = WorksheetFunction.Max(ws.Range("L:L"))
        
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
        
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Range("O1:Q4").EntireColumn.AutoFit
    Next I
End Sub


