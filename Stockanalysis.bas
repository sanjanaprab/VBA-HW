Attribute VB_Name = "stockanalysis"
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim summaryRow As Integer
    
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        summaryRow = 2
        totalVolume = 0
        openPrice = ws.Cells(2, 3).Value
        
        ' Loop through each row of stock data
        For i = 2 To lastRow
            ' Check if we are still on the same ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ticker = ws.Cells(i, 1).Value
                closePrice = ws.Cells(i, 6).Value
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                ' Calculate quarterly change and percentage change
                quarterlyChange = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = (quarterlyChange / openPrice) * 100
                Else
                    percentChange = 0
                End If
                
                ' Output values
                ws.Cells(summaryRow, 9).Value = ticker
                ws.Cells(summaryRow, 10).Value = totalVolume
                ws.Cells(summaryRow, 11).Value = quarterlyChange
                ws.Cells(summaryRow, 12).Value = percentChange
                
                ' Apply conditional formatting for quarterly change
                If quarterlyChange > 0 Then
                    ws.Cells(summaryRow, 11).Interior.Color = RGB(0, 255, 0) ' Green
                Else
                    ws.Cells(summaryRow, 11).Interior.Color = RGB(255, 0, 0) ' Red
                End If
                
                ' Reset for the next ticker
                summaryRow = summaryRow + 1
                totalVolume = 0
                openPrice = ws.Cells(i + 1, 3).Value
            Else
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i

        ' Additional code for finding greatest increase, decrease, and total volume
        Dim maxIncrease As Double
        Dim maxDecrease As Double
        Dim maxVolume As Double
        Dim maxTickerIncrease As String
        Dim maxTickerDecrease As String
        Dim maxTickerVolume As String

        maxIncrease = 0
        maxDecrease = 0
        maxVolume = 0

        For i = 2 To summaryRow - 1
            ' Check for greatest % increase
            If ws.Cells(i, 12).Value > maxIncrease Then
                maxIncrease = ws.Cells(i, 12).Value
                maxTickerIncrease = ws.Cells(i, 9).Value
            End If
            ' Check for greatest % decrease
            If ws.Cells(i, 12).Value < maxDecrease Then
                maxDecrease = ws.Cells(i, 12).Value
                maxTickerDecrease = ws.Cells(i, 9).Value
            End If
            ' Check for greatest total volume
            If ws.Cells(i, 10).Value > maxVolume Then
                maxVolume = ws.Cells(i, 10).Value
                maxTickerVolume = ws.Cells(i, 9).Value
            End If
        Next i

        ' Output the greatest values
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        ws.Cells(2, 15).Value = maxTickerIncrease
        ws.Cells(2, 16).Value = maxIncrease
        ws.Cells(3, 15).Value = maxTickerDecrease
        ws.Cells(3, 16).Value = maxDecrease
        ws.Cells(4, 15).Value = maxTickerVolume
        ws.Cells(4, 16).Value = maxVolume
    Next ws
End Sub



