# VBA-challenge
Sub StockAnalysis()

    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
    
        ' Set initial variables
        Dim ticker As String
        Dim openingPrice As Double
        Dim closingPrice As Double
        Dim yearlyChange As Double
        Dim percentChange As Double
        Dim totalVolume As Double
        Dim summaryTableRowCount As Integer
        
        ' Initialize summary table row count
        summaryTableRowCount = 2
        
        ' Find the last row of data
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Set initial opening price
        openingPrice = ws.Cells(2, 3).Value
        
        ' Loop through all rows of data
        For i = 2 To lastRow
        
            ' Check if the current ticker is the same as the previous ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ' Set the current ticker
                ticker = ws.Cells(i, 1).Value
                
                ' Set the closing price
                closingPrice = ws.Cells(i, 6).Value
                
                ' Calculate yearly change
                yearlyChange = closingPrice - openingPrice
                
                ' Calculate percent change
                If openingPrice <> 0 Then
                    percentChange = (yearlyChange / openingPrice) * 100
                Else
                    percentChange = 0
                End If
                
                ' Calculate total stock volume
                totalVolume = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(summaryTableRowCount, 7), ws.Cells(i, 7)))
                
                ' Print the analysis results in the summary table
                ws.Cells(summaryTableRowCount, 9).Value = ticker
                ws.Cells(summaryTableRowCount, 10).Value = yearlyChange
                ws.Cells(summaryTableRowCount, 11).Value = percentChange / 100 ' Adjusted line
                ws.Cells(summaryTableRowCount, 12).Value = totalVolume
                
                ' Format percent change as percentage
                ws.Cells(summaryTableRowCount, 11).NumberFormat = "0.00%"
                
                ' Conditional formatting for yearly change
                If yearlyChange > 0 Then
                    ws.Cells(summaryTableRowCount, 10).Interior.Color = RGB(0, 255, 0)
                ElseIf yearlyChange < 0 Then
                    ws.Cells(summaryTableRowCount, 10).Interior.Color = RGB(255, 0, 0)
                Else
                    ws.Cells(summaryTableRowCount, 10).Interior.Color = RGB(255, 255, 255)
                End If
                
                ' Reset the opening price for the next ticker
                openingPrice = ws.Cells(i + 1, 3).Value
                
                ' Increment the summary table row count
                summaryTableRowCount = summaryTableRowCount + 1
                
            End If
            
        Next i
        
        ' Find the last row of the summary table
        lastSummaryRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Find the greatest percent increase, greatest percent decrease, and greatest total volume
        Dim maxPercentIncrease As Double
        Dim maxPercentDecrease As Double
        Dim maxTotalVolume As Double
        Dim maxPercentIncreaseTicker As String
        Dim maxPercentDecreaseTicker As String
        Dim maxTotalVolumeTicker As String
        
        maxPercentIncrease = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastSummaryRow).Value)
        maxPercentDecrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastSummaryRow).Value)
        maxTotalVolume = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastSummaryRow).Value)
        
        maxPercentIncreaseTicker = ws.Cells(Application.WorksheetFunction.Match(maxPercentIncrease, ws.Range("K2:K" & lastSummaryRow), 0) + 1, 9).Value
        maxPercentDecreaseTicker = ws.Cells(Application.WorksheetFunction.Match(maxPercentDecrease, ws.Range("K2:K" & lastSummaryRow), 0) + 1, 9).Value
        maxTotalVolumeTicker = ws.Cells(Application.WorksheetFunction.Match(maxTotalVolume, ws.Range("L2:L" & lastSummaryRow), 0) + 1, 9).Value
        
        ' Print the greatest percent increase, greatest percent decrease, and greatest total volume in the summary table
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(2, 17).Value = maxPercentIncreaseTicker
        ws.Cells(2, 18).Value = maxPercentIncrease
        ws.Cells(2, 18).NumberFormat = "0.00%"
        
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(3, 17).Value = maxPercentDecreaseTicker
        ws.Cells(3, 18).Value = maxPercentDecrease
        ws.Cells(3, 18).NumberFormat = "0.00%"
        
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        ws.Cells(4, 17).Value = maxTotalVolumeTicker
        ws.Cells(4, 18).Value = maxTotalVolume
        
        ' Auto-fit columns for better visibility
        ws.Columns.AutoFit
        
    Next ws

End Sub
