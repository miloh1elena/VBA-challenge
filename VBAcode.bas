Attribute VB_Name = "Module1"
Sub StockAnalysis()

    ' Loop through all worksheets (years)
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        
        ' Set up variables
        Dim Ticker As String
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalVolume As Double
        Dim LastRow As Long
        Dim SummaryRow As Long
        
        ' Initialize summary table headers
        SummaryRow = 2
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
        
        ' Find the last row in the worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Initialize variables for summary calculations
        TotalVolume = 0
        OpenPrice = ws.Cells(2, 3).Value
        
        ' Loop through rows in the worksheet
        For i = 2 To LastRow
            ' Check if the current row's ticker is different from the previous row
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set ticker symbol
                Ticker = ws.Cells(i, 1).Value
                ' Set close price
                ClosePrice = ws.Cells(i, 6).Value
                ' Calculate yearly change and percent change
                YearlyChange = Round(ClosePrice - OpenPrice, 2) ' Round to 2 decimal places
                If OpenPrice <> 0 Then
                    PercentChange = Round((YearlyChange / OpenPrice) * 100, 2) ' Round to 2 decimal places
                Else
                    PercentChange = 0
                End If
                ' Add to total volume
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                
                ' Populate summary table
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = YearlyChange
                ws.Cells(SummaryRow, 11).Value = PercentChange
                ws.Cells(SummaryRow, 12).Value = TotalVolume
                
                ' Reset variables for next ticker
                SummaryRow = SummaryRow + 1
                TotalVolume = 0
                OpenPrice = ws.Cells(i + 1, 3).Value
            Else
                ' Add to total volume
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Find the last row in the summary table
        LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Apply conditional formatting for positive and negative changes
        For i = 2 To LastRow
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4 ' Green
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3 ' Red
            End If
        Next i
        
        ' Find greatest % increase, % decrease, and total volume
        Dim MaxPercentIncrease As Double
        Dim MaxPercentDecrease As Double
        Dim MaxTotalVolume As Double
        Dim MaxPercentIncreaseTicker As String
        Dim MaxPercentDecreaseTicker As String
        Dim MaxTotalVolumeTicker As String
        
        MaxPercentIncrease = WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
        MaxPercentDecrease = WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
        MaxTotalVolume = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
        
        MaxPercentIncreaseTicker = ws.Cells(Application.Match(MaxPercentIncrease, ws.Range("K2:K" & LastRow), 0) + 1, 9).Value
        MaxPercentDecreaseTicker = ws.Cells(Application.Match(MaxPercentDecrease, ws.Range("K2:K" & LastRow), 0) + 1, 9).Value
        MaxTotalVolumeTicker = ws.Cells(Application.Match(MaxTotalVolume, ws.Range("L2:L" & LastRow), 0) + 1, 9).Value
        
        ' Populate greatest % increase, % decrease, and total volume
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ws.Cells(2, 16).Value = MaxPercentIncreaseTicker
        ws.Cells(3, 16).Value = MaxPercentDecreaseTicker
        ws.Cells(4, 16).Value = MaxTotalVolumeTicker
        
        ws.Cells(2, 17).Value = MaxPercentIncrease
        ws.Cells(3, 17).Value = MaxPercentDecrease
        ws.Cells(4, 17).Value = MaxTotalVolume
    Next ws

End Sub

