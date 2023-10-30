Attribute VB_Name = "Module1"
Sub StockAnalysis()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim MaxIncrease As Double
    Dim MaxDecrease As Double
    Dim MaxVolTicker As String
    Dim MaxIncTicker As String
    Dim MaxDecTicker As String
    Dim MaxIncValue As Double
    Dim MaxDecValue As Double
    Dim MaxVolValue As Double

    MaxIncrease = 0
    MaxDecrease = 0
    MaxVolValue = -1
    ' starting with a small value to capture the maximum

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Summary" Then
            ' start variables for each worksheet
            YearlyChange = 0
            TotalVolume = 0

            
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

            ' Set headers above columns I, J, K, L, and M
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"

            ' Loop through the data
            For i = 2 To LastRow
                If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                    Ticker = ws.Cells(i, 1).Value
                    OpenPrice = ws.Cells(i, 3).Value
                End If

                ClosePrice = ws.Cells(i, 6).Value
                YearlyChange = ClosePrice - OpenPrice
                If OpenPrice <> 0 Then
                    PercentChange = (YearlyChange / OpenPrice) * 100
                Else
                    PercentChange = 0
                End If

                TotalVolume = TotalVolume + ws.Cells(i, 7).Value

                ws.Cells(i, 9).Value = Ticker
                ws.Cells(i, 10).Value = YearlyChange
                ws.Cells(i, 11).Value = PercentChange
                ws.Cells(i, 12).Value = TotalVolume

                If YearlyChange > MaxIncrease Then
                    MaxIncrease = YearlyChange
                    MaxIncTicker = Ticker
                    MaxIncValue = PercentChange
                End If

                If YearlyChange < MaxDecrease Then
                    MaxDecrease = YearlyChange
                    MaxDecTicker = Ticker
                    MaxDecValue = PercentChange
                End If

                If TotalVolume > MaxVolValue Then
                    MaxVolValue = TotalVolume
                    MaxVolTicker = Ticker
                End If
            Next i

            ' Set the "Ticker" header and "Value" header in columns P and Q
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"

            ' Set the "Greatest % Increase" and "Greatest % Decrease" in column P
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(2, 16).Value = MaxIncTicker

            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(3, 16).Value = MaxDecTicker

            ws.Cells(4, 15).Value = "Greatest Total Volume"
            ws.Cells(4, 16).Value = MaxVolTicker

            ' Display the values for "Greatest % Increase" and "Greatest % Decrease" in column Q
            ws.Cells(2, 17).Value = MaxIncValue
            ws.Cells(3, 17).Value = MaxDecValue
            ws.Cells(4, 17).Value = MaxVolValue

            ' Color change for negatives and positives. Red=negative, Green=positive in column J
            For Each cell In ws.Range("J2:J" & LastRow)
                If cell.Value > 0 Then
                    cell.Interior.Color = RGB(0, 255, 0)
                ElseIf cell.Value < 0 Then
                    cell.Interior.Color = RGB(255, 0, 0)
                End If
            Next cell
        End If
    Next ws
End Sub

