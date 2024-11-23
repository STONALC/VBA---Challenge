Below is the code used to solve the VBA challenge:

Sub StockAnalysis()
    Dim ws As Worksheet
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim lastRow As Long
    Dim i As Long
    Dim outputRow As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
 
    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Initialize variables
        outputRow = 2
        totalVolume = 0
        greatestIncrease = 0
        greatestDecrease = 0
        greatestVolume = 0
 
        ' Add headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
 
        ' Find the last row with data
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
 
        ' Set initial open price
        openPrice = ws.Cells(2, 3).Value
 
        ' Loop through all rows
        For i = 2 To lastRow
            ' Check if we are at the beginning of a new ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                ' Set the ticker
                ticker = ws.Cells(i, 1).Value
 
                ' Set closing price
                closePrice = ws.Cells(i, 6).Value
 
                ' Calculate yearly change
                yearlyChange = closePrice - openPrice
 
                ' Calculate percent change
                If openPrice <> 0 Then
                    percentChange = (yearlyChange / openPrice)
                Else
                    percentChange = 0
                End If
 
                ' Add final volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
 
                ' Output the results
                ws.Cells(outputRow, 9).Value = ticker
                ws.Cells(outputRow, 10).Value = yearlyChange
                ws.Cells(outputRow, 11).Value = percentChange
                ws.Cells(outputRow, 12).Value = totalVolume
 
                ' Format percent change as percentage
                ws.Cells(outputRow, 11).NumberFormat = "0.00%"
 
                ' Color formatting for yearly change
                If yearlyChange < 0 Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 3 ' Red
                ElseIf yearlyChange > 0 Then
                    ws.Cells(outputRow, 10).Interior.ColorIndex = 4 ' Green
                End If
 
                ' Check for greatest increase, decrease, and volume
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    greatestIncreaseTicker = ticker
                ElseIf percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    greatestDecreaseTicker = ticker
                End If
 
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestVolumeTicker = ticker
                End If
 
                ' Reset variables for next ticker
                openPrice = ws.Cells(i + 1, 3).Value
                totalVolume = 0
                outputRow = outputRow + 1
            Else
                ' Add to total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            End If
        Next i
 
        ' Output Greatest % Increase, % Decrease, and Total Volume
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
 
        ws.Range("P2").Value = greatestIncreaseTicker
        ws.Range("Q2").Value = greatestIncrease
        ws.Range("Q2").NumberFormat = "0.00%"
 
        ws.Range("P3").Value = greatestDecreaseTicker
        ws.Range("Q3").Value = greatestDecrease
        ws.Range("Q3").NumberFormat = "0.00%"
 
        ws.Range("P4").Value = greatestVolumeTicker
        ws.Range("Q4").Value = greatestVolume
 
    Next ws
End Sub

