Sub Stocks()
    Dim ws As Worksheet
    
    Dim currentTicker As String
    Dim iterationTicker As String
    Dim nextTicker As String
    
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    
    Dim i As Long
    Dim lastRow As Long
         
    Dim outputCounter As Long
    outputCounter = 1
    
    Dim greatestPercentIncreaseTicker As String
    Dim greatestPercentIncreaseValue As Double
    Dim greatestPercentDecreaseTicker As String
    Dim greatestPercentDecreaseValue As Double
    Dim greatestTotalVolumeTicker As String
    Dim greatestTotalVolumeValue As Double
    
    For Each ws In ThisWorkbook.Worksheets
        With ws
            ' Setup column labels
            Cells(1, 9).Value = "Ticker"
            Cells(1, 10).Value = "Quarterly Change"
            Cells(1, 11).Value = "Percent Change"
            Cells(1, 12).Value = "Total Volume"
            
            ' Determine the last row of the current worksheet
            lastRow = Cells(ws.Rows.Count, "A").End(xlUp).Row
            
            ' Reset total volume for each worksheet
            totalVolume = 0
            
            ' Loop from second row to the last
            For i = 2 To lastRow
            
                ' Get the ticker for the current row
                iterationTicker = Cells(i, 1).Value
                                
                ' Set the opening price for the current ticker
                If currentTicker <> iterationTicker Then
                    currentTicker = iterationTicker
                    openingPrice = Cells(i, 3).Value
                End If
                
                ' Calculate the running sum of the volume after each iteration for this ticker
                nextTicker = Cells(i + 1, 1).Value
                totalVolume = totalVolume + Cells(i, 7).Value
                
                ' Determine if total volume is the greatest for this quarter
                If totalVolume > greatestTotalVolumeValue Then
                    greatestTotalVolumeValue = totalVolume
                    greatestTotalVolumeTicker = iterationTicker
                End If
                
                ' If we detect that ticker in next row is not equal to the current ticker
                ' Finish calculating the value for current ticker
                If nextTicker <> currentTicker Then
                    
                    ' Tracks where to output the current ticker
                    outputCounter = outputCounter + 1
                                       
                    ' Sets the ticker at the current row defined by the output counter
                    Cells(outputCounter, 9).Value = currentTicker
                    
                    ' Grab the closing price of this ticker
                    closingPrice = Cells(i, 6).Value
                    
                    ' Calculate and set the quarterly change
                    quarterlyChange = closingPrice - openingPrice
                    Cells(outputCounter, 10).Value = quarterlyChange
                    
                    ' Sets the Quarterly Change cell color to red, green or white
                    If quarterlyChange < 0 Then
                        Cells(outputCounter, 10).Interior.ColorIndex = 3
                    ElseIf quarterlyChange > 0 Then
                        Cells(outputCounter, 10).Interior.ColorIndex = 4
                    Else
                        Cells(outputCounter, 10).Interior.ColorIndex = 2
                    End If
                    
                    ' Calculate and set percentage change
                    percentChange = quarterlyChange / openingPrice
                    Cells(outputCounter, 11).Value = Format(percentChange, "0.00%")
                                        
                    ' Keeps track of greatest percent increase across all tickers
                    If percentChange > greatestPercentIncreaseValue Then
                        greatestPercentIncreaseValue = percentChange
                        greatestPercentIncreaseTicker = iterationTicker
                    End If
                    
                    ' Keeps track of greatest percent decrease across all tickers
                    If percentChange < greatestPercentDecreaseValue Then
                        greatestPercentDecreaseValue = percentChange
                        greatestPercentDecreaseTicker = iterationTicker
                    End If
                                       
                    ' Output the total volume for this ticker
                    Cells(outputCounter, 12).Value = totalVolume
                    totalVolume = 0
                End If
                
            Next i
        End With
            
        ' Output greatest percent increase
        Cells(2, 17).Value = Format(greatestPercentIncreaseValue, "0.00%")
        Cells(2, 16).Value = greatestPercentIncreaseTicker
        greatestPercentIncreaseValue = 0
        
        ' Output greatest percent decrease
        Cells(3, 17).Value = Format(greatestPercentDecreaseValue, "0.00%")
        Cells(3, 16).Value = greatestPercentDecreaseTicker
        greatestPercentDecreaseValue = 0
        
        ' Output greatest total volume
        Cells(4, 17).Value = greatestTotalVolumeValue
        Cells(4, 16).Value = greatestTotalVolumeTicker
        greatestTotalVolumeValue = 0
        
    Next ws
    
    ' Set up column and row labels
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
End Sub
