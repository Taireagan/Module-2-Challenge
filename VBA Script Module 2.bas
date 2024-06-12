Attribute VB_Name = "Module1"
Sub QuarterlyPercentStockChange()

'Set/Delcare Variables
    Dim ws As Worksheet
    Dim i As Long
    Dim startRow As Long
    Dim tickerRow As Long
    Dim lastRow As Long
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    
' Going through each worksheet

  For Each ws In Worksheets
    
'Set Header titles
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest Percent Increase"
        ws.Cells(3, 15).Value = "Greatest Percent Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

' To find the last row of data
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        tickerRow = 2
        startRow = 2

' Create starting point and initial valuse for the biggest increase, decreas and volume
        maxIncrease = -1
        maxDecrease = 1
        maxVolume = 0
        
' Create a loop for each row
' Starting the loop from the first row (i = 2 is first row) to the last row
        For i = 2 To lastRow

' Check if value in the next row of column A is different from the value in current row
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

' Calculate the quarterly change in stock price. (closing price - opening price)
                quarterlyChange = ws.Cells(i, 6).Value - ws.Cells(startRow, 3).Value

' Create the calculation for the percent change in stock prices
                If ws.Cells(startRow, 3).Value <> 0 Then
                    percentChange = quarterlyChange / ws.Cells(startRow, 3).Value
                Else
                    percentChange = 0
                End If
                
                
 ' Add headlines and data into the columns
 
                ws.Cells(tickerRow, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(tickerRow, 10).Value = quarterlyChange
                ws.Cells(tickerRow, 11).Value = percentChange

' Calculate total stock column and add it to the stock volume column
                ws.Cells(tickerRow, 12).Value = WorksheetFunction.Sum(ws.Range(ws.Cells(startRow, 7), ws.Cells(i, 7)))
                                
                
' Add color to quarterly change column green if increase and red if decrease
                ws.Cells(tickerRow, 10).Interior.ColorIndex = IIf(quarterlyChange < 0, 3, 4)

'If statement to update the max & min values (Cited:Stack overflow)

' Checks if the percent change for the current ticker is greater than max increase
    ' If true update max increase to current percent change and current ticker name
    
                If percentChange > maxIncrease Then
                    maxIncrease = percentChange
                    maxIncreaseTicker = ws.Cells(i, 1).Value
                End If

' Check if percent change for current ticker is less than the max decrease
    ' If true update max decrease to the currrent percent change and current ticker name
    
                If percentChange < maxDecrease Then
                    maxDecrease = percentChange
                    maxDecreaseTicker = ws.Cells(i, 1).Value
                End If
                
' Check if total volume for the current ticker is greater than the max volumne
    ' If true update max volume to the current total volume and to the current ticker name
    
                If ws.Cells(tickerRow, 12).Value > maxVolume Then
                    maxVolume = ws.Cells(tickerRow, 12).Value
                    maxVolumeTicker = ws.Cells(i, 1).Value
                End If
                
' Move the ticker row down by one to start next data set
                tickerRow = tickerRow + 1
        
' Update Startrow to the next row
                startRow = i + 1
            End If
            
' End loop
        Next i

' Summarizing the results

        ws.Cells(2, 16).Value = maxIncreaseTicker
        ws.Cells(2, 17).Value = Format(maxIncrease, "Percent")
        ws.Cells(3, 16).Value = maxDecreaseTicker
        ws.Cells(3, 17).Value = Format(maxDecrease, "Percent")
        ws.Cells(4, 16).Value = maxVolumeTicker
        ws.Cells(4, 17).Value = Format(maxVolume, "Scientific")
        
        ws.Columns("A:Z").AutoFit
    Next ws

End Sub
