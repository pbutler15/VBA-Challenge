Sub MultiYearStock()

'Loop through all columns on each worksheet
For Each ws In Worksheets
    ws.Cells(1, 9).Value = "Ticker Symbol"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
'Set all initial variables to find yearly change, percent change and total stock volume per ticker.
    Dim i As Long
    Dim open_price As Double
    Dim tickerSymbol As String
    Dim percent_change As Double
    percent_change = 0
    Dim yearly_change As Double
    yearly_change = 0
    Dim total_volume As Double
    total_volume = 0
    Dim tickerRow As Long
    tickerRow = 2
    Dim lastRow As Long
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
'Loop to grab ticker and open price information across all worksheets
    For i = 2 To lastRow
    open_price = ws.Cells(tickerRow, 3).Value
    
'If the value of the next cell is different than that of the current cell
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        tickerSymbol = ws.Cells(i, 1).Value
        ws.Range("I" & tickerRow).Value = tickerSymbol
        
        yearly_change = yearly_change + (ws.Cells(i, 6).Value - open_price)
        ws.Range("J" & tickerRow).Value = yearly_change
        
        percent_change = (yearly_change / open_price)
        ws.Range("K" & tickerRow).Value = percent_change
        ws.Range("K" & tickerRow).Style = "Percent"
        
        total_volume = total_volume + ws.Cells(i, 7).Value
        ws.Range("L" & tickerRow).Value = total_volume
        
'Add to the tickerRow counter
        tickerRow = tickerRow + 1
        yearly_change = 0
        total_volume = 0
        open_price = ws.Cells(tickerRow, 3).Value
        
    Else
        total_volume = total_volume + ws.Cells(i, 7).Value
        
    End If
    
Next i


'Create cell conditional formatting. Positive is green, negative is red.
'Set variables for cell format

Dim yearLastRow As Long
yearLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row

'Set loop for cell conditional formatting
    For i = 2 To yearLastRow
        If ws.Cells(i, 10).Value >= 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    
    Else
    
        ws.Cells(i, 10).Interior.ColorIndex = 3
        
    End If
    
Next i

'Set variables to find the maximum and the minimum, start from last row
Dim percentLastRow As Long
percentLastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row

Dim percentMax As Double
percentMax = 0
Dim percentMin As Double
percentMin = 0

    For i = 2 To percentLastRow
     If percentMax < ws.Cells(i, 11).Value Then
     percentMax = ws.Cells(i, 11).Value
     ws.Cells(2, 17).Value = percentMax
     ws.Cells(2, 17).Style = "Percent"
     ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
     
    ElseIf percentMin > ws.Cells(i, 11).Value Then
     percentMin = ws.Cells(i, 11).Value
     ws.Cells(3, 17).Value = percentMin
     ws.Cells(3, 17).Style = "Percent"
     ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
    End If
    
Next i

'Loop through worksheets to find greatest percent increase, decrease and total volume values.
ws.Cells(1, 16).Value = "Ticker Symbol"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

Dim totalVolumeRow As Long
totalVolumeRow = ws.Cells(Rows.Count, 12).End(xlUp).Row
Dim totalVolumeMax As Double
totalVolumeMax = 0

For i = 2 To totalVolumeRow

    If totalVolumeMax < ws.Cells(i, 12).Value Then
    totalVolumeMax = ws.Cells(i, 12).Value
    
    ws.Cells(4, 17).Value = totalVolumeMax
    ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
    
    End If
    
    Next i
    
    Next ws

End Sub

