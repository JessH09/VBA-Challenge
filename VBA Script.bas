Attribute VB_Name = "Module1"
Sub module2challenge()

'define variables

Dim tickername As String
Dim yearlychange As Double
Dim percentagechange As Double
Dim totalstockvol As Long
Dim greatestincrease As Double
Dim greatestdecrease As Double
Dim summary_row As Double
Dim tickervolume As Double
Dim close_price As Double
Dim open_price As Double
Dim lastrow As Double
Dim worksheetname As String

For Each ws In Worksheets

worksheetname = ws.Name


'create column for ticker
ws.Cells(1, 9).Value = "Ticker"

'create column for total stock volume
ws.Cells(1, 12).Value = "Total Stock Volume"

'create column for yearly change
ws.Cells(1, 10).Value = "Yearly Change"

'create column for percent change
ws.Cells(1, 11).Value = "Percent Change"


'use last row formula to work with entire worksheet

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'specify that summary starts on row 2
stockvolume = 0
summary_row = 2

'loop through first column to get the different tickers
For i = 2 To lastrow

 If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    tickername = ws.Cells(i, 1).Value
    stockvolume = stockvolume + ws.Cells(i, 7).Value

'output the ticker and stock volume
ws.Range("I" & summary_row).Value = tickername
ws.Range("L" & summary_row).Value = stockvolume

'reset stock volume
stockvolume = 0


'specify what the close and open price are
close_price = ws.Cells(i, 6).Value

'specify how to calculate the yearly change
If open_price = 0 Then
    yearly_change = 0
    percent_change = 0

    Else:

    yearly_change = (close_price - open_price)
    percent_change = (close_price - open_price) / open_price
    End If


ws.Range("J" & summary_row).Value = yearly_change
ws.Range("K" & summary_row).Value = percent_change
ws.Range("K" & summary_row).NumberFormat = "0.00%"

'at the end of all calculations specify that the summary moves down one row after each output
summary_row = summary_row + 1

ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
    open_price = ws.Cells(i, 3)
    
Else: stockvolume = stockvolume + ws.Cells(i, 7).Value

    
End If
Next i

'create column for "ticker" in increase/decrease
ws.Cells(1, 15).Value = "Ticker"

'create value column for increase/decrease
ws.Cells(1, 16).Value = "Value"

'add column for greatest increase

lastsumrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To lastsumrow

ws.Cells(2, 14).Value = "Greatest % Increase"

If ws.Cells(i, 11).Value = WorksheetFunction.Max(Range("K2:K" & lastsumrow)) Then
    ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
    ws.Cells(2, 16).Value = ws.Cells(i, 11).Value
    ws.Cells(2, 16).NumberFormat = "0.00%"
    
End If

'add column for greatest decrease
ws.Cells(3, 14).Value = "Greatest % Decrease"
If ws.Cells(i, 11).Value = WorksheetFunction.Min(Range("K2:K" & lastsumrow)) Then
    ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
    ws.Cells(3, 16).Value = ws.Cells(i, 11).Value
    ws.Cells(3, 16).NumberFormat = ("0.00%")

End If

 
'add column for greatest total volume
ws.Cells(4, 14).Value = "Greatest Total Volume"
 If ws.Cells(i, 12).Value = WorksheetFunction.Max(Range("L2:L" & lastsumrow)) Then
    ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
    ws.Cells(4, 16).Value = ws.Cells(i, 12).Value
    

End If

Next i


'loop through summary data

For i = 2 To lastsumrow

'create conditional that says negatives are red and positives are green
    If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
    
    Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
        
    End If
    
Next i
Next ws

End Sub
