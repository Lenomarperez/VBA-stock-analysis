# VBA-stock-analysis
Sub stock_analysis()

Dim total As Double
Dim rowindex As Long
Dim change As Double
Dim columnIndex As Integer
Dim start As Long
Dim rowCount As Long
Dim percentchange As Double
Dim days As Integer
Dim dailychange As Single
Dim averageChange As Double
Dim ws As Worksheet

For Each ws In Worksheets

columnIndex = 0
total = 0
change = 0
start = 2
dailychange = 0



ws.Range("i1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"


rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
MsgBox (rowCount)

For rowlndex = 2 To rowCount

'if ticker changes then print results

If ws.Cells(rowlndex + 1, 1).Value <> ws.Cells(rowindex, 1).Value Then
'Store results in variable
total = total + ws.Cells(rowlndex, 7).Value

If total = 0 Then
'print the results
ws.Range("i" & 2 + columnIndex).Value = Cells(rowindex, 1).Value
ws.Range("j" & 2 + columnIndex).Value = 0
ws.Range("K" & 2 + columnIndex).Value = "%" & 0
ws.Range("L" & 2 + columnIndex).Value = 0

Else

If ws.Cells(start, 3) = 0 Then
For find_value = start To rowindex
If ws.Cells(find_value, 3).Value <> 0 Then
   start = find_value
   Exit For
End If
Next find_value
End If
change = (ws.Cells(rowlndex, 6) - ws.Cells(start, 3))
percentchange = change / ws.Cells(start, 3)

start = rowlndex + 1

ws.Range("i" & 2 + columnIndex) = ws.Cells(rowlndex, 1).Value
ws.Range("j" & 2 + columnIndex) = change
ws.Range("J" & 2 + columnIndex).NumberFormat = "0.00"
ws.Range("K" & 2 + columnIndex).Value = percentchange
ws.Range("K" & 2 + columnIndex).NumberFormat = "0.00%"
ws.Range("L" & 2 + columnIndex).Value = total
Select Case change
Case Is > 0
ws.Range("J" & 2 + columnIndex).Interior.ColorIndex = 4
Case Is < 0
ws.Range("J" & 2 + columnIndex).Interior.ColorIndex = 3
Case Else
ws.Range("J" & 2 + columnIndex).Interior.ColorIndex = 0
End Select

End If

total = 0
change = 0
columnIndex = columnIndex + 1
days = 0
dailychange = 0
  
  
  Else
'if ticker is still the same add results
total = total + ws.Cells(rowindex, 7).Value

  End If
  
Next rowlndex

'take the max and min and place them in a separate part in the worksheet
ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))

increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
volume_Number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L:L" & rowCount), 0)

ws.Range("P2") = ws.Cells(increase_number + 1, 9)
ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
ws.Range("P4") = ws.Cells(volume_Number + 1, 9)

Next ws

End Sub





