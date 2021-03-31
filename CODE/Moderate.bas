Attribute VB_Name = "Module2"

Sub TotalStockVolume():

 'Loop Through All Sheets
 For Each ws In Worksheets

 'Set the Initial Variables
 Dim i, j As Long
 Dim Ticker As String
 Dim YearlyChange As Double
 Dim PercentChange As Double
 Dim TotalVolume As Double
 Dim OpenPrice As Double
 Dim ClosePrice As Double
 Dim OpenPrice_Row As Long
 Dim Lastrow As Long
 
 'Add Header Name
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Yearly Change"
 ws.Range("K1").Value = "Percent Change"
 ws.Range("L1").Value = "Total Stock Volume"

 'Set Initial Total
TotalVolume = 0
 j = 2
 OpenPrice_Row = 2

 'Determine the last Row
 Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

 'Loop Through Each Year on Stock Data
 For i = 2 To Lastrow
     
 'Compare Tickers
 'Add to the Total Volume If Tickers are same
If ws.Range("A" & i + 1).Value = ws.Range("A" & i).Value Then
TotalVolume = TotalVolume + ws.Range("G" & i).Value

Else

'Grabbed Ticker when it changed
 Ticker = ws.Range("A" & i).Value

 'Calculate Yearly Change and Percent Change
OpenPrice = ws.Range("C" & OpenPrice_Row)
ClosePrice = ws.Range("F" & i)
YearlyChange = ClosePrice - OpenPrice

 'Calculate Percent Change
 If OpenPrice = 0 Then
 PercentChange = 0

Else

PercentChange = YearlyChange / OpenPrice
        
End If

'Grabbed Ticker,Total Volume,Yearly Change and Percent Change
ws.Range("I" & j).Value = Ticker
ws.Range("J" & j).Value = YearlyChange
ws.Range("K" & j).Value = PercentChange
ws.Range("L" & j).Value = TotalVolume + ws.Range("G" & i).Value
ws.Range("K" & j).NumberFormat = "0.00%"
         
 'Conditional Formating that highlight Positive change in Green and Negative change in Red
If ws.Range("J" & j).Value > 0 Then
ws.Range("J" & j).Interior.ColorIndex = 4
 
 Else

ws.Range("J" & j).Interior.ColorIndex = 3

End If

'Add a New Row  for Next Ticker
'Set New OpenPrice row and Reset Total Volume
j = j + 1
TotalVolume = 0
OpenPrice_Row = i + 1
         
End If
     
 Next i
 Next ws
 
End Sub

