Attribute VB_Name = "Module1"
Sub TotalStockVolume():

'Set the Initial Variables
Dim i, j As Integer
Dim Total As Double
Dim Ticker As String
Dim Lastrow As Long

'Loop Through All Sheets
For Each ws In Worksheets

'Determine the last Row
 Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Add the Ticker and Total Stock Volume to New Columns
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Total Stock Volume"

'Set an Initial Total
 Total = 0
 j = 2

'Loop Through Each Year 's Stock Data and Grab the Total Amount
 For i = 2 To Lastrow
     If ws.Range("A" & i + 1).Value = ws.Range("A" & i).Value Then
         Total = Total + ws.Range("G" & i).Value

     Else
         Ticker = ws.Range("A" & i).Value
         ws.Range("I" & j).Value = Ticker
         ws.Range("J" & j).Value = Total + Range("G" & i).Value
         
         'Add a New Row Next Ticker
         'Reset Total Volume
         j = j + 1
         Total = 0
     End If

   Next i
  Next ws
  
  
End Sub

