VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock_market():

' Declare worksheet
  Dim ws As Worksheet

'Loop through all stocks in one year
  For Each ws In Worksheets

'create the column headings
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("K1").Value = "Percent Change"
  ws.Range("L1").Value = "Total Stock Volume"

  ws.Range("P1").Value = "Ticker"
  ws.Range("Q1").Value = "Value"
  ws.Range("O2").Value = "Greatest % Increase"
  ws.Range("O3").Value = "Greatest % Decrease"
  ws.Range("O4").Value = "Greatest Total Volume"

'Define Ticker variable
  Dim Ticker As String
  Ticker = " "

'Set initial and summary row for worksheet
  Dim lastrow As Long
  Dim i As Long
  Dim j As Integer
  Dim Tickerrow As Long
  Tickerrow = 1
  Dim summary_row As Integer
  summary_row = 2

  Dim close_price As Double
  Dim open_price As Double
  open_price = Cells(2, 3).Value
  Dim ticker_volume As Double
  ticker_volume = 0
  Dim percent_change As Double

'Define last row of worksheet
  lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Looping for current worksheet to last row
  For i = 2 To lastrow

'Ticker symbol output
  If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
    Tickerrow = Tickerrow + 1
    Ticker = ws.Cells(i, 1).Value
    ws.Cells(Tickerrow, "I").Value = Ticker

'Ticker Volume
    ticker_volume = ticker_volume + Cells(i, 7).Value
    ws.Range("L" & summary_row).Value = ticker_volume
        
'Yearly price change
    close_price = ws.Cells(i, 6).Value
    price_change = close_price - open_price
    ws.Range("J" & summary_row).Value = price_change
    
'Percent change
    percent_change = price_change / open_price
    ws.Range("K" & summary_row).Value = percent_change
    'ws.Range("K" & summary_row).NumberFormat = "0.00%"
    
'Greatest % Increase
 If ws.Range("K" & summary_row).Value > ws.Range("Q2").Value Then
    ws.Range("Q2").Value = ws.Range("K" & summary_row).Value
    ws.Range("P2").Value = ws.Range("I" & summary_row).Value
 End If

'Greatest % Decrease
 If ws.Range("K" & summary_row).Value < ws.Range("Q3").Value Then
    ws.Range("Q3").Value = ws.Range("K" & summary_row).Value
    ws.Range("P3").Value = ws.Range("I" & summary_row).Value
 End If

'Greatest total volume
 If ws.Range("L" & summary_row).Value > ws.Range("Q4").Value Then
    ws.Range("Q4").Value = ws.Range("L" & summary_row).Value
    ws.Range("P4").Value = ws.Range("I" & summary_row).Value
 End If


'Conditional Formatting - (Positive - Green, Negative - Red)
   If ws.Range("J" & summary_row).Value > 0 Then
      ws.Range("J" & summary_row).Interior.ColorIndex = 4
   Else
      ws.Range("J" & summary_row).Interior.ColorIndex = 3
   End If
    
'iterate over row
    summary_row = summary_row + 1
'Reset ticker volume
    ticker_volume = 0
'Reset open price
    open_price = ws.Cells(i + 1, 3).Value

    
  Else
     ticker_volume = ticker_volume + Cells(i, 7).Value
     price_change = close_price - open_price

End If
  
Next i

'Format Values to %symbol and two decimal places
   'ws.Range("Q2").NumberFormat = "0.00%"
   'ws.Range("Q3").NumberFormat = "0.00%"
   
'Format table columns to auto fit
   ws.Columns("I:Q").AutoFit
   
Next ws

End Sub

