Attribute VB_Name = "Module1"
Sub VBAofWallStreet()

For Each ws In Worksheets
Dim Ticker As String

Dim TotalStockVolume As Double
    TotalStockVolume = 0

'Tracking each ticker in the summary ticker table
Dim TickerTable As Integer
    TickerTable = 2

'Variables
Dim YrOpen As Double
Dim YrClose As Double
Dim YrChange As Double
Dim PctYrChange As Double
Dim OpenPriceRow As Long
OpenPriceRow = 2

'Set Greatest Increase/Decrease/Total Volume Variables
Dim GreatestInc As Double
    GreatestInc = 0
Dim GreatestDec As Double
    GreatestDec = 0
Dim GreatestTtotVol As Double
    GreatestTotVol = 0

'Set Ticker Table Headings including Greatest Inc/Dec/Total Volume
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

'Last Row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loop through all the ticker volumes
For i = 2 To LastRow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      'Ticker name
      Ticker = ws.Cells(i, 1).Value

      'Year Open Price
      YrOpen = ws.Range("C" & OpenPriceRow).Value

      'Year Close Price
      YrClose = ws.Cells(i, 6).Value

      'Year change calc and put into the ticker table
      YrChange = YrClose - YrOpen
            ws.Range("J" & TickerTable).Value = YrChange

      'Formatting
      ws.Range("J" & TickerTable).NumberFormat = "0.00"
      ws.Range("K" & TickerTable).NumberFormat = "0.00%"

      'Percent Year Change calc and put into the ticker table
      If YrOpen = 0 Then
        PctYrChange = 0
        ws.Range("K" & TickerTable).Value = PctYrChange
      Else
        PctYrChange = (YrClose - YrOpen) / YrOpen
        ws.Range("K" & TickerTable).Value = PctYrChange
      End If

      ' Stock Volume & Total Summary Table
      TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        ws.Range("I" & TickerTable).Value = Ticker
        ws.Range("L" & TickerTable).Value = TotalStockVolume

      'Conditonal formatting
      If ws.Range("J" & TickerTable).Value >= 0 Then
        ws.Range("J" & TickerTable).Interior.ColorIndex = 4
      Else
        ws.Range("J" & TickerTable).Interior.ColorIndex = 3
      End If

      ' Add one to the ticker table table
      TickerTable = TickerTable + 1
        TotalStockVolume = 0

      'Remaning Price
      OpenPriceRow = i + 1

    Else

      ' Add to the Total Volume
      TotalStockVolume = TotalStockVolume + Cells(i, 7).Value

    End If

  Next i

'Find greatest increase, decrease, and total volume & place in the Summary Ticker Table
  SummaryLastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

  For j = 2 To SummaryLastRow

    If ws.Range("K" & j).Value > ws.Range("Q2").Value Then
        ws.Range("Q2").Value = ws.Range("K" & j).Value
        ws.Range("P2").Value = ws.Range("I" & j).Value
    End If

    If ws.Range("K" & j).Value < ws.Range("Q3").Value Then
        ws.Range("Q3").Value = ws.Range("K" & j).Value
        ws.Range("P3").Value = ws.Range("I" & j).Value
    End If

    If ws.Range("L" & j).Value > ws.Range("Q4").Value Then
        ws.Range("Q4").Value = ws.Range("L" & j).Value
        ws.Range("P4").Value = ws.Range("I" & j).Value
    End If

  Next j

'Formatting
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"

Next ws

End Sub
