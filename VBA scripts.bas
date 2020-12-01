Attribute VB_Name = "Module1"
Sub stock()

Dim ws As Worksheet

For Each ws In Worksheets

ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent change"
ws.Range("L1") = "Total Stock Volume"

ws.Range("O2") = "Greatest% increase"
ws.Range("O3") = "Greatest% decrease"
ws.Range("O4") = "Greatest Total Volume"
ws.Range("P1") = "Ticker"
ws.Range("Q1") = "Value"

Dim increaseticker As String
Dim decreaseticker As String
Dim tickertotalgre As String
Dim valueincrease As Double
Dim valuedecrease As Double
Dim valuetotalgre As Double
Dim lastkrow As Double: lastkrow = ws.Cells(Rows.Count, 11).End(xlUp).Row
Dim lastlrow As Double: lastkrow = ws.Cells(Rows.Count, 12).End(xlUp).Row


Dim tickers As String
Dim total As Double: total = 0
Dim lastrow As Double: lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
Dim summaryrow As Double: summaryrow = 2
Dim yearchange, yearopen, yearclose As Double

yearopen = 0
yearclose = 0

For x = 2 To lastrow

If ws.Cells(x + 1, 1) <> ws.Cells(x, 1) Then
tickers = ws.Cells(x, 1)
total = total + ws.Cells(x, 7)

yearclose = Cells(x, 6)
yearchange = yearclose - yearopen


If yearopen = 0 Then
percentchange = 0

Else
percentchange = yearchange / yearopen

End If

ws.Range("I" & summaryrow) = tickers
ws.Range("L" & summaryrow) = total
ws.Range("J" & summaryrow) = yearchange
ws.Range("K" & summaryrow) = Format(percentchange, "percent")


If ws.Range("J" & summaryrow) < 0 Then
ws.Range("J" & summaryrow).Interior.ColorIndex = 3

ElseIf ws.Range("J" & summaryrow) > 0 Then
ws.Range("J" & summaryrow).Interior.ColorIndex = 4


End If

summaryrow = summaryrow + 1
total = 0
yearopen = 0
yearclose = 0

Else

If yearopen = 0 Then
yearopen = Cells(x, 3)

End If

total = total + ws.Cells(x, 7)

End If


Next x



Next ws
End Sub

