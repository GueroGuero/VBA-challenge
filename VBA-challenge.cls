Sub StockData()
'selecting first worksheet
Dim CurrentWs As Worksheet

'table and headers as true and false
Dim Results As Boolean
TableHeader = True

'loop worksheets
For Each CurrentWs In Worksheets

'Variable for tickers
Dim Ticker As String

'Variable for volume
Dim StockVolume As Double
StockVolume = 0

'Variable for opening and closing price
Dim OpenPrice As Double
OpenPrice = 0
Dim ClosePrice As Double
ClosePrice = 0

'Variable for yearly change and percent
Dim YearlyChange As Double
YearlyChange = 0
Dim YearlyPercent As Double
YearlyPercent = 0

'Variable for Greatest values
Dim GreatestIncrease As Double
GreatestIncrease = 0
Dim GreatestDecrease As Double
GreatestDecrease = 0
Dim GreatestVolume As Double
GreatestVolume = 0

'Ticker variables for greatest values
Dim GreatesIncreaseTicker As String
Dim GreatestDecreaseTicker As String
Dim GreatestVolumeTicker As String

'Location for tickers
Dim TableRow As Long
TableRow = 2

'Variable for LastRow
Dim LastRow As Long
Dim i As Long

'LastRow selection
LastRow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

'Headers for table and greatest values
If TableHeader Then
    CurrentWs.Range("I1").Value = "Ticker"
    CurrentWs.Range("J1").Value = "Yearly Change"
    CurrentWs.Range("K1").Value = "Percent Change"
    CurrentWs.Range("L1").Value = "Total Stock Volume"
    CurrentWs.Range("O2").Value = "Greatest % Increase"
    CurrentWs.Range("O3").Value = "Greatest % Decrease"
    CurrentWs.Range("O4").Value = "Greatest Total Volume"
    CurrentWs.Range("P1").Value = "Ticker"
    CurrentWs.Range("Q1").Value = "Value"
Else
    TableHeader = True
End If

'First opening price for the worksheets
OpenPrice = CurrentWs.Cells(2, 3).Value

'Looping through the tickers
For i = 2 To LastRow

'ticker name change
If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
    Ticker = CurrentWs.Cells(i, 1).Value

'Calculation for changes
ClosePrice = CurrentWs.Cells(i, 6).Value
YearlyChange = (ClosePrice - OpenPrice)

    If OpenPrice <> 0 Then
        YearlyPercent = (YearlyChange / OpenPrice) * 100
End If
      
'ticker for StockVolume
StockVolume = StockVolume + CurrentWs.Cells(i, 7).Value

'ticker names, StockVolume, YearlyChange, and YearlyPercent in table
CurrentWs.Range("I" & TableRow).Value = Ticker
CurrentWs.Range("L" & TableRow).Value = StockVolume
CurrentWs.Range("J" & TableRow).Value = YearlyChange
CurrentWs.Range("K" & TableRow).Value = (CStr(YearlyPercent) & "%")

'adding colors
If (YearlyChange > 0) Then
    CurrentWs.Range("J" & TableRow).Interior.ColorIndex = 4
ElseIf (YearlyChange <= 0) Then
    CurrentWs.Range("J" & TableRow).Interior.ColorIndex = 3
End If

'extra one for row count
TableRow = TableRow + 1

'caluclations for greatest values
GreatestIncrease = WorksheetFunction.Max(CurrentWs.Range("K:K"))
GreatestDecrease = WorksheetFunction.Min(CurrentWs.Range("K:K"))
GreatestVolume = WorksheetFunction.Max(CurrentWs.Range("L:L"))

'adding the greatest values to table
CurrentWs.Range("Q2").Value = (CStr(GreatestIncrease) & "%")
CurrentWs.Range("Q3").Value = (CStr(GreatestDecrease) & "%")
CurrentWs.Range("P2").Value = GreatestIncreaseTicker
CurrentWs.Range("P3").Value = GreatestDecreaseTicker
CurrentWs.Range("Q4").Value = GreatestVolume
CurrentWs.Range("P4").Value = GreatestVolumeTicker

'reset values for next ticker values
YearlyPercent = 0
StockVolume = 0
YearlyChange = 0
ClosePrice = 0

'Opening price for next ticker
OpenPrice = CurrentWs.Cells(i + 1, 3).Value
Else
StockVolume = StockVolume + CurrentWs.Cells(i, 7).Value
End If

Next i

Next CurrentWs



End Sub
