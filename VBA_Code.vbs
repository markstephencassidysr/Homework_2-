Attribute VB_Name = "Module1"
Sub Multiple_year_stock_data()
Dim Ticker As String
Ticker = Range("A2")
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim Counter As Long
Dim TotalStockVolume As Long
Dim StockVolume As Double

Counter = 1
Dim lastrow As Long
lastrow = Cells(Rows.Count, "A").End(xlUp).Row
OpeningPrice = Cells(2, 3).Value
StockVolume = Cells(2, 7).Value

Cells(1, 9).Value = ("Ticker")
Cells(1, 10).Value = ("Yearly Change")
Cells(1, 11).Value = ("Percent Change")
Cells(1, 12).Value = ("Total Stock Volume")

For i = 2 To lastrow
If Cells(i, 1).Value <> Ticker Then
    ClosingPrice = Cells(i - 1, 6).Value
Cells(1 + Counter, 10).Value = ClosingPrice - OpeningPrice
Cells(1 + Counter, 9).Value = Ticker
Cells(1 + Counter, 11).Value = (Cells(1 + Counter, 10) / OpeningPrice)
Cells(1 + Counter, 11).NumberFormat = "0.00%"
Cells(1 + Counter, 12).Value = StockVolume
StockVolume = 0
Ticker = Cells(i, 1).Value

If Cells(1 + Counter, 10).Value < 0 Then
    Cells(1 + Counter, 10).Interior.ColorIndex = 3
    ElseIf Cells(1 + Counter, 10).Value > 0 Then
    Cells(1 + Counter, 10).Interior.ColorIndex = 4
    End If
        Counter = Counter + 1
        OpeningPrice = Cells(i, 3).Value
        
Else
    StockVolume = StockVolume + Cells(i, 7).Value
On Error Resume Next
End If

Next i


End Sub
Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Multiple_year_stock_data
    Next
    Application.ScreenUpdating = True

End Sub



