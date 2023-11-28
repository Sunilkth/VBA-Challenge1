Create column for and ticker symbol

create column for yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

create column and calculate percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

create column and calculate the total stock volume of the stock.

Code for the aforementioned scripts

Sub Multiple_year_stock_data()

Dim i As Long
Dim column As Integer
Dim percent_change As Double
Dim yearly_change As Double
Dim closing_price As Double
Dim opening_price As Double
Dim summary_row As Integer
Dim ticker_name As String
Dim total_vol As Double


Dim ws As Worksheet
For Each ws In Worksheets

summary_row = 2
total_vol = 0

column = 1



opening_price = ws.Cells(2, 3).Value

For i = 2 To 753001
total_vol = total_vol + ws.Cells(i, 7).Value

If ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then
ticker_name = ws.Cells(i, column).Value
ws.Cells(summary_row, 9).Value = ticker_name

closing_price = ws.Cells(i, 6).Value
yearly_change = closing_price - opening_pric

ws.Cells(summary_row, 10).Value = yearly_change

If (opening_price > 0) Then
percent_change = (yearly_change / opening_price) * (100)
Else
percent_change = 0

End If
ws.Cells(summary_row, 11).Value = percent_change
ws.Cells(summary_row, 12).Value = total_vol

If yearly_change > 0 Then
ws.Cells(summary_row, 10).Interior.ColorIndex = 4
Else
ws.Cells(summary_row, 10).Interior.ColorIndex = 3
End If

opening_price = ws.Cells(i + 1, 3).Value
total_vol = 0
summary_row = summary_row + 1
End If
Next i
Next ws


End Sub

