Attribute VB_Name = "Module1"

Sub stock_part1()

Dim WS As Worksheet


For Each WS In Worksheets

'Dim WorksheetName As String
LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

Dim Ticker As String
Dim Stock_Total As Double
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

WS.Cells(1, 9).Value =

WS.Cells(1, 10).Value = "Yearly_Change"

WS.Cells(1, 11).Value = "Percentage_Change"

WS.Cells(1, 12).Value = "Ticker_Stock_Total"

'For i = 2 To 753001
For i = 2 To LastRow

If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
Ticker = WS.Cells(i, 1).Value
Stock_Total = Stock_Total + WS.Cells(i, 7).Value

WS.Range("I" & Summary_Table_Row).Value = Ticker
WS.Range("L" & Summary_Table_Row).Value = Stock_Total

Summary_Table_Row = Summary_Table_Row + 1
Stock_Total = 0

Else

Stock_Total = Stock_Total + WS.Cells(i, 7).Value

End If
Next i

    
Next WS

End Sub

