Attribute VB_Name = "Module1"
Sub TickerStats()
Dim WS As Worksheet
For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

    Dim Yearly_Change As Double
    Dim Ticker_Name As String
    Dim Percent_Change As Double
    Dim Volume As Double
    Dim High_S As Double
    Dim Low_S As Double
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Row As Double
    Dim Column As Integer
    Column = 1
    Row = 2
    Volume = 0

 

    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Yearly Change"
    Cells(1, "K").Value = "Percent Change"
    Cells(1, "L").Value = "Total Stock Volume"




Open_Price = Cells(2, Column + 2).Value

 Dim r As Long
    For r = 2 To LastRow
        If Cells(r + 1, Column).Value <> Cells(r, Column).Value Then
            Ticker_Name = Cells(r, Column).Value
            Cells(Row, Column + 8).Value = Ticker_Name
            Close_Price = Cells(r, Column + 5).Value
            Yearly_Change = Close_Price - Open_Price
            Cells(Row, Column + 9).Value = Yearly_Change
        If (Open_Price = 0 And Close_Price = 0) Then
            Percent_Change = 0
        ElseIf (Open_Price = 0 And Close_Price <> 0) Then
            Percent_Change = 1
        Else
            Percent_Change = Yearly_Change / Open_Price
            Cells(Row, Column + 10).Value = Percent_Change
            Cells(Row, Column + 10).NumberFormat = "0.00%"
    End If
        Volume = Volume + Cells(r, Column + 6).Value
        Cells(Row, Column + 11).Value = Volume
        Row = Row + 1
        Open_Price = Cells(r + 1, Column + 2)
        Volume = 0
    Else
        Volume = Volume + Cells(r, Column + 6).Value
    End If
Next r

YearLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row

    For c = 2 To YearLastRow
        If (Cells(c, Column + 9).Value > 0 Or Cells(c, Column + 9).Value = 0) Then
            Cells(c, Column + 9).Interior.ColorIndex = 4
        ElseIf Cells(c, Column + 9).Value < 0 Then
            Cells(c, Column + 9).Interior.ColorIndex = 3
        End If
    Next c

Cells(2, Column + 14).Value = "Greatest % Increase"
Cells(3, Column + 14).Value = "Greatest % Decrease"
Cells(4, Column + 14).Value = "Greatest Total Volume"
Cells(1, Column + 15).Value = "Ticker"
Cells(1, Column + 16).Value = "Value"
    
    For x = 2 To YearLastRow
        If Cells(x, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YearLastRow)) Then
            Cells(2, Column + 15).Value = Cells(x, Column + 8).Value
            Cells(2, Column + 16).Value = Cells(x, Column + 10).Value
            Cells(2, Column + 16).NumberFormat = "0.00%"
        ElseIf Cells(x, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YearLastRow)) Then
            Cells(3, Column + 15).Value = Cells(x, Column + 8).Value
            Cells(3, Column + 16).Value = Cells(x, Column + 10).Value
            Cells(3, Column + 16).NumberFormat = "0.00%"
        ElseIf Cells(x, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YearLastRow)) Then
            Cells(4, Column + 15).Value = Cells(x, Column + 8).Value
            Cells(4, Column + 16).Value = Cells(x, Column + 11).Value
        End If
    Next x
Next WS

       

End Sub
