
Sub StockTicker()
    
Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
        
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

       
    WS.Cells(1, 9).Value = "Ticker"
    WS.Cells(1, 10).Value = "Yearly Change"
    WS.Cells(1, 11).Value = "Percent Change"
    WS.Cells(1, 12).Value = "Total Stock Volume"

        Dim LastRowNum, FirstTickerRowNum, NextTickerRowNum As Long

    Dim TickerYearClose, TickerYearOpen As Double

    Dim TickerNum As Long

    Dim TickerName, NextTickerName  As String

    Dim Volume As Double

    Dim TotalTabNum As Integer
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Integer
        Column = 1
        Dim i As Long
        
        
        Open_Price = Cells(2, 3).Value
         
        
        For i = 2 To LastRow
         
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then
                
                Ticker_Name = Cells(i, Column).Value
                Cells(Row, 9).Value = Ticker_Name
                
                Close_Price = Cells(i, 6).Value
                
                Yearly_Change = Close_Price - Open_Price
                Cells(Row, 10).Value = Yearly_Change
                
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(Row, 11).Value = Percent_Change
                    Cells(Row, 11).NumberFormat = "0.00%"
                End If
                
                Volume = Volume + Cells(i, 7).Value
                Cells(Row, 12).Value = Volume
                
                Row = Row + 1
                
                Open_Price = Cells(i + 1, 3)
                
                Volume = 0
            
            Else
                Volume = Volume + Cells(i, 7).Value
            End If
        Next i
        
        
        YCLastRow = WS.Cells(Rows.Count, 9).End(xlUp).Row
        
        For j = 2 To YCLastRow
            If (Cells(j, 10).Value > 0 Or Cells(j, 10).Value = 0) Then
                Cells(j, 10).Interior.ColorIndex = 10
            ElseIf Cells(j, 10).Value < 0 Then
                Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
        
        
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        For Z = 2 To YCLastRow
            If Cells(Z, 11).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & YCLastRow)) Then
                Cells(2, 16).Value = Cells(Z, 9).Value
                Cells(2, 17).Value = Cells(Z, 11).Value
                Cells(2, 17).NumberFormat = "0.00%"
            ElseIf Cells(Z, 11).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & YCLastRow)) Then
                Cells(3, 16).Value = Cells(Z, 9).Value
                Cells(3, 17).Value = Cells(Z, 11).Value
                Cells(3, 17).NumberFormat = "0.00%"
            ElseIf Cells(Z, 12).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & YCLastRow)) Then
                Cells(4, 16).Value = Cells(Z, 9).Value
                Cells(4, 17).Value = Cells(Z, 12).Value
            End If
        Next Z
        
    Next WS
        
End Sub







