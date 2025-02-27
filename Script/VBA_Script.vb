VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Quarter1()

For Each ws In Worksheets

    Dim i As Long
    Dim tickername As String
    Dim tickervolume As Double
    Dim TickerRow As Long
    Dim quarterly_change As Double
    Dim percent_change As Double
    Dim close_price As Double
    Dim open_price As Double
    Dim LastRow As Long
    Dim lastrow_summary_table As Long
    
    tickervolume = 0
    TickerRow = 2
    open_price = Cells(2, 3).Value
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            tickername = ws.Cells(i, 1).Value
            tickervolume = tickervolume + ws.Cells(i, 7).Value
            
            Range("I" & TickerRow).Value = tickername
            Range("L" & TickerRow).Value = tickervolume
            close_price = ws.Cells(i, 6).Value
            quarterly_change = (close_price - open_price)
            Range("J" & TickerRow).Value = quarterly_change
            If open_price = 0 Then
                percent_change = 0
            Else
                percent_change = quarterly_change / open_price
            End If
            Range("K" & TickerRow).Value = percent_change
            Range("K" & TickerRow).NumberFormat = "0.00%"
            
            TickerRow = TickerRow + 1
            
            tickervolume = 0
            open_price = ws.Cells(i + 1, 3).Value
        Else
            
            tickervolume = tickervolume + ws.Cells(i, 7).Value
        End If
    Next i
    
    lastrow_summary_table = ws.Cells(Rows.Count, 9).End(xlUp).Row
    For i = 2 To lastrow_summary_table
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 10
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
    
    For i = 2 To lastrow_summary_table
        If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrow_summary_table)) Then
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
            ws.Cells(2, 17).NumberFormat = "0.00%"
        ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrow_summary_table)) Then
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
            ws.Cells(3, 17).NumberFormat = "0.00%"
        ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrow_summary_table)) Then
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
        End If
    Next i
    Next ws
End Sub




