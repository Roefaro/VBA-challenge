Attribute VB_Name = "Module1"

Sub module2()
    Dim ws As Worksheet
    Dim lastrow As Long
    Dim summary_row As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim stockvolume As Double
    Dim quarterly_change As Double
    Dim percent_change As Double
    Dim tickername As String
    Dim tickervolume As Double
    Dim greatestincrease As Double
    Dim greatestdecrease As Double
    Dim worksheetname As String

    For Each ws In Worksheets
        worksheetname = ws.Name
        
        ' Clear existing data in columns I, J, K, L
        ws.Range("I:L").Clear
        
        ' Set up headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        stockvolume = 0
        summary_row = 2
        
        For i = 2 To lastrow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                tickername = ws.Cells(i, 1).Value
                stockvolume = stockvolume + ws.Cells(i, 7).Value
                ws.Range("I" & summary_row).Value = tickername
                ws.Range("L" & summary_row).Value = stockvolume
                
                close_price = ws.Cells(i, 6).Value
                
                If open_price = 0 Then
                    quarterly_change = 0
                    percent_change = 0
                Else
                    quarterly_change = (close_price - open_price)
                    percent_change = (close_price - open_price) / open_price
                End If
                
                ws.Range("J" & summary_row).Value = quarterly_change
                ws.Range("K" & summary_row).Value = percent_change
                ws.Range("K" & summary_row).NumberFormat = "0.00%"
                
                summary_row = summary_row + 1
                stockvolume = 0
            ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
                open_price = ws.Cells(i, 3).Value
            Else
                stockvolume = stockvolume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Now, let's populate greatest increase/decrease and total volume columns
        lastrow = ws.Cells(ws.Rows.Count, 9).End(xlUp).Row
        
        ' Add column for greatest increase
        ws.Cells(2, 14).Value = "Greatest % Increase"
        For i = 2 To lastrow
            If ws.Cells(i, 11).Value = WorksheetFunction.Max(ws.Range("K2:K" & lastrow)) Then
                ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 11).Value
                ws.Cells(2, 16).NumberFormat = "0.00%"
            End If
        Next i
        
        ' Add column for greatest decrease
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        For i = 2 To lastrow
            If ws.Cells(i, 11).Value = WorksheetFunction.Min(ws.Range("K2:K" & lastrow)) Then
                ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 11).Value
                ws.Cells(3, 16).NumberFormat = "0.00%"
            End If
        Next i
        
        ' Add column for greatest total volume
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        For i = 2 To lastrow
            If ws.Cells(i, 12).Value = WorksheetFunction.Max(ws.Range("L2:L" & lastrow)) Then
                ws.Cells(4, 15).Value = ws.Cells(i, 9).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 12).Value
            End If
        Next i
        
        ' Conditional formatting
        For i = 2 To lastrow
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
    Next ws
End Sub

