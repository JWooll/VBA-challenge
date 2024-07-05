Attribute VB_Name = "Module1"
Sub Ticker()

Dim i As Long
Dim j As Integer
Dim lastrow As Long
Dim InitValue As Double
Dim FinalValue As Double
Dim total As LongLong
Dim extrema(0 To 2) As Double
Dim tic(0 To 2) As String


Dim ws As Worksheet

For Each ws In Worksheets
        'Headers'
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Quarterly Change"
        ws.Cells(1, 11) = "Percent change"
        ws.Cells(1, 12) = "Total Stock Volume"
        ws.Cells(1, 16) = "Ticker"
        ws.Cells(1, 17) = "Value"
        ws.Cells(2, 15) = "Greatest % increase"
        ws.Cells(3, 15) = "Greatest % Decrease"
        ws.Cells(4, 15) = "Greatest Total Volume"
        
        'Initializing values'
        j = 2
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        InitValue = ws.Cells(2, 3)
        total = 0
        
        'These arrays are used to store the greatest positive change, negative change and volume total.'
        'Note that if for some reason there is no negative or positive change for any ticker in a data set, it will be noted so as none'
        For i = 0 To 2
            extrema(i) = 0
            tic(i) = "none"
        Next i
        
        For i = 2 To lastrow
            total = total + ws.Cells(i, 7)
            If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
                'Ticker names'
                ws.Cells(j, 9) = ws.Cells(i, 1)
                'Setting value at end of Quarter closing'
                FinalValue = ws.Cells(i, 6)
                'Change in Quarterly Value'
                ws.Cells(j, 10) = FinalValue - InitValue
                    'Color coding'
                    If ws.Cells(j, 10) > 0 Then
                    ws.Cells(j, 10).Interior.Color = RGB(0, 255, 0)
                    ElseIf ws.Cells(j, 10) < 0 Then
                    ws.Cells(j, 10).Interior.Color = RGB(255, 0, 0)
                    Else
                    ws.Cells(j, 10).Interior.Color = RGB(255, 255, 255)
                    End If
                
                'Percentage Change'
                ws.Cells(j, 11) = ws.Cells(j, 10) / InitValue
                    'test for if the percentage change is greatest in one direction vs other tickers'
                    If ws.Cells(j, 11) > extrema(2) Then
                        extrema(2) = ws.Cells(j, 11)
                        tic(2) = ws.Cells(j, 9)
                    ElseIf ws.Cells(j, 11) < extrema(1) Then
                        extrema(1) = ws.Cells(j, 11)
                        tic(1) = ws.Cells(j, 9)
                    End If
                 ws.Cells(j, 11).NumberFormat = "0.00%"
                'New opening value for next ticker'
                InitValue = ws.Cells(i + 1, 3)
                'Total Volume and reset volume counter'
                ws.Cells(j, 12) = total
                    'test for if total volume is greatest vs other tickers'
                    If total > extrema(0) Then
                        extrema(0) = total
                        tic(0) = ws.Cells(j, 9)
                    End If
                'Reset total counter'
                total = 0
                
                j = j + 1
            End If
        Next i
        
        'Filling in the largest percentage change'
        ws.Cells(2, 16) = tic(2)
        ws.Cells(3, 16) = tic(1)
        ws.Cells(4, 16) = tic(0)
        ws.Cells(2, 17) = extrema(2)
        ws.Cells(3, 17) = extrema(1)
        ws.Cells(4, 17) = extrema(0)
        
        'formatting cells'
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"
        ws.Cells(4, 17).NumberFormat = "0.00E+00"
        ws.Columns("J:J").EntireColumn.AutoFit
        ws.Columns("K:K").EntireColumn.AutoFit
        ws.Columns("L:L").EntireColumn.AutoFit
        ws.Columns("O:P").EntireColumn.AutoFit
        ws.Columns("P:P").EntireColumn.AutoFit
        ws.Columns("Q:Q").EntireColumn.AutoFit
        
Next ws

End Sub
