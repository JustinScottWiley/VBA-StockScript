# VBA-StockScript
Module 2 Challenge

VBA CODE

Sub StockAnalysis()
 
    Dim Total As Double
    Dim Change As Double
    Dim RowCount As Long
    Dim RowIndex As Long
    Dim Start As Long
    Dim ColumnIndex As Integer
    Dim PercentChange As Double
    Dim Days As Integer
    Dim DailyChange As Single
    Dim AverageChange As Double
     
    For Each ws In Worksheets
        ColumnIndex = 0
        Total = 0
        Change = 0
        Start = 2
        DailyChange = 0
            
        ws.Range("I1").Value = "Ticker"
        ws.Columns("I").ColumnWidth = 10
        ws.Range("J1").Value = "Yearly Change"
        ws.Columns("J").ColumnWidth = 15
        ws.Range("K1").Value = "Percent Change"
        ws.Columns("K").ColumnWidth = 15
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Columns("L").ColumnWidth = 20
        ws.Range("N2").Value = "Greatest % Increase"
        ws.Range("N3").Value = "Greatest % Decrease"
        ws.Range("N4").Value = "Greatest Total Volume"
        ws.Columns("N").ColumnWidth = 20
        ws.Range("O1").Value = "Ticker"
        ws.Columns("O").ColumnWidth = 10
        ws.Range("P1").Value = "Value"
        ws.Columns("P").ColumnWidth = 10
                 
        RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
                
        For RowIndex = 2 To RowCount
            If ws.Cells(RowIndex + 1, 1).Value <> ws.Cells(RowIndex, 1).Value Then
                Total = Total + ws.Cells(RowIndex, 7).Value
                
                If Total = 0 Then
                    ws.Range("I" & 2 + ColumnIndex).Value = Cells(RowIndex, 1).Value
                    ws.Range("J" & 2 + ColumnIndex).Value = 0
                    ws.Range("K" & 2 + ColumnIndex).Value = "%" & 0
                    ws.Range("L" & 2 + ColumnIndex).Value = 0
                Else
                    If ws.Cells(Start, 3) = 0 Then
                        For Find_Value = Start To RowIndex
                            If ws.Cells(Find_Value, 3).Value <> 0 Then
                                Start = Find_Value
                                Exit For
                            End If
                        Next Find_Value
                    End If
                    
                    Change = (ws.Cells(RowIndex, 6) - ws.Cells(Start, 3))
                    PercentChange = Change / ws.Cells(Start, 3)
                    Start = RowIndex + 1
                    
                    ws.Range("I" & 2 + ColumnIndex) = ws.Cells(RowIndex, 1).Value
                    ws.Range("J" & 2 + ColumnIndex) = Change
                    ws.Range("J" & 2 + ColumnIndex).NumberFormat = "0.00"
                    ws.Range("K" & 2 + ColumnIndex).Value = PercentChange
                    ws.Range("K" & 2 + ColumnIndex).NumberFormat = "0.00%"
                    ws.Range("L" & 2 + ColumnIndex).Value = Total
                    
                    Select Case Change
                        Case Is > 0
                            ws.Range("J" & 2 + ColumnIndex).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & 2 + ColumnIndex).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("J" & 2 + ColumnIndex).Interior.ColorIndex = 0
                    End Select
                End If
                Total = 0
                Change = 0
                ColumnIndex = ColumnIndex + 1
                Days = 0
                DailyChange = 0
            Else
                Total = Total + ws.Cells(RowIndex, 7).Value
            End If
        Next RowIndex
        ws.Range("P2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & RowCount)) * 100
        ws.Range("P3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & RowCount)) * 100
        ws.Range("P4") = WorksheetFunction.Max(ws.Range("K2:K" & RowCount))
        
        Increase_Number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
        Decrease_Number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & RowCount)), ws.Range("K2:K" & RowCount), 0)
        Volume_Number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & RowCount)), ws.Range("L2:L" & RowCount), 0)
        
        ws.Range("O2") = ws.Cells(Increase_Number + 1, 9)
        ws.Range("O3") = ws.Cells(Decrease_Number + 1, 9)
        ws.Range("O4") = ws.Cells(Volume_Number + 1, 9)
        
    Next ws
 
End Sub
