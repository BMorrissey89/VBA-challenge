Sub ticker_2018()
    Dim ticker_name As String
    Dim total As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_volume As Double
    Dim start As Long
    Dim open_value As Double
    Dim closing_value As Double
    Dim lastrow As Long
    Dim i As Long
    Dim j As Long
    Dim increase_number As Double
    Dim decrease_number As Double
    Dim volume_number As Double
    
    Dim ws As Worksheet
    
        For Each ws In Worksheets
    
    ' Start
    start = 2
    j = 0
    ' Last Row
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    ' Loop
    For i = 2 To lastrow
        ' Check Ticker Name
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ' Ticker name value
            ticker_name = ws.Cells(i, 1).Value
            
            ' Yearly Change Value
            open_value = ws.Cells(start, 3).Value
            closing_value = ws.Cells(i, 6).Value
            yearly_change = closing_value - open_value
                
            ' Percent Change value
            percent_change = (yearly_change / open_value)
            ws.Range("K" & 2 + j).NumberFormat = "0.00%"
            
            ' Total stock volume
            total_volume = total_volume + ws.Cells(i, 7).Value
            Select Case yearly_change
                Case Is > 0
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                Case Is < 0
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                Case Else
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 0
            End Select
                    
            ' Place the results
           ' With ThisWorkbook.Sheets("2018")
                ws.Range("I" & 2 + j).Value = ticker_name
                ws.Range("J" & 2 + j).Value = yearly_change
                ws.Range("K" & 2 + j).Value = percent_change
                ws.Range("L" & 2 + j).Value = total_volume

            'End With
            start = i + 1
            j = j + 1
            ' Initialize total_volume
            total_volume = 0
        Else
            ' Total stock volume
            total_volume = total_volume + ws.Cells(i, 7).Value
        End If
    Next i
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    ws.Range("O2").Value = "Greatest Percent Increase"
    ws.Range("O3").Value = "Greatest Percent Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    ws.Range("Q2").Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastrow)) * 100
    ws.Range("Q3").Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & lastrow)) * 100
    ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
    
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)
    
    ws.Range("P2") = ws.Cells(increase_number + 1, 9)
    ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
    ws.Range("P4") = ws.Cells(volume_number + 1, 9)
    
    Next ws
    
    
End Sub


