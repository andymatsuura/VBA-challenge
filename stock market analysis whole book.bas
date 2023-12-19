Attribute VB_Name = "Module2"
Sub SuperStonks()

  Dim ws As Worksheet
  
  For Each ws In Worksheets

  Dim TickerName As String
    Dim TickerVolume As Double
    Dim SummaryTable As Integer
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim AnnualChange As Double
    Dim PercentChange As Double
    
        SummaryTable = 2
        TickerVolume = 0
        OpenPrice = ws.Cells(2, 3).Value
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'code source in readme
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastrow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            TickerName = ws.Cells(i, 1).Value
            TickerVolume = TickerVolume + ws.Cells(i, 7).Value
            ClosePrice = ws.Cells(i, 6).Value
            AnnualChange = ClosePrice - OpenPrice
            
            If OpenPrice = 0 Then
            
                PercentChange = 0
                
            Else
            
                PercentChange = AnnualChange / OpenPrice
            End If
            
            ws.Range("I" & SummaryTable).Value = TickerName
            ws.Range("J" & SummaryTable).Value = AnnualChange
            ws.Range("K" & SummaryTable).Value = PercentChange
            ws.Range("L" & SummaryTable).Value = TickerVolume
            
            SummaryTable = SummaryTable + 1
            
            TickerVolume = 0
            OpenPrice = ws.Cells(i + 1, 3).Value
            
        Else
        
            TickerVolume = TickerVolume + ws.Cells(i, 7).Value
            
        End If
        
    Next i

    ws.Columns("K:K").NumberFormat = "0.00%"
    ' Code below source in readme
    lastrowsummary = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    For i = 2 To lastrowsummary
    
        If ws.Cells(i, 10).Value > 0 Then
        
            ws.Cells(i, 10).Interior.ColorIndex = 4
            
        Else
        
            ws.Cells(i, 10).Interior.ColorIndex = 3
            
        End If
        
    Next i
    
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
    For i = 2 To lastrowsummary
        'source used for max/min function in readme
    
        If ws.Cells(i, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrowsummary)) Then
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
            ws.Cells(2, 17).NumberFormat = "0.00%"
            
        ElseIf ws.Cells(i, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrowsummary)) Then
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
            ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ElseIf ws.Cells(i, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrowsummary)) Then
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
            
        End If
    
    Next i
    
    Next ws
    
End Sub
