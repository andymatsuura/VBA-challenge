Attribute VB_Name = "Module2"
Sub Stonks()
    Dim TickerName As String
    Dim TickerVolume As Double
    Dim SummaryTable As Integer
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim AnnualChange As Double
    Dim PercentChange As Double
    
        SummaryTable = 2
        TickerVolume = 0
        OpenPrice = Cells(2, 3).Value
    
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'code source in readme
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastrow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            TickerName = Cells(i, 1).Value
            TickerVolume = TickerVolume + Cells(i, 7).Value
            ClosePrice = Cells(i, 6).Value
            AnnualChange = ClosePrice - OpenPrice
            
            If OpenPrice = 0 Then
            
                PercentChange = 0
                
            Else
            
                PercentChange = AnnualChange / OpenPrice
            End If
            
            Range("I" & SummaryTable).Value = TickerName
            Range("J" & SummaryTable).Value = AnnualChange
            Range("K" & SummaryTable).Value = PercentChange
            Range("L" & SummaryTable).Value = TickerVolume
            
            SummaryTable = SummaryTable + 1
            
            TickerVolume = 0
            OpenPrice = Cells(i + 1, 3).Value
            
        Else
        
            TickerVolume = TickerVolume + Cells(i, 7).Value
            
        End If
        
    Next i

    Columns("K:K").NumberFormat = "0.00%"
    ' Code below source in readme
    lastrowsummary = Cells(Rows.Count, 9).End(xlUp).Row
    
    For i = 2 To lastrowsummary
    
        If Cells(i, 10).Value > 0 Then
        
            Cells(i, 10).Interior.ColorIndex = 4
            
        Else
        
            Cells(i, 10).Interior.ColorIndex = 3
            
        End If
        
    Next i
    
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
    For i = 2 To lastrowsummary
        'source used for max/min function in readme
    
        If Cells(i, 11).Value = Application.WorksheetFunction.Max(Range("K2:K" & lastrowsummary)) Then
            Cells(2, 16).Value = Cells(i, 9).Value
            Cells(2, 17).Value = Cells(i, 11).Value
            Cells(2, 17).NumberFormat = "0.00%"
            
        ElseIf Cells(i, 11).Value = Application.WorksheetFunction.Min(Range("K2:K" & lastrowsummary)) Then
            Cells(3, 16).Value = Cells(i, 9).Value
            Cells(3, 17).Value = Cells(i, 11).Value
            Cells(3, 17).NumberFormat = "0.00%"
        
        ElseIf Cells(i, 12).Value = Application.WorksheetFunction.Max(Range("L2:L" & lastrowsummary)) Then
            Cells(4, 16).Value = Cells(i, 9).Value
            Cells(4, 17).Value = Cells(i, 12).Value
            
        End If
    
    Next i
    
End Sub
