Sub stocks():

    For Each ws In ThisWorkbook.Worksheets
    
        ws.Activate

    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percentage Change"
    Range("L1").Value = "Total Stock Volume"
    
    Dim i As Long
    Dim Ticker As String
    
    Dim StockTotal As LongLong
        StockTotal = 0
        
    Dim SummaryTable As Integer
        SummaryTable = 2
        
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To LastRow
        
        If Cells(i + 1, 1).Value <> Cells(i, 1) Then
        
            Ticker = Cells(i, 1).Value
            StockTotal = StockTotal + Cells(i, 7).Value
            
            Range("I" & SummaryTable).Value = Ticker
            Range("L" & SummaryTable).Value = StockTotal
            
            SummaryTable = SummaryTable + 1
            
            StockTotal = 0
            
        Else
            
            StockTotal = StockTotal + Cells(i, 7).Value
            
        End If
        
    Next i
    
    
    Dim CloseBalance As Double
    Dim OpenBalance As Double
    Dim YearlyChange As Double
        YearlyChange = 0
    
    SummaryTable = 2
    
    For i = 2 To LastRow
        
        If Cells(i + 1, 1).Value <> Cells(i, 1) Then
        
            OpenBalance = Cells(i, 3).Value
        
            CloseBalance = Cells(i, 6).Value
        
            YearlyChange = CloseBalance - OpenBalance
            
            Range("J" & SummaryTable).Value = YearlyChange
            
            If YearlyChange >= 0 Then
                Range("J" & SummaryTable).Interior.ColorIndex = 4
            ElseIf YearlyChange < 0 Then
                Range("J" & SummaryTable).Interior.ColorIndex = 3
            End If
                
            SummaryTable = SummaryTable + 1
            
            YearlyChange = 0
            
        Else
        
            YearlyChange = CloseBalance - OpenBalance
            
        End If
        
    Next i
    
        Dim PercentChange As Double
        PercentChange = 0
        
        SummaryTable = 2
    
    For i = 2 To LastRow
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            OpenBalance = Cells(i, 3).Value
        
            CloseBalance = Cells(i, 6).Value
            
            PercentageChange = YearlyChange / OpenBalance
            
            Range("K" & SummaryTable).Value = PercentageChange
            Range("K" & SummaryTable).NumberFormat = "0.00%"
            
            SummaryTable = SummaryTable + 1
            
            PercentageChange = 0
            
        Else
        
            PercentageChange = YearlyChange / OpenBalance
            
        End If
        
    Next i
   
   SummaryTable = 2
   
   Range("N2").Value = "Greatest % Increase"
   Range("N3").Value = "Greatest % Decrease"
   Range("N4").Value = "Greatest Total Volume"
   
   Range("O1").Value = "Ticker"
   Range("P1").Value = "Value"
   
   Dim MaxValue As Double
   Dim MinValue As Double
   Dim MaxVolume As LongLong
   
   MaxValue = Application.WorksheetFunction.Max(Range("K:K"))
   MinValue = Application.WorksheetFunction.Min(Range("K:K"))
   MaxVolume = Application.WorksheetFunction.Max(Range("L:L"))
   
   For i = 2 To LastRow
   
        If Cells(i, 11).Value = MaxValue Then
        
            Cells(2, 15).Value = Cells(i, 9).Value
            Cells(2, 16).Value = MaxValue
            
        ElseIf Cells(i, 11).Value = MinValue Then
            
            Cells(3, 15).Value = Cells(i, 9).Value
            Cells(3, 16).Value = MinValue
        
        ElseIf Cells(i, 12).Value = MaxVolume Then
            
            Cells(4, 15).Value = Cells(i, 9).Value
            Cells(4, 16).Value = MaxVolume
            
        End If
    
    Next i
   
   Range("P2:P3").NumberFormat = "0.00%"
   
   ws.Columns("I:P").AutoFit
   
   Next ws
    
End Sub

