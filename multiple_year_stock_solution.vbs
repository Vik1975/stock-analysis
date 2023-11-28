Sub All_Year():

    For Each ws In Worksheets
    
        Dim WorksheetName As String
        Dim PercentDiff As Double
        'Maximum increase
        Dim MaxIncrease As Double
        'Maximum decrease
        Dim MaxDecrease As Double
        'Maximum total volume
        Dim MaxVolume As Double
        Dim TotalVolume As Double
        Dim i As Long
        Dim j As Long
        Dim Ticker_count As Long
        'Length for column A
        Dim LengthColumnA As Long
        'Length for column I
        Dim LengthColumnI As Long
        'Percent diff
        WorksheetName = ws.Name
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
       TotalVolume = 0
       Ticker_type_count = 2
        j = 2
        
        'The number of  non-empty cells in column A
        LengthColumnA = ws.Cells(Rows.Count, 1).End(xlUp).Row
            For i = 2 To LengthColumnA
                TotalVolume = TotalVolume + ws.Cells(i, 7).Value
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                  ws.Cells(Ticker_type_count, 12).Value = TotalVolume
                  TotalVolume = 0
                  ws.Cells(Ticker_type_count, 9).Value = ws.Cells(i, 1).Value
                  ws.Cells(Ticker_type_count, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    'Color formating
                    If ws.Cells(Ticker_type_count, 10).Value < 0 Then
                       'Cell becomes red
                       ws.Cells(Ticker_type_count, 10).Interior.Color = RGB(255, 0, 0)
                    Else
                       'Cell becomes green
                       ws.Cells(Ticker_type_count, 10).Interior.Color = RGB(0, 255, 0)
                    End If
                    If ws.Cells(j, 3).Value <> 0 Then
                       PercentDiff = (ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value
                       ws.Cells(Ticker_type_count, 11).Value = Format(PercentDiff, "Percent")
                    Else
                       ws.Cells(Ticker_type_count, 11).Value = Format(0, "Percent")
                    End If
                Ticker_type_count = Ticker_type_count + 1
                j = i + 1
                End If
            Next i
        'Number of non-empty cells in column I
        LengthColumnI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Summary
        MaxVolume = ws.Range("L2").Value
        MaxIncrease = ws.Range("K2").Value
        MaxDecrease = ws.Range("K2").Value
            For i = 2 To LengthColumnI
                If ws.Cells(i, 11).Value > MaxIncrease Then
                   MaxIncrease = ws.Cells(i, 11).Value
                   ws.Range("Q2").Value = ws.Cells(i, 9).Value
                Else
                   MaxIncrease = MaxIncrease
                End If
                If ws.Cells(i, 11).Value < MaxDecrease Then
                   MaxDecrease = ws.Cells(i, 11).Value
                   ws.Range("Q3").Value = ws.Cells(i, 9).Value
                Else
                   MaxDecrease = MaxDecrease
                End If
                If ws.Cells(i, 12).Value > MaxVolume Then
                    MaxVolume = ws.Cells(i, 12).Value
                    ws.Range("Q4").Value = ws.Cells(i, 9).Value
                Else
                    MaxVolume = MaxVolume
                End If
                
            'Summary results formatting
            ws.Range("Q2").Value = Format(MaxIncrease, "Percent")
            ws.Range("Q3").Value = Format(MaxDecrease, "Percent")
            ws.Range("Q4").Value = Format(MaxVolume, "Scientific")
            
            Next i
    
      
     Worksheets(WorksheetName).Columns("A:Q").AutoFit
     
    Next ws
        
End Sub