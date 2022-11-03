Sub stock_market_analysis()
    
    ' Declare variables
    
    Dim NewTickerSymbol As String
    Dim PrevTickerSymbol As String
    Dim ResultsRowCounter As Integer
    Dim OpenDate As String
    Dim PrevTickerCloseDate As String
    Dim NumOfRowsInSheet As Double
    Dim TickerOpenPrice As Double
    Dim TickerClosePrice As Double
    Dim YearlyChange As Double
    Dim TotalVolume As Double
    
    ' Bonus section variables
    Dim GreatestPercentIncrease As Double
    Dim CurrentPercentIncrease As Double
    Dim PreviousPercentIncrease As Double

    Dim GreatestPercentDecrease As Double
    Dim GreatestTotalVolume As Double
    
    For Each ws In Worksheets
    
        ' Initialize variables
        
        NewTickerSymbol = "Start"
        PrevTickerSymbol = "None"
        ResultsRowCounter = 2
        
        YearlyChange = 0
        TotalVolume = 0
        
        
        'Cleanup report area
        ws.Columns("I:U").EntireColumn.Delete
        
        'Build header row
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Open Date"
        ws.Range("K1").Value = "Close Date"
        ws.Range("L1").Value = "Open Price"
        ws.Range("M1").Value = "Close Price"
        ws.Range("N1").Value = "Yearly Change"
        ws.Range("O1").Value = "Percent Change"
        ws.Range("P1").Value = "Total Stock Volume"
        
        ws.Range("I1").Font.Bold = True
        ws.Range("J1").Font.Bold = True
        ws.Range("K1").Font.Bold = True
        ws.Range("L1").Font.Bold = True
        ws.Range("M1").Font.Bold = True
        ws.Range("N1").Font.Bold = True
        ws.Range("O1").Font.Bold = True
        ws.Range("P1").Font.Bold = True
        
        
       ' LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastRow = ws.Range("A2").End(xlDown).Row
    
        
        NumOfRowsInSheet = LastRow  ' 25554     '22026
            
        
        ' Main logic
        
        For i = 2 To NumOfRowsInSheet
            
            'grab a new ticker symbol
            NewTickerSymbol = ws.Cells(i, 1)
            
            'if brand new symbol
            If NewTickerSymbol <> PrevTickerSymbol Then
                
                'write it to worksheet
                ws.Cells(ResultsRowCounter, 9).Value = NewTickerSymbol
                
                'collect additional ticker data
                OpenDate = ws.Cells(i, 2).Value
                TickerOpenPrice = ws.Cells(i, 3).Value
                
                If i > 2 Then
                    PrevTickerCloseDate = ws.Cells((i - 1), 2).Value
                    ws.Cells(ResultsRowCounter - 1, 11).Value = PrevTickerCloseDate
                    
                    TickerClosePrice = ws.Cells(i - 1, 6).Value
                    ws.Cells(ResultsRowCounter - 1, 13).Value = TickerClosePrice
                    ws.Cells(ResultsRowCounter - 1, 16).Value = TotalVolume
                End If
                
                
                'write additional data to worksheet results
                ws.Cells(ResultsRowCounter, 10).Value = OpenDate
                ws.Cells(ResultsRowCounter, 12).Value = TickerOpenPrice
                
                'Reset to start calculating a new ticker symbol total volume
                TotalVolume = 0
                
                'set for next loop
                PrevTickerSymbol = NewTickerSymbol
                ResultsRowCounter = ResultsRowCounter + 1
                
            End If
            
            'keep adding to the TotalVolume
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
        Next i
        
        ' Handling last row
        PrevTickerCloseDate = ws.Cells(NumOfRowsInSheet, 2).Value
        ws.Cells(ResultsRowCounter - 1, 11).Value = PrevTickerCloseDate
    
        TickerClosePrice = ws.Cells(NumOfRowsInSheet, 6).Value
        ws.Cells(ResultsRowCounter - 1, 13).Value = TickerClosePrice
        
        ws.Cells(ResultsRowCounter - 1, 16).Value = TotalVolume
        
        ResultsRowCount = ws.Range("I2").End(xlDown).Row
        
        ' MsgBox ("Results Row Count: " + Str(ResultsRowCount))
        
        GreatestPercentIncrease = 0
        CurrentPercentIncrease = 0
        PreviousPercentIncrease = 0
        GreatestPercentTicker = "none"
    
        GreatestPercentDecrease = 0
        CurrentPercentDecrease = 0
        PreviousPercentDecrease = 0
        GreatestPercentDecreaseTicker = "none"
        
        GreatestTotalVolume = 0
        CurrentTotalVolume = 0
        PreviousTotalVolume = 0
        GreatestTotalVolumeTicker = "none"
        
        ' Compile final calculations
        For i = 2 To ResultsRowCount ' 91
            YearlyChange = ws.Cells(i, 13).Value - ws.Cells(i, 12).Value
            ws.Cells(i, 14).Value = YearlyChange
            
            If YearlyChange > 0 Then
                ws.Cells(i, 14).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 14).Interior.ColorIndex = 3
            End If
            
            'Pecent Change calculation
            ws.Cells(i, 15).Value = FormatPercent((ws.Cells(i, 14).Value / ws.Cells(i, 12).Value), 2)
            
            'Bonus percent increase calculation
            CurrentPercentIncrease = ws.Cells(i, 14).Value / ws.Cells(i, 12).Value
            If CurrentPercentIncrease > PreviousPercentIncrease Then
                GreatestPercentIncrease = CurrentPercentIncrease
                GreatestPercentTicker = ws.Cells(i, 9).Value
                
                PreviousPercentIncrease = CurrentPercentIncrease
            End If
            
            
            'Bonus percent decrease calculation
            CurrentPercentDecrease = ws.Cells(i, 14).Value / ws.Cells(i, 12).Value
            If CurrentPercentDecrease < PreviousPercentDecrease Then
                GreatestPercentDecrease = CurrentPercentDecrease
                GreatestPercentDecreaseTicker = ws.Cells(i, 9).Value
                
                PreviousPercentDecrease = CurrentPercentDecrease
            End If
            
            'Bonus total volume calculation
            CurrentTotalVolume = ws.Cells(i, 16).Value
            If CurrentTotalVolume > PreviousTotalVolume Then
                GreatestTotalVolume = CurrentTotalVolume
                GreatestTotalVolumeTicker = ws.Cells(i, 9).Value
                
                PreviousTotalVolume = CurrentTotalVolume
            End If
            
        Next i
        
        ' Bonus calculations
        ' Headers & titles
        ws.Range("T1").Value = "Ticker"
        ws.Range("U1").Value = "Value"
        ws.Range("S2").Value = "Greatest % Increase"
        ws.Range("S3").Value = "Greatest % Decrease"
        ws.Range("S4").Value = "Greatest Total Volume"
         
        
        ' GreatestPercentIncrease = Application.WorksheetFunction.Max(Range(Cells(2, 15), Cells(ResultsRowCount, 15)))
        ws.Range("T2").Value = GreatestPercentTicker
        ws.Range("U2").Value = FormatPercent(GreatestPercentIncrease, 2)
    
        
        ' GreatestPercentDecrease = Application.WorksheetFunction.Min(Range(Cells(2, 15), Cells(ResultsRowCount, 15)))
        ws.Range("T3").Value = GreatestPercentDecreaseTicker
        ws.Range("U3").Value = FormatPercent(GreatestPercentDecrease, 2)
        
        'GreatestTotalVolume = Application.WorksheetFunction.Sum(Range(Cells(2, 16), Cells(ResultsRowCount, 16)))
        ws.Range("T4").Value = GreatestTotalVolumeTicker
        ws.Range("U4").Value = GreatestTotalVolume
        
        'Cleanup extra columns
        ws.Columns("J:M").EntireColumn.Delete
        ws.Columns("I:Q").AutoFit
    
    Next ws
    
End Sub
