
Sub Stock_Market():

    
    Dim Ticker As String
    Dim OpenRate As Double
    Dim CloseRate As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVol As LongLong
    TotalVol = 0
    
    For Each ws In Worksheets
    
      
      OpenRate = ws.Cells(2, 3).Value
       
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        Dim TableRow As Integer
        TableRow = 2
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
            For i = 2 To LastRow - 1
               
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                 Ticker = ws.Cells(i, 1).Value
                 TotalVal = TotalVal + ws.Cells(i, 7).Value
                 CloseRate = ws.Cells(i, 6).Value
        
                 YearlyChange = (CloseRate - OpenRate)
                
                 PercentChange = YearlyChange / OpenRate
                
                 ' Print Ticker, YearlyChange, PercentChange and TotalVolume in Table
                 ws.Range("I" & TableRow).Value = Ticker
                 ws.Range("J" & TableRow).Value = YearlyChange
                 ws.Range("K" & TableRow).Value = PercentChange
                 ws.Range("L" & TableRow).Value = TotalVol
                    
                    If ws.Range("J" & TableRow).Value >= 0 Then
                    ws.Range("J" & TableRow).Interior.ColorIndex = 4
                    Else
                    ws.Range("J" & TableRow).Interior.ColorIndex = 3
                    End If
                    
                
                 TableRow = TableRow + 1
                 TotalVol = 0
                 OpenRate = ws.Cells(i + 1, 3).Value
                
                Else
                TotalVol = TotalVol + ws.Cells(i, 7).Value
                
                End If
            
            Next i
            
            
            ws.Range("K:K").Style = "Percent"
            ws.Range("K:K").NumberFormat = "0.00%"
            ws.Range("O1").Value = "Ticker"
            ws.Range("P1").Value = "Value"
            ws.Range("N2").Value = "Greatest % Increase"
            ws.Range("N3").Value = "Greatest % Decrease"
            ws.Range("N4").Value = "Greatest Total Volume"
            
            
            Last = ws.Cells(Rows.Count, 9).End(xlUp).Row
            
            For i = 2 To Last
            
            If ws.Cells(i, 11).Value > ws.Range("P2").Value Then
            
            ws.Range("P2").Value = ws.Cells(i, 11).Value
            
            ws.Range("O2").Value = ws.Cells(i, 9).Value
            
            End If
            
        Next i
            
        For i = 2 To Last
            
            If ws.Cells(i, 11).Value < ws.Range("P3").Value Then
            
            ws.Range("P3").Value = ws.Cells(i, 11).Value
            
            ws.Range("O3").Value = ws.Cells(i, 9).Value
            
            End If
            
        Next i
            
       
        For i = 2 To Last
            
            If ws.Cells(i, 12).Value > ws.Range("P4").Value Then
            
            ws.Range("P4").Value = ws.Cells(i, 12).Value
            
            ws.Range("O4").Value = ws.Cells(i, 9).Value
            
            End If
            
        Next i
            
            
            ws.Range("P2:P3").Style = "Percent"
            ws.Range("P2:P3").NumberFormat = "0.00%"
            
            
            ws.Columns("I:P").AutoFit
        Next ws
        
End Sub
