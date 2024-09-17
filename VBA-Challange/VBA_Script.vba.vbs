Sub StockAnalysis():
    'Set Dimensions/Variables
    Dim TotalVolume As Double
    Dim Change As Double
    Dim PercentChange As Double
    Dim i As Long
    Dim j As Integer
    Dim Start As Long
    Dim RowEndCount As Long
    Dim QuarterlyChange As Double
    Dim Ticker_incr As String
    Dim Open_Ticker As Double
    Dim Close_Ticker As Double
    Dim Volume_Total As Double
    Dim Increase_Percents As Double
    Dim Decrease_Percents As Double
    Dim Increase_Ticker As Double
    Dim Decrease_Ticker As Double
    Dim Total_Ticker As Double
    Dim Ticker_dcr As String
    Dim Volume_MAX As Double
    Dim Vol_Max_Ticker As String
    Dim ws As Worksheet
    
    'Loop through each worksheet
    For Each ws In Worksheets
    
    'Set row titles
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    'Initial value
    TotalVolume = 0
    Change = 0
    j = 0
    Start = 2
    Increase_Percents = 0
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    'Find the last row with data in column A
    RowEndCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    'For loop (calculate total volume, change, and percent change), and print out final results
    For i = 2 To RowEndCount
        'If the ticker changes, calculate and print results
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
            'Calculate Change and Percent Change Hint: Change has to do with open and closed
            QuarterlyChange = ws.Cells(i, 6).Value - ws.Cells(Start, 3).Value
            PercentChange = QuarterlyChange / ws.Cells(Start, 3).Value
            
            'Colors positives green and negative red
            Select Case QuarterlyChange
                Case Is > 0
                    'Color the cell Green (ColorIndex 4 is green)
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                Case Is < 0
                    'Color the cell red (ColorIndex 3 is red)
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                Case Else
                    'Reset the color to no fill (ColorIndex 0)
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 0
            End Select
    
            'Print
            ws.Range("L" & 2 + j).Value = TotalVolume
            ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
            ws.Range("J" & 2 + j).Value = QuarterlyChange
            ws.Range("K" & 2 + j).Value = PercentChange
            ws.Range("K" & 2 + j).NumberFormat = "0.00%"
    
            'Reset total volume for the next ticker
            TotalVolume = 0
            j = j + 1
            Start = i + 1
        Else
            'Continue accumulating the total volume
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        End If
        
        Change = 0
        
        PercentChange = ws.Cells(i, 11)
        If (PercentChange > Increase_Percents) Then
            Increase_Percents = PercentChange
            Ticker_incr = ws.Cells(i, 9).Value
        ElseIf (PercentChange < Decrease_Percents) Then
            Decrease_Percents = PercentChange
            Ticker_dcr = ws.Cells(i, 9).Value
        End If
        
        Volume_Total = ws.Cells(i, 12)
        If (Volume_Total > Volume_MAX) Then
            Volume_MAX = Volume_Total
            Vol_Max_Ticker = ws.Cells(i, 9)
        End If
            
    Next i
    
    'Greatest precent increase and decrease and print results
    ws.Range("Q2").Value = Increase_Percents
    ws.Range("Q3").Value = Decrease_Percents
    ws.Range("P2").Value = Ticker_incr
    ws.Range("P3").Value = Ticker_dcr
    ws.Range("Q4").Value = Volume_MAX
    ws.Range("P4").Value = Vol_Max_Ticker
    
    Next ws
End Sub