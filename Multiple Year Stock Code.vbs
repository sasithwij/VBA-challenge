Sub Stocks()

'Loop through all worksheets
    For Each ws In Worksheets

'Paste column title for both summary tables
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"


    'Set variable for ticker name
    Dim Ticker As String
    Ticker = 0

    'Set variable for holding total volume
    Dim Vol As Double
    Vol = 0
    
    'Set Variable for Open Price
    Dim OpenPrice As Double
    OpenPrice = 0
    
    'Set Vairable for Close Price
    Dim ClosePrice As Double
    ClosePrice = 0
    
    'Set Vairable for difference between Open and Close
    Dim Delta As Double
    Delta = 0
    
    'Set Vairable for %Change
    Dim PercentChange As Double
    PercentChange = 0
    
    'Set Variables for second summary table
    Dim MaxTicker As String
    MaxTicker = " "
    
    Dim MinTicker As String
    MinTicker = " "
    
    Dim MaxPercent As Double
    MaxPercent = 0
    
    Dim MinPercent As Double
    MinPercent = 0
    
    Dim MaxVolTicker As String
    MaxVolTicker = " "
    
    Dim MaxVol As Double
    MaxVol = 0
    
    'Set Variable to keep track of summary table
    Dim SummaryRow As Long
    SummaryRow = 2

    'Set Variable i As Long
    Dim i As Long
    
    'Set Open Price as first instance of open price in each worksheet. Loop will loop through each open price for each ticker
    OpenPrice = ws.Cells(2, 3).Value
    
        'Loop through all rows in each sheet until last row
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            'Check whether ticker is still the same
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Set Ticker name
                Ticker = ws.Cells(i, 1).Value
                
                'Calculate Delta
                ClosePrice = ws.Cells(i, 6).Value
                Delta = ClosePrice - OpenPrice
                
                'Calculate Pecentage Change. Make sure Open price is not 0
                If OpenPrice <> 0 Then
                    PercentChange = (Delta / OpenPrice) * 100
                Else
                    ws.Cells(SummaryRow, 11).Value = "Open Price = 0"
                End If
                
                'Add to Ticker Volumne
                Vol = Vol + ws.Cells(i, 7).Value
                
                'Print Ticker Name
                ws.Cells(SummaryRow, 9).Value = Ticker
                
                'Print Ticker delta
                ws.Cells(SummaryRow, 10).Value = Delta
                
                'Print percent change
                ws.Cells(SummaryRow, 11).Value = (Str(PercentChange) & "%")
                
                'Print Ticker Volume
                ws.Cells(SummaryRow, 12).Value = Vol
                
                'Conditional Formating on Delta
                If Delta > 0 Then
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
                
                Else
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
                    
                End If
                
                'Add 1 to summary table count
                SummaryRow = SummaryRow + 1
                
                'Reset Delta
                Delta = 0
                
                'Reset Close Price
                ClosePirce = 0
                
                'Net Ticker's Open Price
                OpenPrice = ws.Cells(i + 1, 3).Value
                
                'Find Ticker Name and Value for Greatest Increse
                If PercentChange > MaxPercent Then
                    MaxPercent = PercentChange
                    MaxTicker = Ticker
                    
                ElseIf PercentChange < MinPercent Then
                    MinPercent = PercentChange
                    MinTicker = Ticker
                    
                End If
                
                If Vol > MaxVol Then
                    MaxVol = Vol
                    MaxVolTicker = Ticker
                    
                End If
                
                'Reset PercentageChange and Ticker Volumn
                Vol = 0
                PercentChange = 0
                
            'If still within the same ticker add volume to volume total
            Else
                Vol = Vol + ws.Cells(i, 7).Value
                
            End If
        
        Next i
        
        'Paste values into 2nd table
            
            ws.Range("P2").Value = MaxTicker
            ws.Range("Q2").Value = (Str(MaxPercent) & "%")
            ws.Range("P3").Value = MaxTicker
            ws.Range("Q3").Value = (Str(MinPercent) & "%")
            ws.Range("P4").Value = MaxVolTicker
            ws.Range("Q4").Value = MaxVol
        
    Next ws

End Sub

