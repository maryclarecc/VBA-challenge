Sub stocks():

    For Each ws In Worksheets
        'Declare variables
        Dim Ticker As String
        Ticker = 0
        Dim YrlyChg As Single
        Dim PctChg As Single
        'PctChg = 0
        Dim StockVol As Single
        StockVol = 0
        Dim sumtable As Integer
        sumtable = 2
        Dim yropen As Double
        Dim yrclose As Double
        Dim openpr As Double
        openpr = 2 'start out in row 2
        
        'Find the last row in the sheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'Loop through to find breakpoints
        For i = 2 To LastRow
            
            'if cell 3 does not equal 2 (meaning a new value) then
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'The ticker symbol & total stock volume
                ws.Range("I1").Value = "Ticker" 'give column heading
                Ticker = ws.Cells(i, 1).Value 'set tickers as column 1
            
                ws.Range("L1").Value = "Total Stock Volume"
                StockVol = StockVol + ws.Cells(i, 7).Value

                ws.Range("I" & sumtable).Value = Ticker
                ws.Range("L" & sumtable).Value = StockVol
               
                'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
                
                
                ws.Range("J1").Value = "Yearly Change"
        
                yropen = ws.Cells(openpr, 3).Value 'row 2 start, gets incremented later
                yrclose = ws.Cells(i, 6).Value
                
                YrlyChg = yrclose - yropen
                ws.Range("J" & sumtable).NumberFormat = "0.00"
                
                 'ws.Range("J" & sumtable).Value = YrlyChg
                'for j in 2 to LastRow
                '    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                '        Set MyRange = Range("a" & i)
                '            LastRow_1 = MyRange.Row + MyRange.Rows.Count - 1
                '            firstRow = MyRange.Row
                '    End if
                'Next j
                
                
                ws.Range("J" & sumtable).Value = YrlyChg
                
                'The percent change from opening price at the beginning of a given year to the closing price at the end of that year
                
                ws.Range("K1").Value = "Percent Change"
                
                ws.Range("K" & sumtable - 1).Value = PctChg
                'i don't know why pctchg is offset by one row
                PctChg = (YrlyChg / (yropen + 0.00000001)) 'gives an overflow error & divide by zero
                ws.Range("K" & sumtable).NumberFormat = "0.00%"
                
                
                'conditional formatting
                If ws.Range("J" & sumtable).Value > 0 Then
                    ws.Range("J" & sumtable).Interior.ColorIndex = 4
                        Else
                            ws.Range("J" & sumtable).Interior.ColorIndex = 3
                End If
                
                'move on and reset totalers
                sumtable = sumtable + 1
                StockVol = 0
                openpr = i + 1 ' increments the open price for future loops
                
               

            Else
                'Add to the volume
                StockVol = StockVol + ws.Cells(i, 7).Value
                
            End If
        Next i

    'Range("K2").Delete Shift:=xlToUp
                
    Next ws
 
End Sub

