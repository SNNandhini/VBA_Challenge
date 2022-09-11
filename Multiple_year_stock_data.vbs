Sub stockanalysis()

    For Each ws In Worksheets
        
        'Headers
         ws.Range("I1").Value = "Ticker"
         ws.Range("J1").Value = "Yearly Change"
         ws.Range("K1").Value = "Percent Change"
         ws.Range("L1").Value = "Total Stock Volume"
         
         ws.Range("O2").Value = "Greatest % Increase"
         ws.Range("O3").Value = "Greatest % Decrease"
         ws.Range("O4").Value = "Greatest Total Volume"
         ws.Range("P1").Value = "Ticker"
         ws.Range("Q1").Value = "Value"
         
        'Variable to calculate the Total Stock Volume
         Dim TotalStockVolume As Double
        
        'Variables to calculate the Yearly Change
         Dim yearly_change, opening_row, opening_price, closing_price As Double
         
        'Variable to identify the row where the results are to be printed for each Ticker
         Dim resultrow As Integer
         
        'Variables to calculate Greatest % Increase, Greatest % Decrease, Greatest Total Volume
         Dim IncVal, DecVal, TotVal As Double
         Dim IncTicker, DecTicker, TotTicker As String
         
        'Initialize the variables for Total Stock Volume and Yearly Change Calculations
         TotalStockVolume = 0
         yearly_change = 0
         opening_price = 0
         closing_price = 0
         
        'Initialize the variable used to identify the row of opening price on the first day
         opening_row = 2
         
        'Initialize the variable used for row while printing the results
         resultrow = 2
         
        'To identify the number of rows in the given data
         lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
         
        'calculation (Column I - Ticker, Column J - Yearly Change, Column K - Percent Change, Column L - Total Stock Volume)
         For i = 2 To lastrow
             If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
                 TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
             Else
                'Column I
                 ws.Cells(resultrow, 9).Value = ws.Cells(i, 1).Value
                 
                'Column J
                 opening_price = ws.Cells(opening_row, 3).Value
                 closing_price = ws.Cells(i, 6).Value
                 yearly_change = closing_price - opening_price
                 ws.Cells(resultrow, 10).Value = yearly_change
                 ws.Cells(resultrow, 10).NumberFormat = "###0.00"
                 If yearly_change < 0 Then
                     ws.Cells(resultrow, 10).Interior.ColorIndex = 3
                 Else
                     ws.Cells(resultrow, 10).Interior.ColorIndex = 4
                 End If
                 
                'Column K
                 ws.Cells(resultrow, 11).Value = FormatPercent(yearly_change / opening_price, 2)
                             
                'Column L
                 TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                 ws.Cells(resultrow, 12).Value = TotalStockVolume
                 
                 yearly_change = 0
                 opening_price = 0
                 closing_price = 0
                 TotalStockVolume = 0
                 resultrow = resultrow + 1
                 opening_row = i + 1
                 
             End If
         Next i
         
        'To identify the number of rows in the summary table (Column I)
         smry_lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
         
        'Initialize the variables with the first values from the summary table (Columns K and L)
         IncVal = ws.Cells(2, 11).Value
         DecVal = ws.Cells(2, 11).Value
         TotVal = ws.Cells(2, 12).Value
             
        'Compare the values from Columns K and L with the values above to identify the Greatest % Increase, Greatest % Decrease, Greatest Total Volume
         For j = 2 To smry_lastrow
         
            'Greatest % Increase
             If ws.Cells(j, 11) > IncVal Then
                 IncVal = ws.Cells(j, 11).Value
                 IncTicker = ws.Cells(j, 9).Value
             End If
             
            'Greatest % Decrease
             If ws.Cells(j, 11) < DecVal Then
                 DecVal = ws.Cells(j, 11).Value
                 DecTicker = ws.Cells(j, 9).Value
             End If
             
            'Greatest Total Volume
             If ws.Cells(j, 12) > TotVal Then
                 TotVal = ws.Cells(j, 12).Value
                 TotTicker = ws.Cells(j, 9).Value
             End If
         
         Next j
             
         ws.Range("P2").Value = IncTicker
         ws.Range("Q2").Value = FormatPercent(IncVal, 2)
         
         ws.Range("P3").Value = DecTicker
         ws.Range("Q3").Value = FormatPercent(DecVal, 2)
             
         ws.Range("P4").Value = TotTicker
         ws.Range("Q4").Value = TotVal
         ws.Range("Q4").NumberFormat = "##0.00E+00"
         
        'Autofit the cells that have data
         ws.UsedRange.EntireColumn.AutoFit
         ws.UsedRange.EntireRow.AutoFit
    
    Next ws
    
End Sub
    
