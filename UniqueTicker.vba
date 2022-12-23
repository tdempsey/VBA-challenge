Sub UniqueTicker()

    'populated headers and set widths
    Cells(1, 9) = "Ticker"
    Cells(1, 9).EntireColumn.AutoFit
    Cells(1, 10) = "Yearly Change"
    Cells(1, 10).EntireColumn.AutoFit
    Cells(1, 11) = "Percent Change"
    Cells(1, 11).EntireColumn.AutoFit
    Cells(1, 12) = "Total Stock Volume"
    Cells(1, 12).EntireColumn.AutoFit
    
    'find the last row of column A
    Dim row_a_last As Long
    row_a_last = Cells(Rows.Count, "A").End(xlUp).row
    
    Dim ticker As String: ticker = Cells(2, 1)
    Dim yearly_change As Double: yearly_change = 0#
    Dim percent_change As Double: percent_change = 0#
    Dim total_volume As Double: total_volume = 0#
    Dim opening_price_year As Double: opening_price_year = Cells(2, 3)
    Dim total_year As Double: total_year = 0#
    Dim y As Long: y = 2
    
    For x = 2 To row_a_last
    'For x = 2 To 500
        If ticker <> Cells(x, 1) Then
            'display cells in columns I-L
            Cells(y, 9).Value = ticker
            total_year = total_year + (Cells(x - 1, 6).Value - opening_price_year)
            Cells(y, 10).Value = Cells(x - 1, 6).Value - opening_price_year
            
            If Cells(y, 10).Value >= 0 Then
                Cells(y, 10).Interior.ColorIndex = 4
            Else
                Cells(y, 10).Interior.ColorIndex = 3
            End If
            
            Cells(y, 11).Value = Cells(y, 10).Value / opening_price_year
            Cells(y, 11).Style = "Percent"
            Cells(y, 12).Value = total_volume
            
            y = y + 1
            
            ticker = Cells(x, 1)
            yearly_change = 0
            percent_change = 0
            total_volume = 0
            opening_price_year = Cells(x, 3)
        Else
            total_volume = total_volume + Cells(x, 7).Value
        End If
    Next x
    
    'display last ticker in columns I-L
    Cells(y, 9).Value = ticker
    total_year = total_year + (Cells(x - 1, 6).Value - opening_price_year)
    Cells(y, 10).Value = Cells(x - 1, 6).Value - opening_price_year
    
    If Cells(y, 10).Value >= 0 Then
        Cells(y, 10).Interior.ColorIndex = 4
    Else
        Cells(y, 10).Interior.ColorIndex = 3
    End If
            
    Cells(y, 11).Value = Cells(y, 10).Value / opening_price_year
    Cells(y, 11).Style = "Percent"
    Cells(y, 12).Value = total_volume
    
    'populated headers and set widths
    Cells(1, 16) = "Ticker"
    Cells(1, 16).EntireColumn.AutoFit
    Cells(1, 17) = "Value"
    Cells(1, 17).EntireColumn.AutoFit
    Cells(2, 15) = "Greatest % Increase"
    Cells(3, 15) = "Greatest % Decrease"
    Cells(4, 15) = "Greatest Total Volume"
    Cells(4, 15).EntireColumn.AutoFit
    
    'find the last row of column I
    Dim row_i_last As Long
    row_i_last = Cells(Rows.Count, "I").End(xlUp).row
    
    Dim greatest_max As Double: greatest_max = 0
    Dim greatest_min As Double: greatest_min = 0
    Dim greatest_vol As Double: greatest_vol = 0
    
    For Z = 2 To row_i_last
        If Cells(Z, 11).Value > greatest_max Then
            Cells(2, 16).Value = Cells(Z, 9).Value
            Cells(2, 17).Value = Cells(Z, 11).Value
            Cells(2, 17).Style = "Percent"
            greatest_max = Cells(Z, 11).Value
        End If
        
        If Cells(Z, 11).Value < greatest_min Then
            Cells(3, 16).Value = Cells(Z, 9).Value
            Cells(3, 17).Value = Cells(Z, 11).Value
            Cells(3, 17).Style = "Percent"
            greatest_min = Cells(Z, 11).Value
        End If
        
        If Cells(Z, 12).Value > greatest_vol Then
            Cells(4, 16).Value = Cells(Z, 9).Value
            Cells(4, 17).Value = Cells(Z, 12).Value
            greatest_vol = Cells(Z, 12).Value
            Cells(1, 17).EntireColumn.AutoFit
        End If
    Next Z

End Sub



