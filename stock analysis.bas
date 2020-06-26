Attribute VB_Name = "Module1"
Sub stock_market_analysis():
 'define variables
    Dim ticker_symbol As String
    Dim opening_price As Double
    Dim price_Change As Double
    Dim closing_price As Double
    Dim percent_change As Double
    Dim total_volume As Double
        total_volume = 0
    Dim i As Long
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    Range("J1") = "Ticker"
    Range("K1") = "Opening price"
    Range("L1") = "Closing price"
    Range("M1") = "Price change"
    Range("N1") = "percent change"
    Range("O1") = "total volume"
                
'Find and store opening, closing and ticker values
    opening_price = Cells(2, 3).Value
    ticker_symbol = Cells(2, 1).Value
    
    For i = 2 To Range("A1").CurrentRegion.End(xlDown).Row
        closing_price = Cells(i, 6).Value
        'Calculate total stock volume
        total_volume = total_volume + Cells(i, 7).Value
     
     'If ticker changes then print results
                
        If ticker_symbol <> Cells(i + 1, 1).Value Then
            'calculate Price_change
            price_Change = closing_price - opening_price
        If opening_price <> 0 Then
            'Calculate percent_change
            percent_change = price_Change / opening_price * 100
        Else
            percent_change = 100
        End If
            Range("J" & Summary_Table_Row).Value = ticker_symbol
            Range("K" & Summary_Table_Row).Value = opening_price
            Range("L" & Summary_Table_Row).Value = closing_price
            Range("M" & Summary_Table_Row).Value = price_Change
                'set formats
                If price_Change >= 0 Then
                    Range("M" & Summary_Table_Row).Interior.ColorIndex = 4
                Else
                    Range("M" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
            Range("N" & Summary_Table_Row).Value = percent_change
            Range("O" & Summary_Table_Row).Value = total_volume
            Summary_Table_Row = Summary_Table_Row + 1
            opening_price = Cells(i + 1, 3).Value
            ticker_symbol = Cells(i + 1, 1).Value
            total_volume = 0
        End If
          
    Next i
    
    Columns("J:O").Select
    Selection.EntireColumn.AutoFit
    
End Sub


