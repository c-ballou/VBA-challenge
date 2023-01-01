Sub vba()

Dim ticker As String
Dim yearly_change As Double
Dim open_price As Double
Dim close_price As Double
Dim percent_change As Double
Dim total_volume As Double
Dim summary_table_row As Integer


yearly_change = 0
open_price = 0
close_price = 0
percent_change = 0
total_volume = 0
summary_table_row = 2


Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"

row_count = Cells(Rows.Count, 1).End(xlUp).Row

For x = 2 To row_count

    If Cells(x + 1, 1).Value <> Cells(x, 1).Value Then
        
        ticker = Cells(x, 1).Value
        open_price = Cells(x - 250, 3).Value
        close_price = Cells(x, 6).Value
        yearly_change = close_price - open_price
        percent_change = (yearly_change / open_price)
        total_volume = total_volume + Cells(x, 7).Value
        
        
        Range("I" & summary_table_row).Value = ticker
        Range("J" & summary_table_row).Value = yearly_change
        Range("K" & summary_table_row).Value = percent_change
        Range("L" & summary_table_row).Value = total_volume
        
        
        Range("K" & summary_table_row).NumberFormat = "0.00%"
        
        
        If yearly_change > 0 Then
            Range("J" & summary_table_row).Interior.ColorIndex = 4
        ElseIf yearly_change = 0 Then
            Range("J" & summary_table_row).Interior.ColorIndex = 0
        Else
            Range("J" & summary_table_row).Interior.ColorIndex = 3
        End If
        
        
        summary_table_row = summary_table_row + 1
    
        total_volume = 0
        
        open_price = Cells(x + 1, 3).Value
        
    Else
        total_volume = total_volume + Cells(x, 7).Value
        
    End If
    
Next x


End Sub
