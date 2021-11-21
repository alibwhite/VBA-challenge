Sub stock_data()

Dim ticker as String
Dim yearly_change as Double
yearly_change = 0
Dim year_open as Double
year_open = 0
Dim year_close as Double
year_close = 0
Dim percent_change as Double
percent_change = 0
Dim total_stock_volume as Double
total_stock_volume = 0
Dim summary_table_row as Integer
summary_table_row = 2

    For i = 2 to 797711

        If Cells(i-1,1).Value <> Cells(i,1).Value and Cells(i,3).Value <>0 Then
            year_open = Cells(i,3).Value
        ElseIf Cells(i-1,3).Value = 0 and Cells(i,3).Value <> 0 then 
            year_open = Cells(i,3).Value
        ElseIf Cells(i+1,1).Value <> Cells(i,1).Value Then
            ticker = Cells(i,1).Value
            year_close = Cells(i,6).Value 
            total_stock_volume = total_stock_volume + Cells(i,7).Value
            yearly_change = year_close - year_open
            Range("I" & summary_table_row).Value = ticker
            Range("J" & summary_table_row).Value = yearly_change
            Range("K" & summary_table_row).Value = yearly_change/year_open
            Range("L" & summary_table_row).Value = total_stock_volume
            summary_table_row = summary_table_row + 1
            total_stock_volume = 0
        Else
            total_stock_volume = total_stock_volume + Cells(i,7).Value
        End if

    Next i 

Dim greatest_increase as Double
Dim greatest_decrease as Double
Dim greatest_stock_volume as Double
Dim percent_rng as Range
Dim volume_rng as Range

    Set percent_rng = Range("K2:K3169")
    Set volume_rng = Range("L2:L3169")

    greatest_increase = Application.WorksheetFunction.Max(percent_rng)
    greatest_decrease = Application.WorksheetFunction.Min(percent_rng)
    greatest_stock_volume = Application.WorksheetFunction.Max(volume_rng)

    Range("P2").Value = greatest_increase
    Range("P3").Value = greatest_decrease
    Range("P4").Value = greatest_stock_volume

End Sub

