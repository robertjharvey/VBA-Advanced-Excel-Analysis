Sub Ticker_Symbol()

Dim vol_total as longlong
Dim brand_index as Integer
vol_total = 0
brand_index = 2

For i = 2 to 798000

    If Cells(i,1).Value = Cells(i+1,1).Value Then
        vol_total = Cells(i,7).Value + vol_total

    Else
        vol_total = Cells(i,7).value + vol_total
        Cells(brand_index,9).value = Cells(i,1).value
        Cells(brand_index,12).value = vol_total
        vol_total = 0
        brand_index = brand_index + 1
    End If

Next i

End Sub
