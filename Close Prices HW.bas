Sub Year_End_Price()

Dim close_price as double 
Dim brand_index as Integer
close_price = 0
brand_index = 2

For i = 2 to 760192

    If Cells(i,1).Value <> Cells(i+1,1).Value Then
        close_price = Cells(i,6).Value
        Cells(brand_index,14).value = close_price
        
        close_price = 0
        brand_index = brand_index + 1
    End If

Next i

End Sub