Sub Year_Open_Price()

Dim open_price as double
Dim close_price as double 
Dim brand_index as Integer
open_price = 0
close_price = 0
brand_index = 2

For i = 1 to 7060192

    If Cells(i,1).Value <> Cells(i+1,1).Value Then
        open_price = Cells(i+1,3).Value
        Cells(brand_index,13).value = open_price
        
        open_price = 0
        brand_index = brand_index + 1
    End If

Next i

End Sub