Sub WallStreetStock():

Dim open1 As Integer
open1 = Cells(2, 3).Value
myindex = 2
total_stock_volume = 0
For i = 2 To 10000
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        'count it as a new value to be input to column J'

        close1 = Cells(i, 6).Value
        Cells(myindex, 9).Value = Cells(i, 1).Value
        Cells(myindex, 10).Value = close1 - open1
        Cells(myindex, 11).Value = Format((((close1 - open1) / open1)), "Percent")
        open1 = Cells(i + 1, 3)

        total_stock_volume = total_stock_volume + Cells(i, 7).Value

        'If Cells(myindex, 11).Value > 0 Then Cells(myindex, 11).Value.Interior.ColorIndex ' = 4
           'ElseIf Cells(myindex, 11).Value < 0 Then Cells(myindex, 11).Interior.ColorIndex = 3
            'Else: Cells(myindex, 11).Interior.ColorIndex = 0
            'End If'
            'i keeping getting error message after running the color format code 
            'i commented the code out because i couldnt fix it 

        Cells(myindex, 12).Value = total_stock_volume
        total_stock_volume = 0
        myindex = myindex + 1

    Else
        total_stock_volume = total_stock_volume + Cells(i, 7).value
    End If
    
End Sub