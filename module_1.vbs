Sub ticker()
Dim total As Double
total = 0
Dim summary_table_row As Integer
summary_table_row = 2

For i = 2 To 753001
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
Name = Cells(i, 1)
total = total + Cells(i, 7)


Range("K" & summary_table_row).Value = Name
Range("n" & summary_table_row).Value = total

summary_table_row = summary_table_row + 1
total = 0

Else
total = total + Cells(i, 7).Value

End If

Next i

End Sub

