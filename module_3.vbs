Sub great_change()

Dim max_increase As Double
Dim max_decrease As Double
Dim great_volume As Double

Dim max_increase_row As Integer
Dim max_decrease_row As Integer
Dim great_volume_row As Integer

max_increase = 0
max_decrease = Cells(2, 13).Value
great_volume = 0

For i = 2 To 753001
If Cells(i, 13).Value > max_increase Then
max_increase = Cells(i, 13).Value
max_increase_row = i
End If
Next i

For j = 2 To 753001
If Cells(j, 13).Value < max_decrease Then
max_decrease = Cells(j, 13).Value
max_decrease_row = j
End If
Next j

For k = 2 To 753001
If Cells(k, 14).Value > great_volume Then
great_volume = Cells(k, 14).Value
great_volume_row = k
End If
Next k


Range("Q2").Value = Cells(max_increase_row, 11).Value
Range("Q3").Value = Cells(max_decrease_row, 11).Value
Range("q4").Value = Cells(great_volume_row, 11).Value

Range("r2").Value = Cells(max_increase_row, 13).Value
Range("r3").Value = Cells(max_decrease_row, 13).Value
Range("r4").Value = Cells(great_volume_row, 13).Value



End Sub
