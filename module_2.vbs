Sub yearly_change()

Dim first_opening As Double
Dim last_closing As Double
Dim yearly_change As Double
Dim percent_change As Double

Dim first_row As Double
first_row = 2
Dim summary_table_row As Integer
summary_table_row = 2

For i = 2 To 753001
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
last_closing = Cells(i, 6).Value
first_opening = Cells(first_row, 3).Value
yearly_change = last_closing - first_opening
percent_change = yearly_change / first_opening


Range("L" & summary_table_row).Value = yearly_change

Range("M" & summary_table_row).Value = percent_change

summary_table_row = summary_table_row + 1
first_row = i + 1

yearly_change = 0

End If

If Cells(summary_table_row, 12).Value > 0 Then
Cells(summary_table_row, 12).Interior.ColorIndex = 4
ElseIf Cells(summary_table_row, 12).Value < 0 Then
Cells(summary_table_row, 12).Interior.ColorIndex = 3
End If

Next i
End Sub

