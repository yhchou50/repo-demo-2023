Sub run_modules_in_ws()
Dim ws As Worksheet

For Each ws In Worksheets
ws.Activate
Call Module1.ticker
Call Module2.yearly_change
Call Module3.great_change


Next ws



End sub