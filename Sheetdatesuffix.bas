Option Explicit
Sub Sheetdatesuffix()

Dim todaysdate As Date
Dim wsh As Worksheet

todaysdate = Date

For Each wsh In ActiveWorkbook.Worksheets
wsh.Visible = xlSheetVisible
wsh.Name = wsh.Name & "_" & todaysdate
Next wsh

End Sub
