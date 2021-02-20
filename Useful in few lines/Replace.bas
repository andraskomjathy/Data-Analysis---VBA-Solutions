Option Explicit
Sub Replace()

Dim rng As Range
Dim lastrow As Long
Dim sht As Worksheet

Set sht = Worksheets("News")
lastrow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row

Set rng = Range(Cells(2, 1), Cells(lastrow, 1))


Cells.Replace What:="e", Replacement:="A"


End Sub
