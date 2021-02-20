Attribute VB_Name = "formatchange"
Option Explicit
Sub formatchange()

Dim lastrow As Long
Dim wb As Workbook
Dim sht As Worksheet
Dim rng As Range

lastrow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row

Set wb = ThisWorkbook
Set sht = wb.Sheets("Random")
sht.Select
Set rng = sht.Range(Range("E1:F1"), Range("E" & lastrow, "F" & lastrow))

With rng
    .NumberFormat = "General"
    .Value = .Value
End With

End Sub
