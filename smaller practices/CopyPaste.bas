Attribute VB_Name = "CopyPaste"
Option Explicit
Sub CopyPaste()

Dim lastrow As Long
Dim copysht As Worksheet
Dim destsht As Worksheet
Dim rng As Range

lastrow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row


Set copysht = Worksheets("Source")
Set destsht = Worksheets("Random")

copysht.Select

Set rng = copysht.Range(Range("E1:F1"), Range("E" & lastrow, "F" & lastrow))


rng.Copy


destsht.Range("E1").PasteSpecial Paste:=xlPasteValues
Application.CutCopyMode = False



End Sub


