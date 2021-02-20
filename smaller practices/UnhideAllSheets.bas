Attribute VB_Name = "UnhideAllSheets"
Option Explicit
Sub UnhideAllSheets()
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Visible = xlSheetVisible
Next ws
End Sub
