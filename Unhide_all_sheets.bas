Attribute VB_Name = "Unhide_all"
Sub UnhideAllSheets()
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Visible = xlSheetVisible
Next ws
End Sub
