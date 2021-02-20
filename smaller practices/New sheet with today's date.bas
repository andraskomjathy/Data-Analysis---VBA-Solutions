Attribute VB_Name = "TodaysSheet"
Option Explicit
Sub TodaysSheet()

Dim datetoday As String

datetoday = Date

Worksheets.Add().Name = datetoday

End Sub
