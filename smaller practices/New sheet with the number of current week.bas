Attribute VB_Name = "Module2"
Option Explicit
Sub ThisWeeksSheet()

Dim weeknum As Integer

weeknum = Format(Now(), "WW")

Worksheets.Add().Name = weeknum

End Sub
