Option Explicit
Sub ThisWeeksSheet()

Dim weeknum As Integer

weeknum = Format(Now(), "WW")

Worksheets.Add().Name = weeknum

End Sub
