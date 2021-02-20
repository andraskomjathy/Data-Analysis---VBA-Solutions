Option Explicit
Sub create_chart()

Dim wb As Workbook
Dim sht As Worksheet
Dim rng As Range
Dim cht As Object



Set wb = ThisWorkbook
Set sht = wb.Sheets("Source")
Set rng = ActiveSheet.Range("A2:A20")
Set cht = ActiveSheet.Shapes.AddChart2
cht.Chart.SetSourceData Source:=rng
cht.Chart.HasTitle = True
cht.Chart.ChartTitle.Text = "Sales"

End Sub
