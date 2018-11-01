Sub matrix_evaluation()
Dim x As Integer
Dim y As Integer
For x = 2 To (Cells(Rows.Count, 1).End(xlUp).Row)
    For y = 2 To (Cells(Rows.Count, 1).End(xlUp).Row)
Rem Of course, this works if we only have our correlation matrix on our sheet, otherwise we may replace "(Cells(Rows.Count, 1).End(xlUp).Row)" with the number of
Rem rows and columns of our matrix. The aformentioned solution might be better, since it adapts to the size of the matrix.
Rem Since a matrix has equal number of rows and columns, we can use the number of rows for the case of columns, as well, because the number of columns will change later.
        If Abs(Cells(x, y).Value) > 0.3 And Abs(Cells(x, y).Value) < 0.5 Then
            Cells(x, y).Interior.Color = vbYellow
        ElseIf Abs(Cells(x, y).Value) > 0.5 And Abs(Cells(x, y).Value) < 0.7 Then
            Cells(x, y).Interior.Color = vbGreen
        ElseIf Abs(Cells(x, y).Value) > 0.7 And Abs(Cells(x, y).Value) < 1 Then
            Cells(x, y).Interior.Color = vbRed
        Else
            Cells(x, y).Interior.Color = vbWhite
        End If
Rem Colors will show the strength of the correlation values in the matrix: yellow - weak, green - moderate and red - strong.
    Next y
Next x
Range("M1").Value = "Remarkable correlations"
Range("N1").Value = "Direction"
Range("O1").Value = "Strength"
Range("P1").Value = "Correlation pair 1"
Range("Q1").Value = "Correlation pair 2"
Dim i As Integer
i = 2
For x = 2 To (Cells(Rows.Count, 1).End(xlUp).Row)
    For y = 2 To (Cells(Rows.Count, 1).End(xlUp).Row)
        If Abs(Cells(x, y).Value) > 0.5 And Abs(Cells(x, y).Value) < 1 Then
            Cells(i, 13).Value = Abs(Cells(x, y).Value)
            i = i + 1
        End If
        If Abs(Cells(x, y).Value) > 0.5 And (Cells(x, y).Value) > 0 And Abs(Cells(x, y).Value) < 1 Then
            Cells((i - 1), 14).Value = "Positive"
        ElseIf Abs(Cells(x, y).Value) > 0.5 And (Cells(x, y).Value) < 0 And Abs(Cells(x, y).Value) < 1 Then
        Cells((i - 1), 14).Value = "Negative"
        End If
        If Abs(Cells(x, y).Value) > 0.5 And Abs(Cells(x, y).Value) < 1 Then
            Cells((i - 1), 16).Value = x
            Cells((i - 1), 17).Value = y
        End If
        If Abs(Cells(x, y).Value) > 0.5 And Abs(Cells(x, y).Value) < 0.7 Then
            Cells((i - 1), 15).Value = "Moderate"
        ElseIf Abs(Cells(x, y).Value) > 0.7 And Abs(Cells(x, y).Value) < 1 Then
            Cells((i - 1), 15).Value = "Strong"
        End If
Rem Listing out the correlation-pairs above 0.5 according to strength, evaluating them by direction.
    Next y
Next x
Range("M2", Range("Q2").End(xlDown)).Sort Key1:=Range("M2"), Order1:=xlDescending, Header:=xlNo
Range("P2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
        Selection.Replace What:="11", Replacement:="TUR"
        Selection.Replace What:="10", Replacement:="SLO"
        Selection.Replace What:="2", Replacement:="ALB"
        Selection.Replace What:="3", Replacement:="BEL"
        Selection.Replace What:="4", Replacement:="BUL"
        Selection.Replace What:="5", Replacement:="DEN"
        Selection.Replace What:="6", Replacement:="HUN"
        Selection.Replace What:="7", Replacement:="ROM"
        Selection.Replace What:="8", Replacement:="RUS"
        Selection.Replace What:="9", Replacement:="SVK"
Rem Printing out the country codes between the correlations exist.
        
End Sub
