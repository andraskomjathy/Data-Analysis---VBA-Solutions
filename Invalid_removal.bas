Option Explicit
Sub Invalid_removal()

Dim wb As Workbook
Dim sht As Worksheet
Dim lastrow As Long

'We have to define long once we have
'expectedly more than 32767 rows, because
'integer cannot handle them afterwards.

Dim i As Long



Set wb = ThisWorkbook
Set sht = wb.Sheets("Transportation")

'It is advised to avoid select commands for many reasons,
'however it might be useful when the file is usually being
'saved at another sheet.


sht.Select

'Defining the number of last used cell.

lastrow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row

'It is advised to loop from the last row upwards, because
'row deletion might mix rownumbers and thus ruin the loop.

For i = lastrow To 2 Step -1
    If Len(Cells(i, 6)) <> 8 Then
    Cells(i, 6).EntireRow.Delete
    End If
Next


End Sub
