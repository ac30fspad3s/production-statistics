Sub автозаполнение()
Dim pge As Integer
pge = ThisWorkbook.Worksheets(35).Cells(8, 4) + 3
For i = 3 To 202
If IsEmpty(ThisWorkbook.Worksheets(pge).Cells(i, 3)) = False Then
ThisWorkbook.Worksheets(pge).Cells(i, 7) = ThisWorkbook.Worksheets(35).Cells(9, 4)
ThisWorkbook.Worksheets(pge).Cells(i, 10) = ThisWorkbook.Worksheets(35).Cells(10, 4)
End If
If IsEmpty(ThisWorkbook.Worksheets(pge).Cells(i, 16)) = False Then
ThisWorkbook.Worksheets(pge).Cells(i, 20) = ThisWorkbook.Worksheets(35).Cells(9, 4)
ThisWorkbook.Worksheets(pge).Cells(i, 23) = ThisWorkbook.Worksheets(35).Cells(10, 4)
End If
Next i
End Sub

