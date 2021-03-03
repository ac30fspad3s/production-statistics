Attribute VB_Name = "Module2"
Sub автозаполнение()
Dim pge As Integer
pge = ThisWorkbook.Worksheets(35).Cells(8, 4) + 3
For i = 3 To 202
If IsEmpty(ThisWorkbook.Worksheets(pge).Cells(i, 3)) = False Then
ThisWorkbook.Worksheets(pge).Cells(i, 7) = "Чилонбоев"
ThisWorkbook.Worksheets(pge).Cells(i, 10) = "Мамажонов"
End If
If IsEmpty(ThisWorkbook.Worksheets(pge).Cells(i, 16)) = False Then
ThisWorkbook.Worksheets(pge).Cells(i, 20) = "Чилонбоев"
ThisWorkbook.Worksheets(pge).Cells(i, 23) = "Мамажонов"
End If
Next i
End Sub

