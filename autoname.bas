Attribute VB_Name = "Module2"
Sub ��������������()
Dim pge As Integer
pge = ThisWorkbook.Worksheets(35).Cells(8, 4) + 3
For i = 3 To 202
If IsEmpty(ThisWorkbook.Worksheets(pge).Cells(i, 3)) = False Then
ThisWorkbook.Worksheets(pge).Cells(i, 7) = "���������"
ThisWorkbook.Worksheets(pge).Cells(i, 10) = "���������"
End If
If IsEmpty(ThisWorkbook.Worksheets(pge).Cells(i, 16)) = False Then
ThisWorkbook.Worksheets(pge).Cells(i, 20) = "���������"
ThisWorkbook.Worksheets(pge).Cells(i, 23) = "���������"
End If
Next i
End Sub

