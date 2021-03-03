Attribute VB_Name = "Module1"
Function quality(line As Integer, coloumn As Integer, sht As Integer)

If IsEmpty(ThisWorkbook.Worksheets(sht).Cells(line, coloumn)) = True Then
Exit Function
End If

If InStr(ThisWorkbook.Worksheets(sht).Cells(line, coloumn), "ç") Then
ThisWorkbook.Worksheets(sht).Cells(line, coloumn + 2) = -1
ElseIf InStr(ThisWorkbook.Worksheets(sht).Cells(line, coloumn), "ã") Or InStr(ThisWorkbook.Worksheets(sht).Cells(line, coloumn), "ò") Or InStr(ThisWorkbook.Worksheets(sht).Cells(line, coloumn), "ï") Then
ThisWorkbook.Worksheets(sht).Cells(line, coloumn + 2) = 0
Else: ThisWorkbook.Worksheets(sht).Cells(line, coloumn + 2) = 1
End If



If InStr(ThisWorkbook.Worksheets(sht).Cells(line, coloumn), "ê") Or InStr(ThisWorkbook.Worksheets(sht).Cells(line, coloumn), "í") Then
ThisWorkbook.Worksheets(sht).Cells(line, coloumn + 5) = -1
Else: ThisWorkbook.Worksheets(sht).Cells(line, coloumn + 5) = 1
End If

If InStr(ThisWorkbook.Worksheets(sht).Cells(line, coloumn), "í") Then
ThisWorkbook.Worksheets(sht).Cells(line, coloumn + 8) = -1
ElseIf InStr(ThisWorkbook.Worksheets(sht).Cells(line, coloumn), "ê") Then
ThisWorkbook.Worksheets(sht).Cells(line, coloumn + 8) = 0
Else: ThisWorkbook.Worksheets(sht).Cells(line, coloumn + 8) = 1
End If

If InStr(ThisWorkbook.Worksheets(sht).Cells(line, coloumn), "ý") Then
ThisWorkbook.Worksheets(sht).Cells(line, coloumn + 2) = 1
ThisWorkbook.Worksheets(sht).Cells(line, coloumn + 5) = 1
ThisWorkbook.Worksheets(sht).Cells(line, coloumn + 8) = 1
End If

If InStr(ThisWorkbook.Worksheets(sht).Cells(line, coloumn), "?") Then
ThisWorkbook.Worksheets(sht).Cells(line, coloumn + 2) = 0
ThisWorkbook.Worksheets(sht).Cells(line, coloumn + 5) = 0
ThisWorkbook.Worksheets(sht).Cells(line, coloumn + 8) = 0
End If


End Function

Function ZP(sht As Integer)
Dim coloumn As Integer

For i = 3 To 202

If IsEmpty(ThisWorkbook.Worksheets(sht).Cells(i, 2)) = False Then

coloumn = 5
ThisWorkbook.Worksheets(sht).Cells(i, coloumn + 1) = ThisWorkbook.Worksheets(sht).Cells(i, coloumn) * ThisWorkbook.Worksheets(2).Cells(23, 2)

coloumn = 8
ThisWorkbook.Worksheets(sht).Cells(i, coloumn + 1) = ThisWorkbook.Worksheets(sht).Cells(i, coloumn) * ThisWorkbook.Worksheets(2).Cells(24, 2)

coloumn = 11
ThisWorkbook.Worksheets(sht).Cells(i, coloumn + 1) = ThisWorkbook.Worksheets(sht).Cells(i, coloumn) * ThisWorkbook.Worksheets(2).Cells(25, 2)

End If

If IsEmpty(ThisWorkbook.Worksheets(sht).Cells(i, 15)) = False Then

coloumn = 18
ThisWorkbook.Worksheets(sht).Cells(i, coloumn + 1) = ThisWorkbook.Worksheets(sht).Cells(i, coloumn) * ThisWorkbook.Worksheets(2).Cells(23, 3)

coloumn = 21
ThisWorkbook.Worksheets(sht).Cells(i, coloumn + 1) = ThisWorkbook.Worksheets(sht).Cells(i, coloumn) * ThisWorkbook.Worksheets(2).Cells(24, 3)

coloumn = 24
ThisWorkbook.Worksheets(sht).Cells(i, coloumn + 1) = ThisWorkbook.Worksheets(sht).Cells(i, coloumn) * ThisWorkbook.Worksheets(2).Cells(25, 3)
End If

Next i

End Function

Function cash(sht As Integer)
Dim cash_day As Integer
Dim coloumn As Integer
cash_day = 0
For i = 2 To 100

If IsEmpty(ThisWorkbook.Worksheets(2).Cells(i, 1)) = True Then
Exit For
End If

For j = 3 To 202
coloumn = 4
If ThisWorkbook.Worksheets(sht).Cells(j, coloumn) = ThisWorkbook.Worksheets(2).Cells(i, 1) Then
cash_day = cash_day + ThisWorkbook.Worksheets(sht).Cells(j, coloumn + 2)
End If

coloumn = 7
If ThisWorkbook.Worksheets(sht).Cells(j, coloumn) = ThisWorkbook.Worksheets(2).Cells(i, 1) Then
cash_day = cash_day + ThisWorkbook.Worksheets(sht).Cells(j, coloumn + 2)
End If

coloumn = 10
If ThisWorkbook.Worksheets(sht).Cells(j, coloumn) = ThisWorkbook.Worksheets(2).Cells(i, 1) Then
cash_day = cash_day + ThisWorkbook.Worksheets(sht).Cells(j, coloumn + 2)
End If

coloumn = 17
If ThisWorkbook.Worksheets(sht).Cells(j, coloumn) = ThisWorkbook.Worksheets(2).Cells(i, 1) Then
cash_day = cash_day + ThisWorkbook.Worksheets(sht).Cells(j, coloumn + 2)
End If

coloumn = 20
If ThisWorkbook.Worksheets(sht).Cells(j, coloumn) = ThisWorkbook.Worksheets(2).Cells(i, 1) Then
cash_day = cash_day + ThisWorkbook.Worksheets(sht).Cells(j, coloumn + 2)
End If

coloumn = 23
If ThisWorkbook.Worksheets(sht).Cells(j, coloumn) = ThisWorkbook.Worksheets(2).Cells(i, 1) Then
cash_day = cash_day + ThisWorkbook.Worksheets(sht).Cells(j, coloumn + 2)
End If

Next j
ThisWorkbook.Worksheets(2).Cells(i, sht - 2) = cash_day
cash_day = 0
Next i

End Function

Function Tables(table As Integer)

Dim count As Integer, count_s As Integer, count_k As Integer, count_p As Integer, count_z As Integer, count_g As Integer, count_t As Integer, count_n As Integer, count_e As Integer

count = 0
count_s = 0
count_k = 0
count_p = 0
count_z = 0
count_g = 0
count_t = 0
count_n = 0
count_e = 0

Dim coloumn As Integer

For pge = 4 To 34

For i = 3 To 202
coloumn = 2
If ThisWorkbook.Worksheets(pge).Cells(i, coloumn) = table Then
count = count + 1

If InStr(ThisWorkbook.Worksheets(pge).Cells(i, coloumn + 1), "ñ") Then
count_s = count_s + 1
End If
If InStr(ThisWorkbook.Worksheets(pge).Cells(i, coloumn + 1), "ê") Then
count_k = count_k + 1
End If
If InStr(ThisWorkbook.Worksheets(pge).Cells(i, coloumn + 1), "ï") Then
count_p = count_p + 1
End If
If InStr(ThisWorkbook.Worksheets(pge).Cells(i, coloumn + 1), "ç") Then
count_z = count_z + 1
End If
If InStr(ThisWorkbook.Worksheets(pge).Cells(i, coloumn + 1), "ã") Then
count_g = count_g + 1
End If
If InStr(ThisWorkbook.Worksheets(pge).Cells(i, coloumn + 1), "ò") Then
count_t = count_t + 1
End If
If InStr(ThisWorkbook.Worksheets(pge).Cells(i, coloumn + 1), "í") Then
count_n = count_n + 1
End If
If InStr(ThisWorkbook.Worksheets(pge).Cells(i, coloumn + 1), "ý") Then
count_e = count_e + 1
End If
End If



coloumn = 15


If ThisWorkbook.Worksheets(pge).Cells(i, coloumn) = table Then

count = count + 1
If InStr(ThisWorkbook.Worksheets(pge).Cells(i, coloumn + 1), "ñ") Then
count_s = count_s + 1
End If
If InStr(ThisWorkbook.Worksheets(pge).Cells(i, coloumn + 1), "ê") Then
count_k = count_k + 1
End If
If InStr(ThisWorkbook.Worksheets(pge).Cells(i, coloumn + 1), "ï") Then
count_p = count_p + 1
End If
If InStr(ThisWorkbook.Worksheets(pge).Cells(i, coloumn + 1), "ç") Then
count_z = count_z + 1
End If
If InStr(ThisWorkbook.Worksheets(pge).Cells(i, coloumn + 1), "ã") Then
count_g = count_g + 1
End If
If InStr(ThisWorkbook.Worksheets(pge).Cells(i, coloumn + 1), "ò") Then
count_t = count_t + 1
End If
If InStr(ThisWorkbook.Worksheets(pge).Cells(i, coloumn + 1), "í") Then
count_n = count_n + 1
End If
If InStr(ThisWorkbook.Worksheets(pge).Cells(i, coloumn + 1), "ý") Then
count_e = count_e + 1
End If

End If

Next i
Next pge

ThisWorkbook.Worksheets(1).Cells(2, table + 1) = count
ThisWorkbook.Worksheets(1).Cells(4, table + 1) = count_s
ThisWorkbook.Worksheets(1).Cells(5, table + 1) = count_k
ThisWorkbook.Worksheets(1).Cells(6, table + 1) = count_p
ThisWorkbook.Worksheets(1).Cells(7, table + 1) = count_z
ThisWorkbook.Worksheets(1).Cells(8, table + 1) = count_g
ThisWorkbook.Worksheets(1).Cells(9, table + 1) = count_t
ThisWorkbook.Worksheets(1).Cells(10, table + 1) = count_n
ThisWorkbook.Worksheets(1).Cells(11, table + 1) = count_e



End Function

Function forms(form As Integer, ftype As Integer)

Dim count As Integer, count_s As Integer, count_k As Integer, count_p As Integer, count_z As Integer, count_g As Integer, count_t As Integer, count_n As Integer, count_e As Integer

count = 0
count_s = 0
count_k = 0
count_p = 0
count_z = 0
count_g = 0
count_t = 0
count_n = 0
count_e = 0

If ftype = 1210 Then
For sht = 4 To 34
If IsEmpty(ThisWorkbook.Worksheets(sht).Cells(form + 2, 3)) = False Then
count = count + 1

If InStr(ThisWorkbook.Worksheets(sht).Cells(form + 2, 3), "ñ") Then
count_s = count_s + 1
End If
If InStr(ThisWorkbook.Worksheets(sht).Cells(form + 2, 3), "ê") Then
count_k = count_k + 1
End If
If InStr(ThisWorkbook.Worksheets(sht).Cells(form + 2, 3), "ï") Then
count_p = count_p + 1
End If
If InStr(ThisWorkbook.Worksheets(sht).Cells(form + 2, 3), "ç") Then
count_z = count_z + 1
End If
If InStr(ThisWorkbook.Worksheets(sht).Cells(form + 2, 3), "ã") Then
count_g = count_g + 1
End If
If InStr(ThisWorkbook.Worksheets(sht).Cells(form + 2, 3), "ò") Then
count_t = count_t + 1
End If
If InStr(ThisWorkbook.Worksheets(sht).Cells(form + 2, 3), "í") Then
count_n = count_n + 1
End If
If InStr(ThisWorkbook.Worksheets(sht).Cells(form + 2, 3), "ý") Then
count_e = count_e + 1
End If

End If
Next sht

ThisWorkbook.Worksheets(1).Cells(14, form + 1) = count
ThisWorkbook.Worksheets(1).Cells(16, form + 1) = count_s
ThisWorkbook.Worksheets(1).Cells(17, form + 1) = count_k
ThisWorkbook.Worksheets(1).Cells(18, form + 1) = count_p
ThisWorkbook.Worksheets(1).Cells(19, form + 1) = count_z
ThisWorkbook.Worksheets(1).Cells(20, form + 1) = count_g
ThisWorkbook.Worksheets(1).Cells(21, form + 1) = count_t
ThisWorkbook.Worksheets(1).Cells(22, form + 1) = count_n
ThisWorkbook.Worksheets(1).Cells(23, form + 1) = count_e

End If

If ftype = 1540 Then
For sht = 4 To 34
If IsEmpty(ThisWorkbook.Worksheets(sht).Cells(form + 2, 16)) = False Then
count = count + 1

If InStr(ThisWorkbook.Worksheets(sht).Cells(form + 2, 16), "ñ") Then
count_s = count_s + 1
End If
If InStr(ThisWorkbook.Worksheets(sht).Cells(form + 2, 16), "ê") Then
count_k = count_k + 1
End If
If InStr(ThisWorkbook.Worksheets(sht).Cells(form + 2, 16), "ï") Then
count_p = count_p + 1
End If
If InStr(ThisWorkbook.Worksheets(sht).Cells(form + 2, 16), "ç") Then
count_z = count_z + 1
End If
If InStr(ThisWorkbook.Worksheets(sht).Cells(form + 2, 16), "ã") Then
count_g = count_g + 1
End If
If InStr(ThisWorkbook.Worksheets(sht).Cells(form + 2, 16), "ò") Then
count_t = count_t + 1
End If
If InStr(ThisWorkbook.Worksheets(sht).Cells(form + 2, 16), "í") Then
count_n = count_n + 1
End If
If InStr(ThisWorkbook.Worksheets(sht).Cells(form + 2, 16), "ý") Then
count_e = count_e + 1
End If

End If
Next sht

ThisWorkbook.Worksheets(1).Cells(27, form + 1) = count
ThisWorkbook.Worksheets(1).Cells(29, form + 1) = count_s
ThisWorkbook.Worksheets(1).Cells(30, form + 1) = count_k
ThisWorkbook.Worksheets(1).Cells(31, form + 1) = count_p
ThisWorkbook.Worksheets(1).Cells(32, form + 1) = count_z
ThisWorkbook.Worksheets(1).Cells(33, form + 1) = count_g
ThisWorkbook.Worksheets(1).Cells(34, form + 1) = count_t
ThisWorkbook.Worksheets(1).Cells(35, form + 1) = count_n
ThisWorkbook.Worksheets(1).Cells(36, form + 1) = count_e

End If





End Function

Function percent()

Dim standart_1540 As Integer, _
bubble_1540 As Integer, _
zamyatie_1540 As Integer, _
treshina_1540 As Integer, _
form_1540 As Integer, _
nedoliv_1540 As Integer, _
kaverna_1540 As Integer, _
standart_1210 As Integer, _
bubble_1210 As Integer, _
zamyatie_1210 As Integer, _
treshina_1210 As Integer, _
form_1210 As Integer, _
nedoliv_1210 As Integer, _
kaverna_1210 As Integer

For pge = 4 To 34

standart_1540 = 0
bubble_1540 = 0
zamyatie_1540 = 0
treshina_1540 = 0
form_1540 = 0
nedoliv_1540 = 0
kaverna_1540 = 0
standart_1210 = 0
bubble_1210 = 0
zamyatie_1210 = 0
treshina_1210 = 0
form_1210 = 0
nedoliv_1210 = 0
kaverna_1210 = 0

For line = 3 To 202

If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 3), "ñ") = 1 Then
standart_1210 = standart_1210 + 1
End If

If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 3), "ï") = 1 Then
bubble_1210 = bubble_1210 + 1
End If

If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 3), "ç") = 1 Then
zamyatie_1210 = zamyatie_1210 + 1
End If

If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 3), "ò") = 1 Then
treshina_1210 = treshina_1210 + 1
End If

If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 3), "ã") = 1 Then
form_1210 = form_1210 + 1
End If

If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 3), "ê") = 1 Then
kaverna_1210 = kaverna_1210 + 1
End If

If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 3), "í") = 1 Then
nedoliv_1210 = nedoliv_1210 + 1
End If

If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 3), "ý") = 1 Then

    If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 3), "ñ") = 2 Then
    standart_1210 = standart_1210 + 1
    End If

    If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 3), "ï") = 2 Then
    bubble_1210 = bubble_1210 + 1
    End If

    If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 3), "ç") = 2 Then
    zamyatie_1210 = zamyatie_1210 + 1
    End If

    If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 3), "ò") = 2 Then
    treshina_1210 = treshina_1210 + 1
    End If

    If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 3), "ã") = 2 Then
    form_1210 = form_1210 + 1
    End If

    If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 3), "ê") = 2 Then
    kaverna_1210 = kaverna_1210 + 1
    End If

    If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 3), "í") = 2 Then
    nedoliv_1210 = nedoliv_1210 + 1
    End If

End If

' __________1540


If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 16), "ñ") = 1 Then
standart_1540 = standart_1540 + 1
End If

If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 16), "ï") = 1 Then
bubble_1540 = bubble_1540 + 1
End If

If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 16), "ç") = 1 Then
zamyatie_1540 = zamyatie_1540 + 1
End If

If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 16), "ò") = 1 Then
treshina_1540 = treshina_1540 + 1
End If

If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 16), "ã") = 1 Then
form_1540 = form_1540 + 1
End If

If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 16), "ê") = 1 Then
kaverna_1540 = kaverna_1540 + 1
End If

If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 16), "í") = 1 Then
nedoliv_1540 = nedoliv_1540 + 1
End If

If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 16), "ý") = 1 Then

    If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 16), "ñ") = 2 Then
    standart_1540 = standart_1540 + 1
    End If

    If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 16), "ï") = 2 Then
    bubble_1540 = bubble_1540 + 1
    End If

    If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 16), "ç") = 2 Then
    zamyatie_1540 = zamyatie_1540 + 1
    End If

    If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 16), "ò") = 2 Then
    treshina_1540 = treshina_1540 + 1
    End If

    If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 16), "ã") = 2 Then
    form_1540 = form_1540 + 1
    End If

    If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 16), "ê") = 2 Then
    kaverna_1540 = kaverna_1540 + 1
    End If

    If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 16), "í") = 2 Then
    nedoliv_1540 = nedoliv_1540 + 1
    End If

End If



'If IsEmpty(ThisWorkbook.Worksheets(pge).Cells(line, 3)) = False And InStr(ThisWorkbook.Worksheets(pge).Cells(line, 3), "?") = 0 Then
'count_all1210 = count_all1210 + 1
'If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 3), "ñ") Then
'count_s1210 = count_s1210 + 1
'End If
'End If
'
'If IsEmpty(ThisWorkbook.Worksheets(pge).Cells(line, 16)) = False And InStr(ThisWorkbook.Worksheets(pge).Cells(line, 16), "?") = 0 Then
'count_all1540 = count_all1540 + 1
'If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 16), "ñ") Then
'count_s1540 = count_s1540 + 1
'End If
'End If
'
'If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 3), "í") Then
'count_n = count_n + 1
'End If
'
'If InStr(ThisWorkbook.Worksheets(pge).Cells(line, 16), "í") Then
'count_n = count_n + 1
'End If

Next line

'ThisWorkbook.Worksheets(3).Cells(4, pge - 2) = count_s1540
'ThisWorkbook.Worksheets(3).Cells(5, pge - 2) = count_all1540 - count_s1540
'ThisWorkbook.Worksheets(3).Cells(8, pge - 2) = count_s1210
'ThisWorkbook.Worksheets(3).Cells(9, pge - 2) = count_all1210 - count_s1210
'
'ThisWorkbook.Worksheets(3).Cells(22, pge - 2) = count_n + ThisWorkbook.Worksheets(3).Cells(17, pge - 2)

'standart_1540 = 0
'bubble_1540 = 0
'zamyatie_1540 = 0
'treshina_1540 = 0
'form_1540 = 0
'nedoliv_1540 = 0
'kaverna_1540 = 0
'standart_1210 = 0
'bubble_1210 = 0
'zamyatie_1210 = 0
'treshina_1210 = 0
'form_1210 = 0
'nedoliv_1210 = 0
'kaverna_1210 = 0

ThisWorkbook.Worksheets(3).Cells(4, pge - 2) = standart_1540
ThisWorkbook.Worksheets(3).Cells(5, pge - 2) = bubble_1540
ThisWorkbook.Worksheets(3).Cells(6, pge - 2) = zamyatie_1540
ThisWorkbook.Worksheets(3).Cells(7, pge - 2) = kaverna_1540
ThisWorkbook.Worksheets(3).Cells(8, pge - 2) = treshina_1540
ThisWorkbook.Worksheets(3).Cells(9, pge - 2) = form_1540
ThisWorkbook.Worksheets(3).Cells(10, pge - 2) = nedoliv_1540


ThisWorkbook.Worksheets(3).Cells(14, pge - 2) = standart_1210
ThisWorkbook.Worksheets(3).Cells(15, pge - 2) = bubble_1210
ThisWorkbook.Worksheets(3).Cells(16, pge - 2) = zamyatie_1210
ThisWorkbook.Worksheets(3).Cells(17, pge - 2) = kaverna_1210
ThisWorkbook.Worksheets(3).Cells(18, pge - 2) = treshina_1210
ThisWorkbook.Worksheets(3).Cells(19, pge - 2) = form_1210
ThisWorkbook.Worksheets(3).Cells(20, pge - 2) = nedoliv_1210


Next pge


End Function

Function Full_Zach(day As Integer)

Dim count_all As Integer, count_1 As Integer, count0 As Integer, count1 As Integer
Dim coloumn As Integer

For i = 30 To 100

If IsEmpty(ThisWorkbook.Worksheets(2).Cells(i, 1)) = True Then
Exit For
End If

count_all = 0
count_1 = 0
count0 = 0
count1 = 0


For j = 3 To 202
coloumn = 18


If (ThisWorkbook.Worksheets(day + 3).Cells(j, coloumn - 1) = ThisWorkbook.Worksheets(2).Cells(i, 1)) And (ThisWorkbook.Worksheets(day + 3).Cells(j, coloumn - 1).Interior.Color = vbWhite Or ThisWorkbook.Worksheets(day + 3).Cells(j, coloumn - 1).Interior.Color = vbYellow) Then
count_all = count_all + 1
End If

If ThisWorkbook.Worksheets(day + 3).Cells(j, coloumn - 1) = ThisWorkbook.Worksheets(2).Cells(i, 1) And ThisWorkbook.Worksheets(day + 3).Cells(j, coloumn) = -1 Then
count_1 = count_1 + 1
ElseIf ThisWorkbook.Worksheets(day + 3).Cells(j, coloumn - 1) = ThisWorkbook.Worksheets(2).Cells(i, 1) And ThisWorkbook.Worksheets(day + 3).Cells(j, coloumn) = "0" Then
count0 = count0 + 1
ElseIf ThisWorkbook.Worksheets(day + 3).Cells(j, coloumn - 1) = ThisWorkbook.Worksheets(2).Cells(i, 1) And ThisWorkbook.Worksheets(day + 3).Cells(j, coloumn) = 1 Then
count1 = count1 + 1
End If

Next j

ThisWorkbook.Worksheets(2).Cells(i, 2 + (day - 1) * 11) = count_all
ThisWorkbook.Worksheets(2).Cells(i, 3 + (day - 1) * 11) = count_1
ThisWorkbook.Worksheets(2).Cells(i, 4 + (day - 1) * 11) = count0
ThisWorkbook.Worksheets(2).Cells(i, 5 + (day - 1) * 11) = count1
ThisWorkbook.Worksheets(2).Cells(i, 6 + (day - 1) * 11) = count1 + count_1 + count0


count_all = 0
count_1 = 0
count0 = 0
count1 = 0

For j = 3 To 202
coloumn = 5
If (ThisWorkbook.Worksheets(day + 3).Cells(j, coloumn - 1) = ThisWorkbook.Worksheets(2).Cells(i, 1)) And (ThisWorkbook.Worksheets(day + 3).Cells(j, coloumn - 1).Interior.Color = vbWhite Or ThisWorkbook.Worksheets(day + 3).Cells(j, coloumn - 1).Interior.Color = vbYellow) Then
count_all = count_all + 1
End If


If ThisWorkbook.Worksheets(day + 3).Cells(j, coloumn - 1) = ThisWorkbook.Worksheets(2).Cells(i, 1) And ThisWorkbook.Worksheets(day + 3).Cells(j, coloumn) = -1 Then
count_1 = count_1 + 1
ElseIf ThisWorkbook.Worksheets(day + 3).Cells(j, coloumn - 1) = ThisWorkbook.Worksheets(2).Cells(i, 1) And ThisWorkbook.Worksheets(day + 3).Cells(j, coloumn) = "0" Then
count0 = count0 + 1
ElseIf ThisWorkbook.Worksheets(day + 3).Cells(j, coloumn - 1) = ThisWorkbook.Worksheets(2).Cells(i, 1) And ThisWorkbook.Worksheets(day + 3).Cells(j, coloumn) = 1 Then
count1 = count1 + 1
End If

Next j

ThisWorkbook.Worksheets(2).Cells(i, 7 + (day - 1) * 11) = count_all
ThisWorkbook.Worksheets(2).Cells(i, 8 + (day - 1) * 11) = count_1
ThisWorkbook.Worksheets(2).Cells(i, 9 + (day - 1) * 11) = count0
ThisWorkbook.Worksheets(2).Cells(i, 10 + (day - 1) * 11) = count1
ThisWorkbook.Worksheets(2).Cells(i, 11 + (day - 1) * 11) = count1 + count_1 + count0

Next i



For i = 2 To 20
If IsEmpty(ThisWorkbook.Worksheets(2).Cells(i, 1)) = True Then
Exit For
End If
For j = 30 To 100
If ThisWorkbook.Worksheets(2).Cells(i, 1) = ThisWorkbook.Worksheets(2).Cells(j, 1) Then
ThisWorkbook.Worksheets(2).Cells(j, 12 + (day - 1) * 11) = ThisWorkbook.Worksheets(2).Cells(i, day + 1)
End If
Next j
Next i




End Function



Sub main()

Dim coloumn As Integer
Dim sht As Integer
Dim line As Integer


coloumn = 3

For s = 4 To 34
sht = s
For l = 3 To 202
line = l
quality line, coloumn, sht
Next l
Next s

coloumn = 16

For s = 4 To 34
sht = s
For l = 3 To 202
line = l
quality line, coloumn, sht
Next l
Next s

For s = 4 To 34
sht = s
ZP sht
Next s

For s = 4 To 34
sht = s
cash sht
Next s

For table = 1 To 10
sht = table
Tables sht
Next table

For form = 1 To 200
sht = form
forms sht, 1210
Next form

For form = 1 To 200
sht = form
forms sht, 1540
Next form

percent

For i = 1 To 31
sht = i
Full_Zach sht
Next i

End Sub


