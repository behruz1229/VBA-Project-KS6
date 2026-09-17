Attribute VB_Name = "sep_O"
Sub ConcatColumnsToO()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Определяем последнюю заполненную строку по столбцам A, C, D, F
    Dim lastRow As Long
    lastRow = Application.Max( _
        ws.Cells(ws.Rows.Count, "A").End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, "C").End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, "D").End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, "F").End(xlUp).Row)
    
    ' Если данные заканчиваются раньше 3-й строки, выходим
    If lastRow < 3 Then
        MsgBox "Данные начинаются выше 3-й строки или отсутствуют.", vbInformation
        Exit Sub
    End If
    
    ' Считываем нужные диапазоны в массивы для быстродействия
    Dim arrA, arrB, arrC, arrD, arrE, arrF, arrOut
    arrA = ws.Range("A3:A" & lastRow).Value
    arrC = ws.Range("C3:C" & lastRow).Value
    arrD = ws.Range("D3:D" & lastRow).Value
    arrF = ws.Range("F3:F" & lastRow).Value
    
    Dim i As Long, numRows As Long
    numRows = lastRow - 2               ' количество строк, начиная с 3-й
    ReDim arrOut(1 To numRows, 1 To 1)  ' массив для результатов
    
    ' Склеиваем значения (без разделителей)
    For i = 1 To numRows
        arrOut(i, 1) = CStr(arrA(i, 1)) & "_" & CStr(arrC(i, 1)) & _
                       "_" & CStr(arrD(i, 1)) & "_" & CStr(arrF(i, 1))
    Next i
    
    ' Вставляем результат как значения в столбец O
    ws.Range("O3:O" & lastRow).Value = arrOut
    
    MsgBox "Готово! Сцепка вставлена в столбец O.", vbInformation
End Sub
