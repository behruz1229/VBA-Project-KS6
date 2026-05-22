Attribute VB_Name = "Dublicate"
Sub CountDuplicatesInO()
    ' Отключаем обновление экрана и пересчёт для максимальной скорости
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "O").End(xlUp).Row
    
    If lastRow < 3 Then
        ws.Range("O2").Value = 0
        GoTo CleanUp
    End If
    
    ' 1. Считываем весь диапазон в массив (в 1000+ раз быстрее построчного чтения)
    Dim data As Variant
    data = ws.Range("O3:O" & lastRow).Value
    
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long, key As Variant, totalDup As Long
    
    ' 2. Подсчёт частоты каждого значения
    For i = 1 To UBound(data, 1)
        If Not IsEmpty(data(i, 1)) And data(i, 1) <> "" Then
            key = CStr(data(i, 1)) ' Приводим к строке для корректной работы
            If dict.Exists(key) Then
                dict(key) = dict(key) + 1
            Else
                dict(key) = 1
            End If
        End If
    Next i
    
    ' 3. Суммируем только те значения, которые встречаются >1 раза
    For Each key In dict.Keys
        If dict(key) > 1 Then totalDup = totalDup + dict(key)
    Next key
    
    ' 4. Вставляем результат в O2
    ws.Range("O2").Value = totalDup
    
CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "Готово! Найдено " & totalDup & " повторяющихся ячеек.", vbInformation
End Sub
