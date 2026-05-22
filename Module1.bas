Attribute VB_Name = "Module1"
Option Explicit

Public gAddedCount As Long
Public gOutdatedCount As Long
Public wbRcv As Workbook ' Глобальная ссылка на файл-приёмник

' ================= ГЛАВНЫЙ ЗАПУСК =================
Public Sub MainUpdate()
    Dim tStart As Double: tStart = CDbl(Now)
    
    ' Сохраняем ссылку на приёмник сразу, чтобы не зависеть от ActiveWorkbook
    On Error Resume Next
    Set wbRcv = ActiveWorkbook
    On Error GoTo 0
    If wbRcv Is Nothing Then
        MsgBox "Не удалось определить активный файл-приёмник. Запустите макрос из него.", vbCritical
        Exit Sub
    End If
    
    With Application
        .ScreenUpdating = False
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .StatusBar = False
    End With
    
    On Error GoTo ErrHandler
    
    UpdateProgress 0, "Инициализация процесса"
    Debug.Print "[Main] Приёмник: " & wbRcv.Name
    Module1_Sync
    Module2_UpdateInspections
    
    Dim elapsedSec As Double: elapsedSec = (CDbl(Now) - tStart) * 86400
    Dim timeStr As String
    timeStr = Int(elapsedSec / 3600) & ":" & Format((elapsedSec Mod 3600) / 60, "00") & ":" & Format(elapsedSec Mod 60, "00")
    
    MsgBox "Обновление завершено!" & vbCrLf & _
           "? Время выполнения: " & timeStr & vbCrLf & _
           "? Добавлено строк: " & gAddedCount & vbCrLf & _
           "?? Устаревших строк: " & gOutdatedCount, vbInformation, "Результат"
    
    GoTo CleanUp
    
ErrHandler:
    MsgBox "Произошла ошибка: " & Err.Description & vbCrLf & "Модуль: " & Err.Source, vbCritical, "Ошибка"
    
CleanUp:
    With Application
        .ScreenUpdating = True
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .StatusBar = False
    End With
End Sub

' ================= МОДУЛЬ 1: СИНХРОНИЗАЦИЯ =================
Private Sub Module1_Sync()
    UpdateProgress 5, "Чтение файла-приёмника"
    Dim wsRcv As Worksheet
    Set wsRcv = GetSheetSafe(wbRcv, "ТСБиМОТ")
    If wsRcv Is Nothing Then Exit Sub
    
    Dim lrRcv As Long: lrRcv = wsRcv.Cells(wsRcv.Rows.Count, 1).End(xlUp).Row
    If lrRcv < 3 Then Exit Sub
    
    Dim arrRcv As Variant: arrRcv = wsRcv.Range("A3:P" & lrRcv).Value2
    Dim dictRcv As Object: Set dictRcv = CreateObject("Scripting.Dictionary")
    Dim i As Long, k As String
    Dim totalRcv As Long: totalRcv = UBound(arrRcv, 1)
    
    For i = 1 To totalRcv
        k = MakeKey(arrRcv(i, 1), arrRcv(i, 3), arrRcv(i, 4), arrRcv(i, 6), arrRcv(i, 7))
        If k <> "" And Not dictRcv.Exists(k) Then dictRcv.Add k, i + 2
        If i Mod 5000 = 0 Then UpdateProgress 5 + Int(i / totalRcv * 5), "Чтение приёмника: " & i
    Next i
    UpdateProgress 10, "Словарь приёмника построен (" & dictRcv.Count & " ключей)"
    Debug.Print "[Module1] Приёмник: " & dictRcv.Count & " уникальных ключей."
    
    Dim srcFiles(1 To 2) As String, srcSheets(1 To 2) As String
    srcFiles(1) = "Ведомость элементов МК МОТ.xlsb": srcSheets(1) = "База данных по элементам"
    srcFiles(2) = "Ведомость элементов МК ТСБ.xlsb": srcSheets(2) = "База данных по элементам"
    Const srcDir As String = "\\vls.lan\ULVZG-DFS\ПТС\1.14. ТСБ и МОТ_Исполнительная документация КМ\!Ведомость элементов\"
    
    Dim dictSrc As Object: Set dictSrc = CreateObject("Scripting.Dictionary")
    Dim newRows As New Collection
    Dim s As Integer
    
    For s = 1 To 2
        UpdateProgress 10 + s * 10, "Подготовка источника: " & srcFiles(s)
        Dim fPath As String: fPath = srcDir & srcFiles(s)
        
        If Dir(fPath) = "" Then
            fPath = Application.GetOpenFilename(FileFilter:="Excel файлы (*.xlsb), *.xlsb", Title:="Не найден: " & srcFiles(s) & ". Выберите вручную.")
            If fPath = "False" Then GoTo SkipSrc
        End If
        
        Dim wbSrc As Workbook
        Set wbSrc = Workbooks.Open(fPath, ReadOnly:=True)
        Dim wsSrc As Worksheet
        On Error Resume Next
        Set wsSrc = wbSrc.Sheets(srcSheets(s))
        On Error GoTo 0
        If wsSrc Is Nothing Then wbSrc.Close False: GoTo SkipSrc
        
        Dim lrSrc As Long: lrSrc = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
        If lrSrc < 3 Then wbSrc.Close False: GoTo SkipSrc
        
        Dim arrSrc As Variant: arrSrc = wsSrc.Range("A3:P" & lrSrc).Value2
        Dim totalSrc As Long: totalSrc = UBound(arrSrc, 1)
        Dim r As Long
        
        UpdateProgress 15 + s * 10, "Сканирование: " & srcFiles(s)
        For r = 1 To totalSrc
            If InStr(1, UCase(Trim(CStr(arrSrc(r, 1)))), "TQ", vbTextCompare) > 0 Then GoTo NextRow
            k = MakeKey(arrSrc(r, 1), arrSrc(r, 3), arrSrc(r, 4), arrSrc(r, 6), arrSrc(r, 7))
            If k = "" Then GoTo NextRow
            
            If Not dictSrc.Exists(k) Then
                dictSrc.Add k, r
                If Not dictRcv.Exists(k) Then
                    Dim tmpRow(1 To 10) As Variant, c As Integer
                    For c = 1 To 9: tmpRow(c) = arrSrc(r, c): Next c
                    tmpRow(10) = arrSrc(r, 16) ' P -> J
                    newRows.Add tmpRow
                End If
            End If
            If r Mod 5000 = 0 Then UpdateProgress 20 + s * 10 + Int(r / totalSrc * 15), "Источник " & s & ": " & r & "/" & totalSrc
NextRow:
        Next r
        wbSrc.Close False
SkipSrc:
    Next s
    
    gAddedCount = newRows.Count
    If gAddedCount > 0 Then
        UpdateProgress 55, "Вставка " & gAddedCount & " новых строк"
        Dim newArr() As Variant: ReDim newArr(1 To gAddedCount, 1 To 10)
        Dim idx As Long
        For idx = 1 To gAddedCount
            Dim tmp As Variant: tmp = newRows(idx)
            For c = 1 To 10: newArr(idx, c) = tmp(c): Next c
        Next idx
        wsRcv.Range("A" & lrRcv + 1).Resize(gAddedCount, 10).Value = newArr
        wsRcv.Range("A" & lrRcv + 1 & ":P" & lrRcv + gAddedCount).Interior.color = RGB(102, 255, 102)
        lrRcv = wsRcv.Cells(wsRcv.Rows.Count, 1).End(xlUp).Row
    End If
    
    UpdateProgress 70, "Поиск устаревших строк"
    Dim outdatedRows As New Collection, key
    For Each key In dictRcv.Keys
        If Not dictSrc.Exists(key) Then outdatedRows.Add dictRcv(key)
    Next key
    gOutdatedCount = outdatedRows.Count
    
    If gOutdatedCount > 0 Then
        Dim ou, stepOut As Long
        For Each ou In outdatedRows
            stepOut = stepOut + 1
            wsRcv.Range("A" & ou & ":P" & ou).Interior.color = RGB(255, 153, 0)
            If stepOut Mod 5000 = 0 Then UpdateProgress 80, "Покраска устаревших: " & stepOut
        Next ou
    End If
    
    UpdateProgress 95, "Применение формата даты"
    wsRcv.Range("I3:I" & lrRcv).NumberFormat = "dd.mm.yyyy"
    
    UpdateProgress 100, "Модуль 1 завершён"
    Debug.Print "[Module1] Готово. Добавлено: " & gAddedCount & ", Устаревших: " & gOutdatedCount
End Sub

' ================= МОДУЛЬ 2: ИНСПЕКЦИИ =================
Private Sub Module2_UpdateInspections()
    Dim wsInsp As Worksheet, wbInsp As Workbook, wsRcv As Worksheet
    Dim lrInsp As Long, lrRcv As Long
    Dim vC As Variant, vAH As Variant, vJ As Variant, vK As Variant
    Dim totalInsp As Long, totalMatch As Long, i As Long, matchCnt As Long
    Dim keyC As String, jVal As String
    Dim dictInsp As Object
    Dim inspDir As String, latestFile As String, fName As String
    
    On Error GoTo ErrHandler
    
    UpdateProgress 5, "Поиск файла инспекций"
    inspDir = "\\vls.lan\ULVZG-DFS\ПТС\1.14. ТСБ и МОТ_Исполнительная документация КМ\Выгрузки\RFI\"
    latestFile = GetLatestModifiedFile(inspDir, "Инспекции на *.xlsx")
    
    If latestFile = "" Then
        latestFile = Application.GetOpenFilename(FileFilter:="Excel файлы (*.xlsx), *.xlsx", Title:="Файл инспекций не найден. Выберите вручную.")
        If latestFile = "False" Then Exit Sub
    Else
        latestFile = inspDir & latestFile
    End If
    
    fName = Mid(latestFile, InStrRev(latestFile, "\") + 1)
    On Error Resume Next
    Set wbInsp = Workbooks(fName)
    On Error GoTo ErrHandler
    If wbInsp Is Nothing Then Set wbInsp = Workbooks.Open(latestFile, ReadOnly:=True)
    
    UpdateProgress 15, "Чтение данных инспекций"
    Set wsInsp = wbInsp.Sheets(1)
    lrInsp = wsInsp.Cells(wsInsp.Rows.Count, 1).End(xlUp).Row
    Debug.Print "[M2] Конец по столбцу A: " & lrInsp
    If lrInsp < 3 Then wbInsp.Close False: Exit Sub
    
    vC = wsInsp.Range("C3:C" & lrInsp).Value2
    vAH = wsInsp.Range("AH3:AH" & lrInsp).Value2
    If Not IsArray(vC) Then ReDim vC(1 To 1, 1 To 1): vC(1, 1) = wsInsp.Range("C3").Value2
    If Not IsArray(vAH) Then ReDim vAH(1 To 1, 1 To 1): vAH(1, 1) = wsInsp.Range("AH3").Value2
    
    totalInsp = UBound(vC, 1)
    If UBound(vAH, 1) < totalInsp Then totalInsp = UBound(vAH, 1)
    Debug.Print "[M2] vC строк: " & UBound(vC, 1) & " | vAH строк: " & UBound(vAH, 1) & " | Цикл до: " & totalInsp
    
    Set dictInsp = CreateObject("Scripting.Dictionary")
    For i = 1 To totalInsp
        keyC = Trim(CStr(vC(i, 1)))
        If keyC <> "" Then dictInsp(keyC) = vAH(i, 1)
        If i Mod 5000 = 0 Then UpdateProgress 20 + Int(i / totalInsp * 15), "Чтение инспекций: " & i
    Next i
    Debug.Print "[M2] Словарь инспекций готов. Ключей: " & dictInsp.Count
    
    UpdateProgress 40, "Сопоставление с приёмником"
    Debug.Print "[M2] Поиск листа приёмника..."
    Set wsRcv = GetSheetSafe(wbRcv, "ТСБиМОТ") ' Используем сохранённую ссылку!
    If wsRcv Is Nothing Then
        MsgBox "Лист 'ТСБиМОТ' не найден в файле-приёмнике!", vbCritical
        Exit Sub
    End If
    Debug.Print "[M2] Лист найден. Вычисление последней строки..."
    
    lrRcv = wsRcv.Cells(wsRcv.Rows.Count, 1).End(xlUp).Row
    Debug.Print "[M2] Приёмник строк (по A): " & lrRcv
    If lrRcv < 3 Then Exit Sub
    
    vJ = wsRcv.Range("J3:J" & lrRcv).Value2
    vK = wsRcv.Range("K3:K" & lrRcv).Value2
    If Not IsArray(vJ) Then ReDim vJ(1 To 1, 1 To 1): vJ(1, 1) = wsRcv.Range("J3").Value2
    If Not IsArray(vK) Then ReDim vK(1 To 1, 1 To 1): vK(1, 1) = wsRcv.Range("K3").Value2
    
    totalMatch = UBound(vJ, 1)
    If UBound(vK, 1) < totalMatch Then totalMatch = UBound(vK, 1)
    Debug.Print "[M2] vJ строк: " & UBound(vJ, 1) & " | vK строк: " & UBound(vK, 1) & " | Цикл до: " & totalMatch
    
    matchCnt = 0
    For i = 1 To totalMatch
        If i > UBound(vJ, 1) Or i > UBound(vK, 1) Then
            Debug.Print "[M2] ?? SAFETY EXIT: i=" & i & " превышает границы массивов"
            Exit For
        End If
        
        jVal = Trim(CStr(vJ(i, 1)))
        If dictInsp.Exists(jVal) Then
            vK(i, 1) = dictInsp(jVal)
            matchCnt = matchCnt + 1
        End If
        
        If i Mod 5000 = 0 Then UpdateProgress 50 + Int(i / totalMatch * 45), "Сопоставление: " & i & "/" & totalMatch
    Next i
    Debug.Print "[M2] Цикл завершён. Найдено совпадений: " & matchCnt
    
    If totalMatch > 0 Then wsRcv.Range("K3").Resize(totalMatch, 1).Value = vK
    wbInsp.Close False
    UpdateProgress 100, "Модуль 2 завершён"
    Exit Sub

ErrHandler:
    Dim logMsg As String
    logMsg = "=== ОТЛАДОЧНЫЙ ЛОГ ОШИБКИ ===" & vbCrLf & _
             "Err: " & Err.Number & " - " & Err.Description & vbCrLf & _
             "Текущий шаг i: " & i & vbCrLf & _
             "vJ граница(строк): " & IIf(IsArray(vJ), UBound(vJ, 1), "Скаляр") & vbCrLf & _
             "vK граница(строк): " & IIf(IsArray(vK), UBound(vK, 1), "Скаляр") & vbCrLf & _
             "Цикл должен идти до: " & totalMatch & vbCrLf & _
             "dictInsp.Count: " & IIf(Not dictInsp Is Nothing, dictInsp.Count, "Не создан") & vbCrLf & _
             "Последнее jVal: """ & jVal & """" & vbCrLf & _
             "Ключ в словаре: " & IIf(Not dictInsp Is Nothing, dictInsp.Exists(jVal), "N/A")
             
    Debug.Print logMsg
    MsgBox logMsg, vbCritical, "Детальная диагностика Module2"
End Sub

' ================= ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ =================
Public Sub UpdateProgress(pct As Integer, msg As String)
    Application.StatusBar = "Выполняется: " & msg & " [" & pct & "%]"
    Debug.Print "[" & Format(Now, "hh:mm:ss") & "] " & msg & " (" & pct & "%)"
    DoEvents
End Sub

' Безопасный поиск листа (игнорирует регистр и пробелы)
Private Function GetSheetSafe(wb As Workbook, sheetName As String) As Worksheet
    Dim ws As Worksheet, cleanName As String, cleanSearch As String
    cleanSearch = Trim(LCase(sheetName))
    For Each ws In wb.Sheets
        cleanName = Trim(LCase(ws.Name))
        If cleanName = cleanSearch Then
            Set GetSheetSafe = ws
            Exit Function
        End If
    Next ws
    ' Fallback: пробуем прямой доступ, если цикл не сработал
    On Error Resume Next
    Set GetSheetSafe = wb.Sheets(sheetName)
    On Error GoTo 0
End Function

Private Function MakeKey(ParamArray vals() As Variant) As String
    Dim res As String, v As Variant, i As Long
    For i = 0 To UBound(vals)
        v = vals(i)
        If IsError(v) Or IsNull(v) Then v = ""
        res = res & Trim(CStr(v)) & "_"
    Next i
    If Len(res) > 0 Then res = Left(res, Len(res) - 1)
    MakeKey = res
End Function

Private Function GetLatestModifiedFile(folderPath As String, fileMask As String) As String
    On Error Resume Next
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then Exit Function
    Dim f As Object, maxDate As Date: maxDate = #1/1/1900#
    Dim latest As Object
    For Each f In fso.GetFolder(folderPath).Files
        If f.Name Like fileMask Then
            If f.DateLastModified > maxDate Then
                maxDate = f.DateLastModified
                Set latest = f
            End If
        End If
    Next f
    If Not latest Is Nothing Then GetLatestModifiedFile = latest.Name
End Function
