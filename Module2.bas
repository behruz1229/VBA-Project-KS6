Attribute VB_Name = "Module2"
Sub ExportAllVBA()
    Dim vbComp As Object
    Dim filePath As String
    
    ' Папка, куда сохранятся файлы. Можно изменить на свою, например "C:\VBA_Export\"
    filePath = ActiveWorkbook.Path & "\VBA_Export_" & Format(Now, "yyyymmdd_hhnnss") & "\"
    
    ' Создаем папку, если её нет
    On Error Resume Next
    MkDir filePath
    On Error GoTo 0
    
    ' Проходим по всем компонентам VBA проекта
    For Each vbComp In ActiveWorkbook.VBProject.VBComponents
        Dim exportName As String
        exportName = filePath & vbComp.Name
        
        ' Добавляем правильное расширение в зависимости от типа компонента
        Select Case vbComp.Type
            Case 1: exportName = exportName & ".bas"  ' Обычный модуль
            Case 2: exportName = exportName & ".cls"  ' Модуль класса
            Case 3: exportName = exportName & ".frm"  ' Форма (UserForm)
            Case 100: exportName = exportName & ".cls" ' Модуль листа или книги (тоже класс)
            Case Else: exportName = exportName & ".bas"
        End Select
        
        ' Экспортируем компонент
        On Error Resume Next
        vbComp.Export exportName
        If Err.Number <> 0 Then
            Debug.Print "Ошибка экспорта " & vbComp.Name & ": " & Err.Description
        End If
        On Error GoTo 0
    Next vbComp
    
    MsgBox "Готово! Все модули экспортированы в папку:" & vbCrLf & filePath
End Sub
