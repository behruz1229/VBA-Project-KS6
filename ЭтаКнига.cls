VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЭтаКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Создание кнопки при открытии книги
Private Sub Workbook_Open()
    AddContextMenuButton
End Sub

' Удаление кнопки перед закрытием книги
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    RemoveContextMenuButton
End Sub

Public Sub AddContextMenuButton()
    On Error Resume Next
    Application.CommandBars("Cell").Controls("Обновить таблицу из ТСБ и МОТ").Delete
    On Error GoTo 0
    
    Dim btn As Object
    ' Type:=1 безопаснее, чем msoControlButton (не требует подключения библиотек Office)
    Set btn = Application.CommandBars("Cell").Controls.Add(Type:=1, Before:=1)
    With btn
        .Caption = "Обновить таблицу из ТСБ и МОТ"
        .FaceId = 2608
        .OnAction = "MainUpdate"          ' Путь к макросу (см. примечание ниже)
        .Tag = "SyncBtn_Rcv"              ' Метка для надёжного поиска
    End With
End Sub

Public Sub RemoveContextMenuButton()
    On Error Resume Next
    ' Удаляем по Caption
    Application.CommandBars("Cell").Controls("Обновить таблицу из ТСБ и МОТ").Delete
    ' Дублируем удаление по Tag (надёжнее, если пользователь переименует кнопку вручную)
    Application.CommandBars("Cell").FindControl(Tag:="SyncBtn_Rcv").Delete
    On Error GoTo 0
End Sub









' старая кнопка
' _________________________________________________________________________________________




'Private Sub Workbook_Open()
'    AddContextMenuButton
'End Sub
'
'Public Sub AddContextMenuButton()
'    On Error Resume Next
'    Application.CommandBars("Cell").Controls("Обновить таблицу из ТСБ и МОТ").Delete
'    On Error GoTo 0
'
'    Dim btn As Object
'    Set btn = Application.CommandBars("Cell").Controls.Add(Type:=msoControlButton, Before:=1)
'    With btn
'        .Caption = "Обновить таблицу из ТСБ и МОТ"
'        .FaceId = 2608
'        .OnAction = "MainUpdate"
'        .Tag = "SyncBtn_Rcv"
'    End With
'    Debug.Print "[UI] Кнопка контекстного меню добавлена."
'End Sub
