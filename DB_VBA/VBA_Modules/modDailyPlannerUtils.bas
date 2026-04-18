Option Compare Database
Option Explicit

'################################################################
'########      Применить UI режим админ/пользователь      ########
'################################################################
Public Sub ApplyAdminUiMode()
    On Error GoTo ErrHandler

    If IsAdminModeEnabled() Then
        ' Админ-режим: ничего не скрываем
        Call SetAccessSpecialKeysEnabled(True)
        DoCmd.SelectObject acTable, , True
        DoCmd.ShowToolbar "Ribbon", acToolbarYes
        Debug.Print "[f_daily_planner] UI mode -> ADMIN (Ribbon/Navigation visible)"
    Else
        ' Пользовательский режим: скрываем навигацию и ленту
        Call SetAccessSpecialKeysEnabled(False)
        DoCmd.NavigateTo "acNavigationCategoryObjectType"
        DoCmd.RunCommand acCmdWindowHide
        DoCmd.ShowToolbar "Ribbon", acToolbarNo
        Debug.Print "[f_daily_planner] UI mode -> USER (Ribbon/Navigation hidden)"
    End If

    Exit Sub

ErrHandler:
    Debug.Print "[f_daily_planner][ERR][ApplyAdminUiMode] " & Err.Number & " - " & Err.description
End Sub

'################################################################
'########      Разрешить/запретить спецклавиши Access    ########
'################################################################
Public Sub SetAccessSpecialKeysEnabled(ByVal isEnabled As Boolean)
    On Error GoTo PropertyMissing
    CurrentDb.Properties("AllowSpecialKeys") = isEnabled
    Exit Sub

PropertyMissing:
    If Err.Number = 3270 Then
        Dim db As DAO.Database
        Dim prop As DAO.Property

        Set db = CurrentDb
        Set prop = db.CreateProperty("AllowSpecialKeys", dbBoolean, isEnabled)
        db.Properties.Append prop
    Else
        Debug.Print "[f_daily_planner][ERR][SetAccessSpecialKeysEnabled] " & Err.Number & " - " & Err.description
    End If
End Sub

'################################################################
'########      Проверка: включен ли режим админа          ########
'################################################################
Public Function IsAdminModeEnabled() As Boolean
    On Error GoTo PropertyMissing
    IsAdminModeEnabled = CBool(CurrentDb.Properties("AllowByPassKey"))
    Exit Function

PropertyMissing:
    ' Если свойство отсутствует, считаем, что админ-режим выключен.
    IsAdminModeEnabled = False
End Function

'################################################################
'########      Получить UsableWidth/UsableHeight         ########
'################################################################
Public Function GetAccessUsableSize(ByVal propertyName As String) As Long
    On Error GoTo ExitFn
    GetAccessUsableSize = CLng(CallByName(Application, propertyName, VbGet))
    Exit Function
ExitFn:
    GetAccessUsableSize = 0
End Function

'################################################################
'########      Подсчет событий для выбранной даты         ########
'################################################################
Public Function CountEventsForPanelDate(ByVal targetDate As Date, ByVal ExecutorID As Variant) As Long
    Dim whereText As String

    whereText = "EventDate = #" & Format(DateValue(targetDate), "mm\/dd\/yyyy") & "#"
    If Not IsNull(ExecutorID) Then
        whereText = whereText & " AND ExecutorID = " & CLng(ExecutorID)
    End If

    CountEventsForPanelDate = DCount("*", "tbEventInstances", whereText)
End Function

'################################################################
'########   Безопасная установка цвета для контрола      ########
'################################################################
Public Sub SetControlColorIfSupported(ByVal targetControl As Object, ByVal propertyName As String, ByVal colorValue As Long)
    On Error GoTo ExitSub
    CallByName targetControl, propertyName, VbLet, colorValue
ExitSub:
End Sub

'################################################################
'########           Проверка выходного дня               ########
'################################################################
Public Function IsWeekend(checkDate As Date) As Boolean
    Dim dayOfWeek As Integer
    dayOfWeek = weekday(checkDate, vbMonday) ' Понедельник=1, Воскресенье=7
    IsWeekend = (dayOfWeek = 6) Or (dayOfWeek = 7) ' Суббота или Воскресенье
End Function

'################################################################
'########            Затемнение цвета                    ########
'################################################################
Public Function DarkenColor(originalColor As Long, factor As Double) As Long
    Dim r As Integer, g As Integer, b As Integer
    r = originalColor Mod 256
    g = (originalColor \ 256) Mod 256
    b = (originalColor \ 65536) Mod 256

    r = r * factor
    g = g * factor
    b = b * factor

    DarkenColor = RGB(r, g, b)
End Function

