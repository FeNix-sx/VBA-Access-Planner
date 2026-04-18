Option Compare Database
Option Explicit

'################################################################
'########     КНОПКА «ИГРЫ» — ТЕКСТЫ И СЧЁТЧИКИ          ########
'################################################################
Public Const PLN_GAMES_SETTING_COUNT As String = "games_click_count"
Public Const PLN_GAMES_SETTING_CAPTION As String = "games_button_caption"
Public Const PLN_GAMES_SETTING_LAST_MESSAGE_KEY As String = "games_last_message_key"
Private Const PLN_GAMES_RANDOM_POOL_SIZE As Long = 20
Public Function PlnGames_GetClickCount() As Long
    On Error GoTo Err_Handler
    Dim rawValue As String
    Dim parsedValue As Long

    rawValue = PlnSettings_GetValue(PLN_GAMES_SETTING_COUNT, "0")
    parsedValue = CLng(Val(Nz(rawValue, "0")))

    If parsedValue < 0 Or parsedValue > 99 Then
        parsedValue = 0
    End If

    PlnGames_GetClickCount = parsedValue
    Exit Function

Err_Handler:
    PlnGames_GetClickCount = 0
End Function

'################################################################
'########        Определить титул кнопки "Игры"          ########
'################################################################
Public Function PlnGames_GetTitleByCount(ByVal clickCount As Long) As String
    Select Case clickCount
        Case 0 To 9
            PlnGames_GetTitleByCount = "Игры"
        Case 10 To 19
            PlnGames_GetTitleByCount = "Разминатор"
        Case 20 To 29
            PlnGames_GetTitleByCount = "Уверенный нажиматор"
        Case 30 To 39
            PlnGames_GetTitleByCount = "Серийный нажиматор"
        Case 40 To 49
            PlnGames_GetTitleByCount = "Мастер-нажиматор"
        Case 50 To 59
            PlnGames_GetTitleByCount = "ЛКМ III степени"
        Case 60 To 69
            PlnGames_GetTitleByCount = "ЛКМ II степени"
        Case 70 To 79
            PlnGames_GetTitleByCount = "ЛКМ I степени"
        Case 80 To 89
            PlnGames_GetTitleByCount = "Ветеран кликов"
        Case Else
            PlnGames_GetTitleByCount = "Заслуженный НАЖИМАТОР"
    End Select
End Function

'################################################################
'########      Сформировать caption кнопки "Игры"        ########
'################################################################
Public Function PlnGames_BuildCaption(ByVal titleText As String, ByVal clickCount As Long) As String
    If clickCount <= 0 Then
        PlnGames_BuildCaption = titleText
    Else
        PlnGames_BuildCaption = titleText & " (" & CStr(clickCount) & ")"
    End If
End Function

'################################################################
'########          Построить сообщение кнопки            ########
'################################################################
Public Function PlnGames_BuildMessage(ByVal clickCount As Long, ByVal titleText As String, ByVal lastMessageKey As Long, ByRef newMessageKey As Long) As String
    Dim bodyText As String

    bodyText = GetGamesFixedMessage(clickCount)
    If Len(bodyText) > 0 Then
        newMessageKey = 0
        PlnGames_BuildMessage = bodyText
        Exit Function
    End If

    If IsGamesSpecialCount(clickCount) Then
        newMessageKey = 0
        PlnGames_BuildMessage = "Присвоен титул «" & titleText & "» !!!" & vbCrLf & GetGamesSpecialBody(clickCount)
        Exit Function
    End If

    newMessageKey = GetRandomMessageKeyExcluding(lastMessageKey)
    PlnGames_BuildMessage = GetGamesRandomMessageByKey(newMessageKey)
End Function

'################################################################
'########       Фиксированные сообщения 1..9             ########
'################################################################
Private Function GetGamesFixedMessage(ByVal clickCount As Long) As String
    Select Case clickCount
        Case 1: GetGamesFixedMessage = "Первый клик принят. Официально: разминаемся."
        Case 2: GetGamesFixedMessage = "Второй клик. Рабочий настрой сделал шаг назад."
        Case 3: GetGamesFixedMessage = "Третий клик. Вы действуете уверенно и с огоньком."
        Case 4: GetGamesFixedMessage = "Четвертый клик. Планер сделал вид, что ничего не заметил."
        Case 5: GetGamesFixedMessage = "Пятый клик. Пауза оформлена по всем правилам."
        Case 6: GetGamesFixedMessage = "Шестой клик. Концентрация ушла за кофе."
        Case 7: GetGamesFixedMessage = "Седьмой клик. Режим ""еще чуть-чуть"" активирован."
        Case 8: GetGamesFixedMessage = "Восьмой клик. Кнопка уже узнает ваш почерк."
        Case 9: GetGamesFixedMessage = "Девятый клик. На горизонте юбилейное нажатие."
        Case Else
            GetGamesFixedMessage = vbNullString
    End Select
End Function

'################################################################
'########          Проверка специальных чисел            ########
'################################################################
Private Function IsGamesSpecialCount(ByVal clickCount As Long) As Boolean
    Select Case clickCount
        Case 10, 15, 20, 30, 40, 50, 60, 70, 80, 90, 99
            IsGamesSpecialCount = True
        Case Else
            IsGamesSpecialCount = False
    End Select
End Function

'################################################################
'########         Спец-тексты для юбилейных точек        ########
'################################################################
Private Function GetGamesSpecialBody(ByVal clickCount As Long) As String
    Select Case clickCount
        Case 10
            GetGamesSpecialBody = "10 нажатий! Первый юбилей достигнут, это заявка на стиль."
        Case 15
            GetGamesSpecialBody = "15 нажатий! Полуторадесяток отвлечения выполнен образцово."
        Case 20
            GetGamesSpecialBody = "20 нажатий! Уверенный нажиматор официально на посту."
        Case 30
            GetGamesSpecialBody = "30 нажатий! Серийный режим работает без сбоев."
        Case 40
            GetGamesSpecialBody = "40 нажатий! Мастер-нажиматор набирает обороты."
        Case 50
            GetGamesSpecialBody = "50 нажатий! ЛКМ III степени присвоена."
        Case 60
            GetGamesSpecialBody = "60 нажатий! ЛКМ II степени получена уверенно."
        Case 70
            GetGamesSpecialBody = "70 нажатий! ЛКМ I степени — уже почти легенда."
        Case 80
            GetGamesSpecialBody = "80 нажатий! Ветеран кликов в прекрасной форме."
        Case 90
            GetGamesSpecialBody = "90 нажатий! До заслуженного титула совсем немного."
        Case 99
            GetGamesSpecialBody = "Финальная отметка взята! Вы в ЭЛИТЕ НАЖИМАТОРОВ. Передайте разработчику, что система проверена на максимум."
        Case Else
            GetGamesSpecialBody = "Юбилейный клик принят."
    End Select
End Function

'################################################################
'########   Случайный ключ сообщения без повтора подряд  ########
'################################################################
Private Function GetRandomMessageKeyExcluding(ByVal lastMessageKey As Long) As Long
    Dim nextKey As Long

    If PLN_GAMES_RANDOM_POOL_SIZE <= 1 Then
        GetRandomMessageKeyExcluding = 1
        Exit Function
    End If

    Do
        nextKey = Int(PLN_GAMES_RANDOM_POOL_SIZE * Rnd) + 1
    Loop While nextKey = lastMessageKey

    GetRandomMessageKeyExcluding = nextKey
End Function

'################################################################
'########        Случайные сообщения для пула            ########
'################################################################
Private Function GetGamesRandomMessageByKey(ByVal messageKey As Long) As String
    Select Case messageKey
        Case 1: GetGamesRandomMessageByKey = "Клик засчитан. Список задач слегка напрягся, но держится."
        Case 2: GetGamesRandomMessageByKey = "Нажатие принято. План работ сделал глубокий вдох."
        Case 3: GetGamesRandomMessageByKey = "Еще один клик — и перерыв официально считается полезным."
        Case 4: GetGamesRandomMessageByKey = "Кнопка одобряет ваш стиль микропауз."
        Case 5: GetGamesRandomMessageByKey = "Рабочий ритм на минуту уступил место хорошему настроению."
        Case 6: GetGamesRandomMessageByKey = "Планер молчит, но статистика помнит все."
        Case 7: GetGamesRandomMessageByKey = "Совесть получила уведомление: ""Вернусь через секундочку""."
        Case 8: GetGamesRandomMessageByKey = "Красиво нажато. Даже слишком красиво."
        Case 9: GetGamesRandomMessageByKey = "Это был уверенный клик человека с опытом."
        Case 10: GetGamesRandomMessageByKey = "Кнопка на месте, чувство юмора тоже."
        Case 11: GetGamesRandomMessageByKey = "Небольшое отвлечение выполнено без потери качества."
        Case 12: GetGamesRandomMessageByKey = "Пауза принята. Возврат к делам рекомендован, но не навязывается."
        Case 13: GetGamesRandomMessageByKey = "Планер прикинулся невозмутимым и пропустил нажатие."
        Case 14: GetGamesRandomMessageByKey = "Календарь серьезен, а кнопка — в отличном настроении."
        Case 15: GetGamesRandomMessageByKey = "Вы нажимаете так, будто это отдельный вид спорта."
        Case 16: GetGamesRandomMessageByKey = "Отмечено: еще один стратегический клик."
        Case 17: GetGamesRandomMessageByKey = "Задачи подождут минуту. Возможно, две."
        Case 18: GetGamesRandomMessageByKey = "Кнопка в тонусе. Рабочий настрой догоняет."
        Case 19: GetGamesRandomMessageByKey = "Статистика пополнилась. Настроение тоже."
        Case 20: GetGamesRandomMessageByKey = "Отличный клик: коротко, точно, с душой."
        Case Else
            GetGamesRandomMessageByKey = "Нажатие принято."
    End Select
End Function

