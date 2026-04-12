'################################################################
'########              ОЧИСТКА ТАБЛИЦ БАЗЫ ДАННЫХ       ########
'################################################################

Public Sub ClearAllTables()
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim tableName As String
    Dim db As DAO.Database
    Dim logText As String
    
    ' НАЧАЛО ЛОГА
    logText = "ЛОГ ОЧИСТКИ ТАБЛИЦ БАЗЫ ДАННЫХ" & vbCrLf & _
              "Создан: " & Now & vbCrLf & vbCrLf & _
              "Время операции" & vbTab & "Таблица" & vbTab & "Статус" & vbTab & "Записей удалено" & vbTab & "Ошибка" & vbCrLf & _
              String(80, "-") & vbCrLf
    
    ' ОТКРЫВАЕМ ТЕКУЩУЮ БАЗУ ДАННЫХ
    Set db = CurrentDb
    
    ' МАССИВ ТАБЛИЦ ДЛЯ ОЧИСТКИ (ТОЛЬКО ПОЛЬЗОВАТЕЛЬСКИЕ ДАННЫЕ)
    Dim tablesToClear(1 To 3) As String
    tablesToClear(1) = "tbEventInstances"  ' Основные события
    tablesToClear(2) = "tbExecutors"       ' Исполнители (кроме администратора)
    tablesToClear(3) = "tbTempEvents"      ' Временные события
    
    ' ОЧИСТКА ПОЛЬЗОВАТЕЛЬСКИХ ТАБЛИЦ
    For i = 1 To 3
        tableName = tablesToClear(i)
        logText = logText & ClearSingleTable(db, tableName)
    Next i
    
    ' ЗАПИСЬ О СИСТЕМНЫХ ТАБЛИЦАХ (НЕ ОЧИЩАЛИСЬ)
    logText = logText & Format(Now, "dd.mm.yyyy hh:nn:ss") & vbTab & "tbThemes" & vbTab & "СОХРАНЕНО" & vbTab & "Системная таблица-словарь" & vbTab & "" & vbCrLf
    logText = logText & Format(Now, "dd.mm.yyyy hh:nn:ss") & vbTab & "tbPeriodicity" & vbTab & "СОХРАНЕНО" & vbTab & "Системная таблица-словарь" & vbTab & "" & vbCrLf
    logText = logText & Format(Now, "dd.mm.yyyy hh:nn:ss") & vbTab & "tbRules" & vbTab & "СОХРАНЕНО" & vbTab & "Системная таблица-словарь" & vbTab & "" & vbCrLf
    
    ' ФИНАЛЬНОЕ СООБЩЕНИЕ
    logText = logText & vbCrLf & Format(Now, "dd.mm.yyyy hh:nn:ss") & vbTab & "ВСЕ ТАБЛИЦЫ" & vbTab & "ЗАВЕРШЕНО" & vbTab & "Пользовательские данные очищены" & vbTab & "" & vbCrLf
    
    ' ПОКАЗЫВАЕМ ЛОГ В СООБЩЕНИИ
    MsgBox "Очистка таблиц завершена успешно!" & vbCrLf & vbCrLf & _
           "Системные таблицы-словари сохранены" & vbCrLf & vbCrLf & _
           "Детальный лог:" & vbCrLf & logText, vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при очистке таблиц: " & Err.description, vbCritical
End Sub

'################################################################
'########           ОЧИСТКА ОДНОЙ ТАБЛИЦЫ              ########
'################################################################

Private Function ClearSingleTable(db As DAO.Database, tableName As String) As String
    On Error GoTo ErrorHandler
    
    Dim recordsCount As Long
    Dim sql As String
    Dim logLine As String
    
    ' ПРОВЕРЯЕМ СУЩЕСТВОВАНИЕ ТАБЛИЦЫ
    If Not TableExists(db, tableName) Then
        logLine = Format(Now, "dd.mm.yyyy hh:nn:ss") & vbTab & tableName & vbTab & "ПРОПУЩЕНО" & vbTab & "" & vbTab & "Таблица не существует" & vbCrLf
        ClearSingleTable = logLine
        Exit Function
    End If
    
    ' ПОДСЧИТЫВАЕМ КОЛИЧЕСТВО ЗАПИСЕЙ ДО ОЧИСТКИ
    recordsCount = GetRecordCount(db, tableName)
    
    ' ИНТЕЛЛЕКТУАЛЬНАЯ ОЧИСТКА В ЗАВИСИМОСТИ ОТ ТАБЛИЦЫ
    Select Case tableName
        Case "tbEventInstances"
            sql = "DELETE FROM tbEventInstances"  ' Полная очистка событий
            
        Case "tbExecutors"
            ' Сохраняем только администратора
            sql = "DELETE FROM tbExecutors WHERE LastName <> 'Администратор'"
            
        Case "tbTempEvents"
            sql = "DELETE FROM tbTempEvents"  ' Полная очистка временных событий
            
    End Select
    
    ' ВЫПОЛНЯЕМ SQL ЗАПРОС
    db.Execute sql, dbFailOnError
    
    ' ПОДСЧИТЫВАЕМ СКОЛЬКО ЗАПИСЕЙ УДАЛЕНО
    Dim remainingCount As Long
    remainingCount = GetRecordCount(db, tableName)
    Dim deletedCount As Long
    deletedCount = recordsCount - remainingCount
    
    ' ФОРМИРУЕМ СТРОКУ ЛОГА
    logLine = Format(Now, "dd.mm.yyyy hh:nn:ss") & vbTab & tableName & vbTab & "УСПЕХ" & vbTab & deletedCount & " записей удалено" & vbTab & "" & vbCrLf
    ClearSingleTable = logLine
    
    Exit Function
    
ErrorHandler:
    logLine = Format(Now, "dd.mm.yyyy hh:nn:ss") & vbTab & tableName & vbTab & "ОШИБКА" & vbTab & "" & vbTab & Err.description & vbCrLf
    ClearSingleTable = logLine
End Function

'################################################################
'########           ПРОВЕРКА СУЩЕСТВОВАНИЯ ТАБЛИЦЫ     ########
'################################################################

Private Function TableExists(db As DAO.Database, tableName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim tdf As DAO.TableDef
    TableExists = False
    
    For Each tdf In db.TableDefs
        If tdf.Name = tableName Then
            TableExists = True
            Exit For
        End If
    Next tdf
    
    Exit Function
    
ErrorHandler:
    TableExists = False
End Function

'################################################################
'########         ПОДСЧЕТ КОЛИЧЕСТВА ЗАПИСЕЙ           ########
'################################################################

Private Function GetRecordCount(db As DAO.Database, tableName As String) As Long
    On Error GoTo ErrorHandler
    
    Dim rs As DAO.Recordset
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS RecCount FROM " & tableName)
    
    GetRecordCount = rs!RecCount
    rs.Close
    
    Exit Function
    
ErrorHandler:
    GetRecordCount = 0
End Function

'################################################################
'########        ОТКЛЮЧЕНИЕ РЕЖИМА КОНСТРУКТОРА        ########
'################################################################

Public Sub DisableDesignMode()
    On Error GoTo ErrorHandler
    
    Dim frm As AccessObject
    Dim formsCount As Integer
    Dim processedCount As Integer
    Dim logText As String
    
    ' НАЧАЛО ЛОГА
    logText = "ЛОГ ОТКЛЮЧЕНИЯ РЕЖИМА КОНСТРУКТОРА ФОРМ" & vbCrLf & _
              "Выполнено: " & Now & vbCrLf & vbCrLf & _
              "Форма" & vbTab & "Статус" & vbTab & "Ошибка" & vbCrLf & _
              String(60, "-") & vbCrLf
    
    ' ПОДСЧИТЫВАЕМ КОЛИЧЕСТВО ФОРМ
    formsCount = CurrentProject.allForms.count
    processedCount = 0
    
    ' ОБРАБАТЫВАЕМ КАЖДУЮ ФОРМУ
    For Each frm In CurrentProject.allForms
        processedCount = processedCount + 1
        logText = logText & ProcessFormDesignMode(frm.Name)
    Next frm
    
    ' ФИНАЛЬНОЕ СООБЩЕНИЕ
    logText = logText & vbCrLf & String(60, "-") & vbCrLf & _
              "Обработано форм: " & processedCount & " из " & formsCount & vbCrLf & _
              "Режим конструктора отключен для всех форм"
    
    ' ПОКАЗЫВАЕМ РЕЗУЛЬТАТ
    MsgBox "Режим конструктора отключен для всех форм!" & vbCrLf & vbCrLf & _
           "Детальный лог:" & vbCrLf & logText, vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при отключении режима конструктора: " & Err.description, vbCritical
End Sub

'################################################################
'########      ОБРАБОТКА ОДНОЙ ФОРМЫ                  ########
'################################################################

Private Function ProcessFormDesignMode(formName As String) As String
    On Error GoTo ErrorHandler
    
    Dim logLine As String
    Dim frm As Form
    
    ' ПРОПУСКАЕМ СИСТЕМНЫЕ И СЛУЖЕБНЫЕ ФОРМЫ
    If IsSystemForm(formName) Then
        logLine = formName & vbTab & "ПРОПУЩЕНО" & vbTab & "Системная форма" & vbCrLf
        ProcessFormDesignMode = logLine
        Exit Function
    End If
    
    ' ПЫТАЕМСЯ ОТКРЫТЬ ФОРМУ В РЕЖИМЕ ДИЗАЙНА
    DoCmd.OpenForm formName, acDesign
    
    ' ПОЛУЧАЕМ ОБЪЕКТ ФОРМЫ
    Set frm = Forms(formName)
    
    ' ОТКЛЮЧАЕМ РАЗРЕШЕНИЕ НА ИЗМЕНЕНИЯ
    frm.AllowDesignChanges = False
    
    ' СОХРАНЯЕМ И ЗАКРЫВАЕМ ФОРМУ
    DoCmd.Close acForm, formName, acSaveYes
    
    logLine = formName & vbTab & "УСПЕХ" & vbTab & "Режим конструктора отключен" & vbCrLf
    ProcessFormDesignMode = logLine
    
    Exit Function
    
ErrorHandler:
    logLine = formName & vbTab & "ОШИБКА" & vbTab & Err.description & vbCrLf
    ProcessFormDesignMode = logLine
End Function

'################################################################
'########      ПРОВЕРКА СИСТЕМНОЙ ФОРМЫ               ########
'################################################################
Private Function IsSystemForm(formName As String) As Boolean
    ' СПИСОК СИСТЕМНЫХ ФОРМ, КОТОРЫЕ НЕ НУЖНО ОБРАБАТЫВАТЬ
    Dim systemForms As Variant
    systemForms = Array("MSys", "~", "frmDemo") ' frmDemo оставляем для возможных изменений
    
    Dim i As Integer
    For i = LBound(systemForms) To UBound(systemForms)
        If InStr(1, formName, systemForms(i), vbTextCompare) > 0 Then
            IsSystemForm = True
            Exit Function
        End If
    Next i
    
    IsSystemForm = False
End Function

'################################################################
'########        ВКЛЮЧЕНИЕ РЕЖИМА КОНСТРУКТОРА         ########
'################################################################
Public Sub EnableDesignMode()
    On Error GoTo ErrorHandler
    
    Dim frm As AccessObject
    Dim formsCount As Integer
    Dim processedCount As Integer
    Dim logText As String
    
    ' НАЧАЛО ЛОГА
    logText = "ЛОГ ВКЛЮЧЕНИЯ РЕЖИМА КОНСТРУКТОРА ФОРМ" & vbCrLf & _
              "Выполнено: " & Now & vbCrLf & vbCrLf & _
              "Форма" & vbTab & "Статус" & vbTab & "Ошибка" & vbCrLf & _
              String(60, "-") & vbCrLf
    
    ' ПОДСЧИТЫВАЕМ КОЛИЧЕСТВО ФОРМ
    formsCount = CurrentProject.allForms.count
    processedCount = 0
    
    ' ОБРАБАТЫВАЕМ КАЖДУЮ ФОРМУ
    For Each frm In CurrentProject.allForms
        processedCount = processedCount + 1
        logText = logText & ProcessFormDesignModeEnable(frm.Name)
    Next frm
    
    ' ФИНАЛЬНОЕ СООБЩЕНИЕ
    logText = logText & vbCrLf & String(60, "-") & vbCrLf & _
              "Обработано форм: " & processedCount & " из " & formsCount & vbCrLf & _
              "Режим конструктора включен для всех форм"
    
    ' ПОКАЗЫВАЕМ РЕЗУЛЬТАТ
    MsgBox "Режим конструктора включен для всех форм!" & vbCrLf & vbCrLf & _
           "Детальный лог:" & vbCrLf & logText, vbInformation
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при включении режима конструктора: " & Err.description, vbCritical
End Sub

'################################################################
'########      ВКЛЮЧЕНИЕ РЕЖИМА ДЛЯ ОДНОЙ ФОРМЫ       ########
'################################################################

Private Function ProcessFormDesignModeEnable(formName As String) As String
    On Error GoTo ErrorHandler
    
    Dim logLine As String
    Dim frm As Form
    
    ' ПРОПУСКАЕМ СИСТЕМНЫЕ ФОРМЫ
    If IsSystemForm(formName) Then
        logLine = formName & vbTab & "ПРОПУЩЕНО" & vbTab & "Системная форма" & vbCrLf
        ProcessFormDesignModeEnable = logLine
        Exit Function
    End If
    
    ' ОТКРЫВАЕМ ФОРМУ В РЕЖИМЕ ДИЗАЙНА
    DoCmd.OpenForm formName, acDesign
    
    ' ПОЛУЧАЕМ ОБЪЕКТ ФОРМЫ
    Set frm = Forms(formName)
    
    ' ВКЛЮЧАЕМ РАЗРЕШЕНИЕ НА ИЗМЕНЕНИЯ
    frm.AllowDesignChanges = True
    
    ' СОХРАНЯЕМ И ЗАКРЫВАЕМ ФОРМУ
    DoCmd.Close acForm, formName, acSaveYes
    
    logLine = formName & vbTab & "УСПЕХ" & vbTab & "Режим конструктора включен" & vbCrLf
    ProcessFormDesignModeEnable = logLine
    
    Exit Function
    
ErrorHandler:
    logLine = formName & vbTab & "ОШИБКА" & vbTab & Err.description & vbCrLf
    ProcessFormDesignModeEnable = logLine
End Function

