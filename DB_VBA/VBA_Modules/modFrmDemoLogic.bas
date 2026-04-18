Option Compare Database
Option Explicit

'################################################################
'########   ЛОГИКА ДЕМО-ПОШАГОВ (ВЫЗОВ ИЗ Form_frmDemo)  ########
'################################################################
Public Function FrmDemo_IsFormLoaded(ByVal formName As String) As Boolean
    FrmDemo_IsFormLoaded = CurrentProject.allForms(formName).IsLoaded
End Function

'################################################################
'########            ДЕМО-РЕЖИМ НАВИГАЦИИ                ########
'################################################################
Public Sub FrmDemo_ExecuteNavigationDemo(ByVal frm As Form)
    ' ОБЪЯВЛЕНИЕ ПЕРЕМЕННЫХ
    Dim originalNextBackColor As Long
    Dim originalNextForeColor As Long
    Dim originalPrevBackColor As Long
    Dim originalPrevForeColor As Long
    Dim originalCurrentBackColor As Long
    Dim originalCurrentForeColor As Long
    Dim startTime As Double

    ' 1. ОТКРЫВАЕМ ФОРМУ ЕСЛИ ЗАКРЫТА
    If Not FrmDemo_IsFormLoaded("f_daily_planner") Then
        DoCmd.OpenForm "f_daily_planner"
        DoEvents ' Ждем загрузки формы
    End If

    ' === ПЕРВОЕ НАЖАТИЕ "СЛЕДУЮЩИЙ МЕСЯЦ" ===

    ' 2. СОХРАНЯЕМ ТЕКУЩИЕ ЦВЕТА КНОПКИ "СЛЕДУЮЩИЙ МЕСЯЦ"
    originalNextBackColor = Forms!f_daily_planner.btn_next.backColor
    originalNextForeColor = Forms!f_daily_planner.btn_next.ForeColor

    ' 3. НАЖАТИЕ КНОПКИ "СЛЕДУЮЩИЙ МЕСЯЦ" (1)
    frm.txtCurrentAction.value = "Нажатие: Следующий месяц (1)"
    Debug.Print "Подсветка кнопки: Следующий месяц (1)"
    Forms!f_daily_planner.btn_next.backColor = RGB(255, 255, 0)
    Forms!f_daily_planner.btn_next.ForeColor = RGB(0, 0, 0)
    DoEvents

    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    Forms!f_daily_planner.GoToNextMonth
    Forms!f_daily_planner.btn_next.backColor = originalNextBackColor
    Forms!f_daily_planner.btn_next.ForeColor = originalNextForeColor
    DoEvents

    ' ЗАДЕРЖКА 1 СЕКУНДА
    Debug.Print "Задержка 1 секунда..."
    startTime = Timer
    Do While Timer < startTime + 1
        DoEvents
    Loop

    ' === ВТОРОЕ НАЖАТИЕ "СЛЕДУЮЩИЙ МЕСЯЦ" ===

    ' 4. НАЖАТИЕ КНОПКИ "СЛЕДУЮЩИЙ МЕСЯЦ" (2)
    frm.txtCurrentAction.value = "Нажатие: Следующий месяц (2)"
    Debug.Print "Подсветка кнопки: Следующий месяц (2)"
    Forms!f_daily_planner.btn_next.backColor = RGB(255, 255, 0)
    Forms!f_daily_planner.btn_next.ForeColor = RGB(0, 0, 0)
    DoEvents

    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    Forms!f_daily_planner.GoToNextMonth
    Forms!f_daily_planner.btn_next.backColor = originalNextBackColor
    Forms!f_daily_planner.btn_next.ForeColor = originalNextForeColor
    DoEvents

    ' ЗАДЕРЖКА 1 СЕКУНДА
    Debug.Print "Задержка 1 секунда..."
    startTime = Timer
    Do While Timer < startTime + 1
        DoEvents
    Loop

    ' === НАЖАТИЕ "ПРЕДЫДУЩИЙ МЕСЯЦ" ===

    ' 5. СОХРАНЯЕМ ТЕКУЩИЕ ЦВЕТА КНОПКИ "ПРЕДЫДУЩИЙ МЕСЯЦ"
    originalPrevBackColor = Forms!f_daily_planner.btn_previous.backColor
    originalPrevForeColor = Forms!f_daily_planner.btn_previous.ForeColor

    ' 6. НАЖАТИЕ КНОПКИ "ПРЕДЫДУЩИЙ МЕСЯЦ"
    frm.txtCurrentAction.value = "Нажатие: Предыдущий месяц"
    Debug.Print "Подсветка кнопки: Предыдущий месяц"
    Forms!f_daily_planner.btn_previous.backColor = RGB(255, 255, 0)
    Forms!f_daily_planner.btn_previous.ForeColor = RGB(0, 0, 0)
    DoEvents

    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    Forms!f_daily_planner.GoToPreviousMonth
    Forms!f_daily_planner.btn_previous.backColor = originalPrevBackColor
    Forms!f_daily_planner.btn_previous.ForeColor = originalPrevForeColor
    DoEvents

    ' ЗАДЕРЖКА 1 СЕКУНДА
    Debug.Print "Задержка 1 секунда..."
    startTime = Timer
    Do While Timer < startTime + 1
        DoEvents
    Loop

    ' === ВОЗВРАТ К ТЕКУЩЕМУ МЕСЯЦУ ===

    ' 7. СОХРАНЯЕМ ТЕКУЩИЕ ЦВЕТА КНОПКИ "ТЕКУЩИЙ МЕСЯЦ"
    originalCurrentBackColor = Forms!f_daily_planner.btn_current.backColor
    originalCurrentForeColor = Forms!f_daily_planner.btn_current.ForeColor

    ' 8. НАЖАТИЕ КНОПКИ "ТЕКУЩИЙ МЕСЯЦ"
    frm.txtCurrentAction.value = "Нажатие: Текущий месяц"
    Debug.Print "Подсветка кнопки: Текущий месяц"
    Forms!f_daily_planner.btn_current.backColor = RGB(255, 255, 0)
    Forms!f_daily_planner.btn_current.ForeColor = RGB(0, 0, 0)
    DoEvents

    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    Forms!f_daily_planner.GoToCurrentMonth
    Forms!f_daily_planner.btn_current.backColor = originalCurrentBackColor
    Forms!f_daily_planner.btn_current.ForeColor = originalCurrentForeColor
    DoEvents
    ' ЗАВЕРШАЮЩЕЕ СООБЩЕНИЕ
    frm.txtCurrentAction.value = "Демонстрация навигации завершена" & vbCrLf & "Нажмите ""Далее"""
    DoEvents
End Sub

'################################################################
'########            ДЕМО-РЕЖИМ СОБЫТИЙ                  ########
'################################################################

Public Sub FrmDemo_ExecuteEventsDemo(ByVal frm As Form)
    Call FrmDemo_DemoHighlightDay(frm)
    Call FrmDemo_DemoOpenEventsForm(frm)
    Call FrmDemo_DemoEditAndFillEvent(frm)
    Call FrmDemo_DemoCloseEventsForm(frm)  ' < ДОБАВЛЯЕМ ЗАКРЫТИЕ
    Call FrmDemo_DemoRestoreDay(frm)
    ' ЗАВЕРШАЮЩЕЕ СООБЩЕНИЕ
    frm.txtCurrentAction.value = "Демонстрация работы с событиями завершена" & vbCrLf & "Нажмите ""Далее"""
    DoEvents
End Sub

'################################################################
'########        1. ПОДСВЕТКА ДНЯ В КАЛЕНДАРЕ           ########
'################################################################

Public Sub FrmDemo_DemoHighlightDay(ByVal frm As Form)
    ' ОБЪЯВЛЕНИЕ ПЕРЕМЕННЫХ
    Dim originalDayBackColor As Long
    Dim originalDayBorderColor As Long
    Dim originalBorderWidth As Integer
    Dim startTime As Double
    Dim startDate As Date
    Dim DayNumber As Integer
    Dim monthYear As String

    ' 1. ВЫЧИСЛЯЕМ НОМЕР ПОЛЯ ДЛЯ СЕГОДНЯШНЕЙ ДАТЫ
    monthYear = Forms!f_daily_planner.lbl_MonthYear.Caption
    startDate = DateSerial(Year(monthYear), Month(monthYear), 1)
    startDate = startDate - weekday(startDate, vbMonday) + 1
    DayNumber = DateDiff("d", startDate, Date) + 1

    ' 2. СОХРАНЯЕМ ТЕКУЩИЕ СВОЙСТВА ПОЛЯ ДНЯ
    originalDayBackColor = Forms!f_daily_planner.Controls("fld_day_" & DayNumber).backColor
    originalDayBorderColor = Forms!f_daily_planner.Controls("fld_day_" & DayNumber).borderColor
    originalBorderWidth = Forms!f_daily_planner.Controls("fld_day_" & DayNumber).borderWidth

    ' 3. ПОДСВЕТКА ПОЛЯ ТЕКУЩЕГО ДНЯ
    frm.txtCurrentAction.value = "Двойной клик: Открытие событий дня"
    Forms!f_daily_planner.Controls("fld_day_" & DayNumber).backColor = RGB(255, 255, 0)
    Forms!f_daily_planner.Controls("fld_day_" & DayNumber).borderColor = RGB(255, 0, 0)
    Forms!f_daily_planner.Controls("fld_day_" & DayNumber).borderWidth = 3
    DoEvents

    ' 4. ЗАДЕРЖКА 0.5 СЕКУНДЫ
    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    ' Сохраняем данные для следующих процедур
    Call FrmDemo_SaveDayData(frm, DayNumber, originalDayBackColor, originalDayBorderColor, originalBorderWidth)
End Sub

'################################################################
'########         2. ОТКРЫТИЕ ФОРМЫ СОБЫТИЙ              ########
'################################################################

Public Sub FrmDemo_DemoOpenEventsForm(ByVal frm As Form)
    ' Получаем сохраненные данные
    Dim dayData As Variant
    dayData = FrmDemo_GetDayData(frm)
    Dim DayNumber As Integer
    DayNumber = dayData(0)

    ' ОТКРЫВАЕМ ФОРМУ СОБЫТИЙ
    Forms!f_daily_planner.OpenDayEvents DayNumber
    DoEvents
End Sub

'################################################################
'########        3.1 НАВИГАЦИЯ ПО ДНЯМ                   ########
'################################################################

Public Sub FrmDemo_DemoNavigateDays(ByVal frm As Form)
    If Not FrmDemo_IsFormLoaded("frmDayEvents") Then Exit Sub

    Dim originalNextDayBackColor As Long
    Dim originalNextDayForeColor As Long
    Dim originalPrevDayBackColor As Long
    Dim originalPrevDayForeColor As Long
    Dim startTime As Double

    ' ПОДСВЕТКА И НАЖАТИЕ "СЛЕДУЮЩИЙ ДЕНЬ"
    frm.txtCurrentAction.value = "Навигация: Следующий день"

    originalNextDayBackColor = Forms!frmDayEvents.cmdNextDay.backColor
    originalNextDayForeColor = Forms!frmDayEvents.cmdNextDay.ForeColor

    Forms!frmDayEvents.cmdNextDay.backColor = RGB(255, 255, 0)
    Forms!frmDayEvents.cmdNextDay.ForeColor = RGB(0, 0, 0)
    DoEvents

    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    Forms!frmDayEvents.GoToNextDay
    DoEvents

    Forms!frmDayEvents.cmdNextDay.backColor = originalNextDayBackColor
    Forms!frmDayEvents.cmdNextDay.ForeColor = originalNextDayForeColor
    DoEvents

    ' ЗАДЕРЖКА 1 СЕКУНДА
    startTime = Timer
    Do While Timer < startTime + 1
        DoEvents
    Loop

    ' ПОДСВЕТКА И НАЖАТИЕ "ПРЕДЫДУЩИЙ ДЕНЬ"
    frm.txtCurrentAction.value = "Навигация: Предыдущий день"

    originalPrevDayBackColor = Forms!frmDayEvents.cmdPrevDay.backColor
    originalPrevDayForeColor = Forms!frmDayEvents.cmdPrevDay.ForeColor

    Forms!frmDayEvents.cmdPrevDay.backColor = RGB(255, 255, 0)
    Forms!frmDayEvents.cmdPrevDay.ForeColor = RGB(0, 0, 0)
    DoEvents

    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    Forms!frmDayEvents.GoToPreviousDay
    DoEvents

    Forms!frmDayEvents.cmdPrevDay.backColor = originalPrevDayBackColor
    Forms!frmDayEvents.cmdPrevDay.ForeColor = originalPrevDayForeColor
    DoEvents

    ' ЗАДЕРЖКА 1 СЕКУНДА
    startTime = Timer
    Do While Timer < startTime + 1
        DoEvents
    Loop
End Sub

'################################################################
'########   3.2 РЕДАКТИРОВАНИЕ И СОХРАНЕНИЕ           ########
'################################################################

Public Sub FrmDemo_DemoEditAndFillEvent(ByVal frm As Form)
    If Not FrmDemo_IsFormLoaded("frmDayEvents") Then Exit Sub

    Dim originalEditBackColor As Long
    Dim originalEditForeColor As Long
    Dim originalEventNoteBackColor As Long
    Dim originalEventNoteForeColor As Long
    Dim startTime As Double

    ' 1. КНОПКА РЕДАКТИРОВАНИЯ
    Debug.Print "ДЕМО: Режим редактирования"

    originalEditBackColor = Forms!frmDayEvents.cmdEdit.backColor
    originalEditForeColor = Forms!frmDayEvents.cmdEdit.ForeColor

    ' ПОДСВЕТКА
    Forms!frmDayEvents.cmdEdit.backColor = RGB(255, 255, 0)
    Forms!frmDayEvents.cmdEdit.ForeColor = RGB(0, 0, 0)
    DoEvents

    ' ЗАДЕРЖКА ДЛЯ ВИЗУАЛИЗАЦИИ ПОДСВЕТКИ
    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    ' ВОССТАНАВЛИВАЕМ ЦВЕТА СРАЗУ ПОСЛЕ ПОДСВЕТКИ
    Forms!frmDayEvents.cmdEdit.backColor = originalEditBackColor
    Forms!frmDayEvents.cmdEdit.ForeColor = originalEditForeColor
    DoEvents

    ' ЗАДЕРЖКА ПЕРЕД ВЫПОЛНЕНИЕМ ДЕЙСТВИЯ
    startTime = Timer
    Do While Timer < startTime + 0.3
        DoEvents
    Loop

    ' ВЫПОЛНЯЕМ ДЕЙСТВИЕ (кнопка уже в исходном состоянии)
    Forms!frmDayEvents.StartEditMode
    DoEvents

    ' ЗАДЕРЖКА 1 СЕКУНДА
    startTime = Timer
    Do While Timer < startTime + 1
        DoEvents
    Loop

    ' 2. СОЗДАНИЕ НОВЫХ СОБЫТИЙ В БАЗЕ
    Debug.Print "ДЕМО: Создание тестовых событий в базе"

    Dim currentDate As Date
    currentDate = CDate(Replace(Forms!frmDayEvents.lblDate.Caption, " г.", ""))

    ' Получаем первого исполнителя для демо
    Dim demoExecutorID As Variant
    demoExecutorID = Null
    If Forms!frmDayEvents.cboExecutor.ListCount > 0 Then
        demoExecutorID = Forms!frmDayEvents.cboExecutor.ItemData(0)
    End If

    ' СОБЫТИЕ 1: НЕВЫПОЛНЕННОЕ
    If Not IsNull(demoExecutorID) Then
        CurrentDb.Execute "INSERT INTO tbEventInstances (EventDate, EventNote, Notes, ExecutorID) " & _
                         "VALUES (#" & Format(currentDate, "yyyy\/mm\/dd") & "#, " & _
                         "'Тестовое событие из демо-режима', " & _
                         "'Тестовое примечание из демо-режима', " & demoExecutorID & ")"
    Else
        CurrentDb.Execute "INSERT INTO tbEventInstances (EventDate, EventNote, Notes) " & _
                         "VALUES (#" & Format(currentDate, "yyyy\/mm\/dd") & "#, " & _
                         "'Тестовое событие из демо-режима', " & _
                         "'Тестовое примечание из демо-режима')"
    End If

    ' СОБЫТИЕ 2: ВЫПОЛНЕННОЕ (с отметкой CompletionMark)
    If Not IsNull(demoExecutorID) Then
        CurrentDb.Execute "INSERT INTO tbEventInstances (EventDate, EventNote, Notes, ExecutorID, CompletionMark, CompletionDate) " & _
                         "VALUES (#" & Format(currentDate, "yyyy\/mm\/dd") & "#, " & _
                         "'Тестовое событие', " & _
                         "'Это событие уже выполнено', " & demoExecutorID & ", " & _
                         "'Выполнено', #" & Format(Date, "yyyy\/mm\/dd") & "#)"
    Else
        CurrentDb.Execute "INSERT INTO tbEventInstances (EventDate, EventNote, Notes, CompletionMark, CompletionDate) " & _
                         "VALUES (#" & Format(currentDate, "yyyy\/mm\/dd") & "#, " & _
                         "'ВЫПОЛНЕНО: Тестовое событие', " & _
                         "'Это событие уже выполнено', " & _
                         "'Выполнено', #" & Format(Date, "yyyy\/mm\/dd") & "#)"
    End If

    ' Обновляем форму чтобы показать обе записи
    Forms!frmDayEvents.Requery
    DoEvents

    ' 3. ПОДСВЕТКА НОВОЙ ЗАПИСИ В ФОРМЕ
    Debug.Print "ДЕМО: Подсветка созданной записи"

    ' Переходим к последней записи (нашей новой)
    If Forms!frmDayEvents.Recordset.recordCount > 0 Then
        Forms!frmDayEvents.Recordset.MoveLast

        ' Подсвечиваем поле события
        originalEventNoteBackColor = Forms!frmDayEvents.txtEventNote.backColor
        originalEventNoteForeColor = Forms!frmDayEvents.txtEventNote.ForeColor

        Forms!frmDayEvents.txtEventNote.backColor = RGB(255, 255, 0)
        Forms!frmDayEvents.txtEventNote.ForeColor = RGB(0, 0, 0)
        DoEvents

        startTime = Timer
        Do While Timer < startTime + 0.5
            DoEvents
        Loop

        ' Восстанавливаем цвет
        Forms!frmDayEvents.txtEventNote.backColor = originalEventNoteBackColor
        Forms!frmDayEvents.txtEventNote.ForeColor = originalEventNoteForeColor
        DoEvents

        ' ПОДСВЕТКА КОМБОБОКСА ИСПОЛНИТЕЛЯ (если есть исполнитель)
        If Not IsNull(demoExecutorID) Then
            Dim originalExecutorBackColor As Long
            Dim originalExecutorForeColor As Long

            originalExecutorBackColor = Forms!frmDayEvents.cboExecutor.backColor
            originalExecutorForeColor = Forms!frmDayEvents.cboExecutor.ForeColor

            Forms!frmDayEvents.cboExecutor.backColor = RGB(255, 255, 0)
            Forms!frmDayEvents.cboExecutor.ForeColor = RGB(0, 0, 0)
            DoEvents

            startTime = Timer
            Do While Timer < startTime + 0.5
                DoEvents
            Loop

            Forms!frmDayEvents.cboExecutor.backColor = originalExecutorBackColor
            Forms!frmDayEvents.cboExecutor.ForeColor = originalExecutorForeColor
            DoEvents
        End If
    End If

    ' ФИНАЛЬНАЯ ЗАДЕРЖКА
    Debug.Print "ДЕМО: Событие создано в базе! Форма остается открытой"
    startTime = Timer
    Do While Timer < startTime + 3
        DoEvents
    Loop
End Sub

'################################################################
'########     3.3 ЗАПОЛНЕНИЕ ПОЛЕЙ СОБЫТИЯ             ########
'################################################################

Public Sub FrmDemo_DemoFillEventFields(ByVal frm As Form)
    If Not FrmDemo_IsFormLoaded("frmDayEvents") Then Exit Sub

    Dim originalEventNoteBackColor As Long
    Dim originalEventNoteForeColor As Long
    Dim originalNotesBackColor As Long
    Dim originalNotesForeColor As Long
    Dim startTime As Double

    ' ДОБАВЛЯЕМ/РЕДАКТИРУЕМ ЗАПИСЬ
    If Forms!frmDayEvents.Recordset.recordCount = 0 Then
        Forms!frmDayEvents.Recordset.AddNew
    Else
        Forms!frmDayEvents.Recordset.Edit
    End If

    ' ЗАПОЛНЕНИЕ ОСНОВНОГО СОБЫТИЯ
    frm.txtCurrentAction.value = "Заполнение тестового события"

    originalEventNoteBackColor = Forms!frmDayEvents.txtEventNote.backColor
    originalEventNoteForeColor = Forms!frmDayEvents.txtEventNote.ForeColor

    Forms!frmDayEvents.txtEventNote.backColor = RGB(255, 255, 0)
    Forms!frmDayEvents.txtEventNote.ForeColor = RGB(0, 0, 0)
    DoEvents

    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    Forms!frmDayEvents!EventNote = "Тестовое событие из демо-режима"
    DoEvents
    Forms!frmDayEvents.cboExecutor.SetFocus
    DoEvents

    Forms!frmDayEvents.txtEventNote.backColor = originalEventNoteBackColor
    Forms!frmDayEvents.txtEventNote.ForeColor = originalEventNoteForeColor
    DoEvents

    ' ЗАДЕРЖКА 0.5 СЕКУНДЫ
    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    ' ЗАПОЛНЕНИЕ ПРИМЕЧАНИЙ
    frm.txtCurrentAction.value = "Заполнение примечаний"

    originalNotesBackColor = Forms!frmDayEvents.txtNotes.backColor
    originalNotesForeColor = Forms!frmDayEvents.txtNotes.ForeColor

    Forms!frmDayEvents.txtNotes.backColor = RGB(255, 255, 0)
    Forms!frmDayEvents.txtNotes.ForeColor = RGB(0, 0, 0)
    DoEvents

    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    Forms!frmDayEvents!Notes = "Тестовое примечание из демо-режима"
    DoEvents
    Forms!frmDayEvents.cboExecutor.SetFocus
    DoEvents

    Forms!frmDayEvents.txtNotes.backColor = originalNotesBackColor
    Forms!frmDayEvents.txtNotes.ForeColor = originalNotesForeColor
    DoEvents

    ' ЗАДЕРЖКА 0.5 СЕКУНДЫ
    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    ' ЗАПОЛНЕНИЕ ИСПОЛНИТЕЛЯ
    If Forms!frmDayEvents.cboExecutor.ListCount > 0 Then
        Forms!frmDayEvents!cboExecutor = Forms!frmDayEvents.cboExecutor.ItemData(0)
        DoEvents
        Forms!frmDayEvents.cmdSave.SetFocus
        DoEvents
    End If
End Sub

'################################################################
'########     3.4 СОХРАНЕНИЕ И ЗАКРЫТИЕ                ########
'################################################################

Public Sub FrmDemo_DemoSaveAndClose(ByVal frm As Form)
    If Not FrmDemo_IsFormLoaded("frmDayEvents") Then Exit Sub

    Dim originalSaveBackColor As Long
    Dim originalSaveForeColor As Long
    Dim startTime As Double

    ' СОХРАНЕНИЕ ИЗМЕНЕНИЙ
    frm.txtCurrentAction.value = "Сохранение изменений"

    originalSaveBackColor = Forms!frmDayEvents.cmdSave.backColor
    originalSaveForeColor = Forms!frmDayEvents.cmdSave.ForeColor

    Forms!frmDayEvents.cmdSave.backColor = RGB(255, 255, 0)
    Forms!frmDayEvents.cmdSave.ForeColor = RGB(0, 0, 0)
    DoEvents

    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    Forms!frmDayEvents.SaveChanges
    DoEvents

    Forms!frmDayEvents.cmdSave.backColor = originalSaveBackColor
    Forms!frmDayEvents.cmdSave.ForeColor = originalSaveForeColor
    DoEvents

    ' ЗАДЕРЖКА 2 СЕКУНДЫ
    startTime = Timer
    Do While Timer < startTime + 2
        DoEvents
    Loop

    ' ЗАКРЫВАЕМ ФОРМУ
    DoCmd.Close acForm, "frmDayEvents"
    DoEvents
End Sub

'################################################################
'########        3.5 ЗАКРЫТИЕ ФОРМЫ СОБЫТИЙ              ########
'################################################################

Public Sub FrmDemo_DemoCloseEventsForm(ByVal frm As Form)
    If Not FrmDemo_IsFormLoaded("frmDayEvents") Then Exit Sub

    Dim originalCloseBackColor As Long
    Dim originalCloseForeColor As Long
    Dim startTime As Double

    Debug.Print "ДЕМО: Закрытие формы событий"

    originalCloseBackColor = Forms!frmDayEvents.cmdClose.backColor
    originalCloseForeColor = Forms!frmDayEvents.cmdClose.ForeColor

    Forms!frmDayEvents.cmdClose.backColor = RGB(255, 255, 0)
    Forms!frmDayEvents.cmdClose.ForeColor = RGB(0, 0, 0)
    DoEvents

    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    ' ИСПОЛЬЗУЕМ ПУБЛИЧНЫЙ МЕТОД CloseForm (а не cmdClose_Click)
    Forms!frmDayEvents.CloseForm
    DoEvents

    ' Восстанавливаем цвета (форма уже может быть закрыта, поэтому с проверкой)
    If FrmDemo_IsFormLoaded("frmDayEvents") Then
        Forms!frmDayEvents.cmdClose.backColor = originalCloseBackColor
        Forms!frmDayEvents.cmdClose.ForeColor = originalCloseForeColor
        DoEvents
    End If

    ' ЗАДЕРЖКА
    startTime = Timer
    Do While Timer < startTime + 1
        DoEvents
    Loop

    Debug.Print "ДЕМО: Форма событий закрыта"
End Sub

'################################################################
'########        4. ВОССТАНОВЛЕНИЕ ДНЯ                 ########
'################################################################

Public Sub FrmDemo_DemoRestoreDay(ByVal frm As Form)
    ' Получаем сохраненные данные
    Dim dayData As Variant
    dayData = FrmDemo_GetDayData(frm)
    Dim DayNumber As Integer
    Dim originalDayBackColor As Long
    Dim originalDayBorderColor As Long
    Dim originalBorderWidth As Integer

    DayNumber = dayData(0)
    originalDayBackColor = dayData(1)
    originalDayBorderColor = dayData(2)
    originalBorderWidth = dayData(3)

    ' ВОССТАНАВЛИВАЕМ СВОЙСТВА
    Forms!f_daily_planner.Controls("fld_day_" & DayNumber).backColor = originalDayBackColor
    Forms!f_daily_planner.Controls("fld_day_" & DayNumber).borderColor = originalDayBorderColor
    Forms!f_daily_planner.Controls("fld_day_" & DayNumber).borderWidth = originalBorderWidth
    DoEvents
End Sub

'################################################################
'########        ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ                 ########
'################################################################

' Сохранение данных дня в глобальные переменные
Public Sub FrmDemo_SaveDayData(ByVal frm As Form, DayNumber As Integer, backColor As Long, borderColor As Long, borderWidth As Integer)
    ' Используем форму для хранения временных данных
    frm.Tag = DayNumber & "|" & backColor & "|" & borderColor & "|" & borderWidth
End Sub

' Получение сохраненных данных дня
Public Function FrmDemo_GetDayData(ByVal frm As Form) As Variant
    Dim dataArray() As String
    dataArray = Split(frm.Tag, "|")
    FrmDemo_GetDayData = dataArray
End Function

'################################################################
'########             ДЕМО-РЕЖИМ ФИЛЬТРАЦИИ              ########
'################################################################

Public Sub FrmDemo_ExecuteFilterDemo(ByVal frm As Form)
    ' ОСНОВНАЯ ПРОЦЕДУРА - ВЫЗЫВАЕТ ЧАСТИ
    Call FrmDemo_DemoHighlightExecutorFilter(frm)
    Call FrmDemo_DemoSelectExecutor(frm)
    Call FrmDemo_DemoHideCompletedEvents(frm)
    Call FrmDemo_DemoResetExecutorFilter(frm)
    Call FrmDemo_DemoResetCompletedFilter(frm)
    ' ЗАВЕРШАЮЩЕЕ СООБЩЕНИЕ
    frm.txtCurrentAction.value = "Демонстрация фильтрации завершена" & vbCrLf & "Нажмите ""Далее"""
    DoEvents
End Sub

'################################################################
'########     4.1 ПОДСВЕТКА ФИЛЬТРА ИСПОЛНИТЕЛЕЙ         ########
'################################################################

Public Sub FrmDemo_DemoHighlightExecutorFilter(ByVal frm As Form)
    Dim originalComboBackColor As Long
    Dim originalComboForeColor As Long
    Dim startTime As Double

    Debug.Print "ДЕМО: Подсветка фильтра исполнителей"

    ' СООБЩЕНИЕ ПЕРЕД ПОДСВЕТКОЙ
    frm.txtCurrentAction.value = "Фильтрация: Выбор исполнителя"
    DoEvents

    ' 1. ОТКРЫВАЕМ ФОРМУ ЕСЛИ ЗАКРЫТА
    If Not FrmDemo_IsFormLoaded("f_daily_planner") Then
        DoCmd.OpenForm "f_daily_planner"
        DoEvents
    End If

    ' 2. СОХРАНЯЕМ ОРИГИНАЛЬНЫЕ ЦВЕТА
    originalComboBackColor = Forms!f_daily_planner.cboExecutorFilter.backColor
    originalComboForeColor = Forms!f_daily_planner.cboExecutorFilter.ForeColor

    ' 3. ПОДСВЕТКА КОМБОБОКСА
    Forms!f_daily_planner.cboExecutorFilter.backColor = RGB(255, 255, 0)
    Forms!f_daily_planner.cboExecutorFilter.ForeColor = RGB(0, 0, 0)
    DoEvents

    ' 4. ЗАДЕРЖКА ДЛЯ ВИЗУАЛИЗАЦИИ
    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    ' 5. ВОССТАНАВЛИВАЕМ ЦВЕТА
    Forms!f_daily_planner.cboExecutorFilter.backColor = originalComboBackColor
    Forms!f_daily_planner.cboExecutorFilter.ForeColor = originalComboForeColor
    DoEvents

    ' 6. ЗАДЕРЖКА ПЕРЕД СЛЕДУЮЩИМ ДЕЙСТВИЕМ
    startTime = Timer
    Do While Timer < startTime + 0.3
        DoEvents
    Loop

    Debug.Print "ДЕМО: Фильтр исполнителей подсвечен"
End Sub

'################################################################
'########   4.2 ВЫБОР ИСПОЛНИТЕЛЯ ИЗ ФИЛЬТРА          ########
'################################################################

Public Sub FrmDemo_DemoSelectExecutor(ByVal frm As Form)
    Dim startTime As Double

    Debug.Print "ДЕМО: Выбор исполнителя из фильтра"

    ' 1. ПРОВЕРЯЕМ ЧТО В СПИСКЕ ЕСТЬ ИСПОЛНИТЕЛИ
    If Forms!f_daily_planner.cboExecutorFilter.ListCount = 0 Then
        Debug.Print "ДЕМО: Нет исполнителей для выбора"
        Exit Sub
    End If

    ' 2. ВЫБИРАЕМ ПЕРВОГО ИСПОЛНИТЕЛЯ ИЗ СПИСКА
    Forms!f_daily_planner.cboExecutorFilter = Forms!f_daily_planner.cboExecutorFilter.ItemData(0)
    DoEvents

    ' 3. ПРИНУДИТЕЛЬНО ВЫЗЫВАЕМ ПУБЛИЧНЫЙ МЕТОД ПРИМЕНЕНИЯ ФИЛЬТРА
    Forms!f_daily_planner.ApplyExecutorFilter
    DoEvents

    ' 4. ЗАДЕРЖКА ДЛЯ ОБРАБОТКИ ВЫБОРА И ПЕРЕСТРОЙКИ КАЛЕНДАРЯ
    startTime = Timer
    Do While Timer < startTime + 2
        DoEvents
    Loop

    Debug.Print "ДЕМО: Исполнитель выбран и фильтр применен"
End Sub

'################################################################
'########     4.3 ФИЛЬТРАЦИЯ ПО ВЫПОЛНЕННЫМ СОБЫТИЯМ     ########
'################################################################
Public Sub FrmDemo_DemoHideCompletedEvents(ByVal frm As Form)
    Dim originalLabelBackColor As Long
    Dim originalLabelForeColor As Long
    Dim originalLabelBackStyle As Integer
    Dim startTime As Double

    Debug.Print "ДЕМО: Фильтрация выполненных событий"

    ' СООБЩЕНИЕ ПЕРЕД ПОДСВЕТКОЙ
    frm.txtCurrentAction.value = "Фильтрация: Скрыть выполненные события"
    DoEvents

    ' 1. СОХРАНЯЕМ ОРИГИНАЛЬНЫЕ СВОЙСТВА ПОДПИСИ
    originalLabelBackColor = Forms!f_daily_planner.lblChkHideCompleted.backColor
    originalLabelForeColor = Forms!f_daily_planner.lblChkHideCompleted.ForeColor
    originalLabelBackStyle = Forms!f_daily_planner.lblChkHideCompleted.BackStyle

    ' 2. ПОДСВЕТКА ПОДПИСИ - ДЕЛАЕМ НЕПРОЗРАЧНЫЙ ЖЕЛТЫЙ ФОН
    Forms!f_daily_planner.lblChkHideCompleted.backColor = RGB(255, 255, 0)
    Forms!f_daily_planner.lblChkHideCompleted.ForeColor = RGB(0, 0, 0)
    Forms!f_daily_planner.lblChkHideCompleted.BackStyle = 1
    DoEvents

    ' 3. ЗАДЕРЖКА ДЛЯ ВИЗУАЛИЗАЦИИ ПОДСВЕТКИ
    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    ' 4. ВОССТАНАВЛИВАЕМ ОРИГИНАЛЬНЫЕ СВОЙСТВА ПОДПИСИ
    Forms!f_daily_planner.lblChkHideCompleted.backColor = originalLabelBackColor
    Forms!f_daily_planner.lblChkHideCompleted.ForeColor = originalLabelForeColor
    Forms!f_daily_planner.lblChkHideCompleted.BackStyle = originalLabelBackStyle
    DoEvents

    ' 5. ЗАДЕРЖКА ПЕРЕД ВКЛЮЧЕНИЕМ
    startTime = Timer
    Do While Timer < startTime + 0.3
        DoEvents
    Loop

    ' 6. ВКЛЮЧАЕМ ЧЕКБОКС (если еще не включен)
    If Not Forms!f_daily_planner.chkHideCompleted Then
        Forms!f_daily_planner.chkHideCompleted = True
        DoEvents

        ' 7. ВЫЗЫВАЕМ ПУБЛИЧНЫЙ МЕТОД ПРИМЕНЕНИЯ ФИЛЬТРА
        Forms!f_daily_planner.ApplyHideCompletedFilter
        DoEvents
    End If

    ' 8. ЗАДЕРЖКА ДЛЯ ПЕРЕСТРОЙКИ КАЛЕНДАРЯ
    startTime = Timer
    Do While Timer < startTime + 2
        DoEvents
    Loop

    Debug.Print "ДЕМО: Фильтр выполненных событий применен"
End Sub

'################################################################
'########      4.4 СБРОС ФИЛЬТРА ИСПОЛНИТЕЛЯ             ########
'################################################################

Public Sub FrmDemo_DemoResetExecutorFilter(ByVal frm As Form)
    Dim startTime As Double

    Debug.Print "ДЕМО: Сброс фильтра исполнителя"

    ' СООБЩЕНИЕ ПЕРЕД СБРОСОМ
    frm.txtCurrentAction.value = "Сброс фильтра: Все исполнители"
    DoEvents

    ' 1. ПОДСВЕТКА КОМБОБОКСА ПЕРЕД СБРОСОМ
    Dim originalComboBackColor As Long
    Dim originalComboForeColor As Long

    originalComboBackColor = Forms!f_daily_planner.cboExecutorFilter.backColor
    originalComboForeColor = Forms!f_daily_planner.cboExecutorFilter.ForeColor

    Forms!f_daily_planner.cboExecutorFilter.backColor = RGB(255, 255, 0)
    Forms!f_daily_planner.cboExecutorFilter.ForeColor = RGB(0, 0, 0)
    DoEvents

    ' 2. ЗАДЕРЖКА ДЛЯ ВИЗУАЛИЗАЦИИ
    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    ' 3. ВОССТАНАВЛИВАЕМ ЦВЕТА
    Forms!f_daily_planner.cboExecutorFilter.backColor = originalComboBackColor
    Forms!f_daily_planner.cboExecutorFilter.ForeColor = originalComboForeColor
    DoEvents

    ' 4. ЗАДЕРЖКА ПЕРЕД СБРОСОМ
    startTime = Timer
    Do While Timer < startTime + 0.3
        DoEvents
    Loop

    ' 5. СБРАСЫВАЕМ ФИЛЬТР ИСПОЛНИТЕЛЯ (устанавливаем пустое значение)
    Forms!f_daily_planner.cboExecutorFilter = Null
    DoEvents

    ' 6. ВЫЗЫВАЕМ ПРИМЕНЕНИЕ ФИЛЬТРА
    Forms!f_daily_planner.ApplyExecutorFilter
    DoEvents

    ' 7. ЗАДЕРЖКА ДЛЯ ПЕРЕСТРОЙКИ КАЛЕНДАРЯ
    startTime = Timer
    Do While Timer < startTime + 2
        DoEvents
    Loop

    Debug.Print "ДЕМО: Фильтр исполнителя сброшен"
End Sub

'################################################################
'########      4.5 СБРОС ФИЛЬТРА ВЫПОЛНЕННЫХ             ########
'################################################################

Public Sub FrmDemo_DemoResetCompletedFilter(ByVal frm As Form)
    Dim originalLabelBackColor As Long
    Dim originalLabelForeColor As Long
    Dim originalLabelBackStyle As Integer
    Dim startTime As Double

    Debug.Print "ДЕМО: Сброс фильтра выполненных событий"

    ' СООБЩЕНИЕ ПЕРЕД СБРОСОМ
    frm.txtCurrentAction.value = "Сброс фильтра: Показать все события"
    DoEvents

    ' 1. ПОДСВЕТКА ПОДПИСИ ЧЕКБОКСА
    originalLabelBackColor = Forms!f_daily_planner.lblChkHideCompleted.backColor
    originalLabelForeColor = Forms!f_daily_planner.lblChkHideCompleted.ForeColor
    originalLabelBackStyle = Forms!f_daily_planner.lblChkHideCompleted.BackStyle

    Forms!f_daily_planner.lblChkHideCompleted.backColor = RGB(255, 255, 0)
    Forms!f_daily_planner.lblChkHideCompleted.ForeColor = RGB(0, 0, 0)
    Forms!f_daily_planner.lblChkHideCompleted.BackStyle = 1
    DoEvents

    ' 2. ЗАДЕРЖКА ДЛЯ ВИЗУАЛИЗАЦИИ
    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    ' 3. ВОССТАНАВЛИВАЕМ СВОЙСТВА ПОДПИСИ
    Forms!f_daily_planner.lblChkHideCompleted.backColor = originalLabelBackColor
    Forms!f_daily_planner.lblChkHideCompleted.ForeColor = originalLabelForeColor
    Forms!f_daily_planner.lblChkHideCompleted.BackStyle = originalLabelBackStyle
    DoEvents

    ' 4. ЗАДЕРЖКА ПЕРЕД СБРОСОМ
    startTime = Timer
    Do While Timer < startTime + 0.3
        DoEvents
    Loop

    ' 5. ВЫКЛЮЧАЕМ ЧЕКБОКС (если включен)
    If Forms!f_daily_planner.chkHideCompleted Then
        Forms!f_daily_planner.chkHideCompleted = False
        DoEvents

        ' 6. ВЫЗЫВАЕМ ПРИМЕНЕНИЕ ФИЛЬТРА
        Forms!f_daily_planner.ApplyHideCompletedFilter
        DoEvents
    End If

    ' 7. ЗАДЕРЖКА ДЛЯ ПЕРЕСТРОЙКИ КАЛЕНДАРЯ
    startTime = Timer
    Do While Timer < startTime + 2
        DoEvents
    Loop

    Debug.Print "ДЕМО: Фильтр выполненных событий сброшен"
End Sub

'################################################################
'########              ДЕМО-РЕЖИМ ПОИСКА                 ########
'################################################################

Public Sub FrmDemo_ExecuteSearchDemo(ByVal frm As Form)
    ' ОСНОВНАЯ ПРОЦЕДУРА - ВЫЗЫВАЕТ ЧАСТИ
    Call FrmDemo_DemoOpenSearchForm(frm)
    Call FrmDemo_DemoFillSearchFields(frm)
    Call FrmDemo_DemoApplySearch(frm)        ' < ПРИМЕНЕНИЕ ПОИСКА
    Call FrmDemo_DemoResetSearch(frm)        ' < СБРОС ФИЛЬТРОВ
    Call FrmDemo_DemoCloseSearchForm(frm)    ' < ЗАКРЫТИЕ ФОРМЫ
    ' ЗАВЕРШАЮЩЕЕ СООБЩЕНИЕ
    frm.txtCurrentAction.value = "Демонстрация поиска завершена" & vbCrLf & "Нажмите ""Далее"""
    DoEvents
End Sub

'################################################################
'########          5.1 ОТКРЫТИЕ ФОРМЫ ПОИСКА             ########
'################################################################
Public Sub FrmDemo_DemoOpenSearchForm(ByVal frm As Form)
    Dim originalSearchBtnBackColor As Long
    Dim originalSearchBtnForeColor As Long
    Dim startTime As Double

    Debug.Print "ДЕМО: Открытие формы поиска"

    ' УБИРАЕМ СООБЩЕНИЕ НА ФОРМУ - ТОЛЬКО DEBUG

    ' 1. ОТКРЫВАЕМ ФОРМУ ЕСЛИ ЗАКРЫТА
    If Not FrmDemo_IsFormLoaded("f_daily_planner") Then
        DoCmd.OpenForm "f_daily_planner"
        DoEvents
    End If

    ' 2. ПОДСВЕТКА КНОПКИ ПОИСКА НА ГЛАВНОЙ ФОРМЕ
    originalSearchBtnBackColor = Forms!f_daily_planner.cmdSearchEvents.backColor
    originalSearchBtnForeColor = Forms!f_daily_planner.cmdSearchEvents.ForeColor

    Forms!f_daily_planner.cmdSearchEvents.backColor = RGB(255, 255, 0)
    Forms!f_daily_planner.cmdSearchEvents.ForeColor = RGB(0, 0, 0)
    DoEvents

    ' 3. ЗАДЕРЖКА ДЛЯ ВИЗУАЛИЗАЦИИ ПОДСВЕТКИ
    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    ' 4. ВОССТАНАВЛИВАЕМ ЦВЕТА КНОПКИ
    Forms!f_daily_planner.cmdSearchEvents.backColor = originalSearchBtnBackColor
    Forms!f_daily_planner.cmdSearchEvents.ForeColor = originalSearchBtnForeColor
    DoEvents

    ' 5. ЗАДЕРЖКА ПЕРЕД ОТКРЫТИЕМ ФОРМЫ
    startTime = Timer
    Do While Timer < startTime + 0.3
        DoEvents
    Loop

    ' 6. ОТКРЫВАЕМ ФОРМУ ПОИСКА
    DoCmd.OpenForm "frmSearch"
    DoEvents

    ' 7. ЗАДЕРЖКА ДЛЯ ЗАГРУЗКИ ФОРМЫ
    startTime = Timer
    Do While Timer < startTime + 1
        DoEvents
    Loop

    Debug.Print "ДЕМО: Форма поиска открыта"
End Sub

'################################################################
'########          5.2 ЗАПОЛНЕНИЕ ПОЛЕЙ ПОИСКА           ########
'################################################################

Public Sub FrmDemo_DemoFillSearchFields(ByVal frm As Form)
    Dim startTime As Double
    Dim originalTextBackColor As Long
    Dim originalTextForeColor As Long
    Dim originalComboBackColor As Long
    Dim originalComboForeColor As Long

    Debug.Print "ДЕМО: Заполнение полей поиска"

    ' ПРОВЕРЯЕМ ЧТО ФОРМА ПОИСКА ОТКРЫТА
    If Not FrmDemo_IsFormLoaded("frmSearch") Then
        Debug.Print "ДЕМО: Форма поиска не открыта"
        Exit Sub
    End If

    ' 1. ПОДСВЕТКА И ЗАПОЛНЕНИЕ ТЕКСТОВОГО ПОЛЯ ПОИСКА
    originalTextBackColor = Forms!frmSearch.txtSearchText.backColor
    originalTextForeColor = Forms!frmSearch.txtSearchText.ForeColor

    Forms!frmSearch.txtSearchText.backColor = RGB(255, 255, 0)
    Forms!frmSearch.txtSearchText.ForeColor = RGB(0, 0, 0)
    DoEvents

    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    Forms!frmSearch.txtSearchText = "тест"
    DoEvents

    ' ВЫХОДИМ ИЗ ПОЛЯ
    Forms!frmSearch.dtStartDate.SetFocus
    DoEvents

    Forms!frmSearch.txtSearchText.backColor = originalTextBackColor
    Forms!frmSearch.txtSearchText.ForeColor = originalTextForeColor
    DoEvents

    ' ЗАДЕРЖКА
    startTime = Timer
    Do While Timer < startTime + 0.3
        DoEvents
    Loop

    ' 2. ПОДСВЕТКА И ВЫБОР ИСПОЛНИТЕЛЯ
    If Forms!frmSearch.cboSearchExecutor.ListCount > 1 Then
        originalComboBackColor = Forms!frmSearch.cboSearchExecutor.backColor
        originalComboForeColor = Forms!frmSearch.cboSearchExecutor.ForeColor

        Forms!frmSearch.cboSearchExecutor.backColor = RGB(255, 255, 0)
        Forms!frmSearch.cboSearchExecutor.ForeColor = RGB(0, 0, 0)
        DoEvents

        startTime = Timer
        Do While Timer < startTime + 0.5
            DoEvents
        Loop

        Forms!frmSearch.cboSearchExecutor = Forms!frmSearch.cboSearchExecutor.ItemData(1)
        DoEvents

        Forms!frmSearch.cboSearchExecutor.backColor = originalComboBackColor
        Forms!frmSearch.cboSearchExecutor.ForeColor = originalComboForeColor
        DoEvents

        ' ЗАДЕРЖКА
        startTime = Timer
        Do While Timer < startTime + 0.3
            DoEvents
        Loop
    End If

    ' 3. ПОДСВЕТКА И ВЫБОР СТАТУСА
    originalComboBackColor = Forms!frmSearch.cboCompletionStatus.backColor
    originalComboForeColor = Forms!frmSearch.cboCompletionStatus.ForeColor

    Forms!frmSearch.cboCompletionStatus.backColor = RGB(255, 255, 0)
    Forms!frmSearch.cboCompletionStatus.ForeColor = RGB(0, 0, 0)
    DoEvents

    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    Forms!frmSearch.cboCompletionStatus = "Не выполнено"
    DoEvents

    Forms!frmSearch.cboCompletionStatus.backColor = originalComboBackColor
    Forms!frmSearch.cboCompletionStatus.ForeColor = originalComboForeColor
    DoEvents

    Debug.Print "ДЕМО: Поля поиска заполнены с подсветкой"
End Sub

'################################################################
'########          5.3 ПРИМЕНЕНИЕ ПОИСКА                 ########
'################################################################

Public Sub FrmDemo_DemoApplySearch(ByVal frm As Form)
    Dim originalSearchBtnBackColor As Long
    Dim originalSearchBtnForeColor As Long
    Dim startTime As Double

    Debug.Print "ДЕМО: Применение поиска"

    ' ПРОВЕРЯЕМ ЧТО ФОРМА ПОИСКА ОТКРЫТА
    If Not FrmDemo_IsFormLoaded("frmSearch") Then
        Debug.Print "ДЕМО: Форма поиска не открыта"
        Exit Sub
    End If

    ' 1. ПОДСВЕТКА КНОПКИ "НАЙТИ"
    originalSearchBtnBackColor = Forms!frmSearch.cmdSearch.backColor
    originalSearchBtnForeColor = Forms!frmSearch.cmdSearch.ForeColor

    Forms!frmSearch.cmdSearch.backColor = RGB(255, 255, 0)
    Forms!frmSearch.cmdSearch.ForeColor = RGB(0, 0, 0)
    DoEvents

    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    ' 2. ВОССТАНАВЛИВАЕМ ЦВЕТА
    Forms!frmSearch.cmdSearch.backColor = originalSearchBtnBackColor
    Forms!frmSearch.cmdSearch.ForeColor = originalSearchBtnForeColor
    DoEvents

    ' 3. ЗАДЕРЖКА ПЕРЕД НАЖАТИЕМ
    startTime = Timer
    Do While Timer < startTime + 0.3
        DoEvents
    Loop

    ' 4. НАЖИМАЕМ КНОПКУ "НАЙТИ" ЧЕРЕЗ ПУБЛИЧНЫЙ МЕТОД
    Forms!frmSearch.ExecuteSearch
    DoEvents

    ' 5. ЗАДЕРЖКА ДЛЯ ВЫПОЛНЕНИЯ ПОИСКА
    startTime = Timer
    Do While Timer < startTime + 2
        DoEvents
    Loop

    Debug.Print "ДЕМО: Поиск применен"
End Sub

'################################################################
'########          5.4 СБРОС ФИЛЬТРОВ ПОИСКА             ########
'################################################################

Public Sub FrmDemo_DemoResetSearch(ByVal frm As Form)
    Dim originalResetBtnBackColor As Long
    Dim originalResetBtnForeColor As Long
    Dim startTime As Double

    Debug.Print "ДЕМО: Сброс фильтров поиска"

    ' ПРОВЕРЯЕМ ЧТО ФОРМА ПОИСКА ОТКРЫТА
    If Not FrmDemo_IsFormLoaded("frmSearch") Then
        Debug.Print "ДЕМО: Форма поиска не открыта"
        Exit Sub
    End If

    ' 1. ПОДСВЕТКА КНОПКИ "СБРОСИТЬ"
    originalResetBtnBackColor = Forms!frmSearch.cmdReset.backColor
    originalResetBtnForeColor = Forms!frmSearch.cmdReset.ForeColor

    Forms!frmSearch.cmdReset.backColor = RGB(255, 255, 0)
    Forms!frmSearch.cmdReset.ForeColor = RGB(0, 0, 0)
    DoEvents

    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    ' 2. ВОССТАНАВЛИВАЕМ ЦВЕТА
    Forms!frmSearch.cmdReset.backColor = originalResetBtnBackColor
    Forms!frmSearch.cmdReset.ForeColor = originalResetBtnForeColor
    DoEvents

    ' 3. ЗАДЕРЖКА ПЕРЕД НАЖАТИЕМ
    startTime = Timer
    Do While Timer < startTime + 0.3
        DoEvents
    Loop

    ' 4. НАЖИМАЕМ КНОПКУ "СБРОСИТЬ" ЧЕРЕЗ ПУБЛИЧНЫЙ МЕТОД
    Forms!frmSearch.ResetSearch
    DoEvents

    ' 5. ЗАДЕРЖКА ДЛЯ СБРОСА ФИЛЬТРОВ
    startTime = Timer
    Do While Timer < startTime + 2
        DoEvents
    Loop

    Debug.Print "ДЕМО: Фильтры поиска сброшены"
End Sub

'################################################################
'########          5.5 ЗАКРЫТИЕ ФОРМЫ ПОИСКА             ########
'################################################################

Public Sub FrmDemo_DemoCloseSearchForm(ByVal frm As Form)
    Dim originalCloseBtnBackColor As Long
    Dim originalCloseBtnForeColor As Long
    Dim startTime As Double

    Debug.Print "ДЕМО: Закрытие формы поиска"

    ' ПРОВЕРЯЕМ ЧТО ФОРМА ПОИСКА ОТКРЫТА
    If Not FrmDemo_IsFormLoaded("frmSearch") Then
        Debug.Print "ДЕМО: Форма поиска не открыта"
        Exit Sub
    End If

    ' 1. ПОДСВЕТКА КНОПКИ "ЗАКРЫТЬ"
    originalCloseBtnBackColor = Forms!frmSearch.cmdClose.backColor
    originalCloseBtnForeColor = Forms!frmSearch.cmdClose.ForeColor

    Forms!frmSearch.cmdClose.backColor = RGB(255, 255, 0)
    Forms!frmSearch.cmdClose.ForeColor = RGB(0, 0, 0)
    DoEvents

    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    ' 2. ВОССТАНАВЛИВАЕМ ЦВЕТА
    Forms!frmSearch.cmdClose.backColor = originalCloseBtnBackColor
    Forms!frmSearch.cmdClose.ForeColor = originalCloseBtnForeColor
    DoEvents

    ' 3. ЗАДЕРЖКА ПЕРЕД ЗАКРЫТИЕМ
    startTime = Timer
    Do While Timer < startTime + 0.3
        DoEvents
    Loop

    ' 4. ЗАКРЫВАЕМ ФОРМУ ПОИСКА ЧЕРЕЗ ПУБЛИЧНЫЙ МЕТОД
    Forms!frmSearch.CloseSearchForm
    DoEvents

    Debug.Print "ДЕМО: Форма поиска закрыта"
End Sub

'################################################################
'########              ДЕМО-РЕЖИМ ТЕМ                    ########
'################################################################

Public Sub FrmDemo_ExecuteThemeDemo(ByVal frm As Form)
    ' ОСНОВНАЯ ПРОЦЕДУРА - ВЫЗЫВАЕТ ЧАСТИ
    Call FrmDemo_DemoOpenThemeSelector(frm)
    Call FrmDemo_DemoSelectFirstTheme(frm)    ' < ПЕРВАЯ ТЕМА
    Call FrmDemo_DemoApplyFirstTheme(frm)     ' < ПРИМЕНЕНИЕ ПЕРВОЙ
    Call FrmDemo_DemoSelectSecondTheme(frm)   ' < ВТОРАЯ ТЕМА
    Call FrmDemo_DemoApplySecondTheme(frm)    ' < ПРИМЕНЕНИЕ ВТОРОЙ
    Call FrmDemo_DemoCloseThemeForm(frm)      ' < ЗАКРЫТИЕ ФОРМЫ
    ' ЗАВЕРШАЮЩЕЕ СООБЩЕНИЕ
    frm.txtCurrentAction.value = "Демонстрация смены тем завершена" & vbCrLf & "Нажмите ""Далее"""
    DoEvents
End Sub

'################################################################
'########         6.1 ОТКРЫТИЕ ВЫБОРА ТЕМЫ               ########
'################################################################

Public Sub FrmDemo_DemoOpenThemeSelector(ByVal frm As Form)
    Dim originalThemeBtnBackColor As Long
    Dim originalThemeBtnForeColor As Long
    Dim startTime As Double

    Debug.Print "ДЕМО: Открытие выбора темы"

    ' 1. ОТКРЫВАЕМ ФОРМУ ЕСЛИ ЗАКРЫТА
    If Not FrmDemo_IsFormLoaded("f_daily_planner") Then
        DoCmd.OpenForm "f_daily_planner"
        DoEvents
    End If

    ' 2. ПОДСВЕТКА КНОПКИ "СМЕНИТЬ ОФОРМЛЕНИЕ"
    originalThemeBtnBackColor = Forms!f_daily_planner.btn_theme.backColor
    originalThemeBtnForeColor = Forms!f_daily_planner.btn_theme.ForeColor

    Forms!f_daily_planner.btn_theme.backColor = RGB(255, 255, 0)
    Forms!f_daily_planner.btn_theme.ForeColor = RGB(0, 0, 0)
    DoEvents

    ' 3. ЗАДЕРЖКА ДЛЯ ВИЗУАЛИЗАЦИИ ПОДСВЕТКИ
    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    ' 4. ВОССТАНАВЛИВАЕМ ЦВЕТА КНОПКИ
    Forms!f_daily_planner.btn_theme.backColor = originalThemeBtnBackColor
    Forms!f_daily_planner.btn_theme.ForeColor = originalThemeBtnForeColor
    DoEvents

    ' 5. ЗАДЕРЖКА ПЕРЕД ОТКРЫТИЕМ ФОРМЫ
    startTime = Timer
    Do While Timer < startTime + 0.3
        DoEvents
    Loop

    ' 6. ОТКРЫВАЕМ ФОРМУ ВЫБОРА ТЕМЫ
    DoCmd.OpenForm "frmThemeSelector"
    DoEvents

    ' 7. ЗАДЕРЖКА ДЛЯ ЗАГРУЗКИ ФОРМЫ
    startTime = Timer
    Do While Timer < startTime + 1
        DoEvents
    Loop

    Debug.Print "ДЕМО: Форма выбора темы открыта"
End Sub

'################################################################
'########          6.2 ПЕРВЫЙ ВЫБОР ТЕМЫ                 ########
'################################################################

Public Sub FrmDemo_DemoSelectFirstTheme(ByVal frm As Form)
    Dim startTime As Double
    Dim ThemeName As String

    Debug.Print "ДЕМО: Первый выбор темы"

    ' ПРОВЕРЯЕМ ЧТО ФОРМА ТЕМ ОТКРЫТА
    If Not FrmDemo_IsFormLoaded("frmThemeSelector") Then
        Debug.Print "ДЕМО: Форма тем не открыта"
        Exit Sub
    End If

    ' 1. ПРОВЕРЯЕМ ЧТО ЕСТЬ ТЕМЫ В СПИСКЕ
    If Forms!frmThemeSelector.lstThemes.ListCount = 0 Then
        Debug.Print "ДЕМО: Нет тем в списке"
        Exit Sub
    End If

    ' 2. ВЫБИРАЕМ ПЕРВУЮ ТЕМУ В СПИСКЕ
    Forms!frmThemeSelector.lstThemes = Forms!frmThemeSelector.lstThemes.ItemData(0)
    DoEvents

    ' 3. ЗАДЕРЖКА ДЛЯ ВИЗУАЛИЗАЦИИ ВЫБОРА
    startTime = Timer
    Do While Timer < startTime + 1
        DoEvents
    Loop

    ThemeName = Forms!frmThemeSelector.lstThemes.value
    Debug.Print "ДЕМО: Первая тема выбрана: " & ThemeName
End Sub

'################################################################
'########          6.3 ПРИМЕНЕНИЕ ПЕРВОЙ ТЕМЫ            ########
'################################################################

Public Sub FrmDemo_DemoApplyFirstTheme(ByVal frm As Form)
    Dim originalApplyBtnBackColor As Long
    Dim originalApplyBtnForeColor As Long
    Dim startTime As Double

    Debug.Print "ДЕМО: Применение первой темы"

    ' ПРОВЕРЯЕМ ЧТО ФОРМА ТЕМ ОТКРЫТА
    If Not FrmDemo_IsFormLoaded("frmThemeSelector") Then
        Debug.Print "ДЕМО: Форма тем не открыта"
        Exit Sub
    End If

    ' 1. ПОДСВЕТКА КНОПКИ "ПРИМЕНИТЬ"
    originalApplyBtnBackColor = Forms!frmThemeSelector.btnApply.backColor
    originalApplyBtnForeColor = Forms!frmThemeSelector.btnApply.ForeColor

    Forms!frmThemeSelector.btnApply.backColor = RGB(255, 255, 0)
    Forms!frmThemeSelector.btnApply.ForeColor = RGB(0, 0, 0)
    DoEvents

    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    ' 2. ВОССТАНАВЛИВАЕМ ЦВЕТА
    Forms!frmThemeSelector.btnApply.backColor = originalApplyBtnBackColor
    Forms!frmThemeSelector.btnApply.ForeColor = originalApplyBtnForeColor
    DoEvents

    ' 3. ЗАДЕРЖКА ПЕРЕД ПРИМЕНЕНИЕМ
    startTime = Timer
    Do While Timer < startTime + 0.3
        DoEvents
    Loop

    ' 4. ПРИМЕНЯЕМ ПЕРВУЮ ТЕМУ
    Forms!frmThemeSelector.ApplySelectedTheme
    DoEvents

    ' 5. ЗАДЕРЖКА ДЛЯ ПРИМЕНЕНИЯ ТЕМЫ
    startTime = Timer
    Do While Timer < startTime + 2
        DoEvents
    Loop

    Debug.Print "ДЕМО: Первая тема применена"
End Sub

'################################################################
'########          6.4 ВТОРОЙ ВЫБОР ТЕМЫ                 ########
'################################################################

Public Sub FrmDemo_DemoSelectSecondTheme(ByVal frm As Form)
    Dim startTime As Double
    Dim ThemeName As String

    Debug.Print "ДЕМО: Второй выбор темы"

    ' 1. СНОВА ОТКРЫВАЕМ ФОРМУ ТЕМ (ЕСЛИ ЗАКРЫЛАСЬ)
    If Not FrmDemo_IsFormLoaded("frmThemeSelector") Then
        DoCmd.OpenForm "frmThemeSelector"
        DoEvents

        ' ЗАДЕРЖКА ДЛЯ ЗАГРУЗКИ ФОРМЫ
        startTime = Timer
        Do While Timer < startTime + 1
            DoEvents
        Loop
    End If

    ' 2. ПРОВЕРЯЕМ ЧТО ЕСТЬ ХОТЯ БЫ 2 ТЕМЫ В СПИСКЕ
    If Forms!frmThemeSelector.lstThemes.ListCount < 2 Then
        Debug.Print "ДЕМО: Меньше 2 тем в списке"
        Exit Sub
    End If

    ' 3. ВЫБИРАЕМ ВТОРУЮ ТЕМУ В СПИСКЕ
    Forms!frmThemeSelector.lstThemes = Forms!frmThemeSelector.lstThemes.ItemData(1)
    DoEvents

    ' 4. ЗАДЕРЖКА ДЛЯ ВИЗУАЛИЗАЦИИ ВЫБОРА
    startTime = Timer
    Do While Timer < startTime + 1
        DoEvents
    Loop

    ThemeName = Forms!frmThemeSelector.lstThemes.value
    Debug.Print "ДЕМО: Вторая тема выбрана: " & ThemeName
End Sub

'################################################################
'########          6.5 ПРИМЕНЕНИЕ ВТОРОЙ ТЕМЫ            ########
'################################################################

Public Sub FrmDemo_DemoApplySecondTheme(ByVal frm As Form)
    Dim originalApplyBtnBackColor As Long
    Dim originalApplyBtnForeColor As Long
    Dim startTime As Double

    Debug.Print "ДЕМО: Применение второй темы"

    ' ПРОВЕРЯЕМ ЧТО ФОРМА ТЕМ ОТКРЫТА
    If Not FrmDemo_IsFormLoaded("frmThemeSelector") Then
        Debug.Print "ДЕМО: Форма тем не открыта"
        Exit Sub
    End If

    ' 1. ПОДСВЕТКА КНОПКИ "ПРИМЕНИТЬ"
    originalApplyBtnBackColor = Forms!frmThemeSelector.btnApply.backColor
    originalApplyBtnForeColor = Forms!frmThemeSelector.btnApply.ForeColor

    Forms!frmThemeSelector.btnApply.backColor = RGB(255, 255, 0)
    Forms!frmThemeSelector.btnApply.ForeColor = RGB(0, 0, 0)
    DoEvents

    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    ' 2. ВОССТАНАВЛИВАЕМ ЦВЕТА
    Forms!frmThemeSelector.btnApply.backColor = originalApplyBtnBackColor
    Forms!frmThemeSelector.btnApply.ForeColor = originalApplyBtnForeColor
    DoEvents

    ' 3. ЗАДЕРЖКА ПЕРЕД ПРИМЕНЕНИЕМ
    startTime = Timer
    Do While Timer < startTime + 0.3
        DoEvents
    Loop

    ' 4. ПРИМЕНЯЕМ ВТОРУЮ ТЕМУ
    Forms!frmThemeSelector.ApplySelectedTheme
    DoEvents

    ' 5. ЗАДЕРЖКА ДЛЯ ПРИМЕНЕНИЯ ТЕМЫ
    startTime = Timer
    Do While Timer < startTime + 2
        DoEvents
    Loop

    Debug.Print "ДЕМО: Вторая тема применена"
End Sub

'################################################################
'########           6.6 ЗАКРЫТИЕ ФОРМЫ ТЕМ               ########
'################################################################

Public Sub FrmDemo_DemoCloseThemeForm(ByVal frm As Form)
    Dim originalCloseBtnBackColor As Long
    Dim originalCloseBtnForeColor As Long
    Dim startTime As Double

    Debug.Print "ДЕМО: Закрытие формы тем"

    ' ПРОВЕРЯЕМ ЧТО ФОРМА ТЕМ ОТКРЫТА
    If Not FrmDemo_IsFormLoaded("frmThemeSelector") Then
        Debug.Print "ДЕМО: Форма тем уже закрыта"
        Exit Sub
    End If

    ' 1. ПОДСВЕТКА КНОПКИ "ЗАКРЫТЬ"
    originalCloseBtnBackColor = Forms!frmThemeSelector.btnClose.backColor
    originalCloseBtnForeColor = Forms!frmThemeSelector.btnClose.ForeColor

    Forms!frmThemeSelector.btnClose.backColor = RGB(255, 255, 0)
    Forms!frmThemeSelector.btnClose.ForeColor = RGB(0, 0, 0)
    DoEvents

    startTime = Timer
    Do While Timer < startTime + 0.5
        DoEvents
    Loop

    ' 2. ВОССТАНАВЛИВАЕМ ЦВЕТА
    Forms!frmThemeSelector.btnClose.backColor = originalCloseBtnBackColor
    Forms!frmThemeSelector.btnClose.ForeColor = originalCloseBtnForeColor
    DoEvents

    ' 3. ЗАДЕРЖКА ПЕРЕД ЗАКРЫТИЕМ
    startTime = Timer
    Do While Timer < startTime + 0.3
        DoEvents
    Loop

    ' 4. ЗАКРЫВАЕМ ФОРМУ ТЕМ
    Forms!frmThemeSelector.CloseThemeForm
    DoEvents

    Debug.Print "ДЕМО: Форма тем закрыта"
End Sub


