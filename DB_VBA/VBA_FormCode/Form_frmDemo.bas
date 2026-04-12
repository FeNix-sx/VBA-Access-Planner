Option Compare Database

'################################################################
'########            ОБЪЯВЛЕНИЕ ПЕРЕМЕННЫХ               ########
'################################################################

Dim DemoStep As Integer
Dim DemoText As String
Dim DemoChar As Integer
Dim DemoInterval As Integer
Dim TimerMode As Integer ' 0-печать, 1-задержка

'################################################################
'########            ТЕКСТЫ ДЛЯ ДЕМО-РЕЖИМА              ########
'################################################################

' ШАГ 1 - ПРИВЕТСТВИЕ
Const DEMO_STEP1_TITLE As String = "ДЕМО-РЕЖИМ: ЗНАКОМСТВО С ПЛАНИРОВЩИКОМ"
Const DEMO_STEP1_TEXT As String = "Добро пожаловать в демонстрацию планировщика!" & vbCrLf & vbCrLf & _
                                 "Познакомимся с основными функциями:" & vbCrLf & _
                                 "• Навигация по календарю" & vbCrLf & _
                                 "• Работа с событиями" & vbCrLf & _
                                 "• Фильтрация и поиск" & vbCrLf & _
                                 "• Смена оформления" & vbCrLf & vbCrLf & _
                                 "Для продолжения нажмите кнопку 'Далее'."

' ШАГ 2 - НАВИГАЦИЯ (новый ШАГ 1)
Const DEMO_STEP2_TITLE As String = "ШАГ 1: НАВИГАЦИЯ ПО МЕСЯЦАМ"
Const DEMO_STEP2_TEXT As String = "Переход между месяцами с помощью кнопок:" & vbCrLf & vbCrLf & _
                                 "• 'Предыдущий месяц' - перейти назад" & vbCrLf & _
                                 "• 'Следующий месяц' - перейти вперед" & vbCrLf & _
                                 "• 'Текущий месяц' - вернуться к актуальной дате" & vbCrLf & vbCrLf & _
                                 "Кнопка 'Текущий месяц' скрывается когда вы находитесь в текущем месяце."

' ШАГ 3 - СОБЫТИЯ (новый ШАГ 2)
Const DEMO_STEP3_TITLE As String = "ШАГ 2: СОЗДАНИЕ И РЕДАКТИРОВАНИЕ СОБЫТИЙ"
Const DEMO_STEP3_TEXT As String = "Работа с событиями дня:" & vbCrLf & vbCrLf & _
                                 "• Двойной клик по дню - открыть события" & vbCrLf & _
                                 "• В форме событий можно: добавлять, редактировать, удалять" & vbCrLf & _
                                 "• Отмечать выполненные задачи" & vbCrLf & _
                                 "• Указывать исполнителя и периодичность" & vbCrLf & vbCrLf & _
                                 "Цветовая индикация показывает статус событий:" & vbCrLf & _
                                 "• Красный - просроченные" & vbCrLf & _
                                 "• Синий - ближайшие (1-3 дня)" & vbCrLf & _
                                 "• Серый - выполненные" & vbCrLf & _
                                 "• Обычный - текущие задачи"

' ШАГ 4 - ФИЛЬТРАЦИЯ (новый ШАГ 3)
Const DEMO_STEP4_TITLE As String = "ШАГ 3: ФИЛЬТРАЦИЯ СОБЫТИЙ"
Const DEMO_STEP4_TEXT As String = "Фильтрация событий по различным критериям:" & vbCrLf & vbCrLf & _
                                 "• Выпадающий список - фильтр по исполнителям" & vbCrLf & _
                                 "• Чекбокс 'Скрыть выполненные' - скрывает завершенные задачи" & vbCrLf & _
                                 "• Двойной клик по фильтру - сброс настроек" & vbCrLf & vbCrLf & _
                                 "Фильтры применяются мгновенно к отображению календаря."

' ШАГ 5 - ПОИСК (новый ШАГ 4)
Const DEMO_STEP5_TITLE As String = "ШАГ 4: ПОИСК ПО СИСТЕМЕ"
Const DEMO_STEP5_TEXT As String = "Поиск событий по ключевым словам:" & vbCrLf & vbCrLf & _
                                 "• Кнопка 'Поиск событий' открывает форму поиска" & vbCrLf & _
                                 "• Поиск работает по всем событиям во всех месяцах" & vbCrLf & _
                                 "• Результаты можно редактировать прямо из поиска" & vbCrLf & vbCrLf & _
                                 "Поиск не зависит от текущего фильтра исполнителей."

' ШАГ 6 - ТЕМЫ (новый ШАГ 5)
Const DEMO_STEP6_TITLE As String = "ШАГ 5: ПЕРСОНАЛИЗАЦИЯ ИНТЕРФЕЙСА"
Const DEMO_STEP6_TEXT As String = "Смена цветового оформления:" & vbCrLf & vbCrLf & _
                                 "• Кнопка 'Сменить оформление' открывает выбор тем" & vbCrLf & _
                                 "• Доступно 6 цветовых схем на любой вкус" & vbCrLf & _
                                 "• Тема применяется сразу ко всему интерфейсу" & vbCrLf & vbCrLf & _
                                 "Выбранная тема сохраняется между запусками программы."

'################################################################
'########          ЗАГРУЗКА ФОРМЫ ДЕМО-РЕЖИМА         ########
'################################################################

Private Sub Form_Load()
    ' Инициализация демо-режима
    DemoStep = 1
    DemoInterval = 10
    Call UpdateDemoStep
End Sub

'################################################################
'########           ТАЙМЕР АНИМАЦИИ И ДЕЙСТВИЙ           ########
'################################################################

Private Sub Form_Timer()
    If TimerMode = 0 Then
        ' РЕЖИМ ПЕЧАТИ ТЕКСТА
        If Len(DemoText & "") = 0 Then
            Me.TimerInterval = 0
            Exit Sub
        End If
        
        If DemoChar <= Len(DemoText) Then
            Me.txtActionDescription.value = Me.txtActionDescription.value & Mid(DemoText, DemoChar, 1)
            DemoChar = DemoChar + 1
        Else
            ' Текст напечатан - переходим к задержке
            TimerMode = 1
            Me.TimerInterval = 1000 ' Задержка 1 секунда
        End If
        
    ElseIf TimerMode = 1 Then
        ' РЕЖИМ ЗАДЕРЖКИ - выполняем действие
        TimerMode = 0
        Call ExecuteDemoAction
        Me.TimerInterval = 0 ' Останавливаем таймер
    End If
End Sub


'################################################################
'########          ОБНОВЛЕНИЕ ШАГА ДЕМО-РЕЖИМА       ########
'################################################################

Private Sub UpdateDemoStep()
    ' Очищаем поля и устанавливаем текст для текущего шага
    DemoChar = 1
    Me.txtActionDescription.value = ""
    Me.txtCurrentAction.value = "" ' Очищаем поле действий
    TimerMode = 0 ' Сбрасываем режим таймера
    
    Select Case DemoStep
        Case 1
            Me.lblProcessName.Caption = DEMO_STEP1_TITLE
            DemoText = DEMO_STEP1_TEXT
        Case 2
            Me.lblProcessName.Caption = DEMO_STEP2_TITLE
            DemoText = DEMO_STEP2_TEXT
        Case 3
            Me.lblProcessName.Caption = DEMO_STEP3_TITLE
            DemoText = DEMO_STEP3_TEXT
        Case 4
            Me.lblProcessName.Caption = DEMO_STEP4_TITLE
            DemoText = DEMO_STEP4_TEXT
        Case 5
            Me.lblProcessName.Caption = DEMO_STEP5_TITLE
            DemoText = DEMO_STEP5_TEXT
        Case 6
            Me.lblProcessName.Caption = DEMO_STEP6_TITLE
            DemoText = DEMO_STEP6_TEXT
    End Select
    
    ' Запускаем анимацию печати
    Me.TimerInterval = DemoInterval
End Sub

'################################################################
'########             КНОПКА "ДАЛЕЕ"                     ########
'################################################################

Private Sub cmdNext_Click()
    ' Переходим к следующему шагу (действия выполняются автоматически после печати)
    If DemoStep < 6 Then
        DemoStep = DemoStep + 1
        Call UpdateDemoStep
    Else
        ' Последний шаг - закрываем демо-режим
        DoCmd.Close acForm, "frmDemo"
    End If
End Sub

'################################################################
'########             КНОПКА "НАЗАД"                     ########
'################################################################

Private Sub cmdBack_Click()
    If DemoStep > 1 Then
        DemoStep = DemoStep - 1
        Call UpdateDemoStep
    End If
End Sub

'################################################################
'########             ПРОВЕРКА ЗАГРУЗКИ ФОРМЫ            ########
'################################################################

Private Function IsFormLoaded(formName As String) As Boolean
    IsFormLoaded = CurrentProject.allForms(formName).IsLoaded
End Function

'################################################################
'########          ВЫПОЛНЕНИЕ ДЕЙСТВИЙ ДЕМО-РЕЖИМА       ########
'################################################################

Private Sub ExecuteDemoAction()
    ' Выполняем действие для текущего шага
    Select Case DemoStep
        Case 2 ' ШАГ 1 - НАВИГАЦИЯ
            Call ExecuteNavigationDemo
            
        Case 3 ' ШАГ 2 - СОБЫТИЯ
            Call ExecuteEventsDemo
            
        Case 4 ' ШАГ 3 - ФИЛЬТРАЦИЯ
            Call ExecuteFilterDemo
            
        Case 5 ' ШАГ 4 - ПОИСК
            Call ExecuteSearchDemo
            
        Case 6 ' ШАГ 5 - ТЕМЫ  < ДОБАВЛЯЕМ
            Call ExecuteThemeDemo
    End Select
End Sub

'################################################################
'########            ДЕМО-РЕЖИМ НАВИГАЦИИ                ########
'################################################################
Private Sub ExecuteNavigationDemo()
    ' ОБЪЯВЛЕНИЕ ПЕРЕМЕННЫХ
    Dim originalNextBackColor As Long
    Dim originalNextForeColor As Long
    Dim originalPrevBackColor As Long
    Dim originalPrevForeColor As Long
    Dim originalCurrentBackColor As Long
    Dim originalCurrentForeColor As Long
    Dim startTime As Double
    
    ' 1. ОТКРЫВАЕМ ФОРМУ ЕСЛИ ЗАКРЫТА
    If Not IsFormLoaded("f_daily_planner") Then
        DoCmd.OpenForm "f_daily_planner"
        DoEvents ' Ждем загрузки формы
    End If
    
    ' === ПЕРВОЕ НАЖАТИЕ "СЛЕДУЮЩИЙ МЕСЯЦ" ===
    
    ' 2. СОХРАНЯЕМ ТЕКУЩИЕ ЦВЕТА КНОПКИ "СЛЕДУЮЩИЙ МЕСЯЦ"
    originalNextBackColor = Forms!f_daily_planner.btn_next.backColor
    originalNextForeColor = Forms!f_daily_planner.btn_next.ForeColor
    
    ' 3. НАЖАТИЕ КНОПКИ "СЛЕДУЮЩИЙ МЕСЯЦ" (1)
    Me.txtCurrentAction.value = "Нажатие: Следующий месяц (1)"
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
    Me.txtCurrentAction.value = "Нажатие: Следующий месяц (2)"
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
    Me.txtCurrentAction.value = "Нажатие: Предыдущий месяц"
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
    Me.txtCurrentAction.value = "Нажатие: Текущий месяц"
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
    Me.txtCurrentAction.value = "Демонстрация навигации завершена" & vbCrLf & "Нажмите ""Далее"""
    DoEvents
End Sub

'################################################################
'########            ДЕМО-РЕЖИМ СОБЫТИЙ                  ########
'################################################################

Private Sub ExecuteEventsDemo()
    Call Demo_HighlightDay
    Call Demo_OpenEventsForm
    Call Demo_EditAndFillEvent
    Call Demo_CloseEventsForm  ' < ДОБАВЛЯЕМ ЗАКРЫТИЕ
    Call Demo_RestoreDay
    ' ЗАВЕРШАЮЩЕЕ СООБЩЕНИЕ
    Me.txtCurrentAction.value = "Демонстрация работы с событиями завершена" & vbCrLf & "Нажмите ""Далее"""
    DoEvents
End Sub

'################################################################
'########        1. ПОДСВЕТКА ДНЯ В КАЛЕНДАРЕ           ########
'################################################################

Private Sub Demo_HighlightDay()
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
    Me.txtCurrentAction.value = "Двойной клик: Открытие событий дня"
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
    Call SaveDayData(DayNumber, originalDayBackColor, originalDayBorderColor, originalBorderWidth)
End Sub

'################################################################
'########        2. ОТКРЫТИЕ ФОРМЫ СОБЫТИЙ              ########
'################################################################

Private Sub Demo_OpenEventsForm()
    ' Получаем сохраненные данные
    Dim dayData As Variant
    dayData = GetDayData()
    Dim DayNumber As Integer
    DayNumber = dayData(0)
    
    ' ОТКРЫВАЕМ ФОРМУ СОБЫТИЙ
    Forms!f_daily_planner.OpenDayEvents DayNumber
    DoEvents
End Sub

'################################################################
'########      3.1 НАВИГАЦИЯ ПО ДНЯМ                   ########
'################################################################

Private Sub Demo_NavigateDays()
    If Not IsFormLoaded("frmDayEvents") Then Exit Sub
    
    Dim originalNextDayBackColor As Long
    Dim originalNextDayForeColor As Long
    Dim originalPrevDayBackColor As Long
    Dim originalPrevDayForeColor As Long
    Dim startTime As Double
    
    ' ПОДСВЕТКА И НАЖАТИЕ "СЛЕДУЮЩИЙ ДЕНЬ"
    Me.txtCurrentAction.value = "Навигация: Следующий день"
    
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
    Me.txtCurrentAction.value = "Навигация: Предыдущий день"
    
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

Private Sub Demo_EditAndFillEvent()
    If Not IsFormLoaded("frmDayEvents") Then Exit Sub
    
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

Private Sub Demo_FillEventFields()
    If Not IsFormLoaded("frmDayEvents") Then Exit Sub
    
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
    Me.txtCurrentAction.value = "Заполнение тестового события"
    
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
    Me.txtCurrentAction.value = "Заполнение примечаний"
    
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

Private Sub Demo_SaveAndClose()
    If Not IsFormLoaded("frmDayEvents") Then Exit Sub
    
    Dim originalSaveBackColor As Long
    Dim originalSaveForeColor As Long
    Dim startTime As Double
    
    ' СОХРАНЕНИЕ ИЗМЕНЕНИЙ
    Me.txtCurrentAction.value = "Сохранение изменений"
    
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

Private Sub Demo_CloseEventsForm()
    If Not IsFormLoaded("frmDayEvents") Then Exit Sub
    
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
    If IsFormLoaded("frmDayEvents") Then
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

Private Sub Demo_RestoreDay()
    ' Получаем сохраненные данные
    Dim dayData As Variant
    dayData = GetDayData()
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
'########      ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ                 ########
'################################################################

' Сохранение данных дня в глобальные переменные
Private Sub SaveDayData(DayNumber As Integer, backColor As Long, borderColor As Long, borderWidth As Integer)
    ' Используем форму для хранения временных данных
    Me.Tag = DayNumber & "|" & backColor & "|" & borderColor & "|" & borderWidth
End Sub

' Получение сохраненных данных дня
Private Function GetDayData() As Variant
    Dim dataArray() As String
    dataArray = Split(Me.Tag, "|")
    GetDayData = dataArray
End Function

'################################################################
'########             ДЕМО-РЕЖИМ ФИЛЬТРАЦИИ              ########
'################################################################

Private Sub ExecuteFilterDemo()
    ' ОСНОВНАЯ ПРОЦЕДУРА - ВЫЗЫВАЕТ ЧАСТИ
    Call Demo_HighlightExecutorFilter
    Call Demo_SelectExecutor
    Call Demo_HideCompletedEvents
    Call Demo_ResetExecutorFilter
    Call Demo_ResetCompletedFilter
    ' ЗАВЕРШАЮЩЕЕ СООБЩЕНИЕ
    Me.txtCurrentAction.value = "Демонстрация фильтрации завершена" & vbCrLf & "Нажмите ""Далее"""
    DoEvents
End Sub

'################################################################
'########     4.1 ПОДСВЕТКА ФИЛЬТРА ИСПОЛНИТЕЛЕЙ         ########
'################################################################

Private Sub Demo_HighlightExecutorFilter()
    Dim originalComboBackColor As Long
    Dim originalComboForeColor As Long
    Dim startTime As Double
    
    Debug.Print "ДЕМО: Подсветка фильтра исполнителей"
    
    ' СООБЩЕНИЕ ПЕРЕД ПОДСВЕТКОЙ
    Me.txtCurrentAction.value = "Фильтрация: Выбор исполнителя"
    DoEvents
    
    ' 1. ОТКРЫВАЕМ ФОРМУ ЕСЛИ ЗАКРЫТА
    If Not IsFormLoaded("f_daily_planner") Then
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

Private Sub Demo_SelectExecutor()
    Dim startTime As Double
    
    Debug.Print "ДЕМО: Выбор исполнителя из фильтра"
    
    ' 1. ПРОВЕРЯЕМ ЧТО В СПИСКЕ ЕСТЬ ИСПОЛНИТЕЛИ
    If Forms!f_daily_planner.cboExecutorFilter.ListCount = 0 Then
        Debug.Print "ДЕМО: Нет исполнителей для выбора"
        Exit Sub
    End If
    
    ' 2. ВЫБИРАЕМ ПЕРВОГО ИСПОЛНИТЕЛА ИЗ СПИСКА
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
Private Sub Demo_HideCompletedEvents()
    Dim originalLabelBackColor As Long
    Dim originalLabelForeColor As Long
    Dim originalLabelBackStyle As Integer
    Dim startTime As Double
    
    Debug.Print "ДЕМО: Фильтрация выполненных событий"
    
    ' СООБЩЕНИЕ ПЕРЕД ПОДСВЕТКОЙ
    Me.txtCurrentAction.value = "Фильтрация: Скрыть выполненные события"
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

Private Sub Demo_ResetExecutorFilter()
    Dim startTime As Double
    
    Debug.Print "ДЕМО: Сброс фильтра исполнителя"
    
    ' СООБЩЕНИЕ ПЕРЕД СБРОСОМ
    Me.txtCurrentAction.value = "Сброс фильтра: Все исполнители"
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

Private Sub Demo_ResetCompletedFilter()
    Dim originalLabelBackColor As Long
    Dim originalLabelForeColor As Long
    Dim originalLabelBackStyle As Integer
    Dim startTime As Double
    
    Debug.Print "ДЕМО: Сброс фильтра выполненных событий"
    
    ' СООБЩЕНИЕ ПЕРЕД СБРОСОМ
    Me.txtCurrentAction.value = "Сброс фильтра: Показать все события"
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

Private Sub ExecuteSearchDemo()
    ' ОСНОВНАЯ ПРОЦЕДУРА - ВЫЗЫВАЕТ ЧАСТИ
    Call Demo_OpenSearchForm
    Call Demo_FillSearchFields
    Call Demo_ApplySearch        ' < ПРИМЕНЕНИЕ ПОИСКА
    Call Demo_ResetSearch        ' < СБРОС ФИЛЬТРОВ
    Call Demo_CloseSearchForm    ' < ЗАКРЫТИЕ ФОРМЫ
    ' ЗАВЕРШАЮЩЕЕ СООБЩЕНИЕ
    Me.txtCurrentAction.value = "Демонстрация поиска завершена" & vbCrLf & "Нажмите ""Далее"""
    DoEvents
End Sub

'################################################################
'########          5.1 ОТКРЫТИЕ ФОРМЫ ПОИСКА             ########
'################################################################
Private Sub Demo_OpenSearchForm()
    Dim originalSearchBtnBackColor As Long
    Dim originalSearchBtnForeColor As Long
    Dim startTime As Double
    
    Debug.Print "ДЕМО: Открытие формы поиска"
    
    ' УБИРАЕМ СООБЩЕНИЕ НА ФОРМУ - ТОЛЬКО DEBUG
    
    ' 1. ОТКРЫВАЕМ ФОРМУ ЕСЛИ ЗАКРЫТА
    If Not IsFormLoaded("f_daily_planner") Then
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

Private Sub Demo_FillSearchFields()
    Dim startTime As Double
    Dim originalTextBackColor As Long
    Dim originalTextForeColor As Long
    Dim originalComboBackColor As Long
    Dim originalComboForeColor As Long
    
    Debug.Print "ДЕМО: Заполнение полей поиска"
    
    ' ПРОВЕРЯЕМ ЧТО ФОРМА ПОИСКА ОТКРЫТА
    If Not IsFormLoaded("frmSearch") Then
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

Private Sub Demo_ApplySearch()
    Dim originalSearchBtnBackColor As Long
    Dim originalSearchBtnForeColor As Long
    Dim startTime As Double
    
    Debug.Print "ДЕМО: Применение поиска"
    
    ' ПРОВЕРЯЕМ ЧТО ФОРМА ПОИСКА ОТКРЫТА
    If Not IsFormLoaded("frmSearch") Then
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

Private Sub Demo_ResetSearch()
    Dim originalResetBtnBackColor As Long
    Dim originalResetBtnForeColor As Long
    Dim startTime As Double
    
    Debug.Print "ДЕМО: Сброс фильтров поиска"
    
    ' ПРОВЕРЯЕМ ЧТО ФОРМА ПОИСКА ОТКРЫТА
    If Not IsFormLoaded("frmSearch") Then
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

Private Sub Demo_CloseSearchForm()
    Dim originalCloseBtnBackColor As Long
    Dim originalCloseBtnForeColor As Long
    Dim startTime As Double
    
    Debug.Print "ДЕМО: Закрытие формы поиска"
    
    ' ПРОВЕРЯЕМ ЧТО ФОРМА ПОИСКА ОТКРЫТА
    If Not IsFormLoaded("frmSearch") Then
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

Private Sub ExecuteThemeDemo()
    ' ОСНОВНАЯ ПРОЦЕДУРА - ВЫЗЫВАЕТ ЧАСТИ
    Call Demo_OpenThemeSelector
    Call Demo_SelectFirstTheme    ' < ПЕРВАЯ ТЕМА
    Call Demo_ApplyFirstTheme     ' < ПРИМЕНЕНИЕ ПЕРВОЙ
    Call Demo_SelectSecondTheme   ' < ВТОРАЯ ТЕМА
    Call Demo_ApplySecondTheme    ' < ПРИМЕНЕНИЕ ВТОРОЙ
    Call Demo_CloseThemeForm      ' < ЗАКРЫТИЕ ФОРМЫ
    ' ЗАВЕРШАЮЩЕЕ СООБЩЕНИЕ
    Me.txtCurrentAction.value = "Демонстрация смены тем завершена" & vbCrLf & "Нажмите ""Далее"""
    DoEvents
End Sub

'################################################################
'########         6.1 ОТКРЫТИЕ ВЫБОРА ТЕМЫ               ########
'################################################################

Private Sub Demo_OpenThemeSelector()
    Dim originalThemeBtnBackColor As Long
    Dim originalThemeBtnForeColor As Long
    Dim startTime As Double
    
    Debug.Print "ДЕМО: Открытие выбора темы"
    
    ' 1. ОТКРЫВАЕМ ФОРМУ ЕСЛИ ЗАКРЫТА
    If Not IsFormLoaded("f_daily_planner") Then
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

Private Sub Demo_SelectFirstTheme()
    Dim startTime As Double
    Dim ThemeName As String
    
    Debug.Print "ДЕМО: Первый выбор темы"
    
    ' ПРОВЕРЯЕМ ЧТО ФОРМА ТЕМ ОТКРЫТА
    If Not IsFormLoaded("frmThemeSelector") Then
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

Private Sub Demo_ApplyFirstTheme()
    Dim originalApplyBtnBackColor As Long
    Dim originalApplyBtnForeColor As Long
    Dim startTime As Double
    
    Debug.Print "ДЕМО: Применение первой темы"
    
    ' ПРОВЕРЯЕМ ЧТО ФОРМА ТЕМ ОТКРЫТА
    If Not IsFormLoaded("frmThemeSelector") Then
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

Private Sub Demo_SelectSecondTheme()
    Dim startTime As Double
    Dim ThemeName As String
    
    Debug.Print "ДЕМО: Второй выбор темы"
    
    ' 1. СНОВА ОТКРЫВАЕМ ФОРМУ ТЕМ (ЕСЛИ ЗАКРЫЛАСЬ)
    If Not IsFormLoaded("frmThemeSelector") Then
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

Private Sub Demo_ApplySecondTheme()
    Dim originalApplyBtnBackColor As Long
    Dim originalApplyBtnForeColor As Long
    Dim startTime As Double
    
    Debug.Print "ДЕМО: Применение второй темы"
    
    ' ПРОВЕРЯЕМ ЧТО ФОРМА ТЕМ ОТКРЫТА
    If Not IsFormLoaded("frmThemeSelector") Then
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

Private Sub Demo_CloseThemeForm()
    Dim originalCloseBtnBackColor As Long
    Dim originalCloseBtnForeColor As Long
    Dim startTime As Double
    
    Debug.Print "ДЕМО: Закрытие формы тем"
    
    ' ПРОВЕРЯЕМ ЧТО ФОРМА ТЕМ ОТКРЫТА
    If Not IsFormLoaded("frmThemeSelector") Then
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
