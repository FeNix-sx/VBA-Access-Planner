Option Compare Database

'################################################################
'########              КАЛЕНДАРЬ СОБЫТИЙ                 ########
'################################################################
' Настройки tbSettings и «Игры»: modDailyPlannerSettings, modDailyPlannerGames.
Dim CurrentMonth As Date
Private m_SelectedPanelDate As Date

' Переменные для хранения цветов текущей темы
Dim CurrentTheme_Text As Long
Dim CurrentTheme_Back As Long
Dim CurrentTheme_Border As Long
Dim OtherTheme_Text As Long
Dim OtherTheme_Back As Long
Dim OtherTheme_Border As Long
Dim TodayTheme_Back As Long
Dim TodayTheme_Border As Long
Dim HeaderTheme_Text As Long
Dim HeaderTheme_Back As Long
Dim HeaderTheme_Border As Long
Dim FormTheme_Back As Long
' Константы позиции окна и размеров формы «Дни рождения» (twips)
Private Const TWIPS_DAILY_PLANNER_LEFT As Long = 500
Private Const TWIPS_DAILY_PLANNER_TOP As Long = 1000
Private Const TWIPS_FORM_BASE_WIDTH As Long = 21800
Private Const TWIPS_FORM_HEIGHT As Long = 14000
Private Const TWIPS_BIRTHDAYS_PANEL_EXTRA_WIDTH As Long = 3300

' Состояние режима окна (оконный / полноэкранный)
Private m_IsFullscreen As Boolean
Private m_WindowRectCaptured As Boolean
Private m_WindowLeft As Long
Private m_WindowTop As Long
Private m_WindowWidth As Long
Private m_WindowHeight As Long

'################################################################
'########            Кнопка "Текущий месяц"              ########
'################################################################
Private Sub btn_current_Click()
    ' Переход к текущему месяцу без изменения даты
    Me.btn_next.SetFocus
    ' Устанавливаем текущий месяц
    CurrentMonth = DateSerial(Year(Date), Month(Date), 1)
    ' Перестраиваем календарь
    Call BuildCalendar
End Sub

'################################################################
'########            Фильтр исполнителей                 ########
'################################################################
Private Sub cboExecutorFilter_DblClick(Cancel As Integer)
    Me.cboExecutorFilter.value = Null
    Call cboExecutorFilter_AfterUpdate
End Sub

'################################################################
'########         Кнопка генерации событий               ########
'################################################################
Private Sub cmdEvengGenerate_Click()
    DoCmd.OpenForm "frmEventGenerator"
End Sub
'################################################################
'########         Кнопка справочника исполнителей        ########
'################################################################
Private Sub cmdExecutors_Click()
    DoCmd.OpenForm "frmExecutors"
End Sub

'################################################################
'########         Управление днями рождения              ########
'################################################################
Private Sub btn_BirthdaysManage_Click()
    On Error GoTo ErrorHandler
    DoCmd.OpenForm "frmBirthdaysList"
    Exit Sub
ErrorHandler:
    MsgBox "Не удалось открыть справочник дней рождения: " & Err.description, vbExclamation
End Sub

'################################################################
'########             Кнопка демо-версии                 ########
'################################################################
Private Sub cmdRunDemo_Click()
    DoCmd.OpenForm "frmDemo", acNormal
End Sub

'################################################################
'########           События формы при загрузке           ########
'################################################################

Private Sub Form_Load()

    ' Проверка лицензии перед запуском
    Call CheckLicenseOnStartup

    Call AutoConnectOnStartup
    Call NotifyTodaysBirthdaysOncePerDay
    ' Устанавливаем текущий месяц
    CurrentMonth = DateSerial(Year(Date), Month(Date), 1)

    ' Загрузка темы из настроек или базы
    Call LoadDefaultTheme

    ' Загрузка настройки скрытия выполненных
    Call LoadHideCompletedSetting

    ' Загрузка настройки фиксации панели на текущем дне
    Call LoadCurrentDaySetting

    ' Настройка «Дни рождения» (из настроек базы, см. п.8)
    Call LoadShowBirthdaysPanelSetting

    ' Инициализация фильтра исполнителей
    Call InitializeExecutorFilter

    ' Построение сетки календаря
    Call BuildCalendar

    ' Инициализация состояния выбранной даты панели событий дня
    m_SelectedPanelDate = Date

    ' Первичное обновление панели событий дня
    Call RefreshEventPanel

    ' Применяем UI-режим (админ/пользователь): лента и панель навигации
    Call ApplyAdminUiMode

    ' Применить размер и видимость панели дней рождения
    Call ApplyBirthdaysPanelLayout
    Call InitializeWindowMode
    Call ApplyWindowModeFromSetting
    Call InitializeGamesButtonState

End Sub

'################################################################
'########      Активация формы (обновление UI режима)    ########
'################################################################
Private Sub Form_Activate()
    On Error GoTo ExitSub
    Call ApplyAdminUiMode
ExitSub:
End Sub

'################################################################
'########      Инициализация режима окна формы           ########
'################################################################
Private Sub InitializeWindowMode()
    On Error GoTo ErrHandler

    m_IsFullscreen = False
    m_WindowRectCaptured = False

    Call CaptureWindowRect
    Call UpdateWindowModeButtonState

ExitSub:
    Exit Sub
ErrHandler:
    Debug.Print "[f_daily_planner][ERR][InitializeWindowMode] " & Err.Number & " - " & Err.description
End Sub

'################################################################
'########      Переключение оконный/полноэкранный        ########
'################################################################
Public Sub ToggleWindowMode()
    On Error GoTo ErrHandler
    Dim newMode As String

    If m_IsFullscreen Then
        newMode = "windowed"
    Else
        newMode = "fullscreen"
    End If

    Call PlnSettings_SaveWindowMode(newMode)

    Application.Run "ReopenPlannerFormGlobal", Me.Name
    Exit Sub

ErrHandler:
    Debug.Print "[f_daily_planner][ERR][ToggleWindowMode] " & Err.Number & " - " & Err.description
End Sub

'################################################################
'########       Перевод формы в полноэкранный режим      ########
'################################################################
Public Sub SwitchToFullscreenMode()
    On Error GoTo ErrHandler

    If Not m_WindowRectCaptured Then
        Call CaptureWindowRect
    End If

    Call ApplyFullscreenFormSize

    m_IsFullscreen = True
    Call UpdateWindowModeButtonState

ExitSub:
    Exit Sub
ErrHandler:
    Debug.Print "[f_daily_planner][ERR][SwitchToFullscreenMode] " & Err.Number & " - " & Err.description
End Sub

'################################################################
'########         Перевод формы в оконный режим          ########
'################################################################
Public Sub SwitchToWindowedMode()
    On Error GoTo ErrHandler

    DoCmd.SelectObject acForm, Me.Name, False
    DoCmd.Restore

    m_IsFullscreen = False

    ' Оконный режим всегда применяем по фиксированному макету формы.
    Call ApplyBirthdaysPanelLayout

    Call UpdateWindowModeButtonState

ExitSub:
    Exit Sub
ErrHandler:
    Debug.Print "[f_daily_planner][ERR][SwitchToWindowedMode] " & Err.Number & " - " & Err.description
End Sub

'################################################################
'########      Запомнить текущее положение окна          ########
'################################################################
Private Sub CaptureWindowRect()
    On Error GoTo ErrHandler

    m_WindowLeft = Me.WindowLeft
    m_WindowTop = Me.WindowTop
    m_WindowWidth = Abs(Me.WindowWidth)
    m_WindowHeight = Abs(Me.WindowHeight)
    m_WindowRectCaptured = (m_WindowWidth > 0 And m_WindowHeight > 0)

ExitSub:
    Exit Sub
ErrHandler:
    Debug.Print "[f_daily_planner][ERR][CaptureWindowRect] " & Err.Number & " - " & Err.description
End Sub

'################################################################
'########      Обновить подпись кнопки режима             ########
'################################################################
Private Sub UpdateWindowModeButtonState()
    On Error GoTo ErrHandler
    Dim captionText As String

    If m_IsFullscreen Then
        captionText = "В оконный режим"
    Else
        captionText = "В полноэкранный режим"
    End If

    Call SetControlCaptionIfExists("btn_window_mode", captionText)

ExitSub:
    Exit Sub
ErrHandler:
    Debug.Print "[f_daily_planner][ERR][UpdateWindowModeButtonState] " & Err.Number & " - " & Err.description
End Sub

'################################################################
'########      Установить Caption, если контрол есть      ########
'################################################################
Private Sub SetControlCaptionIfExists(ByVal controlName As String, ByVal captionText As String)
    On Error GoTo ErrHandler
    Me.Controls(controlName).Caption = captionText
ExitSub:
    Exit Sub
ErrHandler:
    Debug.Print "[f_daily_planner][ERR][SetControlCaptionIfExists] " & Err.Number & " - " & Err.description & "; control=" & controlName
End Sub

'################################################################
'########        Кнопка переключения режима окна          ########
'################################################################
Private Sub btn_window_mode_Click()
    Call ToggleWindowMode
End Sub

'################################################################
'########               Кнопка "Игры"                    ########
'################################################################
Private Sub btn_games_Click()
    On Error GoTo Err_Handler
    Dim clickCount As Long
    Dim nextCount As Long
    Dim lastMessageKey As Long
    Dim newMessageKey As Long
    Dim titleText As String
    Dim captionText As String
    Dim messageText As String

    clickCount = PlnGames_GetClickCount()
    nextCount = (clickCount + 1) Mod 100
    titleText = PlnGames_GetTitleByCount(nextCount)
    captionText = PlnGames_BuildCaption(titleText, nextCount)

    lastMessageKey = CLng(Val(Nz(PlnSettings_GetValue(PLN_GAMES_SETTING_LAST_MESSAGE_KEY, "0"), "0")))
    messageText = PlnGames_BuildMessage(nextCount, titleText, lastMessageKey, newMessageKey)

    Call PlnSettings_SaveValue(PLN_GAMES_SETTING_COUNT, CStr(nextCount))
    Call PlnSettings_SaveValue(PLN_GAMES_SETTING_CAPTION, captionText)
    Call PlnSettings_SaveValue(PLN_GAMES_SETTING_LAST_MESSAGE_KEY, CStr(newMessageKey))

    Me.btn_games.Caption = captionText
    MsgBox messageText, vbInformation, "Игры"
    Exit Sub

Err_Handler:
    MsgBox "Ошибка в кнопке «Игры»: " & Err.Description, vbExclamation
End Sub

'################################################################
'########       Применить режим окна из настроек         ########
'################################################################
Private Sub ApplyWindowModeFromSetting()
    On Error GoTo ErrHandler
    Dim modeValue As String

    modeValue = PlnSettings_GetWindowMode()
    m_IsFullscreen = (LCase$(modeValue) = "fullscreen")

    If m_IsFullscreen Then
        Call ApplyFullscreenFormSize
    Else
        DoCmd.SelectObject acForm, Me.Name, False
        DoCmd.Restore
        Call ApplyBirthdaysPanelLayout
    End If

    Call UpdateWindowModeButtonState
    Exit Sub

ErrHandler:
    Debug.Print "[f_daily_planner][ERR][ApplyWindowModeFromSetting] " & Err.Number & " - " & Err.description
End Sub

'################################################################
'########      Инициализация состояния кнопки "Игры"     ########
'################################################################
Private Sub InitializeGamesButtonState()
    On Error GoTo Err_Handler
    Dim clickCount As Long
    Dim titleText As String
    Dim captionText As String

    Randomize
    clickCount = PlnGames_GetClickCount()
    titleText = PlnGames_GetTitleByCount(clickCount)
    captionText = PlnGames_BuildCaption(titleText, clickCount)

    Me.btn_games.Caption = captionText
    Call PlnSettings_SaveValue(PLN_GAMES_SETTING_CAPTION, captionText)
    Exit Sub

Err_Handler:
    Debug.Print "[f_daily_planner][ERR][InitializeGamesButtonState] " & Err.Number & " - " & Err.Description
End Sub

'################################################################
'########      Применить полноэкранный размер формы       ########
'################################################################
Private Sub ApplyFullscreenFormSize()
    On Error GoTo FallbackMaximize
    Dim usableW As Long
    Dim usableH As Long

    DoCmd.SelectObject acForm, Me.Name, False
    usableW = GetAccessUsableSize("UsableWidth")
    usableH = GetAccessUsableSize("UsableHeight")

    If usableW > 0 And usableH > 0 Then
        DoCmd.MoveSize 0, 0, usableW, usableH
    Else
        DoCmd.Maximize
    End If

    Exit Sub

FallbackMaximize:
    On Error Resume Next
    DoCmd.Maximize
End Sub

'################################################################
'########       Целевая дата панели событий дня          ########
'################################################################
Private Function GetPanelTargetDate() As Date
    If Nz(Me.chk_CurrentDay, True) Then
        GetPanelTargetDate = Date
    ElseIf IsDate(m_SelectedPanelDate) Then
        GetPanelTargetDate = DateValue(m_SelectedPanelDate)
    Else
        GetPanelTargetDate = Date
    End If
End Function

'################################################################
'########      Построение фильтра панели событий         ########
'################################################################
Private Function BuildEventPanelFilter(ByVal targetDate As Date) As String
    Dim filterText As String

    filterText = "EventDate = #" & Format(DateValue(targetDate), "mm\/dd\/yyyy") & "#"

    If Not IsNull(Me.cboExecutorFilter.value) And Me.cboExecutorFilter.value <> "" Then
        filterText = filterText & " AND ExecutorID = " & CLng(Me.cboExecutorFilter.value)
    End If

    BuildEventPanelFilter = filterText
End Function

'################################################################
'########      Единое обновление панели событий          ########
'################################################################
Private Sub RefreshEventPanel()
    On Error GoTo ErrorHandler

    Dim targetDate As Date
    Dim filterText As String
    Dim dataPhaseEnabled As Boolean
    Dim ExecutorID As Variant
    Dim expectedRows As Long

    targetDate = GetPanelTargetDate()
    filterText = BuildEventPanelFilter(targetDate)
    dataPhaseEnabled = True
    ExecutorID = Null
    If Not IsNull(Me.cboExecutorFilter.value) And Me.cboExecutorFilter.value <> "" Then
        ExecutorID = CLng(Me.cboExecutorFilter.value)
    End If
    expectedRows = CountEventsForPanelDate(targetDate, ExecutorID)

    If Not IsEventPanelReportReady() Then
        Exit Sub
    End If

    With Me.sub_rptEventInstances.Report

        .lbl_Day.Caption = Format(targetDate, "d mmmm yyyy ""г.""")

        If dataPhaseEnabled Then

            On Error Resume Next
            TempVars.Remove "EventPanelDate"
            TempVars.Remove "EventPanelExecutorID"
            On Error GoTo ErrorHandler

            TempVars.Add "EventPanelDate", DateValue(targetDate)
            TempVars.Add "EventPanelExecutorID", ExecutorID

            .Requery

        End If
    End With

    Call ApplyEventPanelTheme

    On Error Resume Next
    Me.sub_rptEventInstances.Report.lbl_Day.Caption = Format(targetDate, "d mmmm yyyy ""г.""")
    On Error GoTo ErrorHandler

    Exit Sub

ErrorHandler:
    ' Подотчет может быть недоступен во время ранней инициализации формы
End Sub

'################################################################
'########     Применение темы к панели событий дня       ########
'################################################################
Private Sub ApplyEventPanelTheme()
    On Error GoTo ErrorHandler

    If Not IsEventPanelReportReady() Then Exit Sub

    ' Делегируем оформление самому отчету, чтобы не дублировать палитру.
    Me.sub_rptEventInstances.Report.ApplyThemeFromHost

    Exit Sub

ErrorHandler:
    ' Панель может быть недоступна до полной инициализации отчета
End Sub

Private Function IsEventPanelReportReady() As Boolean
    On Error GoTo NotReady

    Dim reportName As String
    reportName = Me.sub_rptEventInstances.Report.Name

    IsEventPanelReportReady = (Len(reportName) > 0)
    Exit Function

NotReady:
    IsEventPanelReportReady = False
End Function

'################################################################
'########     Загрузка настройки «Текущий день»          ########
'################################################################
Private Sub LoadCurrentDaySetting()
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb

    Set rs = db.OpenRecordset("SELECT SettingValue FROM tbSettings WHERE SettingName = 'CurrentDay'")

    If Not rs.EOF Then
        Me.chk_CurrentDay = (rs!settingValue = 1)
    Else
        Me.chk_CurrentDay = True
    End If

    rs.Close
    Exit Sub
ErrorHandler:
    On Error Resume Next
    Me.chk_CurrentDay = True
    On Error GoTo 0
End Sub

'################################################################
'########     Сохранение настройки «Текущий день»        ########
'################################################################
Private Sub SaveCurrentDaySetting()
    On Error GoTo ErrorHandler
    Dim db As DAO.Database

    Set db = CurrentDb
    db.Execute "DELETE FROM tbSettings WHERE SettingName = 'CurrentDay'"
    db.Execute "INSERT INTO tbSettings (SettingName, SettingValue) VALUES ('CurrentDay', " & IIf(Me.chk_CurrentDay, "1", "0") & ")"
    Exit Sub
ErrorHandler:
End Sub

'################################################################
'########         Загрузка настройки скрытия            ########
'################################################################
Private Sub LoadHideCompletedSetting()
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb

    Set rs = db.OpenRecordset("SELECT SettingValue FROM tbSettings WHERE SettingName = 'HideCompleted'")

    If Not rs.EOF Then
        Me.chkHideCompleted = (rs!settingValue = 1)
    Else
        Me.chkHideCompleted = False
    End If

    rs.Close
    Exit Sub
ErrorHandler:
    Me.chkHideCompleted = False
End Sub

'################################################################
'########    Загрузка настройки панели «Дни рождения»    ########
'################################################################
Private Sub LoadShowBirthdaysPanelSetting()
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb

    Set rs = db.OpenRecordset("SELECT SettingValue FROM tbSettings WHERE SettingName = 'ShowBirthdaysPanel'")

    If Not rs.EOF Then
        Me.chk_ShowBirthdays = (rs!settingValue = 1)
    Else
        Me.chk_ShowBirthdays = False
    End If

    rs.Close
    Exit Sub
ErrorHandler:
    On Error Resume Next
    Me.chk_ShowBirthdays = False
    On Error GoTo 0
End Sub

'################################################################
'########   Сохранение настройки панели «Дни рождения»  ########
'################################################################
Private Sub SaveShowBirthdaysPanelSetting()
    On Error GoTo ErrorHandler
    Dim db As DAO.Database

    Set db = CurrentDb
    db.Execute "DELETE FROM tbSettings WHERE SettingName = 'ShowBirthdaysPanel'"
    db.Execute "INSERT INTO tbSettings (SettingName, SettingValue) VALUES ('ShowBirthdaysPanel', " & IIf(Me.chk_ShowBirthdays, "1", "0") & ")"
    Exit Sub
ErrorHandler:
End Sub

'################################################################
'########    Применить размеры, видимость, REQUERY      ########
'################################################################
Private Sub ApplyBirthdaysPanelLayout()
    On Error GoTo BirthdaysLayoutSkip

    If Nz(Me.chk_ShowBirthdays, False) Then
        Me.sub_rptBirthdays.Visible = True
        If m_IsFullscreen Then
            Call ApplyFullscreenFormSize
        Else
            DoCmd.MoveSize TWIPS_DAILY_PLANNER_LEFT, TWIPS_DAILY_PLANNER_TOP, TWIPS_FORM_BASE_WIDTH + TWIPS_BIRTHDAYS_PANEL_EXTRA_WIDTH, TWIPS_FORM_HEIGHT
        End If
    Else
        Me.sub_rptBirthdays.Visible = False
        If m_IsFullscreen Then
            Call ApplyFullscreenFormSize
        Else
            DoCmd.MoveSize TWIPS_DAILY_PLANNER_LEFT, TWIPS_DAILY_PLANNER_TOP, TWIPS_FORM_BASE_WIDTH, TWIPS_FORM_HEIGHT
        End If
    End If

    On Error GoTo 0
    On Error Resume Next
    If Nz(Me.chk_ShowBirthdays, False) Then
        Me.sub_rptBirthdays.Report.Requery
    End If
    On Error GoTo 0
    Exit Sub

BirthdaysLayoutSkip:
    Debug.Print "[f_daily_planner][ERR][ApplyBirthdaysPanelLayout] " & Err.Number & " - " & Err.description
End Sub

'################################################################
'########               Построение календаря             ########
'########                Основная логика                 ########
'################################################################
Public Sub BuildCalendar()
    Dim startDate As Date
    Dim dayCounter As Integer
    Dim ctrlDay As Control
    Dim ctrlEvent As Control

    ' Устанавливаем заголовок с текущим месяцем и годом
    Me.lbl_MonthYear.Caption = Format(CurrentMonth, "mmmm yyyy")

    ' Применяем стили заголовка
    Call ApplyFormHeaderStyle

    ' Вычисляем начальную дату: первый понедельник перед первым днем месяца
    ' Weekday с vbMonday возвращает 1 для понедельника, 7 для воскресенья
    startDate = CurrentMonth - weekday(CurrentMonth, vbMonday) + 1

    ' Проходим по всем 42 ячейкам (6 недель * 7 дней)
    For dayCounter = 1 To 42
        Set ctrlDay = Me.Controls("lbl_day_" & dayCounter)
        Set ctrlEvent = Me.Controls("fld_day_" & dayCounter)

        ' Устанавливаем число дня
        ctrlDay.Caption = Day(startDate)

        ' Настройка доступности поля ввода для текущего месяца
        Call SetEventFieldAccess(ctrlEvent, startDate)
        Call ApplyDayStyling(ctrlDay, ctrlEvent, startDate)
        Call HighlightToday(ctrlDay, ctrlEvent, startDate)
        Call LoadEventData(ctrlEvent, startDate)

        ' Переходим к следующему дню
        startDate = DateAdd("d", 1, startDate)
    Next dayCounter

    ' Отображаем кнопку "Текущий месяц"
    Me.btn_current.Visible = (Month(CurrentMonth) <> Month(Date)) Or (Year(CurrentMonth) <> Year(Date))

    ' Панель событий дня синхронизируется после отрисовки календаря
    Call RefreshEventPanel
End Sub

'################################################################
'########           1. Настройка доступности             ########
'########                поля событий                    ########
'################################################################
Private Sub SetEventFieldAccess(ctrlEvent As Control, currentDate As Date)
    ' Для дней текущего месяца поле доступно для ввода (только для чтения)
    ' Для других месяцев поле недоступно
    If Month(currentDate) = Month(CurrentMonth) Then
        ctrlEvent.Enabled = True
        ctrlEvent.Locked = True
    Else
        ctrlEvent.Enabled = False
        ctrlEvent.Locked = True
    End If
End Sub

'################################################################
'########           2. Применение стилей                 ########
'########                  для дня                       ########
'################################################################
Private Sub ApplyDayStyling(ctrlDay As Control, ctrlEvent As Control, currentDate As Date)
    ' Применение стиля в зависимости от месяца и дня недели
    If Month(currentDate) = Month(CurrentMonth) Then
        ' Дни текущего месяца
        If IsWeekend(currentDate) Then
            ' Суббота/воскресенье текущего месяца
            ApplyWeekendStyle ctrlDay, ctrlEvent
        Else
            ' Будни текущего месяца
            ApplyCurrentMonthStyle ctrlDay, ctrlEvent
        End If
    Else
        ' Дни других месяцев (прошлого/следующего)
        ApplyOtherMonthStyle ctrlDay, ctrlEvent
    End If
End Sub

'################################################################
'########           2.1 Стиль дня                        ########
'########             текущего месяца                    ########
'################################################################
Private Sub ApplyCurrentMonthStyle(ctrlDay As Control, ctrlEvent As Control)
    ' Label и поле - цвета текущего месяца
    ctrlDay.ForeColor = CurrentTheme_Text
    ctrlDay.backColor = CurrentTheme_Back
    ctrlDay.borderColor = CurrentTheme_Border
    ctrlDay.borderWidth = 1

    ' TextBox и поле - цвета текущего месяца
    ctrlEvent.backColor = CurrentTheme_Back
    ctrlEvent.ForeColor = CurrentTheme_Text
    ctrlEvent.borderColor = CurrentTheme_Border
    ctrlEvent.borderWidth = 1
End Sub

'################################################################
'########           2.2 Стиль дня                        ########
'########               других месяцев                   ########
'################################################################
Private Sub ApplyOtherMonthStyle(ctrlDay As Control, ctrlEvent As Control)
    ' Label и поле - цвета других месяцев
    ctrlDay.ForeColor = OtherTheme_Text
    ctrlDay.backColor = OtherTheme_Back
    ctrlDay.borderColor = OtherTheme_Border
    ctrlDay.borderWidth = 1

    ' TextBox и поле - цвета других месяцев
    ctrlEvent.backColor = OtherTheme_Back
    ctrlEvent.ForeColor = OtherTheme_Text
    ctrlEvent.borderColor = OtherTheme_Border
    ctrlEvent.borderWidth = 1
End Sub

'################################################################
'########            2.3 Стиль выходного дня             ########
'########                текущего месяца                 ########
'################################################################
Private Sub ApplyWeekendStyle(ctrlDay As Control, ctrlEvent As Control)
    Const WEEKEND_DARKEN_FACTOR As Double = 0.9 ' 10% затемнение

    ' Label и поле - затемненные цвета
    ctrlDay.ForeColor = CurrentTheme_Text
    ctrlDay.backColor = DarkenColor(CurrentTheme_Back, WEEKEND_DARKEN_FACTOR)
    ctrlDay.borderColor = DarkenColor(CurrentTheme_Border, WEEKEND_DARKEN_FACTOR)
    ctrlDay.borderWidth = 2

    ' TextBox и поле - затемненные цвета
    ctrlEvent.backColor = DarkenColor(CurrentTheme_Back, WEEKEND_DARKEN_FACTOR)
    ctrlEvent.ForeColor = CurrentTheme_Text
    ctrlEvent.borderColor = DarkenColor(CurrentTheme_Border, WEEKEND_DARKEN_FACTOR)
    ctrlEvent.borderWidth = 2
End Sub

'################################################################
'########           3. Выделение сегодняшнего дня        ########
'################################################################
Private Sub HighlightToday(ctrlDay As Control, ctrlEvent As Control, currentDate As Date)

    ' Проверяем, является ли день сегодняшним
    If DateValue(currentDate) = DateValue(Date) Then
        ' Label и поле - выделение сегодняшнего дня
        ctrlDay.backColor = TodayTheme_Back
        ctrlDay.borderColor = TodayTheme_Border
        ctrlDay.borderWidth = 2

        ' TextBox и поле - выделение сегодняшнего дня
        ctrlEvent.borderColor = TodayTheme_Border
        ctrlEvent.borderWidth = 2
    End If

End Sub

'################################################################
'########           4. Загрузка данных событий           ########
'################################################################
Private Sub LoadEventData(ctrlEvent As Control, currentDate As Date)

    On Error GoTo ErrorHandler

    ' Проверка на корректность даты
    If Not IsDate(currentDate) Then
        ctrlEvent.value = ""
        Exit Sub
    End If

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim pendingEvents As String
    Dim completedEvents As String
    Dim pendingCounter As Integer
    Dim completedCounter As Integer
    Dim allCompleted As Boolean
    Dim hasOverdue As Boolean
    Dim hasPending As Boolean
    Dim sqlWhere As String

    Set db = CurrentDb
    pendingEvents = ""
    completedEvents = ""
    pendingCounter = 1
    completedCounter = 1
    allCompleted = False
    hasOverdue = False
    hasPending = False

    ' Формируем SQL условие
    sqlWhere = "WHERE EventDate=#" & Format(currentDate, "yyyy-mm-dd") & "#"

    ' Если включена настройка "Скрыть выполненные"
    If Nz(Me.chkHideCompleted, False) Then
        sqlWhere = sqlWhere & " AND (CompletionMark IS NULL OR CompletionMark = '')"
    End If

    ' Фильтрация по исполнителю
    If Not IsNull(Me.cboExecutorFilter.value) And Me.cboExecutorFilter.value <> "" Then
        sqlWhere = sqlWhere & " AND ExecutorID = " & Me.cboExecutorFilter.value
    End If

    ' Запрос для получения данных
    Set rs = db.OpenRecordset("SELECT EventNote, CompletionMark FROM tbEventInstances " & sqlWhere & " ORDER BY CompletionMark")

    ' Разделение на выполненные и невыполненные
    Do While Not rs.EOF
        If Not IsNull(rs!CompletionMark) And rs!CompletionMark <> "" Then
            ' Выполненные события
            If completedEvents = "" Then
                completedEvents = completedCounter & ". " & rs!EventNote
            Else
                completedEvents = completedEvents & vbCrLf & completedCounter & ". " & rs!EventNote
            End If
            completedCounter = completedCounter + 1
        Else
            ' Невыполненные события
            hasPending = True

            If pendingEvents = "" Then
                pendingEvents = pendingCounter & ". " & rs!EventNote
            Else
                pendingEvents = pendingEvents & vbCrLf & pendingCounter & ". " & rs!EventNote
            End If
            pendingCounter = pendingCounter + 1

            ' Проверка просроченности (если дата уже прошла или сегодня)
            If currentDate <= Date Then
                hasOverdue = True
            End If
        End If
        rs.MoveNext
    Loop

    rs.Close

    ' Определяем, все ли события выполнены
    allCompleted = (pendingEvents = "") And (completedEvents <> "")

    ' Формируем текст и применяем форматирование
    If pendingEvents <> "" And completedEvents <> "" Then
        ctrlEvent.value = pendingEvents & vbCrLf & "----- Выполненные -----" & vbCrLf & completedEvents
        ctrlEvent.FontItalic = False
    ElseIf pendingEvents <> "" Then
        ctrlEvent.value = pendingEvents
        ctrlEvent.FontItalic = False
    ElseIf completedEvents <> "" Then
        ' Если включена настройка "Скрыть выполненные" - не показываем выполненные события
        If Nz(Me.chkHideCompleted, False) Then
            ctrlEvent.value = ""
        Else
            ctrlEvent.value = "----- Выполненные -----" & vbCrLf & completedEvents
        End If
        ctrlEvent.FontItalic = True
    Else
        ctrlEvent.value = ""
        ctrlEvent.FontItalic = False
    End If

    ' Применяем цветовое форматирование
    Call ApplyEventStatusFormatting(ctrlEvent, allCompleted, hasOverdue, hasPending, currentDate)

    Exit Sub

ErrorHandler:
    ctrlEvent.value = "Ошибка загрузки"

End Sub
'################################################################
'########           5. Оформление                        ########
'########              заголовка формы                   ########
'################################################################
Private Sub ApplyFormHeaderStyle()
    On Error Resume Next ' на случай если какие-то элементы отсутствуют

    ' Фон формы (все секции)
    Me.Section(0).backColor = FormTheme_Back
    Me.Section(1).backColor = FormTheme_Back
    Me.Section(2).backColor = FormTheme_Back

    ' Контейнеры подотчетов: применяем только поддерживаемые свойства.
    Call SetControlColorIfSupported(Me.sub_rptEventInstances, "BorderColor", HeaderTheme_Border)
    Call SetControlColorIfSupported(Me.sub_rptEventInstances, "BackColor", FormTheme_Back)
    Call SetControlColorIfSupported(Me.sub_rptBirthdays, "BorderColor", HeaderTheme_Border)
    Call SetControlColorIfSupported(Me.sub_rptBirthdays, "BackColor", FormTheme_Back)

    ' Заголовок месяца и года в шапке
    Me.lbl_MonthYear.ForeColor = HeaderTheme_Text
    Me.lbl_MonthYear.backColor = HeaderTheme_Back

    ' Кнопки навигации по месяцам
    Me.btn_previous.backColor = HeaderTheme_Back
    Me.btn_previous.ForeColor = HeaderTheme_Text
    Me.btn_previous.borderColor = HeaderTheme_Border

    Me.btn_next.backColor = HeaderTheme_Back
    Me.btn_next.ForeColor = HeaderTheme_Text
    Me.btn_next.borderColor = HeaderTheme_Border

    ' Кнопка выбора темы
    Me.btn_theme.backColor = HeaderTheme_Back
    Me.btn_theme.ForeColor = HeaderTheme_Text
    Me.btn_theme.borderColor = HeaderTheme_Border

    ' Кнопка переключения оконного/полноэкранного режима
    Me.btn_window_mode.backColor = HeaderTheme_Back
    Me.btn_window_mode.ForeColor = HeaderTheme_Text
    Me.btn_window_mode.borderColor = HeaderTheme_Border

    ' Кнопка "Текущий месяц"
    Me.btn_current.backColor = HeaderTheme_Back
    Me.btn_current.ForeColor = HeaderTheme_Text
    Me.btn_current.borderColor = HeaderTheme_Border

    ' Кнопка генерации событий
    Me.cmdEvengGenerate.backColor = HeaderTheme_Back
    Me.cmdEvengGenerate.ForeColor = HeaderTheme_Text
    Me.cmdEvengGenerate.borderColor = HeaderTheme_Border

    ' Кнопка исполнителей
    Me.cmdExecutors.backColor = HeaderTheme_Back
    Me.cmdExecutors.ForeColor = HeaderTheme_Text
    Me.cmdExecutors.borderColor = HeaderTheme_Border

    ' Кнопка поиска событий
    Me.cmdSearchEvents.backColor = HeaderTheme_Back
    Me.cmdSearchEvents.ForeColor = HeaderTheme_Text
    Me.cmdSearchEvents.borderColor = HeaderTheme_Border

    ' Кнопка игр
    Me.btn_games.backColor = HeaderTheme_Back
    Me.btn_games.ForeColor = HeaderTheme_Text
    Me.btn_games.borderColor = HeaderTheme_Border

    ' Кнопка демо-режима - скрыта в релизе
    Me.cmdRunDemo.backColor = HeaderTheme_Back
    Me.cmdRunDemo.ForeColor = HeaderTheme_Text
    Me.cmdRunDemo.borderColor = HeaderTheme_Border

    ' Подпись чекбокса "Скрыть выполненные"
    Me.lblChkHideCompleted.ForeColor = HeaderTheme_Text
    Me.lblChkHideCompleted.backColor = FormTheme_Back

    ' Надпись «Дни рождения»
    Me.lblChkShowBirthdays.ForeColor = HeaderTheme_Text
    Me.lblChkShowBirthdays.backColor = FormTheme_Back
    Me.lblChkCurrentDay.ForeColor = HeaderTheme_Text
    Me.lblChkCurrentDay.backColor = FormTheme_Back

    Me.btn_BirthdaysManage.backColor = HeaderTheme_Back
    Me.btn_BirthdaysManage.ForeColor = HeaderTheme_Text
    Me.btn_BirthdaysManage.borderColor = HeaderTheme_Border

    ' Заголовки дней недели (Пн, Вт, Ср...)
    Dim i As Integer
    For i = 1 To 7
        Me.Controls("lbl_weekday_" & i).backColor = HeaderTheme_Back
        Me.Controls("lbl_weekday_" & i).ForeColor = HeaderTheme_Text
    Next i

    ' Альтернативные имена заголовков дней недели (текущий макет формы)
    Me.lbl_Monday.backColor = HeaderTheme_Back
    Me.lbl_Monday.ForeColor = HeaderTheme_Text
    Me.lbl_Tuesday.backColor = HeaderTheme_Back
    Me.lbl_Tuesday.ForeColor = HeaderTheme_Text
    Me.lbl_Wednesday.backColor = HeaderTheme_Back
    Me.lbl_Wednesday.ForeColor = HeaderTheme_Text
    Me.lbl_Thursday.backColor = HeaderTheme_Back
    Me.lbl_Thursday.ForeColor = HeaderTheme_Text
    Me.lbl_Friday.backColor = HeaderTheme_Back
    Me.lbl_Friday.ForeColor = HeaderTheme_Text
    Me.lbl_Saturday.backColor = HeaderTheme_Back
    Me.lbl_Saturday.ForeColor = HeaderTheme_Text
    Me.lbl_Sunday.backColor = HeaderTheme_Back
    Me.lbl_Sunday.ForeColor = HeaderTheme_Text
End Sub

'################################################################
'########           6. Применение темы                   ########
'########              на все элементы                   ########
'################################################################
Public Sub ApplyTheme(ThemeName As String, Optional showMessage As Boolean = False)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb

    ' Сбрасываем активную тему на всех
    db.Execute "UPDATE tbThemes SET IsActive = False"

    ' Устанавливаем активную тему для выбранной
    db.Execute "UPDATE tbThemes SET IsActive = True WHERE ThemeName = '" & ThemeName & "'"

    ' Загружаем новую тему
    Set rs = db.OpenRecordset("SELECT * FROM tbThemes WHERE ThemeName = '" & ThemeName & "'")

    If rs.EOF Then
        MsgBox "Тема '" & ThemeName & "' не найдена!", vbExclamation
        Exit Sub
    End If

    ' Сохраняем значения цветов в переменные модуля
    CurrentTheme_Text = rs!CurrentMonth_Text
    CurrentTheme_Back = rs!CurrentMonth_Back
    CurrentTheme_Border = rs!CurrentMonth_Border

    OtherTheme_Text = rs!OtherMonth_Text
    OtherTheme_Back = rs!OtherMonth_Back
    OtherTheme_Border = rs!OtherMonth_Border

    TodayTheme_Back = rs!Today_Back
    TodayTheme_Border = rs!Today_Border

    HeaderTheme_Text = rs!Header_Text
    HeaderTheme_Back = rs!Header_Back
    HeaderTheme_Border = rs!Header_Border

    FormTheme_Back = rs!Form_Back

    Theme_WritePalette _
        CurrentTheme_Text, CurrentTheme_Back, CurrentTheme_Border, _
        OtherTheme_Text, OtherTheme_Back, OtherTheme_Border, _
        TodayTheme_Back, TodayTheme_Border, _
        HeaderTheme_Text, HeaderTheme_Back, HeaderTheme_Border, _
        FormTheme_Back

    rs.Close

    ' Перестраиваем календарь с новой темой
    Call BuildCalendar
    Call ApplyEventPanelTheme
    On Error Resume Next
    Me.sub_rptEventInstances.Report.ApplyThemeFromHost True
    On Error GoTo ErrorHandler
    On Error Resume Next
    Me.sub_rptBirthdays.Report.ApplyThemeFromHost
    On Error GoTo ErrorHandler

    ' Показываем сообщение если нужно
    If showMessage Then
        MsgBox "Тема '" & ThemeName & "' применена успешно!", vbInformation
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка применения темы: " & Err.description, vbCritical
End Sub

'################################################################
'########           7. Загрузка темы                     ########
'########                  по умолчанию                  ########
'################################################################
Private Sub LoadDefaultTheme()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb

    ' Поиск активной темы (если есть установленная)
    Set rs = db.OpenRecordset("SELECT * FROM tbThemes WHERE IsActive = True")

    If Not rs.EOF Then
        ' Применяем активную тему без сообщения
        ApplyTheme rs!ThemeName, False
    Else
        ' Если активной темы нет, берем первую из списка
        Set rs = db.OpenRecordset("SELECT * FROM tbThemes ORDER BY ThemeID")
        If Not rs.EOF Then
            ApplyTheme rs!ThemeName, False
        Else
            MsgBox "Темы не найдены в базе данных!", vbExclamation
        End If
    End If

    rs.Close
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка загрузки темы: " & Err.description, vbCritical
End Sub

'################################################################
'########          9. Форматирование по статусу         ########
'################################################################
Private Sub ApplyEventStatusFormatting(ctrlEvent As Control, allCompleted As Boolean, hasOverdue As Boolean, hasPending As Boolean, currentDate As Date)
    ' Форматирование текста в ячейке
    ctrlEvent.FontBold = False
    ctrlEvent.ForeColor = CurrentTheme_Text

    ' Если все события выполнены - серый цвет
    If allCompleted Then
        ctrlEvent.ForeColor = RGB(128, 128, 128) ' Серый
        ctrlEvent.FontItalic = True
    ' Если есть просроченные события - красный цвет
    ElseIf hasOverdue And currentDate = Date Then
        ctrlEvent.ForeColor = RGB(255, 0, 0)     ' Красный
        ctrlEvent.FontBold = True
        ctrlEvent.FontItalic = False
    ' Если есть предстоящие в ближайшие 1-3 дня - синий цвет
    ElseIf hasPending And currentDate > Date And currentDate <= Date + 3 Then
        ctrlEvent.ForeColor = RGB(0, 0, 255)     ' Синий
        ctrlEvent.FontBold = True
        ctrlEvent.FontItalic = False
    ' Если есть просроченные события (прошлые) - красный цвет
    ElseIf hasOverdue And currentDate < Date Then
        ctrlEvent.ForeColor = RGB(255, 0, 0)     ' Красный
        ctrlEvent.FontBold = True
        ctrlEvent.FontItalic = False
    ' Иначе стандартный цвет
    Else
        ctrlEvent.ForeColor = CurrentTheme_Text
        ctrlEvent.FontBold = False
        ctrlEvent.FontItalic = False
    End If
End Sub

'################################################################
'########            8. Навигация                        ########
'########                между месяцами                  ########
'################################################################

'################################################################
'########           Кнопка "Следующий месяц"             ########
'################################################################
Private Sub btn_next_Click()
    CurrentMonth = DateAdd("m", 1, CurrentMonth)
    Call BuildCalendar
End Sub

'################################################################
'########           Кнопка "Предыдущий месяц"            ########
'################################################################
Private Sub btn_previous_Click()
    CurrentMonth = DateAdd("m", -1, CurrentMonth)
    Call BuildCalendar
End Sub

'################################################################
'########           Кнопка "Выбор темы"                  ########
'################################################################
Private Sub btn_theme_Click()
    DoCmd.OpenForm "frmThemeSelector", acNormal, , , , acDialog
End Sub

'################################################################
'########          Обработчики клика для выбора          ########
'########              дня панели событий                ########
'################################################################
Private Sub fld_day_1_Click()
    SelectDayForPanelByControl "fld_day_1"
End Sub
Private Sub fld_day_2_Click()
    SelectDayForPanelByControl "fld_day_2"
End Sub
Private Sub fld_day_3_Click()
    SelectDayForPanelByControl "fld_day_3"
End Sub
Private Sub fld_day_4_Click()
    SelectDayForPanelByControl "fld_day_4"
End Sub
Private Sub fld_day_5_Click()
    SelectDayForPanelByControl "fld_day_5"
End Sub
Private Sub fld_day_6_Click()
    SelectDayForPanelByControl "fld_day_6"
End Sub
Private Sub fld_day_7_Click()
    SelectDayForPanelByControl "fld_day_7"
End Sub
Private Sub fld_day_8_Click()
    SelectDayForPanelByControl "fld_day_8"
End Sub
Private Sub fld_day_9_Click()
    SelectDayForPanelByControl "fld_day_9"
End Sub
Private Sub fld_day_10_Click()
    SelectDayForPanelByControl "fld_day_10"
End Sub
Private Sub fld_day_11_Click()
    SelectDayForPanelByControl "fld_day_11"
End Sub
Private Sub fld_day_12_Click()
    SelectDayForPanelByControl "fld_day_12"
End Sub
Private Sub fld_day_13_Click()
    SelectDayForPanelByControl "fld_day_13"
End Sub
Private Sub fld_day_14_Click()
    SelectDayForPanelByControl "fld_day_14"
End Sub
Private Sub fld_day_15_Click()
    SelectDayForPanelByControl "fld_day_15"
End Sub
Private Sub fld_day_16_Click()
    SelectDayForPanelByControl "fld_day_16"
End Sub
Private Sub fld_day_17_Click()
    SelectDayForPanelByControl "fld_day_17"
End Sub
Private Sub fld_day_18_Click()
    SelectDayForPanelByControl "fld_day_18"
End Sub
Private Sub fld_day_19_Click()
    SelectDayForPanelByControl "fld_day_19"
End Sub
Private Sub fld_day_20_Click()
    SelectDayForPanelByControl "fld_day_20"
End Sub
Private Sub fld_day_21_Click()
    SelectDayForPanelByControl "fld_day_21"
End Sub
Private Sub fld_day_22_Click()
    SelectDayForPanelByControl "fld_day_22"
End Sub
Private Sub fld_day_23_Click()
    SelectDayForPanelByControl "fld_day_23"
End Sub
Private Sub fld_day_24_Click()
    SelectDayForPanelByControl "fld_day_24"
End Sub
Private Sub fld_day_25_Click()
    SelectDayForPanelByControl "fld_day_25"
End Sub
Private Sub fld_day_26_Click()
    SelectDayForPanelByControl "fld_day_26"
End Sub
Private Sub fld_day_27_Click()
    SelectDayForPanelByControl "fld_day_27"
End Sub
Private Sub fld_day_28_Click()
    SelectDayForPanelByControl "fld_day_28"
End Sub
Private Sub fld_day_29_Click()
    SelectDayForPanelByControl "fld_day_29"
End Sub
Private Sub fld_day_30_Click()
    SelectDayForPanelByControl "fld_day_30"
End Sub
Private Sub fld_day_31_Click()
    SelectDayForPanelByControl "fld_day_31"
End Sub
Private Sub fld_day_32_Click()
    SelectDayForPanelByControl "fld_day_32"
End Sub
Private Sub fld_day_33_Click()
    SelectDayForPanelByControl "fld_day_33"
End Sub
Private Sub fld_day_34_Click()
    SelectDayForPanelByControl "fld_day_34"
End Sub
Private Sub fld_day_35_Click()
    SelectDayForPanelByControl "fld_day_35"
End Sub
Private Sub fld_day_36_Click()
    SelectDayForPanelByControl "fld_day_36"
End Sub
Private Sub fld_day_37_Click()
    SelectDayForPanelByControl "fld_day_37"
End Sub
Private Sub fld_day_38_Click()
    SelectDayForPanelByControl "fld_day_38"
End Sub
Private Sub fld_day_39_Click()
    SelectDayForPanelByControl "fld_day_39"
End Sub
Private Sub fld_day_40_Click()
    SelectDayForPanelByControl "fld_day_40"
End Sub
Private Sub fld_day_41_Click()
    SelectDayForPanelByControl "fld_day_41"
End Sub
Private Sub fld_day_42_Click()
    SelectDayForPanelByControl "fld_day_42"
End Sub

'################################################################
'########           Обработчики двойного клика           ########
'########               для всех 42 дней                 ########
'################################################################

Private Sub fld_day_1_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_2_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_3_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_4_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_5_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_6_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_7_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_8_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_9_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_10_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_11_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_12_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_13_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_14_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_15_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_16_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_17_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_18_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_19_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_20_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_21_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_22_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_23_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_24_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_25_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_26_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_27_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_28_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_29_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_30_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_31_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_32_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_33_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_34_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_35_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_36_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_37_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_38_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_39_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_40_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_41_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

Private Sub fld_day_42_DblClick(Cancel As Integer)
    OpenDayEventsByControl Me.ActiveControl.Name
End Sub

'################################################################
'########           Чекбокс "Скрыть выполненные"         ########
'################################################################
Private Sub chkHideCompleted_AfterUpdate()
    Call BuildCalendar
    SaveHideCompletedSetting
End Sub

'################################################################
'########      Чекбокс «Показывать дни рождения»         ########
'################################################################
Private Sub chk_ShowBirthdays_AfterUpdate()
    Call SaveShowBirthdaysPanelSetting
    Call ApplyBirthdaysPanelLayout
End Sub

'################################################################
'########      Чекбокс «Текущий день»                    ########
'################################################################
Private Sub chk_CurrentDay_AfterUpdate()
    Call SaveCurrentDaySetting

    If Nz(Me.chk_CurrentDay, True) Then
        m_SelectedPanelDate = Date
    End If

    Call RefreshEventPanel
End Sub

'################################################################
'########          Сохранение настройки фильтра          ########
'################################################################
Private Sub SaveHideCompletedSetting()
    On Error GoTo ErrorHandler
    Dim db As DAO.Database

    Set db = CurrentDb

    db.Execute "DELETE FROM tbSettings WHERE SettingName = 'HideCompleted'"
    db.Execute "INSERT INTO tbSettings (SettingName, SettingValue) VALUES ('HideCompleted', " & IIf(Me.chkHideCompleted, "1", "0") & ")"

    Exit Sub
ErrorHandler:
End Sub

'################################################################
'########      Открытие формы дня по клику               ########
'################################################################
Private Function ResolveCalendarDateByControlName(ByVal controlName As String) As Date
    Dim DayNumber As Integer
    Dim firstVisibleDate As Date

    DayNumber = CInt(Mid(controlName, 9))
    firstVisibleDate = CurrentMonth - weekday(CurrentMonth, vbMonday) + 1
    ResolveCalendarDateByControlName = DateAdd("d", DayNumber - 1, firstVisibleDate)
End Function

'################################################################
'########     Выбор дня календаря для панели             ########
'################################################################
Private Sub SelectDayForPanelByControl(ByVal controlName As String)
    On Error GoTo ExitSub

    If Nz(Me.chk_CurrentDay, True) Then Exit Sub

    m_SelectedPanelDate = ResolveCalendarDateByControlName(controlName)

    Call RefreshEventPanel

ExitSub:
End Sub

'################################################################
'########      Открытие формы дня по клику               ########
'################################################################
Private Sub OpenDayEventsByControl(controlName As String)
    On Error GoTo ErrorHandler

    Dim clickDate As Date
    Dim executorFilter As String

    clickDate = ResolveCalendarDateByControlName(controlName)
    m_SelectedPanelDate = clickDate
    Call RefreshEventPanel

    ' Формируем фильтр по исполнителю
    If Not IsNull(Me.cboExecutorFilter.value) And Me.cboExecutorFilter.value <> "" Then
        executorFilter = " AND ExecutorID = " & Me.cboExecutorFilter.value
    Else
        executorFilter = ""
    End If

    ' Открываем форму и передаем дату и фильтр
    DoCmd.OpenForm "frmDayEvents"
    Forms!frmDayEvents.RecordSource = "SELECT * FROM tbEventInstances WHERE EventDate = " & _
                                      Format(clickDate, "\#mm\/dd\/yyyy\#") & executorFilter
    Forms!frmDayEvents.lblDate.Caption = Format(clickDate, "d mmmm yyyy ""г.""")

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка открытия формы дня: " & Err.description, vbCritical
End Sub

'################################################################
'########       Инициализация фильтра исполнителей       ########
'################################################################
Public Sub InitializeExecutorFilter()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb

    ' Заполнение списка исполнителей
    Me.cboExecutorFilter.rowSource = "SELECT ID, LastName & ' ' & Left(FirstName,1) & '.' & Left(MiddleName,1) & '.' AS FullName " & _
                                    "FROM tbExecutors WHERE ID IS NOT NULL ORDER BY LastName, FirstName"

    Me.cboExecutorFilter.ColumnCount = 2
    Me.cboExecutorFilter.BoundColumn = 1
    Me.cboExecutorFilter.ColumnWidths = "0;5см"

    ' Загрузка сохраненного значения
    Set rs = db.OpenRecordset("SELECT SettingValue FROM tbSettings WHERE SettingName = 'SelectedExecutor'")

    If Not rs.EOF And Not IsNull(rs!settingValue) Then
        Me.cboExecutorFilter.value = rs!settingValue
    Else
        Me.cboExecutorFilter.value = ""
    End If

    rs.Close
    Exit Sub

ErrorHandler:
    Me.cboExecutorFilter.value = ""
    If Not rs Is Nothing Then rs.Close
End Sub

'################################################################
'########      Сохранение выбранного исполнителя         ########
'################################################################
Private Sub SaveExecutorSetting()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim ExecutorID As Variant

    Set db = CurrentDb
    ExecutorID = Me.cboExecutorFilter.value

    ' Удаляем старую запись
    db.Execute "DELETE FROM tbSettings WHERE SettingName = 'SelectedExecutor'"

    ' Сохраняем новое значение
    If Not IsNull(ExecutorID) And ExecutorID <> "" Then
        db.Execute "INSERT INTO tbSettings (SettingName, SettingValue) VALUES ('SelectedExecutor', " & ExecutorID & ")"
    Else
        db.Execute "INSERT INTO tbSettings (SettingName, SettingValue) VALUES ('SelectedExecutor', '')"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка сохранения выбранного исполнителя: " & Err.description, vbExclamation
End Sub

'################################################################
'########       Применение фильтра исполнителя           ########
'################################################################
Private Sub cboExecutorFilter_AfterUpdate()
    On Error GoTo ErrorHandler

    ' Сохраняем настройку
    Call SaveExecutorSetting

    ' Перестраиваем календарь с новым фильтром
    Call BuildCalendar

    ' Синхронно обновляем панель событий дня тем же фильтром исполнителя
    Call RefreshEventPanel

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка применения фильтра: " & Err.description, vbExclamation
End Sub

'################################################################
'########            Кнопка выхода из БД                 ########
'################################################################
Private Sub cmdCloseDataBase_Click()
    On Error GoTo ErrorHandler

    ' Выход из приложения
    Application.Quit

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка при закрытии базы данных: " & Err.description, vbCritical
End Sub

'################################################################
'########             Кнопка "Поиск событий"             ########
'################################################################
Private Sub cmdSearchEvents_Click()
    On Error GoTo ErrorHandler

    ' Открываем форму поиска
    DoCmd.OpenForm "frmSearch"

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка открытия формы поиска: " & Err.description, vbCritical
End Sub

'################################################################
'########          Публичные методы для демо-режима      ########
'################################################################

Public Sub GoToNextMonth()
    Call btn_next_Click
End Sub

Public Sub GoToPreviousMonth()
    Call btn_previous_Click
End Sub

Public Sub GoToCurrentMonth()
    Call btn_current_Click
End Sub

Public Sub OpenDayEvents(DayNumber As Integer)
    Call OpenDayEventsByControl("fld_day_" & DayNumber)
End Sub

Public Sub ApplyExecutorFilter()
    Call cboExecutorFilter_AfterUpdate
End Sub

Public Sub ApplyHideCompletedFilter()
    Call chkHideCompleted_AfterUpdate
End Sub
