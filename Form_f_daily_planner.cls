VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_f_daily_planner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'################################################################
'########              ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ             ########
'################################################################
Dim CurrentMonth As Date

' Глобальные переменные для хранения цветов темы
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

'################################################################
'########           КНОПКА "ТЕКУЩИЙ МЕСЯЦ"               ########
'################################################################
Private Sub btn_current_Click()
    ' Убираем фокус с кнопки перед скрытием
    Me.btn_next.SetFocus
    ' Устанавливаем текущий месяц
    CurrentMonth = DateSerial(Year(Date), Month(Date), 1)
    ' Перестраиваем календарь
    Call BuildCalendar
End Sub

'################################################################
'########           ДВОЙСНОЙ КЛИК СБРОС ФИЛЬТРА          ########
'################################################################
Private Sub cboExecutorFilter_DblClick(Cancel As Integer)
    Me.cboExecutorFilter.Value = Null
    Call cboExecutorFilter_AfterUpdate
End Sub

'################################################################
'########        ОТКРЫТИЕ ФОРМЫ ДОБАЛЕНИЯ СОБЫТИЙ        ########
'################################################################
Private Sub cmdEvengGenerate_Click()
    DoCmd.OpenForm "frmEventGenerator"
End Sub
'################################################################
'########        ОТКРЫТИЕ ФОРМЫ ДОБАЛЕНИЯ СОБЫТИЙ        ########
'################################################################
Private Sub cmdExecutors_Click()
    DoCmd.OpenForm "frmExecutors"
End Sub

'################################################################
'########          ЗАГРУЗКА ФОРМЫ ПРИ ЗАПУСКЕ            ########
'################################################################

Private Sub Form_Load()
    
    ' ПРОВЕРКА ЛИЦЕНЗИИ ПРИ ЗАПУСКЕ
    Call CheckLicenseOnStartup

    Call AutoConnectOnStartup
    ' Устанавливаем текущий месяц
    CurrentMonth = DateSerial(Year(Date), Month(Date), 1)
    
    ' Загружаем тему по умолчанию при старте
    Call LoadDefaultTheme
    
    ' Загружаем настройку фильтра
    Call LoadHideCompletedSetting
    
    ' ИНИЦИАЛИЗИРУЕМ ФИЛЬТР ИСПОЛНИТЕЛЕЙ
    Call InitializeExecutorFilter
    
    ' Перестраиваем календарь с учетом фильтра
    Call BuildCalendar
    
    ' СКРЫВАЕМ БОКОВУЮ ПАНЕЛЬ (NAVIGATION PANE)
    DoCmd.NavigateTo "acNavigationCategoryObjectType"
    DoCmd.RunCommand acCmdWindowHide
    
    ' СКРЫВАЕМ ВЕРХНЮЮ ПАНЕЛЬ (RIBBON)
    DoCmd.ShowToolbar "Ribbon", acToolbarNo
    
    ' УСТАНАВЛИВАЕМ РАЗМЕР И ПОЛОЖЕНИЕ ФОРМЫ
    DoCmd.MoveSize 4500, 600, 18700, 14400
    
End Sub

'################################################################
'########         ЗАГРУЗКА НАСТРОЙКИ ФИЛЬТРА             ########
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
'########              ПОСТРОЕНИЕ КАЛЕНДАРЯ              ########
'########               ОСНОВНАЯ ПРОЦЕДУРА               ########
'################################################################
Public Sub BuildCalendar()
    Dim startDate As Date
    Dim dayCounter As Integer
    Dim ctrlDay As Control
    Dim ctrlEvent As Control
    
    ' Устанавливаем заголовок с названием месяца и года
    Me.lbl_MonthYear.Caption = Format(CurrentMonth, "mmmm yyyy")
    
    ' Применяем стиль заголовка формы
    Call ApplyFormHeaderStyle
    
    ' ВЫЧИСЛЯЕМ ПРАВИЛЬНО: первый понедельник перед месяцем
    ' Weekday с vbMonday возвращает 1 для понедельника, 7 для воскресенья
    startDate = CurrentMonth - weekday(CurrentMonth, vbMonday) + 1
    
    ' Проходим по всем 42 ячейкам календаря (6 недель ? 7 дней)
    For dayCounter = 1 To 42
        Set ctrlDay = Me.Controls("lbl_day_" & dayCounter)
        Set ctrlEvent = Me.Controls("fld_day_" & dayCounter)
        
        ' Устанавливаем число дня
        ctrlDay.Caption = Day(startDate)
        
        ' Вызов отдельных процедур для каждого аспекта оформления
        Call SetEventFieldAccess(ctrlEvent, startDate)
        Call ApplyDayStyling(ctrlDay, ctrlEvent, startDate)
        Call HighlightToday(ctrlDay, ctrlEvent, startDate)
        Call LoadEventData(ctrlEvent, startDate)
        
        ' Переходим к следующему дню
        startDate = DateAdd("d", 1, startDate)
    Next dayCounter
    
    ' Управление видимостью кнопки "Текущий месяц"
    Me.btn_current.Visible = (Month(CurrentMonth) <> Month(Date)) Or (Year(CurrentMonth) <> Year(Date))
End Sub

'################################################################
'########          1. НАСТРОЙКА ДОСТУПНОСТИ              ########
'########               ПОЛЕЙ СОБЫТИЙ                    ########
'################################################################
Private Sub SetEventFieldAccess(ctrlEvent As Control, currentDate As Date)
    ' Для текущего месяца включаем поля (только чтение)
    ' Для других месяцев полностью отключаем
    If Month(currentDate) = Month(CurrentMonth) Then
        ctrlEvent.Enabled = True
        ctrlEvent.Locked = True
    Else
        ctrlEvent.Enabled = False
        ctrlEvent.Locked = True
    End If
End Sub

'################################################################
'########          2. ОФОРМЛЕНИЕ СТИЛЕЙ                  ########
'########                 ДЛЯ ДНЕЙ                       ########
'################################################################
Private Sub ApplyDayStyling(ctrlDay As Control, ctrlEvent As Control, currentDate As Date)
    ' Применяем разные стили для текущего и других месяцев
    If Month(currentDate) = Month(CurrentMonth) Then
        ' Стиль для дней текущего месяца
        If IsWeekend(currentDate) Then
            ' Выходные дни текущего месяца
            ApplyWeekendStyle ctrlDay, ctrlEvent
        Else
            ' Будние дни текущего месяца
            ApplyCurrentMonthStyle ctrlDay, ctrlEvent
        End If
    Else
        ' Стиль для дней других месяцев (прошлых/будущих)
        ApplyOtherMonthStyle ctrlDay, ctrlEvent
    End If
End Sub

'################################################################
'########          2.1 СТИЛЬ ДНЕЙ                        ########
'########             ТЕКУЩЕГО МЕСЯЦА                    ########
'################################################################
Private Sub ApplyCurrentMonthStyle(ctrlDay As Control, ctrlEvent As Control)
    ' Label с числом - текущий месяц
    ctrlDay.ForeColor = CurrentTheme_Text
    ctrlDay.BackColor = CurrentTheme_Back
    ctrlDay.BorderColor = CurrentTheme_Border
    ctrlDay.BorderWidth = 1
    
    ' TextBox с событием - текущий месяц
    ctrlEvent.BackColor = CurrentTheme_Back
    ctrlEvent.ForeColor = CurrentTheme_Text
    ctrlEvent.BorderColor = CurrentTheme_Border
    ctrlEvent.BorderWidth = 1
End Sub

'################################################################
'########          2.2 СТИЛЬ ДНЕЙ                        ########
'########               ДРУГИХ МЕСЯЦЕВ                   ########
'################################################################
Private Sub ApplyOtherMonthStyle(ctrlDay As Control, ctrlEvent As Control)
    ' Label с числом - другие месяцы
    ctrlDay.ForeColor = OtherTheme_Text
    ctrlDay.BackColor = OtherTheme_Back
    ctrlDay.BorderColor = OtherTheme_Border
    ctrlDay.BorderWidth = 1
    
    ' TextBox с событием - другие месяцы
    ctrlEvent.BackColor = OtherTheme_Back
    ctrlEvent.ForeColor = OtherTheme_Text
    ctrlEvent.BorderColor = OtherTheme_Border
    ctrlEvent.BorderWidth = 1
End Sub

'################################################################
'########          2.3 СТИЛЬ ВЫХОДНЫХ ДНЕЙ               ########
'########               ТЕКУЩЕГО МЕСЯЦА                  ########
'################################################################
Private Sub ApplyWeekendStyle(ctrlDay As Control, ctrlEvent As Control)
    Const WEEKEND_DARKEN_FACTOR As Double = 0.9 ' 10% затемнение
    
    ' Label с числом - выходные дни
    ctrlDay.ForeColor = CurrentTheme_Text
    ctrlDay.BackColor = DarkenColor(CurrentTheme_Back, WEEKEND_DARKEN_FACTOR)
    ctrlDay.BorderColor = DarkenColor(CurrentTheme_Border, WEEKEND_DARKEN_FACTOR)
    ctrlDay.BorderWidth = 2
    
    ' TextBox с событием - выходные дни
    ctrlEvent.BackColor = DarkenColor(CurrentTheme_Back, WEEKEND_DARKEN_FACTOR)
    ctrlEvent.ForeColor = CurrentTheme_Text
    ctrlEvent.BorderColor = DarkenColor(CurrentTheme_Border, WEEKEND_DARKEN_FACTOR)
    ctrlEvent.BorderWidth = 2
End Sub

'################################################################
'########          3. ВЫДЕЛЕНИЕ СЕГОДНЯШНЕЙ ДАТЫ        ########
'################################################################
Private Sub HighlightToday(ctrlDay As Control, ctrlEvent As Control, currentDate As Date)
    
    ' Особое оформление для сегодняшней даты
    If DateValue(currentDate) = DateValue(Date) Then
        ' Label с числом - выделение сегодня
        ctrlDay.BackColor = TodayTheme_Back
        ctrlDay.BorderColor = TodayTheme_Border
        ctrlDay.BorderWidth = 2
        
        ' TextBox с событием - выделение сегодня
        ctrlEvent.BorderColor = TodayTheme_Border
        ctrlEvent.BorderWidth = 2
    End If
    
End Sub

'################################################################
'########          4. ЗАГРУЗКА ДАННЫХ СОБЫТИЙ           ########
'################################################################
Private Sub LoadEventData(ctrlEvent As Control, currentDate As Date)
    
    On Error GoTo ErrorHandler
    
    ' Проверяем что дата корректна
    If Not IsDate(currentDate) Then
        ctrlEvent.Value = ""
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
    
    ' Базовый SQL запрос
    sqlWhere = "WHERE EventDate=#" & Format(currentDate, "yyyy-mm-dd") & "#"
    
    ' Если включен фильтр "Скрыть выполненные"
    If Nz(Me.chkHideCompleted, False) Then
        sqlWhere = sqlWhere & " AND (CompletionMark IS NULL OR CompletionMark = '')"
    End If
    
    ' ФИЛЬТРАЦИЯ ПО ИСПОЛНИТЕЛЮ
    If Not IsNull(Me.cboExecutorFilter.Value) And Me.cboExecutorFilter.Value <> "" Then
        sqlWhere = sqlWhere & " AND ExecutorID = " & Me.cboExecutorFilter.Value
    End If
    
    ' Ищем события для этой даты
    Set rs = db.OpenRecordset("SELECT EventNote, CompletionMark FROM tbEventInstances " & sqlWhere & " ORDER BY CompletionMark")
    
    ' Разделяем на выполненные и невыполненные
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
            
            ' Проверяем просроченность (сегодня или уже прошла)
            If currentDate <= Date Then
                hasOverdue = True
            End If
        End If
        rs.MoveNext
    Loop
    
    rs.Close
    
    ' Проверяем все ли события выполнены
    allCompleted = (pendingEvents = "") And (completedEvents <> "")
    
    ' Объединяем с разделителем
    If pendingEvents <> "" And completedEvents <> "" Then
        ctrlEvent.Value = pendingEvents & vbCrLf & "----- ВЫПОЛНЕНО -----" & vbCrLf & completedEvents
        ctrlEvent.FontItalic = False
    ElseIf pendingEvents <> "" Then
        ctrlEvent.Value = pendingEvents
        ctrlEvent.FontItalic = False
    ElseIf completedEvents <> "" Then
        ' Если включен фильтр "Скрыть выполненные" - не показываем выполненные события
        If Nz(Me.chkHideCompleted, False) Then
            ctrlEvent.Value = ""
        Else
            ctrlEvent.Value = "----- ВЫПОЛНЕНО -----" & vbCrLf & completedEvents
        End If
        ctrlEvent.FontItalic = True
    Else
        ctrlEvent.Value = ""
        ctrlEvent.FontItalic = False
    End If
    
    ' Передаем обе переменные
    Call ApplyEventStatusFormatting(ctrlEvent, allCompleted, hasOverdue, hasPending, currentDate)
    
    Exit Sub
    
ErrorHandler:
    ctrlEvent.Value = "Ошибка загрузки"
    
End Sub
'################################################################
'########          5. ОФОРМЛЕНИЕ                         ########
'########             ЗАГОЛОВКА ФОРМЫ                    ########
'################################################################
Private Sub ApplyFormHeaderStyle()
    On Error Resume Next ' На случай если каких-то элементов нет
    
    ' Фон всей формы (все секции)
    Me.Section(0).BackColor = FormTheme_Back
    Me.Section(1).BackColor = FormTheme_Back
    Me.Section(2).BackColor = FormTheme_Back
    
    ' Основной заголовок с названием месяца и года
    Me.lbl_MonthYear.ForeColor = HeaderTheme_Text
    Me.lbl_MonthYear.BackColor = HeaderTheme_Back
    
    ' Кнопки навигации между месяцами
    Me.btn_previous.BackColor = HeaderTheme_Back
    Me.btn_previous.ForeColor = HeaderTheme_Text
    Me.btn_previous.BorderColor = HeaderTheme_Border
    
    Me.btn_next.BackColor = HeaderTheme_Back
    Me.btn_next.ForeColor = HeaderTheme_Text
    Me.btn_next.BorderColor = HeaderTheme_Border
    
    ' Кнопка смены темы
    Me.btn_theme.BackColor = HeaderTheme_Back
    Me.btn_theme.ForeColor = HeaderTheme_Text
    Me.btn_theme.BorderColor = HeaderTheme_Border
    
    ' Кнопка "Текущий месяц"
    Me.btn_current.BackColor = HeaderTheme_Back
    Me.btn_current.ForeColor = HeaderTheme_Text
    Me.btn_current.BorderColor = HeaderTheme_Border
    
    ' Кнопка генератора событий
    Me.cmdEvengGenerate.BackColor = HeaderTheme_Back
    Me.cmdEvengGenerate.ForeColor = HeaderTheme_Text
    Me.cmdEvengGenerate.BorderColor = HeaderTheme_Border
    
    ' НОВЫЕ КНОПКИ - ИСПОЛНИТЕЛИ И ПОИСК
    Me.cmdExecutors.BackColor = HeaderTheme_Back
    Me.cmdExecutors.ForeColor = HeaderTheme_Text
    Me.cmdExecutors.BorderColor = HeaderTheme_Border
    
    Me.cmdSearchEvents.BackColor = HeaderTheme_Back
    Me.cmdSearchEvents.ForeColor = HeaderTheme_Text
    Me.cmdSearchEvents.BorderColor = HeaderTheme_Border
    
    ' Надпись флажка "Скрыть выполненные"
    Me.lblChkHideCompleted.ForeColor = HeaderTheme_Text
    Me.lblChkHideCompleted.BackColor = FormTheme_Back
    
    ' Заголовки дней недели (Пн, Вт, Ср...)
    Dim i As Integer
    For i = 1 To 7
        Me.Controls("lbl_weekday_" & i).BackColor = HeaderTheme_Back
        Me.Controls("lbl_weekday_" & i).ForeColor = HeaderTheme_Text
    Next i
End Sub

'################################################################
'########          6. ПРИМЕНЕНИЕ ТЕМЫ                    ########
'########             ИЗ БАЗЫ ДАННЫХ                     ########
'################################################################
Public Sub ApplyTheme(ThemeName As String, Optional showMessage As Boolean = False)
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    
    ' Сначала снимаем активность со всех тем
    db.Execute "UPDATE tbThemes SET IsActive = False"
    
    ' Устанавливаем активность для выбранной темы
    db.Execute "UPDATE tbThemes SET IsActive = True WHERE ThemeName = '" & ThemeName & "'"
    
    ' Загружаем данные темы
    Set rs = db.OpenRecordset("SELECT * FROM tbThemes WHERE ThemeName = '" & ThemeName & "'")
    
    If rs.EOF Then
        MsgBox "Тема '" & ThemeName & "' не найдена!", vbExclamation
        Exit Sub
    End If
    
    ' Обновляем глобальные переменные с цветами выбранной темы
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
    
    rs.Close
    
    ' Перестраиваем календарь с новой темой
    Call BuildCalendar
    
    ' Показываем сообщение только если явно указано
    If showMessage Then
        MsgBox "Тема '" & ThemeName & "' применена успешно!", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка применения темы: " & Err.Description, vbCritical
End Sub

'################################################################
'########          7. ЗАГРУЗКА ТЕМЫ                      ########
'########                 ПРИ ЗАПУСКЕ                    ########
'################################################################
Private Sub LoadDefaultTheme()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    
    ' Ищем активную тему (которая была выбрана ранее)
    Set rs = db.OpenRecordset("SELECT * FROM tbThemes WHERE IsActive = True")
    
    If Not rs.EOF Then
        ' Применяем активную тему БЕЗ сообщения
        ApplyTheme rs!ThemeName, False
    Else
        ' Если активной темы нет, берем первую и делаем ее активной
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
    MsgBox "Ошибка загрузки темы: " & Err.Description, vbCritical
End Sub

'################################################################
'########          9. ФОРМАТИРОВАНИЕ ПО СТАТУСУ         ########
'################################################################
Private Sub ApplyEventStatusFormatting(ctrlEvent As Control, allCompleted As Boolean, hasOverdue As Boolean, hasPending As Boolean, currentDate As Date)
    ' Сбрасываем стили к базовым
    ctrlEvent.FontBold = False
    ctrlEvent.ForeColor = CurrentTheme_Text
    
    ' Если все события выполнены - светло-серый курсив
    If allCompleted Then
        ctrlEvent.ForeColor = RGB(128, 128, 128) ' Светло-серый
        ctrlEvent.FontItalic = True
    ' Если есть невыполненные СЕГОДНЯ - красный жирный
    ElseIf hasOverdue And currentDate = Date Then
        ctrlEvent.ForeColor = RGB(255, 0, 0)     ' Красный
        ctrlEvent.FontBold = True
        ctrlEvent.FontItalic = False
    ' Если есть невыполненные в ближайшие 1-3 дня - синий жирный
    ElseIf hasPending And currentDate > Date And currentDate <= Date + 3 Then
        ctrlEvent.ForeColor = RGB(0, 0, 255)     ' Синий
        ctrlEvent.FontBold = True
        ctrlEvent.FontItalic = False
    ' Если есть невыполненные просроченные (вчера и ранее) - красный жирный
    ElseIf hasOverdue And currentDate < Date Then
        ctrlEvent.ForeColor = RGB(255, 0, 0)     ' Красный
        ctrlEvent.FontBold = True
        ctrlEvent.FontItalic = False
    ' Во всех остальных случаях - обычный стиль
    Else
        ctrlEvent.ForeColor = CurrentTheme_Text
        ctrlEvent.FontBold = False
        ctrlEvent.FontItalic = False
    End If
End Sub

'################################################################
'########          8. ОБРАБОТЧИКИ                        ########
'########               СОБЫТИЙ ФОРМЫ                    ########
'################################################################

'################################################################
'########          КНОПКА "СЛЕДУЮЩИЙ МЕСЯЦ"              ########
'################################################################
Private Sub btn_next_Click()
    CurrentMonth = DateAdd("m", 1, CurrentMonth)
    Call BuildCalendar
End Sub

'################################################################
'########          КНОПКА "ПРЕДЫДУЩИЙ МЕСЯЦ"             ########
'################################################################
Private Sub btn_previous_Click()
    CurrentMonth = DateAdd("m", -1, CurrentMonth)
    Call BuildCalendar
End Sub

'################################################################
'########          КНОПКА "СМЕНИТЬ ОФОРМЛЕНИЕ"           ########
'################################################################
Private Sub btn_theme_Click()
    ' Открываем форму выбора темы в диалоговом режиме
    DoCmd.OpenForm "frmThemeSelector", , , , , acDialog
End Sub

'################################################################
'########          ОБРАБОТЧИКИ ДВОЙНОГО КЛИКА           ########
'########              ДЛЯ ВСЕХ 42 ПОЛЕЙ                ########
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
'########           ФЛАЖОК "СКРЫТЬ ВЫПОЛНЕННЫЕ"          ########
'################################################################
Private Sub chkHideCompleted_AfterUpdate()
    Call BuildCalendar
    SaveHideCompletedSetting
End Sub

'################################################################
'########         СОХРАНЕНИЕ НАСТРОЙКИ ФИЛЬТРА           ########
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
'########    ОТКРЫТИЕ СОБЫТИЙ ДНЯ С ФИЛЬТРОМ          ########
'################################################################
Private Sub OpenDayEventsByControl(controlName As String)
    On Error GoTo ErrorHandler
    
    Dim dayNumber As Integer
    Dim monthYear As String
    Dim clickDate As Date
    Dim executorFilter As String
    
    ' Извлекаем номер из имени элемента
    dayNumber = CInt(Mid(controlName, 9))
    
    ' Проверяем, что в поле есть число
    If IsNumeric(Me.Controls("lbl_day_" & dayNumber).Caption) Then
        monthYear = Me.lbl_MonthYear.Caption
        clickDate = CStr(Me.Controls("lbl_day_" & dayNumber).Caption) & " " & monthYear
        
        ' Формируем условие фильтра по исполнителю
        If Not IsNull(Me.cboExecutorFilter.Value) And Me.cboExecutorFilter.Value <> "" Then
            executorFilter = " AND ExecutorID = " & Me.cboExecutorFilter.Value
        Else
            executorFilter = ""
        End If
        
        ' Открываем форму и задаем источник данных с фильтром
        DoCmd.OpenForm "frmDayEvents"
        Forms!frmDayEvents.RecordSource = "SELECT * FROM tbEventInstances WHERE EventDate = " & _
                                          Format(clickDate, "\#mm\/dd\/yyyy\#") & executorFilter
        Forms!frmDayEvents.lblDate.Caption = Format(clickDate, "d mmmm yyyy ""г.""")
    Else
        MsgBox "Неверный формат даты", vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка открытия событий дня: " & Err.Description, vbCritical
End Sub

'################################################################
'########          ПРОВЕРКА ВЫХОДНОГО ДНЯ                ########
'################################################################
Private Function IsWeekend(checkDate As Date) As Boolean
    Dim dayOfWeek As Integer
    dayOfWeek = weekday(checkDate, vbMonday) ' Понедельник=1, Воскресенье=7
    IsWeekend = (dayOfWeek = 6) Or (dayOfWeek = 7) ' Суббота или Воскресенье
End Function

'################################################################
'########          ФУНКЦИЯ ЗАТЕМНЕНИЯ ЦВЕТА              ########
'################################################################
Private Function DarkenColor(originalColor As Long, factor As Double) As Long
    Dim R As Integer, G As Integer, B As Integer
    R = originalColor Mod 256
    G = (originalColor \ 256) Mod 256
    B = (originalColor \ 65536) Mod 256
    
    R = R * factor
    G = G * factor
    B = B * factor
    
    DarkenColor = RGB(R, G, B)
End Function

'################################################################
'########      ИНИЦИАЛИЗАЦИЯ ФИЛЬТРА ИСПОЛНИТЕЛЕЙ        ########
'################################################################
Public Sub InitializeExecutorFilter()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    
    ' Устанавливаем источник данных
    Me.cboExecutorFilter.RowSource = "SELECT ID, LastName & ' ' & Left(FirstName,1) & '.' & Left(MiddleName,1) & '.' AS FullName " & _
                                    "FROM tbExecutors WHERE ID IS NOT NULL ORDER BY LastName, FirstName"
    
    Me.cboExecutorFilter.ColumnCount = 2
    Me.cboExecutorFilter.BoundColumn = 1
    Me.cboExecutorFilter.ColumnWidths = "0;5см"
    
    ' Загружаем сохраненную настройку
    Set rs = db.OpenRecordset("SELECT SettingValue FROM tbSettings WHERE SettingName = 'SelectedExecutor'")
    
    If Not rs.EOF And Not IsNull(rs!settingValue) Then
        Me.cboExecutorFilter.Value = rs!settingValue
    Else
        Me.cboExecutorFilter.Value = ""
    End If
    
    rs.Close
    Exit Sub
    
ErrorHandler:
    Me.cboExecutorFilter.Value = ""
    If Not rs Is Nothing Then rs.Close
End Sub

'################################################################
'########    СОХРАНЕНИЕ НАСТРОЙКИ ИСПОЛНИТЕЛЯ         ########
'################################################################
Private Sub SaveExecutorSetting()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim ExecutorID As Variant
    
    Set db = CurrentDb
    ExecutorID = Me.cboExecutorFilter.Value
    
    ' Удаляем старую настройку
    db.Execute "DELETE FROM tbSettings WHERE SettingName = 'SelectedExecutor'"
    
    ' Сохраняем новую настройку
    If Not IsNull(ExecutorID) And ExecutorID <> "" Then
        db.Execute "INSERT INTO tbSettings (SettingName, SettingValue) VALUES ('SelectedExecutor', " & ExecutorID & ")"
    Else
        db.Execute "INSERT INTO tbSettings (SettingName, SettingValue) VALUES ('SelectedExecutor', '')"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка сохранения настройки исполнителя: " & Err.Description, vbExclamation
End Sub

'################################################################
'########    ИЗМЕНЕНИЕ ФИЛЬТРА ИСПОЛНИТЕЛЯ            ########
'################################################################
Private Sub cboExecutorFilter_AfterUpdate()
    On Error GoTo ErrorHandler
    
    ' Сохраняем настройку
    Call SaveExecutorSetting
    
    ' Перестраиваем календарь с новым фильтром
    Call BuildCalendar
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка применения фильтра: " & Err.Description, vbExclamation
End Sub

'################################################################
'########            ЗАКРЫТИЕ БАЗЫ ДАННЫХ                ########
'################################################################
Private Sub cmdCloseDataBase_Click()
    On Error GoTo ErrorHandler
    
    ' Закрываем базу данных
    Application.Quit
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при закрытии базы данных: " & Err.Description, vbCritical
End Sub

'################################################################
'########             КНОПКА "ПОИСК СОБЫТИЙ"             ########
'################################################################
Private Sub cmdSearchEvents_Click()
    On Error GoTo ErrorHandler
    
    ' ОТКРЫВАЕМ ФОРМУ ПОИСКА
    DoCmd.OpenForm "frmSearch"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка открытия формы поиска: " & Err.Description, vbCritical
End Sub
