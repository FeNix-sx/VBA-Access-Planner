Option Compare Database
Option Explicit

'################################################################
'########              МОДУЛЬ ДНИ РОЖДЕНИЯ               ########
'################################################################
' Дата наступления ДР в окне от Date(), текст возраста для отчёта.
' Функции вызываются из запросов Access (см. ТЗ, п.14).

'################################################################
'########       КОНСТАНТЫ ОКНА ПРОСМОТРА И ЮБИЛЕЕВ       ########
'################################################################

' N и M из ТЗ: границы окна [Date()-N .. Date()+M].
Public Const BirthdayWindowDaysBefore As Long = 5
Public Const BirthdayWindowDaysAfter As Long = 10

' Юбилей: возраст не ниже порога и кратен шагу (25, 30, 35, …).
Private Const JubileeAgeMinimum As Long = 25
Private Const JubileeAgeStep As Long = 5

'################################################################
'########      ДАТА НАСТУПЛЕНИЯ ДР В ОКНЕ ПРОСМОТРА      ########
'################################################################
Public Function BirthdayOccurrenceInWindow(ByVal BirthDate As Variant) As Variant
' Назначение: Одна календарная дата наступления дня рождения внутри окна
'             от сегодняшней даты (константы BirthdayWindowDays*).
' Принцип:    Кандидаты в годах Year(Date)-1 … Year(Date)+1; перенос года;
'             29.02 в невисокосный год — 28.02 (ТЗ, п.14).
' Возврат:    Дата наступления или Null (нет даты / вне окна / ошибка приведения).
'################################################################
    Const PROC_NAME As String = "BirthdayOccurrenceInWindow"

    If IsNull(BirthDate) Then
        BirthdayOccurrenceInWindow = Null
        Exit Function
    End If

    On Error GoTo Err_Handler

    Dim bd As Date
    bd = CDate(BirthDate)

    Dim today As Date
    Dim winStart As Date
    Dim winEnd As Date
    Dim y As Long
    Dim cand As Date

    today = DateValue(Date)
    winStart = today - BirthdayWindowDaysBefore
    winEnd = today + BirthdayWindowDaysAfter

    For y = Year(today) - 1 To Year(today) + 1
        cand = BirthdayInCalendarYear(bd, y)
        If cand >= winStart And cand <= winEnd Then
            BirthdayOccurrenceInWindow = cand
            GoTo Exit_Procedure
        End If
    Next y

    BirthdayOccurrenceInWindow = Null

Exit_Procedure:
    Exit Function

Err_Handler:
    Debug.Print "=== Ошибка в процедуре: " & PROC_NAME & " ==="
    Debug.Print "  Описание: " & Err.description
    Debug.Print "  Номер: " & Err.Number
    Debug.Print String(60, "-")
    BirthdayOccurrenceInWindow = Null
    Resume Exit_Procedure
End Function

'################################################################
'########         ТЕКСТ ВОЗРАСТА И ЮБИЛЕЯ ДЛЯ СТРОКИ     ########
'################################################################
Public Function BirthdayAgeText(ByVal BirthDate As Variant, ByVal OccurrenceDate As Variant) As Variant
' Назначение: Строка «исполняется / исполнилось / исполнится … лет» по ТЗ п.6;
'             при наступлении в сегодняшний день — пометка юбилея по правилу модуля.
' Принцип:    Полные годы на дату OccurrenceDate; сравнение с Date();
'             склонение «год / года / лет».
' Возврат:    Текст или Null при пустых аргументах / ошибке приведения.
'################################################################
    Const PROC_NAME As String = "BirthdayAgeText"

    If IsNull(BirthDate) Or IsNull(OccurrenceDate) Then
        BirthdayAgeText = Null
        Exit Function
    End If

    On Error GoTo Err_Handler

    Dim bd As Date
    Dim occ As Date
    Dim age As Long
    Dim occDay As Date
    Dim today As Date
    Dim yrs As String
    Dim t As String

    bd = CDate(BirthDate)
    occ = CDate(OccurrenceDate)
    age = FullYearsAtBirthdayOccurrence(bd, occ)
    occDay = DateValue(occ)
    today = DateValue(Date)
    yrs = RussianYearsWord(age)

    If occDay = today Then
        t = "исполняется " & CStr(age) & " " & yrs
        If IsJubileeAge(age) Then t = t & " (юбилей)"
        BirthdayAgeText = t
    ElseIf occDay < today Then
        BirthdayAgeText = "исполнилось " & CStr(age) & " " & yrs
    Else
        BirthdayAgeText = "исполнится " & CStr(age) & " " & yrs
    End If

Exit_Procedure:
    Exit Function

Err_Handler:
    Debug.Print "=== Ошибка в процедуре: " & PROC_NAME & " ==="
    Debug.Print "  Описание: " & Err.description
    Debug.Print "  Номер: " & Err.Number
    Debug.Print String(60, "-")
    BirthdayAgeText = Null
    Resume Exit_Procedure
End Function

'################################################################
'########       ПОЛНЫЕ ГОДЫ НА ДАТУ НАСТУПЛЕНИЯ ДР       ########
'################################################################
Private Function FullYearsAtBirthdayOccurrence(ByVal BirthDate As Date, ByVal OccurrenceDate As Date) As Long
' Назначение: Возраст в полных годах на календарную дату наступления ДР.
' Принцип:    Сравнение OccurrenceDate с датой ДР в год Year(OccurrenceDate).
' Возврат:    Неотрицательное число лет (при корректной паре дат из запроса).
'################################################################
    Dim yOcc As Long
    Dim bdayThisYear As Date
    Dim age As Long

    yOcc = Year(OccurrenceDate)
    bdayThisYear = BirthdayInCalendarYear(BirthDate, yOcc)
    age = yOcc - Year(BirthDate)
    If DateValue(OccurrenceDate) < DateValue(bdayThisYear) Then
        age = age - 1
    End If
    FullYearsAtBirthdayOccurrence = age
End Function

'################################################################
'########     ДЕНЬ И МЕСЯЦ ДР В ЗАДАННОМ КАЛЕНДАРНОМ     ########
'########                      ГОДУ                      ########
'################################################################
Private Function BirthdayInCalendarYear(ByVal BirthDate As Date, ByVal CalendarYear As Long) As Date
' Назначение: Дата наступления ДР в указанном календарном году.
' Принцип:    DateSerial; особый случай 29.02 при невисокосном годе.
' Возврат:    Календарная дата в году CalendarYear.
'################################################################
    Dim m As Long
    Dim d As Long

    m = Month(BirthDate)
    d = Day(BirthDate)
    If m = 2 And d = 29 And Not IsGregorianLeapYear(CalendarYear) Then
        BirthdayInCalendarYear = DateSerial(CalendarYear, 2, 28)
    Else
        BirthdayInCalendarYear = DateSerial(CalendarYear, m, d)
    End If
End Function

'################################################################
'########         ВИСОКОСНЫЙ ГОД (ГРИГОРИАНСКИЙ)         ########
'################################################################
Private Function IsGregorianLeapYear(ByVal y As Long) As Boolean
' Назначение: Определение високосного года для переноса 29.02.
' Возврат:    True, если в году есть 29 февраля.
'################################################################
    IsGregorianLeapYear = ((y Mod 4 = 0 And y Mod 100 <> 0) Or (y Mod 400 = 0))
End Function

'################################################################
'########          ПРИЗНАК ЮБИЛЕЙНОГО ВОЗРАСТА           ########
'################################################################
Private Function IsJubileeAge(ByVal age As Long) As Boolean
' Назначение: Единое правило пометки «юбилей» в тексте строки.
' Возврат:    True, если возраст >= JubileeAgeMinimum и кратен JubileeAgeStep.
'################################################################
    If age < JubileeAgeMinimum Then Exit Function
    IsJubileeAge = (age Mod JubileeAgeStep = 0)
End Function

'################################################################
'########        СКЛОНЕНИЕ СЛОВА «ГОД» ДЛЯ ЧИСЛА         ########
'################################################################
Private Function RussianYearsWord(ByVal n As Long) As String
' Назначение: Подпись к числу лет (1 год, 2 года, 5 лет, 21 год …).
' Возврат:    Одно из слов: год / года / лет.
'################################################################
    Dim m100 As Long

    m100 = n Mod 100
    If m100 >= 11 And m100 <= 14 Then
        RussianYearsWord = "лет"
        Exit Function
    End If
    Select Case n Mod 10
        Case 1: RussianYearsWord = "год"
        Case 2, 3, 4: RussianYearsWord = "года"
        Case Else: RussianYearsWord = "лет"
    End Select
End Function

'################################################################
'########     ЗАПРОС И ОТЧЁТ ПАНЕЛИ «ДНИ РОЖДЕНИЯ»       ########
'################################################################
Public Sub EnsureQryBirthdaysForPanel()
' Назначение: Создаёт или обновляет сохранённый запрос qryBirthdaysForPanel
'             (окно дат, сортировка по номерам полей в SELECT — см. core-settings).
' Принцип:    DAO QueryDef; SQL с подзапросом и функциями модуля.
'################################################################
    Const PROC_NAME As String = "EnsureQryBirthdaysForPanel"
    Const QRY_NAME As String = "qryBirthdaysForPanel"

    On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim qd As DAO.QueryDef
    Dim sSql As String
    Dim errOpen As Long

    sSql = SqlQryBirthdaysForPanel()
    Set db = CurrentDb
    Set qd = Nothing

    On Error Resume Next
    Set qd = db.QueryDefs(QRY_NAME)
    errOpen = Err.Number
    On Error GoTo Err_Handler

    If errOpen <> 0 Then
        Set qd = db.CreateQueryDef(QRY_NAME, sSql)
    Else
        qd.sql = sSql
    End If

Exit_Procedure:
    Set qd = Nothing
    Set db = Nothing
    Exit Sub

Err_Handler:
    Debug.Print "=== Ошибка в процедуре: " & PROC_NAME & " ==="
    Debug.Print "  Описание: " & Err.description
    Debug.Print "  Номер: " & Err.Number
    Debug.Print String(60, "-")
    MsgBox "Ошибка в процедуре " & PROC_NAME & ":" & vbCrLf & _
           Err.description & vbCrLf & "(номер: " & Err.Number & ")", vbCritical
    Resume Exit_Procedure
End Sub

'################################################################
'########   ЗАГРУЗКА ОТЧЁТА rptBirthdays ИЗ ФАЙЛА        ########
'################################################################
Public Sub EnsureRptBirthdaysFromExportFile()
' Назначение: Импортирует отчёт rptBirthdays из DB_VBA\VBA_Reports\rptBirthdays.txt
'             (группировка по дате наступления ДР, макет панели).
' Принцип:    Application.LoadFromText; путь относительно папки .accdb.
'################################################################
    Const PROC_NAME As String = "EnsureRptBirthdaysFromExportFile"
    Const RPT_NAME As String = "rptBirthdays"
    Const REL_PATH As String = "\DB_VBA\VBA_Reports\rptBirthdays.txt"

    On Error GoTo Err_Handler

    Dim sPath As String
    sPath = CurrentProject.path & REL_PATH

    If Len(Dir(sPath)) = 0 Then
        MsgBox "Файл разметки отчёта не найден:" & vbCrLf & sPath, vbExclamation, PROC_NAME
        GoTo Exit_Procedure
    End If

    Application.LoadFromText acReport, RPT_NAME, sPath

Exit_Procedure:
    Exit Sub

Err_Handler:
    Debug.Print "=== Ошибка в процедуре: " & PROC_NAME & " ==="
    Debug.Print "  Описание: " & Err.description
    Debug.Print "  Номер: " & Err.Number
    Debug.Print String(60, "-")
    MsgBox "Ошибка в процедуре " & PROC_NAME & ":" & vbCrLf & _
           Err.description & vbCrLf & "(номер: " & Err.Number & ")", vbCritical
    Resume Exit_Procedure
End Sub

'################################################################
'########     ЗАПРОС + ОТЧЁТ (ОДНА ТОЧКА ДЛЯ МИГРАЦИИ)    ########
'################################################################
Public Sub EnsureBirthdaysPanelQueryAndReport()
' Назначение: Создать/обновить qryBirthdaysForPanel и загрузить rptBirthdays.
'################################################################
    EnsureQryBirthdaysForPanel
    EnsureRptBirthdaysFromExportFile
End Sub

'################################################################
'########    УВЕДОМЛЕНИЕ О ДР ПРИ ПЕРВОМ ЗАПУСКЕ ДНЯ     ########
'################################################################
Public Sub NotifyTodaysBirthdaysOncePerDay()
' Назначение: Один раз в сутки показать уведомление о сегодняшних днях рождения.
' Принцип:    Проверяет ключ BirthdaysNotifyDate в tbSettings; если на сегодня не
'             показывали, собирает именинников за сегодня и выводит MsgBox.
'             Для каждой записи 2 строки: ФИО+возраст и примечание.
'################################################################
    Const PROC_NAME As String = "NotifyTodaysBirthdaysOncePerDay"
    Const SETTING_NAME As String = "BirthdaysNotifyDate"

    On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim todayKey As String
    Dim savedKey As String
    Dim strSQL As String
    Dim msgText As String
    Dim FullName As String
    Dim AgeText As String
    Dim noteText As String

    Set db = CurrentDb
    todayKey = Format(DateValue(Date), "yyyy-mm-dd")
    savedKey = ""

    Set rs = db.OpenRecordset("SELECT SettingValue FROM tbSettings WHERE SettingName = '" & SETTING_NAME & "'", dbOpenSnapshot)
    If Not rs.EOF Then
        savedKey = Nz(rs!settingValue, "")
    End If
    rs.Close
    Set rs = Nothing

    If savedKey = todayKey Then GoTo Exit_Procedure

    strSQL = "SELECT Trim([LastName] & ' ' & [FirstName] & IIf(Len(Nz([MiddleName],''))=0,'',' ' & [MiddleName])) AS FullName, " & _
             "BirthdayAgeText([BirthDate], BirthdayOccurrenceInWindow([BirthDate])) AS AgeText, " & _
             "Notes " & _
             "FROM tbBirthdays " & _
             "WHERE BirthdayOccurrenceInWindow([BirthDate]) = Date() " & _
             "ORDER BY [LastName], [FirstName], [MiddleName];"

    Set rs = db.OpenRecordset(strSQL, dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        GoTo Exit_Procedure
    End If

    msgText = "Сегодня дни рождения:" & vbCrLf & String(28, "-") & vbCrLf & vbCrLf

    Do While Not rs.EOF
        FullName = Trim(Nz(rs!FullName, ""))
        AgeText = Trim(Nz(rs!AgeText, ""))
        noteText = Trim(Nz(rs!Notes, ""))
        If Len(noteText) = 0 Then noteText = "—"

        msgText = msgText & FullName & " — " & AgeText & vbCrLf & _
                  "Примечание: " & noteText & vbCrLf & vbCrLf
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing

    MsgBox msgText, vbInformation, "Напоминание о днях рождения"

    db.Execute "DELETE FROM tbSettings WHERE SettingName = '" & SETTING_NAME & "'", dbFailOnError
    db.Execute "INSERT INTO tbSettings (SettingName, SettingValue) VALUES ('" & SETTING_NAME & "', '" & todayKey & "')", dbFailOnError

Exit_Procedure:
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

Err_Handler:
    Debug.Print "=== Ошибка в процедуре: " & PROC_NAME & " ==="
    Debug.Print "  Описание: " & Err.description
    Debug.Print "  Номер: " & Err.Number
    Debug.Print String(60, "-")
    Resume Exit_Procedure
End Sub

'################################################################
'########     ДЕМО-ЗАПИСИ В tbBirthdays ДЛЯ ПРОВЕРКИ       ########
'################################################################
Public Sub SeedBirthdaysTestData()
' Назначение: Добавить 10 человек в tbBirthdays для проверки отчёта; двое с ДР в одну дату.
' Принцип:    Даты рождения строятся от Date() (окно N/M), чтобы записи попадали в выборку;
'             сначала удаляются только строки с Notes = маркер демо.
'################################################################
    Const PROC_NAME As String = "SeedBirthdaysTestData"
    Const DEMO_NOTES As String = "Демо-данные"

    On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim today As Date
    Dim dSame As Date

    today = DateValue(Date)
    dSame = DateAdd("d", 3, today)

    Set db = CurrentDb
    db.Execute "DELETE FROM tbBirthdays WHERE Notes = '" & Replace(DEMO_NOTES, "'", "''") & "'", dbFailOnError

    Set rs = db.OpenRecordset("tbBirthdays", dbOpenDynaset)

    SeedBirthdaysDemoRow rs, "Иванов", "Иван", "Иванович", BirthOnCalendarDay(1990, today), DEMO_NOTES
    SeedBirthdaysDemoRow rs, "Петрова", "Мария", "Сергеевна", BirthOnCalendarDay(1988, DateAdd("d", 1, today)), DEMO_NOTES
    SeedBirthdaysDemoRow rs, "Сидоров", "Пётр", "Александрович", BirthOnCalendarDay(1985, dSame), DEMO_NOTES
    SeedBirthdaysDemoRow rs, "Козлова", "Анна", "Дмитриевна", BirthOnCalendarDay(1992, dSame), DEMO_NOTES
    SeedBirthdaysDemoRow rs, "Волков", "Олег", "Николаевич", BirthOnCalendarDay(1975, DateAdd("d", 5, today)), DEMO_NOTES
    SeedBirthdaysDemoRow rs, "Новикова", "Елена", "Викторовна", BirthOnCalendarDay(2001, DateAdd("d", -2, today)), DEMO_NOTES
    SeedBirthdaysDemoRow rs, "Морозов", "Дмитрий", "Павлович", BirthOnCalendarDay(1993, DateAdd("d", 7, today)), DEMO_NOTES
    SeedBirthdaysDemoRow rs, "Соколова", "Ирина", "Олеговна", BirthOnCalendarDay(1980, DateAdd("d", -5, today)), DEMO_NOTES
    SeedBirthdaysDemoRow rs, "Лебедев", "Андрей", "Игоревич", BirthOnCalendarDay(1995, DateAdd("d", 4, today)), DEMO_NOTES
    SeedBirthdaysDemoRow rs, "Орлов", "Сергей", Null, BirthOnCalendarDay(1988, DateAdd("d", -1, today)), DEMO_NOTES

    rs.Close
    Set rs = Nothing
    Set db = Nothing

    MsgBox "Добавлено 10 демо-записей в tbBirthdays (Notes = """ & DEMO_NOTES & """)." & vbCrLf & _
           "Два человека: Сидоров и Козлова — на одну дату наступления ДР (сегодня + 3 дня).", _
           vbInformation, PROC_NAME

Exit_Procedure:
    Exit Sub

Err_Handler:
    Debug.Print "=== Ошибка в процедуре: " & PROC_NAME & " ==="
    Debug.Print "  Описание: " & Err.description
    Debug.Print "  Номер: " & Err.Number
    Debug.Print String(60, "-")
    MsgBox "Ошибка в процедуре " & PROC_NAME & ":" & vbCrLf & _
           Err.description & vbCrLf & "(номер: " & Err.Number & ")", vbCritical
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Resume Exit_Procedure
End Sub

Private Function BirthOnCalendarDay(ByVal birthYear As Long, ByVal calDay As Date) As Date
' Назначение: Полная дата рождения с заданным годом и календарным днём/месяцем calDay.
'################################################################
    BirthOnCalendarDay = DateSerial(birthYear, Month(calDay), Day(calDay))
End Function

Private Sub SeedBirthdaysDemoRow(ByVal rs As DAO.Recordset, _
                                 ByVal LastName As String, _
                                 ByVal FirstName As String, _
                                 ByVal MiddleName As Variant, _
                                 ByVal BirthDate As Date, _
                                 ByVal notesText As String)
' Назначение: Одна строка в tbBirthdays (AddNew/Update).
'################################################################
    rs.AddNew
    rs!LastName = LastName
    rs!FirstName = FirstName
    If IsNull(MiddleName) Then
        rs!MiddleName = Null
    Else
        rs!MiddleName = MiddleName
    End If
    rs!BirthDate = BirthDate
    rs!Notes = notesText
    rs.Update
End Sub

'################################################################
'########              ТЕКСТ SQL ЗАПРОСА                   ########
'################################################################
Private Function SqlQryBirthdaysForPanel() As String
' Назначение: SQL для qryBirthdaysForPanel и свойства RecordSource отчёта rptBirthdays.
' Примечание: OccurrenceSortKey (yyyymmdd) — сортировка и группировка по календарю; при группировке
'             отчёта только по OccurrenceDate с форматом «дд.мм.гггг» Access нередко сортирует
'             группы как текст (апрель оказывается перед мартом). Группируйте по OccurrenceSortKey,
'             в шапке группы выводите =Min([OccurrenceDate]) с нужным форматом.
'################################################################
    Dim s As String
    Dim fmt As String

    fmt = Chr(34) & "yyyymmdd" & Chr(34)

    s = "SELECT q.ID, q.LastName, q.FirstName, q.MiddleName, q.BirthDate, q.Notes, q.OccurrenceDate, " & _
        "Format([q].[OccurrenceDate], " & fmt & ") AS OccurrenceSortKey, " & _
        "BirthdayAgeText([q].[BirthDate],[q].[OccurrenceDate]) AS AgeText, " & _
        "Trim([q].[LastName] & " & Chr(34) & " " & Chr(34) & " & [q].[FirstName] & " & _
        "IIf(Len(Nz([q].[MiddleName], " & Chr(34) & Chr(34) & "))=0, " & Chr(34) & Chr(34) & ", " & _
        Chr(34) & " " & Chr(34) & " & [q].[MiddleName])) AS FullName " & _
        "FROM (" & _
        "SELECT tbBirthdays.ID, tbBirthdays.LastName, tbBirthdays.FirstName, tbBirthdays.MiddleName, " & _
        "tbBirthdays.BirthDate, tbBirthdays.Notes, " & _
        "BirthdayOccurrenceInWindow([BirthDate]) AS OccurrenceDate " & _
        "FROM tbBirthdays" & _
        ") AS q " & _
        "WHERE q.OccurrenceDate Is Not Null " & _
        "ORDER BY Format([q].[OccurrenceDate], " & fmt & "), " & _
        "[q].[LastName], [q].[FirstName], [q].[MiddleName];"

    SqlQryBirthdaysForPanel = s
End Function

'################################################################
'########   ТЕКСТ SQL ДЛЯ ВСТАВКИ В RecordSource ОТЧЁТА    ########
'################################################################
Public Function BirthdaysPanelRecordSourceSql() As String
' Назначение: Тот же SQL, что у сохранённого запроса — для вставки в свойство отчёта вручную.
'################################################################
    BirthdaysPanelRecordSourceSql = SqlQryBirthdaysForPanel()
End Function

'################################################################
'########   ОБНОВЛЕНИЕ СПИСКА И ПАНЕЛИ ДР ПОСЛЕ ПРАВОК     ########
'################################################################
Public Sub RefreshBirthdaysUIAfterEdit()
' Назначение: После сохранения/закрытия карточки или удаления в списке — обновить
'             frmBirthdaysList (если открыта) и подчинённый отчёт на f_daily_planner (если есть и виден).
'################################################################
    On Error Resume Next
    If CurrentProject.allForms("frmBirthdaysList").IsLoaded Then
        Forms!frmBirthdaysList.Requery
    End If
    If CurrentProject.allForms("f_daily_planner").IsLoaded Then
        If Forms!f_daily_planner!sub_rptBirthdays.Visible = True Then
            Forms!f_daily_planner!sub_rptBirthdays.Report.Requery
        End If
    End If
    On Error GoTo 0
End Sub


