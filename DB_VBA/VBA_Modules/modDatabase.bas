Option Compare Database
'################################################################
'########                    ОБЗОР ТАБЛИЦ БАЗЫ ДАННЫХ    ########
'################################################################

' tbThemes              - Цветовые темы оформления календаря
' tbEvents              - Одиночные события (устаревшая, на удаление)
' tbRecurringEvents     - Шаблоны повторяющихся событий
' tbEventInstances      - Конкретные экземпляры событий для календаря
' tbTempEvents          - Временные события для предпросмотра перед сохранением

'################################################################
'########              СТРУКТУРА ОСНОВНЫХ ТАБЛИЦ         ########
'################################################################

' tbThemes:
'   - Хранение цветовых схем (персиковая, медовая, оливковая)
'   - Управление внешним видом календаря

' tbRecurringEvents:
'   - Шаблоны для генерации повторяющихся событий
'   - Параметры: периодичность, интервал, период действия

' tbEventInstances:
'   - Фактические события, отображаемые в календаре
'   - Генерируются из шаблонов или создаются вручную

' tbTempEvents:
'   - Временное хранилище для предпросмотра сгенерированных событий
'   - Позволяет корректировать даты перед сохранением

' tbBirthdays (backend):
'   - Справочник дней рождения (ФИО, дата рождения, примечание)

'################################################################
'########              НАПРАВЛЕНИЯ РАЗВИТИЯ              ########
'################################################################

' 1. Форма для интеллектуального добавления повторяющихся событий
' 2. Алгоритм генерации событий с учетом особенностей месяцев
' 3. Интеграция с основным календарем
' 4. Удаление устаревшей tbEvents

'################################################################
'########           МОДУЛЬ ДЛЯ РАБОТЫ С БАЗОЙ          ########
'################################################################

'################################################################
'########           1. СОЗДАНИЕ ТАБЛИЦЫ ТЕМ           ########
'################################################################
Public Sub CreateThemesTable()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field

    Set db = CurrentDb

    ' Удаляем таблицу если существует
    On Error Resume Next
    db.TableDefs.Delete "tbThemes"
    On Error GoTo ErrorHandler

    ' Создаем новую таблицу
    Set tdf = db.CreateTableDef("tbThemes")

    ' Поля таблицы
    Set fld = tdf.CreateField("ThemeID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    tdf.Fields.Append tdf.CreateField("ThemeName", dbText, 50)
    tdf.Fields.Append tdf.CreateField("IsActive", dbBoolean)

    ' Цвета для текущего месяца
    tdf.Fields.Append tdf.CreateField("CurrentMonth_Text", dbLong)
    tdf.Fields.Append tdf.CreateField("CurrentMonth_Back", dbLong)
    tdf.Fields.Append tdf.CreateField("CurrentMonth_Border", dbLong)

    ' Цвета для других месяцев
    tdf.Fields.Append tdf.CreateField("OtherMonth_Text", dbLong)
    tdf.Fields.Append tdf.CreateField("OtherMonth_Back", dbLong)
    tdf.Fields.Append tdf.CreateField("OtherMonth_Border", dbLong)

    ' Цвета для сегодняшнего дня
    tdf.Fields.Append tdf.CreateField("Today_Back", dbLong)
    tdf.Fields.Append tdf.CreateField("Today_Border", dbLong)

    ' Цвета для заголовка формы
    tdf.Fields.Append tdf.CreateField("Header_Text", dbLong)
    tdf.Fields.Append tdf.CreateField("Header_Back", dbLong)
    tdf.Fields.Append tdf.CreateField("Header_Border", dbLong)

    ' Цвет фона формы
    tdf.Fields.Append tdf.CreateField("Form_Back", dbLong)

    ' Добавляем таблицу в базу
    db.TableDefs.Append tdf

    ' Создаем индекс
    Dim idx As DAO.Index
    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Fields.Append idx.CreateField("ThemeID")
    idx.Primary = True
    tdf.Indexes.Append idx

    MsgBox "Таблица tbThemes создана успешно!", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка при создании таблицы: " & Err.description, vbCritical
End Sub

'################################################################
'########          2. ДОБАВЛЕНИЕ СТАНДАРТНЫХ ТЕМ      ########
'################################################################
Public Sub AddDefaultThemes()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("tbThemes")

    ' 1. ПЕРСИКОВАЯ ТЕМА
    rs.AddNew
    rs!ThemeName = "Персиковая"
    rs!IsActive = True

    ' Цвета текущего месяца
    rs!CurrentMonth_Text = RGB(139, 69, 19)
    rs!CurrentMonth_Back = RGB(255, 240, 230)
    rs!CurrentMonth_Border = RGB(210, 105, 30)

    ' Цвета других месяцев
    rs!OtherMonth_Text = RGB(160, 120, 90)
    rs!OtherMonth_Back = RGB(255, 250, 245)
    rs!OtherMonth_Border = RGB(220, 200, 180)

    ' Сегодня
    rs!Today_Back = RGB(255, 228, 181)
    rs!Today_Border = RGB(255, 140, 0)

    ' Заголовок
    rs!Header_Text = RGB(139, 69, 19)
    rs!Header_Back = RGB(255, 240, 230)
    rs!Header_Border = RGB(210, 105, 30)

    ' Фон формы
    rs!Form_Back = RGB(255, 250, 245)
    rs.Update

    ' 2. МЕДОВАЯ ТЕМА
    rs.AddNew
    rs!ThemeName = "Медовая"
    rs!IsActive = False

    ' Цвета текущего месяца
    rs!CurrentMonth_Text = RGB(101, 67, 33)
    rs!CurrentMonth_Back = RGB(255, 236, 179)
    rs!CurrentMonth_Border = RGB(255, 193, 7)

    ' Цвета других месяцев
    rs!OtherMonth_Text = RGB(150, 120, 80)
    rs!OtherMonth_Back = RGB(255, 248, 225)
    rs!OtherMonth_Border = RGB(255, 224, 130)

    ' Сегодня
    rs!Today_Back = RGB(255, 215, 0)
    rs!Today_Border = RGB(255, 140, 0)

    ' Заголовок
    rs!Header_Text = RGB(101, 67, 33)
    rs!Header_Back = RGB(255, 236, 179)
    rs!Header_Border = RGB(255, 193, 7)

    ' Фон формы
    rs!Form_Back = RGB(255, 248, 225)
    rs.Update

    ' 3. ОЛИВКОВАЯ ТЕМА
    rs.AddNew
    rs!ThemeName = "Оливковая"
    rs!IsActive = False

    ' Цвета текущего месяца
    rs!CurrentMonth_Text = RGB(85, 107, 47)
    rs!CurrentMonth_Back = RGB(240, 255, 240)
    rs!CurrentMonth_Border = RGB(107, 142, 35)

    ' Цвета других месяцев
    rs!OtherMonth_Text = RGB(120, 140, 80)
    rs!OtherMonth_Back = RGB(245, 255, 245)
    rs!OtherMonth_Border = RGB(180, 200, 150)

    ' Сегодня
    rs!Today_Back = RGB(189, 236, 182)
    rs!Today_Border = RGB(85, 160, 70)

    ' Заголовок
    rs!Header_Text = RGB(85, 107, 47)
    rs!Header_Back = RGB(240, 255, 240)
    rs!Header_Border = RGB(107, 142, 35)

    ' Фон формы
    rs!Form_Back = RGB(245, 255, 245)
    rs.Update

    rs.Close
    MsgBox "Добавлено 3 стандартные темы!", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка при добавлении тем: " & Err.description, vbCritical
End Sub

'################################################################
'########          3. ИНИЦИАЛИЗАЦИЯ БАЗЫ ТЕМ          ########
'################################################################
Public Sub InitializeThemes()
    CreateThemesTable
    AddDefaultThemes
End Sub

'################################################################
'########          УСТАНОВКА АКТИВНОЙ ТЕМЫ            ########
'################################################################
Public Sub SetActiveTheme(ThemeName As String)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb

    ' Сначала снимаем активность со всех тем
    db.Execute "UPDATE tbThemes SET IsActive = False"

    ' Устанавливаем активную тему
    db.Execute "UPDATE tbThemes SET IsActive = True WHERE ThemeName = '" & ThemeName & "'"

    MsgBox "Тема '" & ThemeName & "' установлена как активная!", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка установки темы: " & Err.description, vbCritical
End Sub

'################################################################
'########             СОЗДАНИЕ ТАБЛИЦЫ                   ########
'########             tbEventInstances                   ########
'################################################################
Public Sub CreateEventInstancesTable()
    On Error GoTo ErrorHandler
    Dim db As Database: Set db = CurrentDb
    Dim tdf As TableDef
    Dim fld As Field
    Dim idx As Index

    On Error Resume Next
    db.TableDefs.Delete "tbEventInstances"
    On Error GoTo ErrorHandler

    ' Создаем таблицу
    Set tdf = db.CreateTableDef("tbEventInstances")

    ' 1. InstanceID (автоинкремент) - первичный ключ
    Set fld = tdf.CreateField("InstanceID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    ' 2. EventDate (дата события)
    tdf.Fields.Append tdf.CreateField("EventDate", dbDate)

    ' 3. EventNote (текст события)
    tdf.Fields.Append tdf.CreateField("EventNote", dbText, 255)

    ' 4. Notes (пометки)
    tdf.Fields.Append tdf.CreateField("Notes", dbText, 255)

    ' 5. CompletionMark (отметка о выполнении - ТЕКСТ)
    tdf.Fields.Append tdf.CreateField("CompletionMark", dbText, 255)

    ' 6. CompletionDate (дата/время ПЕРВОГО заполнения CompletionMark)
    tdf.Fields.Append tdf.CreateField("CompletionDate", dbDate)

    ' 7. LastModified (дата/время ПОСЛЕДНЕГО изменения CompletionMark)
    tdf.Fields.Append tdf.CreateField("LastModified", dbDate)

    ' 8. AttachmentPath (путь к прикрепленному файлу)
    tdf.Fields.Append tdf.CreateField("AttachmentPath", dbText, 255)

    ' 9. Basis (Основание для проведения мероприятия)
    tdf.Fields.Append tdf.CreateField("Basis", dbText, 255)

    ' 10. BasisAttachment (Приложение для основания)
    tdf.Fields.Append tdf.CreateField("BasisAttachment", dbText, 255)

    ' 11. ExecutorID (Исполнитель мероприятия)
    tdf.Fields.Append tdf.CreateField("ExecutorID", dbLong)

    db.TableDefs.Append tdf

    ' Создаем первичный ключ
    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Fields.Append idx.CreateField("InstanceID")
    idx.Primary = True
    idx.Unique = True
    tdf.Indexes.Append idx

    MsgBox "Таблица tbEventInstances создана с новыми полями!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка: " & Err.description, vbCritical
End Sub

'################################################################
'########             ДОБАВЛЕНИЕ ОПИСАНИЙ ПОЛЕЙ          ########
'########                ДЛЯ tbEventInstances            ########
'################################################################
Public Sub AddEventInstancesFieldDescriptions()
    Dim db As Database: Set db = CurrentDb
    Dim tdf As TableDef
    Dim fld As Field
    Dim prop As Property

    ' Только tbEventInstances
    Set tdf = db.TableDefs("tbEventInstances")
    For Each fld In tdf.Fields
        On Error Resume Next
        Set prop = fld.CreateProperty("Description", dbText, "Описание поля")
        fld.Properties.Append prop
        On Error GoTo 0
    Next fld

    With tdf
        .Fields("InstanceID").Properties("Description") = "Уникальный ID экземпляра события"
        .Fields("EventDate").Properties("Description") = "Дата события"
        .Fields("EventNote").Properties("Description") = "Текст события"
        .Fields("Notes").Properties("Description") = "Пометки к событию"
        .Fields("CompletionMark").Properties("Description") = "Отметка о выполнении (текст)"
        .Fields("CompletionDate").Properties("Description") = "Дата/время первого заполнения отметки о выполнении"
        .Fields("LastModified").Properties("Description") = "Дата/время последнего изменения отметки о выполнении"
        .Fields("AttachmentPath").Properties("Description") = "Путь к прикрепленному файлу"
        .Fields("Basis").Properties("Description") = "Основание для проведения мероприятия"
        .Fields("BasisAttachment").Properties("Description") = "Ссылка на документ-основание"
        .Fields("ExecutorID").Properties("Description") = "Исполнитель мероприятия"
    End With

    MsgBox "Описания полей добавлены в tbEventInstances!", vbInformation
End Sub

'################################################################
'########        СОЗДАНИЕ ТАБЛИЦЫ ДЛЯ ПРЕДПРОСМОТРА      ########
'########                 tbTempEvents                   ########
'################################################################
Public Sub CreateTempEventsTable()
    On Error GoTo ErrorHandler
    Dim db As Database: Set db = CurrentDb
    Dim tdf As TableDef

    On Error Resume Next
    db.TableDefs.Delete "tbTempEvents"
    On Error GoTo ErrorHandler

    Set tdf = db.CreateTableDef("tbTempEvents")
    tdf.Fields.Append tdf.CreateField("TempID", dbLong)
    tdf.Fields.Append tdf.CreateField("EventDate", dbDate)
    tdf.Fields.Append tdf.CreateField("EventNote", dbText, 255)
    tdf.Fields.Append tdf.CreateField("OriginalDay", dbInteger)
    tdf.Fields.Append tdf.CreateField("ExecutorID", dbLong)
    tdf.Fields.Append tdf.CreateField("AttachmentPath", dbText, 255)
    tdf.Fields.Append tdf.CreateField("BasisAttachment", dbText, 255)

    db.TableDefs.Append tdf

    MsgBox "Таблица tbTempEvents создана!", vbInformation

    ' Вызываем отдельную процедуру для описаний
    AddTempEventsFieldDescriptions

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка: " & Err.description, vbCritical
End Sub

'################################################################
'########             ДОБАВЛЕНИЕ ОПИСАНИЙ ПОЛЕЙ          ########
'########                 ДЛЯ tbTempEvents               ########
'################################################################
Public Sub AddTempEventsFieldDescriptions()
    Dim db As Database: Set db = CurrentDb
    Dim tdf As TableDef
    Dim fld As Field
    Dim prop As Property

    Set tdf = db.TableDefs("tbTempEvents")
    For Each fld In tdf.Fields
        On Error Resume Next
        Set prop = fld.CreateProperty("Description", dbText, "Описание поля")
        fld.Properties.Append prop
        On Error GoTo 0
    Next fld

    With tdf
        .Fields("TempID").Properties("Description") = "Временный ID для предпросмотра"
        .Fields("EventDate").Properties("Description") = "Дата события"
        .Fields("EventNote").Properties("Description") = "Текст события"
        .Fields("OriginalDay").Properties("Description") = "Исходный день месяца от пользователя"
        .Fields("ExecutorID").Properties("Description") = "ID исполнителя"
        .Fields("AttachmentPath").Properties("Description") = "Путь к прикрепленному файлу/папке события"
        .Fields("BasisAttachment").Properties("Description") = "Путь к прикрепленному файлу/папке основания"
    End With

    MsgBox "Описания полей добавлены в tbTempEvents!", vbInformation
End Sub

'################################################################
'########             СОЗДАНИЕ ТАБЛИЦЫ                   ########
'########        ПЕРИОДИЧНОСТИ  tbPeriodicity            ########
'########           С ЗАПОЛНЕНИЕМ ДАННЫМИ                ########
'########             И ОПИСАНИЯМИ ПОЛЕЙ                 ########
'################################################################
Public Sub CreatePeriodicityTable()
    On Error GoTo ErrorHandler
    Dim db As Database: Set db = CurrentDb
    Dim tdf As TableDef
    Dim fld As Field
    Dim rs As DAO.Recordset
    Dim prop As Property

    On Error Resume Next
    db.TableDefs.Delete "tbPeriodicity"
    On Error GoTo ErrorHandler

    ' Создаем таблицу с автоинкрементом
    Set tdf = db.CreateTableDef("tbPeriodicity")

    ' Поле с автоинкрементом
    Set fld = tdf.CreateField("PeriodicityID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    ' Остальные поля
    tdf.Fields.Append tdf.CreateField("PeriodicityName", dbText, 50)
    tdf.Fields.Append tdf.CreateField("Description", dbText, 255)

    db.TableDefs.Append tdf

    ' Заполняем данными
    Set rs = db.OpenRecordset("tbPeriodicity")
    With rs
        .AddNew
        !PeriodicityName = "Однократно"
        !description = "Событие происходит один раз в указанную дату"
        .Update

        .AddNew
        !PeriodicityName = "Ежедневно"
        !description = "Событие повторяется каждый день"
        .Update

        .AddNew
        !PeriodicityName = "Еженедельно"
        !description = "Событие повторяется каждую неделю"
        .Update

        .AddNew
        !PeriodicityName = "Ежемесячно"
        !description = "Событие повторяется каждый месяц"
        .Update

        .AddNew
        !PeriodicityName = "Ежеквартально"
        !description = "Событие повторяется каждый квартал (каждые 3 месяца)"
        .Update

        .AddNew
        !PeriodicityName = "Раз в полгода"
        !description = "Событие повторяется каждые 6 месяцев"
        .Update

        .AddNew
        !PeriodicityName = "Ежегодно"
        !description = "Событие повторяется каждый год в ту же дату"
        .Update
    End With
    rs.Close

    ' Добавляем описания полей
    Set tdf = db.TableDefs("tbPeriodicity")
    With tdf
        For Each fld In .Fields
            On Error Resume Next
            Set prop = fld.CreateProperty("Description", dbText, "Описание поля")
            fld.Properties.Append prop
        Next fld

        .Fields("PeriodicityID").Properties("Description") = "Уникальный идентификатор периодичности"
        .Fields("PeriodicityName").Properties("Description") = "Название типа периодичности"
        .Fields("Description").Properties("Description") = "Подробное описание периодичности"
    End With

    MsgBox "Таблица tbPeriodicity создана!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка: " & Err.description, vbCritical
End Sub

'################################################################
'########             СОЗДАНИЕ ТАБЛИЦЫ                   ########
'########             ПРАВИЛ tbRules                     ########
'########           С ЗАПОЛНЕНИЕМ ДАННЫМИ                ########
'########             И ОПИСАНИЯМИ ПОЛЕЙ                 ########
'################################################################
Public Sub CreateRulesTable()
    On Error GoTo ErrorHandler
    Dim db As Database: Set db = CurrentDb
    Dim tdf As TableDef
    Dim fld As Field
    Dim rs As DAO.Recordset
    Dim prop As Property

    On Error Resume Next
    db.TableDefs.Delete "tbRules"
    On Error GoTo ErrorHandler

    ' Создаем таблицу с автоинкрементом
    Set tdf = db.CreateTableDef("tbRules")

    ' Поле с автоинкрементом
    Set fld = tdf.CreateField("RuleID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    ' Остальные поля
    tdf.Fields.Append tdf.CreateField("RuleName", dbText, 100)
    tdf.Fields.Append tdf.CreateField("RuleDescription", dbText, 255)
    tdf.Fields.Append tdf.CreateField("CalculationMethod", dbText, 50)

    db.TableDefs.Append tdf

    ' Заполняем данными
    Set rs = db.OpenRecordset("tbRules")
    With rs
        .AddNew
        !RuleName = "По дате"
        !RuleDescription = "Событие в указанную дату"
        !CalculationMethod = "ByDate"
        .Update

        .AddNew
        !RuleName = "Первый понедельник"
        !RuleDescription = "Событие в первый понедельник месяца"
        !CalculationMethod = "FirstMonday"
        .Update

        .AddNew
        !RuleName = "Последний понедельник"
        !RuleDescription = "Событие в последний понедельник месяца"
        !CalculationMethod = "LastMonday"
        .Update

        .AddNew
        !RuleName = "Первый рабочий день"
        !RuleDescription = "Событие в первый рабочий день месяца"
        !CalculationMethod = "FirstWorkday"
        .Update

        .AddNew
        !RuleName = "Последний рабочий день"
        !RuleDescription = "Событие в последний рабочий день месяца"
        !CalculationMethod = "LastWorkday"
        .Update
    End With
    rs.Close

    ' Добавляем описания полей
    Set tdf = db.TableDefs("tbRules")
    With tdf
        For Each fld In .Fields
            On Error Resume Next
            Set prop = fld.CreateProperty("Description", dbText, "Описание поля")
            fld.Properties.Append prop
        Next fld

        .Fields("RuleID").Properties("Description") = "Уникальный идентификатор правила"
        .Fields("RuleName").Properties("Description") = "Название правила"
        .Fields("RuleDescription").Properties("Description") = "Подробное описание правила"
        .Fields("CalculationMethod").Properties("Description") = "Метод расчета даты события"
    End With

    MsgBox "Таблица tbRules создана!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка: " & Err.description, vbCritical
End Sub

'################################################################
'########           ПОКАЗАТЬ СТРУКТУРУ ТАБЛИЦЫ         ########
'########               TB EVENTINSTANCES              ########
'################################################################
Public Sub ShowTableStructure()
    Dim db As Database: Set db = CurrentDb
    Dim tdf As TableDef
    Dim fld As Field
    Dim structure As String

    Set tdf = db.TableDefs("tbEventInstances")

    structure = "Текущая структура tbEventInstances:" & vbCrLf & vbCrLf

    For Each fld In tdf.Fields
        structure = structure & fld.Name & " (" & GetFieldType(fld.Type) & ")"
        If fld.Attributes And dbAutoIncrField Then
            structure = structure & " [АВТОИНКРЕМЕНТ]"
        End If
        structure = structure & vbCrLf
    Next fld

    Debug.Print structure, vbInformation
End Sub

Private Function GetFieldType(fieldType As Integer) As String
    Select Case fieldType
        Case dbLong: GetFieldType = "Число"
        Case dbDate: GetFieldType = "Дата"
        Case dbText: GetFieldType = "Текст"
        Case dbBoolean: GetFieldType = "Да/Нет"
        Case Else: GetFieldType = "Другой"
    End Select
End Function

'################################################################
'########            ПРОВЕРКА СТРУКТУР ТАБЛИЦ            ########
'################################################################
Public Sub CheckTableStructures()
    Dim db As Database: Set db = CurrentDb
    Dim tdfTemp As TableDef, tdfMain As TableDef
    Dim fld As Field
    Dim msg As String

    ' Проверяем tbTempEvents
    Set tdfTemp = db.TableDefs("tbTempEvents")
    msg = "tbTempEvents:" & vbCrLf
    For Each fld In tdfTemp.Fields
        msg = msg & "  " & fld.Name & vbCrLf
    Next fld

    ' Проверяем tbEventInstances
    Set tdfMain = db.TableDefs("tbEventInstances")
    msg = msg & vbCrLf & "tbEventInstances:" & vbCrLf
    For Each fld In tdfMain.Fields
        msg = msg & "  " & fld.Name & vbCrLf
    Next fld

    Debug.Print msg
    MsgBox msg, vbInformation
End Sub

Private Sub SaveEventsToCalendar()
    Dim db As DAO.Database
    Set db = CurrentDb

    ' Правильный запрос - вставляем только EventDate и EventNote
    db.Execute "INSERT INTO tbEventInstances (EventDate, EventNote) " & _
               "SELECT EventDate, EventNote FROM tbTempEvents"

    MsgBox "События успешно сохранены в календарь!", vbInformation
End Sub

'################################################################
'########            ПРОВЕРКА СТРУКТУР ТАБЛИЦ            ########
'################################################################
Sub ShowTables()
    Dim db As Object
    Dim tdf As Object
    Dim fld As Object

    Set db = CurrentDb

    For Each tdf In db.TableDefs
        If Left(tdf.Name, 4) <> "MSys" Then
            Debug.Print "TABLE: " & tdf.Name
            For Each fld In tdf.Fields
                Debug.Print "  " & fld.Name & " - " & fld.Type
            Next fld
            Debug.Print
        End If
    Next tdf
End Sub

Function GetTypeName(typeNum As Integer) As String
    Select Case typeNum
        Case 1: GetTypeName = "YESNO"
        Case 3: GetTypeName = "LONG"
        Case 4: GetTypeName = "SINGLE"
        Case 5: GetTypeName = "DOUBLE"
        Case 6: GetTypeName = "CURRENCY"
        Case 7: GetTypeName = "DATE"
        Case 10: GetTypeName = "ERROR"
        Case 11: GetTypeName = "YESNO"
        Case 12: GetTypeName = "MEMO"
        Case 15: GetTypeName = "COMPLEX"
        Case 16: GetTypeName = "BIGINT"
        Case 17: GetTypeName = "BINARY"
        Case 18: GetTypeName = "TEXT"
        Case 20: GetTypeName = "LONG"
        Case 21: GetTypeName = "SHORT"
        Case Else: GetTypeName = "UNKNOWN"
    End Select
End Function

Public Sub ShowFormControls()
    Dim ctrl As Control
    For Each ctrl In Forms("f_daily_planner").Controls
        Debug.Print ctrl.Name & " - " & TypeName(ctrl)
    Next ctrl
End Sub

'################################################################
'########     СОЗДАНИЕ ТАБЛИЦЫ НАСТРОЕК tbSettings       ########
'########              С ОПИСАНИЯМИ ПОЛЕЙ                ########
'################################################################
Public Sub CreateSettingsTableWithDescriptions()

    On Error GoTo ErrorHandler

    Dim db As Database
    Dim tdf As TableDef
    Dim fld As Field
    Dim prop As Property
    Dim strSQL As String

    Set db = CurrentDb

    ' Создаем таблицу
    strSQL = "CREATE TABLE tbSettings (" & _
             "SettingName TEXT PRIMARY KEY, " & _
             "SettingValue INTEGER)"

    db.Execute strSQL

    ' Добавляем описания полей
    Set tdf = db.TableDefs("tbSettings")

    For Each fld In tdf.Fields
        On Error Resume Next
        Set prop = fld.CreateProperty("Description", dbText, "Описание поля")
        fld.Properties.Append prop
        On Error GoTo 0
    Next fld

    With tdf
        .Fields("SettingName").Properties("Description") = "Уникальное название настройки"
        .Fields("SettingValue").Properties("Description") = "Значение настройки (числовое)"
    End With

    MsgBox "Таблица tbSettings создана с описаниями полей!", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Таблица уже существует или ошибка: " & Err.description

End Sub

'################################################################
'########        СОЗДАНИЕ ТАБЛИЦЫ ИСПОЛНИТЕЛЕЙ          ########
'########                 tbExecutors                    ########
'################################################################
Public Sub CreateExecutorsTable()
    On Error GoTo ErrorHandler
    Dim db As Database: Set db = CurrentDb
    Dim tdf As TableDef
    Dim fld As Field
    Dim idx As Index

    ' Удаляем таблицу если существует
    On Error Resume Next
    db.TableDefs.Delete "tbExecutors"
    On Error GoTo ErrorHandler

    ' Создаем новую таблицу
    Set tdf = db.CreateTableDef("tbExecutors")

    ' Добавляем поле ID как ключевое с автоинкрементом
    Set fld = tdf.CreateField("ID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    ' Добавляем ФИО и должность (английские названия)
    tdf.Fields.Append tdf.CreateField("LastName", dbText, 100)
    tdf.Fields.Append tdf.CreateField("FirstName", dbText, 100)
    tdf.Fields.Append tdf.CreateField("MiddleName", dbText, 100)
    tdf.Fields.Append tdf.CreateField("Position", dbText, 255)
    tdf.Fields.Append tdf.CreateField("SortOrder", dbLong)
    tdf.Fields.Append tdf.CreateField("Notes", dbText, 255)

    ' Создаем первичный ключ
    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Fields.Append idx.CreateField("ID")
    idx.Primary = True
    idx.Unique = True
    tdf.Indexes.Append idx

    db.TableDefs.Append tdf

    ' Добавляем описания полей
    Set tdf = db.TableDefs("tbExecutors")
    For Each fld In tdf.Fields
        On Error Resume Next
        fld.Properties.Append fld.CreateProperty("Description", dbText, "")
    Next fld

    With tdf
        .Fields("ID").Properties("Description") = "Уникальный ID исполнителя"
        .Fields("LastName").Properties("Description") = "Фамилия исполнителя"
        .Fields("FirstName").Properties("Description") = "Имя исполнителя"
        .Fields("MiddleName").Properties("Description") = "Отчество исполнителя"
        .Fields("Position").Properties("Description") = "Должность исполнителя"
        .Fields("SortOrder").Properties("Description") = "Порядок сортировки в списках"
        .Fields("Notes").Properties("Description") = "Дополнительные заметки"
    End With

    ' Добавляем базовых исполнителей
    AddDefaultExecutors

    MsgBox "Таблица tbExecutors создана успешно!", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка создания таблицы: " & Err.description, vbCritical
End Sub

'################################################################
'########        СОЗДАНИЕ ТАБЛИЦЫ ДНЕЙ РОЖДЕНИЯ          ########
'########                 tbBirthdays                    ########
'################################################################
' Таблица хранится только в backend. Если открыт FE и есть tbTableConnections,
' путь к BE берётся оттуда (как при старте подключения), файл открывается через OpenDatabase.
' Если открыт только BE (в проекте нет локальной tbTableConnections), используется CurrentDb.

Private Function DbHasTable(ByRef db As Database, ByVal tableName As String) As Boolean
    ' Каноническая проверка существования таблицы вынесена в modSchemaSync.DbHasTable.
    ' Здесь оставлен совместимый wrapper без изменения контракта вызовов.
    DbHasTable = modSchemaSync.DbHasTable(db, tableName)
End Function

' Возвращает True и outDb: либо открытый файл backend (outNeedClose=True), либо CurrentDb при работе из BE.
Private Function ResolveBackendDbForBirthdays(ByRef outDb As Database, ByRef outNeedClose As Boolean) As Boolean
    Dim feDb As Database
    Dim rs As Recordset
    Dim p As String

    ResolveBackendDbForBirthdays = False
    outNeedClose = False
    Set outDb = Nothing
    Set feDb = CurrentDb

    If Not DbHasTable(feDb, "tbTableConnections") Then
        Set outDb = feDb
        outNeedClose = False
        ResolveBackendDbForBirthdays = True
        Exit Function
    End If

    On Error GoTo ResolveErr
    Set rs = feDb.OpenRecordset("SELECT TOP 1 TablePath FROM tbTableConnections", dbOpenSnapshot)
    If rs.EOF Then
        rs.Close
        MsgBox "Таблица tbTableConnections пуста. Сначала настройте путь к backend (подключение таблиц).", vbExclamation, "tbBirthdays"
        Exit Function
    End If
    p = Trim(Nz(rs!TablePath, ""))
    rs.Close
    If Len(p) = 0 Then
        MsgBox "В tbTableConnections не задан путь к файлу backend.", vbExclamation, "tbBirthdays"
        Exit Function
    End If
    If Dir(p) = "" Then
        MsgBox "Файл backend не найден:" & vbCrLf & p & vbCrLf & vbCrLf & "Проверьте путь или выполните ConnectAllTables.", vbCritical, "tbBirthdays"
        Exit Function
    End If

    Set outDb = DBEngine(0).OpenDatabase(p, False)
    outNeedClose = True
    ResolveBackendDbForBirthdays = True
    Exit Function

ResolveErr:
    MsgBox "Не удалось открыть backend: " & Err.description, vbCritical, "tbBirthdays"
End Function

Public Sub CreateBirthdaysTable()
    Dim targetDb As Database
    Dim needClose As Boolean

    On Error GoTo ErrorHandler

    If Not ResolveBackendDbForBirthdays(targetDb, needClose) Then Exit Sub

    On Error Resume Next
    targetDb.TableDefs.Delete "tbBirthdays"
    On Error GoTo ErrorHandler

    Call CreateBirthdaysTableAppend(targetDb)

    If needClose Then
        targetDb.Close
        Set targetDb = Nothing
    End If

    MsgBox "Таблица tbBirthdays создана в файле backend.", vbInformation
    Exit Sub

ErrorHandler:
    If needClose Then
        On Error Resume Next
        If Not targetDb Is Nothing Then targetDb.Close
    End If
    MsgBox "Ошибка создания таблицы tbBirthdays: " & Err.description, vbCritical
End Sub

' Однократная миграция: создать tbBirthdays в backend, если её ещё нет (удобно вызывать из FE).
Public Sub MigrateEnsureTbBirthdaysTable()
    Dim targetDb As Database
    Dim needClose As Boolean
    Dim tdf As TableDef

    On Error GoTo ErrorHandler

    If Not ResolveBackendDbForBirthdays(targetDb, needClose) Then Exit Sub

    For Each tdf In targetDb.TableDefs
        If tdf.Name = "tbBirthdays" Then
            If needClose Then targetDb.Close
            MsgBox "Таблица tbBirthdays уже есть в backend — миграция не требуется.", vbInformation
            Exit Sub
        End If
    Next tdf

    Call CreateBirthdaysTableAppend(targetDb)

    If needClose Then
        targetDb.Close
        Set targetDb = Nothing
    End If

    MsgBox "Миграция: таблица tbBirthdays создана в backend.", vbInformation
    Exit Sub

ErrorHandler:
    If needClose Then
        On Error Resume Next
        If Not targetDb Is Nothing Then targetDb.Close
    End If
    MsgBox "Ошибка миграции tbBirthdays: " & Err.description, vbCritical
End Sub

Private Sub CreateBirthdaysTableAppend(db As Database)
    Dim tdf As TableDef
    Dim fld As Field
    Dim idx As Index

    Set tdf = db.CreateTableDef("tbBirthdays")

    Set fld = tdf.CreateField("ID", dbLong)
    fld.Attributes = dbAutoIncrField
    tdf.Fields.Append fld

    tdf.Fields.Append tdf.CreateField("LastName", dbText, 100)
    tdf.Fields.Append tdf.CreateField("FirstName", dbText, 100)
    tdf.Fields.Append tdf.CreateField("MiddleName", dbText, 100)
    tdf.Fields.Append tdf.CreateField("BirthDate", dbDate)
    tdf.Fields.Append tdf.CreateField("Notes", dbText, 255)

    Set idx = tdf.CreateIndex("PrimaryKey")
    idx.Fields.Append idx.CreateField("ID")
    idx.Primary = True
    idx.Unique = True
    tdf.Indexes.Append idx

    db.TableDefs.Append tdf

    Set tdf = db.TableDefs("tbBirthdays")
    For Each fld In tdf.Fields
        On Error Resume Next
        fld.Properties.Append fld.CreateProperty("Description", dbText, "")
    Next fld

    With tdf
        .Fields("ID").Properties("Description") = "Уникальный идентификатор записи"
        .Fields("LastName").Properties("Description") = "Фамилия"
        .Fields("FirstName").Properties("Description") = "Имя"
        .Fields("MiddleName").Properties("Description") = "Отчество"
        .Fields("BirthDate").Properties("Description") = "Дата рождения"
        .Fields("Notes").Properties("Description") = "Примечания"
    End With
End Sub

'################################################################
'########        ДОБАВЛЕНИЕ БАЗОВЫХ ИСПОЛНИТЕЛЕЙ         ########
'################################################################
Private Sub AddDefaultExecutors()
    On Error GoTo ErrorHandler
    Dim db As Database: Set db = CurrentDb

    ' Очищаем таблицу перед добавлением
    db.Execute "DELETE FROM tbExecutors"

    ' Добавляем стандартных исполнителей
    db.Execute "INSERT INTO tbExecutors (LastName, FirstName, MiddleName, Position, SortOrder, Notes) VALUES " & _
               "('', 'General', 'Tasks', 'General Tasks', 1, 'Tasks for all employees')"
    db.Execute "INSERT INTO tbExecutors (LastName, FirstName, MiddleName, Position, SortOrder, Notes) VALUES " & _
               "('Ivanov', 'Ivan', 'Ivanovich', 'Department Manager', 2, '')"
    db.Execute "INSERT INTO tbExecutors (LastName, FirstName, MiddleName, Position, SortOrder, Notes) VALUES " & _
               "('Petrova', 'Maria', 'Sergeevna', 'Sales Manager', 3, '')"
    db.Execute "INSERT INTO tbExecutors (LastName, FirstName, MiddleName, Position, SortOrder, Notes) VALUES " & _
               "('Sidorov', 'Alexey', 'Petrovich', 'Accountant', 4, '')"

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка добавления исполнителей: " & Err.description, vbExclamation
End Sub

'################################################################
'########      ДОБАВЛЕНИЕ ПОЛЕЙ В TBEVENTINSTANCES       ########
'################################################################
Public Sub AddFieldsToEventInstances()
    On Error GoTo ErrorHandler
    Dim db As Database: Set db = CurrentDb
    Dim tdf As TableDef
    Dim fld As Field
    Dim prp As Property

    Set tdf = db.TableDefs("tbEventInstances")

    ' Добавляем поле Basis (Основание) - короткий текст
    On Error Resume Next
    Set fld = tdf.CreateField("Basis", dbText)
    tdf.Fields.Append fld
    On Error GoTo ErrorHandler

    ' Добавляем поле BasisAttachment (Приложение для основания) - короткий текст
    On Error Resume Next
    Set fld = tdf.CreateField("BasisAttachment", dbText)
    tdf.Fields.Append fld
    On Error GoTo ErrorHandler

    ' Добавляем поле ExecutorID (Исполнитель)
    On Error Resume Next
    Set fld = tdf.CreateField("ExecutorID", dbLong)
    tdf.Fields.Append fld
    On Error GoTo ErrorHandler

    ' Добавляем описания полей
    For Each fld In tdf.Fields
        On Error Resume Next
        fld.Properties.Append fld.CreateProperty("Description", dbText, "")
    Next fld

    With tdf
        On Error Resume Next
        .Fields("Basis").Properties("Description") = "Основание для проведения мероприятия"
        .Fields("BasisAttachment").Properties("Description") = "Ссылка на документ-основание"
        .Fields("ExecutorID").Properties("Description") = "Исполнитель мероприятия"
    End With

    ' Обновляем связи если нужно
    RefreshDatabase

    MsgBox "Поля успешно добавлены в tbEventInstances!", vbInformation
    Exit Sub

ErrorHandler:
    If Err.Number = 3211 Then
        MsgBox "Поле уже существует в таблице", vbExclamation
    Else
        MsgBox "Ошибка добавления полей: " & Err.description, vbCritical
    End If
End Sub

'################################################################
'########          ОБНОВЛЕНИЕ СТРУКТУРЫ БАЗЫ             ########
'################################################################
Public Sub RefreshDatabase()
' Назначение: Выполняет пост-обновление структуры таблиц.
' Принцип:    Пересоздаёт связь между tbExecutors и tbEventInstances.
'================================================================
    On Error GoTo ErrorHandler
    CreateRelationship
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка обновления структуры базы: " & Err.description, vbExclamation
End Sub

'################################################################
'########         СОЗДАНИЕ СВЯЗИ МЕЖДУ ТАБЛИЦАМИ         ########
'########           tbExecutors - tbEventInstances       ########
'################################################################
Public Sub CreateRelationship()
    On Error GoTo ErrorHandler
    Dim db As Database: Set db = CurrentDb
    Dim rel As Relation

    ' Удаляем существующую связь если есть
    On Error Resume Next
    db.Relations.Delete "relExecutorsEvents"
    On Error GoTo ErrorHandler

    ' Создаем новую связь
    Set rel = db.CreateRelation("relExecutorsEvents", "tbExecutors", "tbEventInstances")

    ' Добавляем поле связи
    rel.Fields.Append rel.CreateField("ID")
    rel.Fields("ID").ForeignName = "ExecutorID"

    ' Устанавливаем свойства связи
    rel.Attributes = dbRelationUpdateCascade ' Каскадное обновление

    ' Добавляем связь в базу
    db.Relations.Append rel

    MsgBox "Связь между таблицами создана успешно!", vbInformation
    Exit Sub

ErrorHandler:
    If Err.Number = 3211 Then
        MsgBox "Связь уже существует", vbExclamation
    Else
        MsgBox "Ошибка создания связи: " & Err.description, vbCritical
    End If
End Sub

Public Sub ShowThemes()
    Dim rs As Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tbThemes")

    Do While Not rs.EOF
        Debug.Print rs!ThemeName & " - " & rs!ThemeID
        rs.MoveNext
    Loop

    rs.Close
End Sub

'################################################################
'########         ЭКСПОРТ ВСЕГО КОДА В TXT ФАЙЛЫ         ########
'################################################################

Public Sub ExportAllVBAToTxt()
    On Error GoTo ExportAllVBAToTxt_Error

    Dim comp As Object
    Dim exportPath As String
    Dim fso As Object
    Dim txtFile As Object
    Dim i As Integer
    Dim lineCode As String
    Dim filePath As String

    ' ПАПКА ДЛЯ ЭКСПОРТА
    exportPath = "C:\Users\FeNix\Desktop\VBAcode\"
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' СОЗДАЕМ ПАПКУ
    If Not fso.FolderExists(exportPath) Then fso.CreateFolder exportPath

    ' ЭКСПОРТИРУЕМ ВСЕ КОМПОНЕНТЫ В TXT
    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        filePath = exportPath & comp.Name & ".txt"

        ' СОЗДАЕМ TXT ФАЙЛ С КОДИРОВКОЙ UTF-8
        Set txtFile = fso.CreateTextFile(filePath, True, True)

        ' ЗАПИСЫВАЕМ ВСЕ СТРОКИ КОДА
        With comp.CodeModule
            For i = 1 To .CountOfLines
                lineCode = .Lines(i, 1)
                txtFile.WriteLine lineCode
            Next i
        End With

        txtFile.Close
        Set txtFile = Nothing
    Next comp

    MsgBox "Экспорт в TXT завершен! Файлы сохранены в: " & exportPath, vbInformation

    Exit Sub

ExportAllVBAToTxt_Error:
    MsgBox "Ошибка экспорта: " & Err.description, vbCritical
End Sub

'




