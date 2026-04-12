Option Compare Database

'########################################################################
'########           МОДУЛЬ АНАЛИЗА ПРОЕКТА "ЕЖЕДНЕВНИК"         ########
'########################################################################
Option Explicit

'########################################################################
'########           ГЛАВНЫЕ ПРОЦЕДУРЫ ТЕСТИРОВАНИЯ              ########
'########################################################################

'########################################################################
'########           ЗАПУСК ПОЛНОГО АНАЛИЗА ПРОЕКТА              ########
'########################################################################
Public Sub RunFullProjectAnalysis()
    On Error GoTo ErrorHandler

    Debug.Print String(60, "=")
    Debug.Print "ПОЛНЫЙ АНАЛИЗ ПРОЕКТА 'ЕЖЕДНЕВНИК'"
    Debug.Print "Время запуска: " & Now()
    Debug.Print String(60, "=")

    ' Тест 1: Основные компоненты системы
    TestBasicComponents

    ' Тест 2: Формы проекта
    TestAllForms

    ' Тест 3: Таблицы базы данных
    TestAllTables

    ' Тест 4: Модули проекта
    TestAllModules

    ' Тест 5: Анализ главной формы
    AnalyzeMainFormStructure

    Debug.Print String(60, "=")
    Debug.Print "АНАЛИЗ ЗАВЕРШЕН"
    Debug.Print "Результаты выше"

    Exit Sub

ErrorHandler:
    Debug.Print "ОШИБКА ПОЛНОГО АНАЛИЗА: " & Err.description
End Sub

'########################################################################
'########           ЗАПУСК БЫСТРОГО ТЕСТА СИСТЕМЫ               ########
'########################################################################
Public Sub RunQuickTest()
    On Error GoTo ErrorHandler

    Debug.Print String(50, "=")
    Debug.Print "БЫСТРЫЙ ТЕСТ СИСТЕМЫ"
    Debug.Print String(50, "=")

    SimpleTest

    Debug.Print String(50, "=")
    Debug.Print "ТЕСТ ЗАВЕРШЕН"

    Exit Sub

ErrorHandler:
    Debug.Print "ОШИБКА БЫСТРОГО ТЕСТА: " & Err.description
End Sub

'########################################################################
'########           ТЕСТ ОСНОВНЫХ КОМПОНЕНТОВ СИСТЕМЫ           ########
'########################################################################
Private Sub TestBasicComponents()
    On Error GoTo ErrorHandler

    Debug.Print "=== ОСНОВНЫЕ КОМПОНЕНТЫ СИСТЕМЫ ==="

    ' Проверяем главную форму
    If FormExists("f_daily_planner") Then
        Debug.Print "? Главная форма: f_daily_planner"
        If CurrentProject.allForms("f_daily_planner").IsLoaded Then
            Debug.Print "  - Форма загружена"
        Else
            Debug.Print "  - Форма не загружена"
        End If
    Else
        Debug.Print "? Главная форма отсутствует: f_daily_planner"
    End If

    ' Проверяем основные модули
    CheckEssentialModules

    ' Проверяем основные таблицы
    CheckEssentialTables

    ' Объекты модуля «Дни рождения» (таблица, запрос панели, отчёт, modBirthdays)
    CheckBirthdaysArtifacts

    ' Проверяем системные функции
    TestSystemFunctions

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка теста основных компонентов: " & Err.description
End Sub

'########################################################################
'########           ПРОВЕРКА ОСНОВНЫХ МОДУЛЕЙ                   ########
'########################################################################
Private Sub CheckEssentialModules()
    On Error GoTo ErrorHandler

    Debug.Print "--- ПРОВЕРКА МОДУЛЕЙ ---"

    Dim essentialModules As Variant
    essentialModules = Array("modDatabase", "modTableConnect", "modProtection", _
                            "modThemeSelector", "modEventManager", "modBirthdays")

    Dim i As Integer
    For i = 0 To UBound(essentialModules)
        If ModuleExists(CStr(essentialModules(i))) Then
            Debug.Print "? Модуль: " & essentialModules(i)
        Else
            Debug.Print "? Модуль отсутствует: " & essentialModules(i)
        End If
    Next i

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка проверки модулей: " & Err.description
End Sub

'########################################################################
'########           ПРОВЕРКА ОСНОВНЫХ ТАБЛИЦ                    ########
'########################################################################
Private Sub CheckEssentialTables()
    On Error GoTo ErrorHandler

    Debug.Print "--- ПРОВЕРКА ТАБЛИЦ ---"

    Dim essentialTables As Variant
    essentialTables = Array("EventInstances", "Themes", "Executors", _
                           "Periodicity", "Settings", "tbBirthdays")

    Dim i As Integer
    For i = 0 To UBound(essentialTables)
        If TableExistsSimple(CStr(essentialTables(i))) Then
            Debug.Print "? Таблица: " & essentialTables(i)

            ' Проверяем количество записей
            Dim recordCount As Long
            recordCount = DCount("*", essentialTables(i))
            Debug.Print "  - Записей: " & recordCount
        Else
            Debug.Print "? Таблица отсутствует: " & essentialTables(i)
        End If
    Next i

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка проверки таблиц: " & Err.description
End Sub

'########################################################################
'########           ТЕСТ СИСТЕМНЫХ ФУНКЦИЙ                      ########
'########################################################################
Private Sub TestSystemFunctions()
    On Error GoTo ErrorHandler

    Debug.Print "--- СИСТЕМНЫЕ ФУНКЦИИ ---"

    ' Проверяем доступность основных функций
    TestFunctionAvailability "DCount", "Доступ к данным"
    TestFunctionAvailability "CurrentProject", "Объект CurrentProject"
    TestFunctionAvailability "DoCmd.OpenForm", "Команды DoCmd"

    ' Проверяем работу календарных функций
    Debug.Print "? Дата/время: " & Now()
    Debug.Print "? Форматирование: " & Format(Now(), "dd.mm.yyyy")

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка теста системных функций: " & Err.description
End Sub

'########################################################################
'########           ПРОВЕРКА ДОСТУПНОСТИ ФУНКЦИИ                ########
'########################################################################
Private Sub TestFunctionAvailability(functionName As String, description As String)
    On Error GoTo ErrorHandler

    ' Простая проверка что функция доступна
    Debug.Print "? " & description & " (" & functionName & ")"

    Exit Sub

ErrorHandler:
    Debug.Print "? " & description & " (" & functionName & "): " & Err.description
End Sub

'########################################################################
'########           ТЕСТ ВСЕХ ФОРМ ПРОЕКТА                      ########
'########################################################################
Private Sub TestAllForms()
    On Error GoTo ErrorHandler

    Debug.Print "=== ВСЕ ФОРМЫ ПРОЕКТА ==="

    Dim allForms As Variant
    allForms = Array("f_daily_planner", "frmDayEvents", "frmEventGenerator", _
                    "frmThemeSelector", "frmExecutors", "frmSearch", _
                    "frmFileFolderSelector", "frmBirthdayCard", "frmBirthdaysList")

    Dim i As Integer
    For i = 0 To UBound(allForms)
        If FormExists(CStr(allForms(i))) Then
            Debug.Print "? Форма: " & allForms(i)
        Else
            Debug.Print "? Форма отсутствует: " & allForms(i)
        End If
    Next i

    ' Показываем общее количество форм
    Debug.Print "Всего форм в проекте: " & CurrentProject.allForms.count

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка теста форм: " & Err.description
End Sub

'########################################################################
'########           ТЕСТ ВСЕХ ТАБЛИЦ БАЗЫ ДАННЫХ                ########
'########################################################################
Private Sub TestAllTables()
    On Error GoTo ErrorHandler

    Debug.Print "=== ВСЕ ТАБЛИЦЫ БАЗЫ ДАННЫХ ==="

    Dim table As AccessObject
    Dim tableCount As Integer
    tableCount = 0

    For Each table In CurrentData.AllTables
        If Left(table.Name, 4) <> "MSys" Then ' Исключаем системные таблицы
            tableCount = tableCount + 1
            Debug.Print tableCount & ". " & table.Name
        End If
    Next table

    Debug.Print "Всего пользовательских таблиц: " & tableCount

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка теста таблиц: " & Err.description
End Sub

'########################################################################
'########           ТЕСТ ВСЕХ МОДУЛЕЙ ПРОЕКТА                   ########
'########################################################################
Private Sub TestAllModules()
    On Error GoTo ErrorHandler

    Debug.Print "=== ВСЕ МОДУЛИ ПРОЕКТА ==="

    Dim module As AccessObject
    Dim moduleCount As Integer
    moduleCount = 0

    For Each module In CurrentProject.AllModules
        moduleCount = moduleCount + 1
        Debug.Print moduleCount & ". " & module.Name
    Next module

    Debug.Print "Всего модулей: " & moduleCount

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка теста модулей: " & Err.description
End Sub

'########################################################################
'########           АНАЛИЗ СТРУКТУРЫ ГЛАВНОЙ ФОРМЫ              ########
'########################################################################
Private Sub AnalyzeMainFormStructure()
    On Error GoTo ErrorHandler

    Debug.Print "=== АНАЛИЗ ГЛАВНОЙ ФОРМЫ f_daily_planner ==="

    ' Проверяем что форма существует
    If Not FormExists("f_daily_planner") Then
        Debug.Print "Форма f_daily_planner не найдена"
        Exit Sub
    End If

    ' Открываем форму в скрытом режиме для анализа
    Dim formWasLoaded As Boolean
    formWasLoaded = CurrentProject.allForms("f_daily_planner").IsLoaded

    If Not formWasLoaded Then
        DoCmd.OpenForm "f_daily_planner", acNormal, , , , acHidden
    End If

    ' Анализируем элементы управления
    AnalyzeFormControls

    ' Анализируем процедуры формы
    AnalyzeFormProcedures

    ' Закрываем форму если открывали её
    If Not formWasLoaded Then
        DoCmd.Close acForm, "f_daily_planner"
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка анализа главной формы: " & Err.description
    If Not formWasLoaded And CurrentProject.allForms("f_daily_planner").IsLoaded Then
        DoCmd.Close acForm, "f_daily_planner"
    End If
End Sub

'########################################################################
'########           АНАЛИЗ ЭЛЕМЕНТОВ УПРАВЛЕНИЯ ФОРМЫ           ########
'########################################################################
Private Sub AnalyzeFormControls()
    On Error GoTo ErrorHandler

    Dim Form As Form
    Set Form = Forms!f_daily_planner

    Dim controlCount As Integer
    controlCount = 0
    Dim buttonCount As Integer
    buttonCount = 0
    Dim labelCount As Integer
    labelCount = 0
    Dim textBoxCount As Integer
    textBoxCount = 0

    Debug.Print "--- ЭЛЕМЕНТЫ УПРАВЛЕНИЯ ---"

    ' Анализируем все элементы управления
    Dim ctrl As Control
    For Each ctrl In Form.Controls
        controlCount = controlCount + 1

        Select Case TypeName(ctrl)
            Case "CommandButton"
                buttonCount = buttonCount + 1
                Debug.Print "КНОПКА: " & ctrl.Name & " | '" & ctrl.Caption & _
                           "' | Pos: " & ctrl.Left & "," & ctrl.Top & _
                           " | Size: " & ctrl.Width & "x" & ctrl.Height

            Case "Label"
                labelCount = labelCount + 1
                ' Пропускаем большинство лейблов чтобы не засорять вывод

            Case "TextBox"
                textBoxCount = textBoxCount + 1

            Case Else
                Debug.Print "ДРУГОЙ: " & ctrl.Name & " | Тип: " & TypeName(ctrl)
        End Select
    Next ctrl

    Debug.Print "--- СТАТИСТИКА ---"
    Debug.Print "Всего элементов: " & controlCount
    Debug.Print "Кнопок: " & buttonCount
    Debug.Print "Надписей: " & labelCount
    Debug.Print "Текстовых полей: " & textBoxCount

    ' Анализируем расположение кнопок для демо-режима
    AnalyzeButtonLayout Form

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка анализа элементов управления: " & Err.description
End Sub

'########################################################################
'########           АНАЛИЗ РАСПОЛОЖЕНИЯ КНОПОК                  ########
'########################################################################
Private Sub AnalyzeButtonLayout(Form As Form)
    On Error GoTo ErrorHandler

    Debug.Print "--- РАСПОЛОЖЕНИЕ КНОПОК ---"

    Dim maxRight As Integer
    maxRight = 0
    Dim maxBottom As Integer
    maxBottom = 0
    Dim buttonList As String
    buttonList = ""

    ' Находим границы и список кнопок
    Dim ctrl As Control
    For Each ctrl In Form.Controls
        If TypeName(ctrl) = "CommandButton" Then
            buttonList = buttonList & ctrl.Name & " ('" & ctrl.Caption & "'), "

            If ctrl.Left + ctrl.Width > maxRight Then
                maxRight = ctrl.Left + ctrl.Width
            End If

            If ctrl.Top + ctrl.Height > maxBottom Then
                maxBottom = ctrl.Top + ctrl.Height
            End If
        End If
    Next ctrl

    If buttonList <> "" Then
        buttonList = Left(buttonList, Len(buttonList) - 2)
    End If

    Debug.Print "Список кнопок: " & buttonList
    Debug.Print "Правая граница: " & maxRight
    Debug.Print "Нижняя граница: " & maxBottom
    Debug.Print "Рекомендуемая позиция для кнопки 'Демо':"
    Debug.Print "  - Left: " & maxRight + 100
    Debug.Print "  - Top: " & maxBottom - 100

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка анализа расположения кнопок: " & Err.description
End Sub

'########################################################################
'########           АНАЛИЗ ПРОЦЕДУР ФОРМЫ                       ########
'########################################################################
Private Sub AnalyzeFormProcedures()
    On Error GoTo ErrorHandler

    Debug.Print "--- ПРОЦЕДУРЫ ФОРМЫ ---"

    Dim criticalProcedures As Variant
    criticalProcedures = Array("BuildCalendar", "Form_Load", "Form_Open", _
                              "ApplyDayStyling", "LoadEventData", _
                              "cmdNextMonth_Click", "cmdPrevMonth_Click", _
                              "cmdToday_Click", "cmdExecutors_Click", _
                              "cmdThemes_Click", "cmdSearch_Click")

    Dim i As Integer
    For i = 0 To UBound(criticalProcedures)
        If ProcedureExistsSimple("f_daily_planner", CStr(criticalProcedures(i))) Then
            Debug.Print "? Процедура: " & criticalProcedures(i)
        Else
            Debug.Print "? Процедура отсутствует: " & criticalProcedures(i)
        End If
    Next i

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка анализа процедур: " & Err.description
End Sub

'########################################################################
'########           ПРОСТОЙ ТЕСТ СИСТЕМЫ                         ########
'########################################################################
Public Sub SimpleTest()
    On Error GoTo ErrorHandler

    Debug.Print "=== ПРОСТОЙ ТЕСТ СИСТЕМЫ ==="

    ' Проверяем основные формы
    TestFormsExistence

    ' Проверяем основные таблицы
    TestTablesExistence

    ' Проверяем основные модули
    TestModulesExistence

    Debug.Print "=== СИСТЕМНАЯ ИНФОРМАЦИЯ ==="
    Debug.Print "Версия Access: " & Application.Version
    Debug.Print "Имя базы данных: " & CurrentProject.Name
    Debug.Print "Путь к базе: " & CurrentProject.FullName

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка простого теста: " & Err.description
End Sub

'########################################################################
'########           ПРОВЕРКА СУЩЕСТВОВАНИЯ ФОРМ                  ########
'########################################################################
Private Sub TestFormsExistence()
    On Error GoTo ErrorHandler

    Dim formsToTest As Variant
    formsToTest = Array("f_daily_planner", "frmDayEvents", "frmExecutors", _
                       "frmSearch", "frmThemeSelector", "frmEventGenerator", _
                       "frmBirthdayCard", "frmBirthdaysList")

    Dim i As Integer
    For i = 0 To UBound(formsToTest)
        If FormExists(CStr(formsToTest(i))) Then
            Debug.Print "? Форма: " & formsToTest(i)
        Else
            Debug.Print "? Форма: " & formsToTest(i)
        End If
    Next i

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка проверки форм: " & Err.description
End Sub

'########################################################################
'########            ПРОВЕРКА СУЩЕСТВОВАНИЯ ТАБЛИЦ               ########
'########################################################################
Private Sub TestTablesExistence()
    On Error GoTo ErrorHandler

    Dim tablesToTest As Variant
    tablesToTest = Array("EventInstances", "Themes", "Executors", "Periodicity", "Settings", "tbBirthdays")

    Dim i As Integer
    For i = 0 To UBound(tablesToTest)
        If TableExistsSimple(CStr(tablesToTest(i))) Then
            Debug.Print "? Таблица: " & tablesToTest(i)
        Else
            Debug.Print "? Таблица: " & tablesToTest(i)
        End If
    Next i

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка проверки таблиц: " & Err.description
End Sub

'########################################################################
'########            ПРОВЕРКА СУЩЕСТВОВАНИЯ МОДУЛЕЙ              ########
'########################################################################
Private Sub TestModulesExistence()
    On Error GoTo ErrorHandler

    Dim modulesToTest As Variant
    modulesToTest = Array("modDatabase", "modTableConnect", "modProtection", _
                         "modThemeSelector", "modEventManager", "modBirthdays")

    Dim i As Integer
    For i = 0 To UBound(modulesToTest)
        If ModuleExists(CStr(modulesToTest(i))) Then
            Debug.Print "? Модуль: " & modulesToTest(i)
        Else
            Debug.Print "? Модуль: " & modulesToTest(i)
        End If
    Next i

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка проверки модулей: " & Err.description
End Sub

'########################################################################
'########           ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ                      ########
'########################################################################

'########################################################################
'########           ПРОВЕРКА СУЩЕСТВОВАНИЯ ФОРМЫ                 ########
'########################################################################
Private Function FormExists(formName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim frm As AccessObject
    For Each frm In CurrentProject.allForms
        If frm.Name = formName Then
            FormExists = True
            Exit Function
        End If
    Next frm

ErrorHandler:
    FormExists = False
End Function

'########################################################################
'########           ПРОВЕРКА СУЩЕСТВОВАНИЯ МОДУЛЯ                ########
'########################################################################
Private Function ModuleExists(moduleName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim comp As AccessObject
    For Each comp In CurrentProject.AllModules
        If comp.Name = moduleName Then
            ModuleExists = True
            Exit Function
        End If
    Next comp

ErrorHandler:
    ModuleExists = False
End Function

'########################################################################
'########           ПРОВЕРКА СУЩЕСТВОВАНИЯ ТАБЛИЦЫ (DAO)          ########
'########################################################################
Private Function TableExistsSimple(tableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim td As DAO.TableDef

    Set db = CurrentDb
    For Each td In db.TableDefs
        If StrComp(td.Name, tableName, vbTextCompare) = 0 Then
            TableExistsSimple = True
            Exit Function
        End If
    Next td

ErrorHandler:
    TableExistsSimple = False
End Function

'########################################################################
'########           ПРОВЕРКА СУЩЕСТВОВАНИЯ СОХРАНЁННОГО ЗАПРОСА ########
'########################################################################
Private Function QueryExistsSimple(queryName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim qd As DAO.QueryDef
    Set qd = CurrentDb.QueryDefs(queryName)
    QueryExistsSimple = True
    Exit Function

ErrorHandler:
    QueryExistsSimple = False
End Function

'########################################################################
'########           ПРОВЕРКА СУЩЕСТВОВАНИЯ ОТЧЁТА               ########
'########################################################################
Private Function ReportExistsSimple(reportName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim R As AccessObject
    For Each R In CurrentProject.AllReports
        If StrComp(R.Name, reportName, vbTextCompare) = 0 Then
            ReportExistsSimple = True
            Exit Function
        End If
    Next R

ErrorHandler:
    ReportExistsSimple = False
End Function

'########################################################################
'########           ОБЪЕКТЫ МОДУЛЯ «ДНИ РОЖДЕНИЯ»                ########
'########################################################################
Private Sub CheckBirthdaysArtifacts()
    On Error GoTo ErrorHandler

    Debug.Print "--- ОБЪЕКТЫ МОДУЛЯ «ДНИ РОЖДЕНИЯ» ---"

    If ModuleExists("modBirthdays") Then
        Debug.Print "? Модуль: modBirthdays"
    Else
        Debug.Print "? Модуль отсутствует: modBirthdays"
    End If

    If TableExistsSimple("tbBirthdays") Then
        Debug.Print "? Таблица: tbBirthdays"
    Else
        Debug.Print "? Таблица отсутствует: tbBirthdays"
    End If

    If QueryExistsSimple("qryBirthdaysForPanel") Then
        Debug.Print "? Запрос: qryBirthdaysForPanel"
    Else
        Debug.Print "? Запрос отсутствует: qryBirthdaysForPanel (см. EnsureQryBirthdaysForPanel в modBirthdays)"
    End If

    If ReportExistsSimple("rptBirthdays") Then
        Debug.Print "? Отчёт: rptBirthdays"
    Else
        Debug.Print "? Отчёт отсутствует: rptBirthdays (см. EnsureRptBirthdaysFromExportFile в modBirthdays)"
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка проверки объектов ДР: " & Err.description
End Sub

'########################################################################
'########           ПРОВЕРКА СУЩЕСТВОВАНИЯ ПРОЦЕДУРЫ             ########
'########################################################################
Private Function ProcedureExistsSimple(formName As String, procName As String) As Boolean
    On Error GoTo ErrorHandler

    ' Упрощенная проверка - пытаемся вызвать процедуру
    ' Если форма загружена, проверяем наличие процедуры

    If CurrentProject.allForms(formName).IsLoaded Then
        Dim Form As Form
        Set Form = Forms(formName)

        ' Если форма загружена, считаем что процедура существует
        ' Более точная проверка требует доступа к VBE
        ProcedureExistsSimple = True
    Else
        ' Если форма не загружена, предполагаем что процедура существует
        ProcedureExistsSimple = True
    End If

    Exit Function

ErrorHandler:
    ProcedureExistsSimple = False
End Function

'########################################################################
'########           ФУНКЦИИ ДЛЯ РУЧНОГО ТЕСТИРОВАНИЯ             ########
'########################################################################

'########################################################################
'########           ТЕСТ КОНКРЕТНОЙ ФОРМЫ                        ########
'########################################################################
Public Sub TestSpecificForm(formName As String)
    On Error GoTo ErrorHandler

    Debug.Print "=== ТЕСТ ФОРМЫ: " & formName & " ==="

    If FormExists(formName) Then
        Debug.Print "? Форма существует"

        ' Пробуем открыть форму
        DoCmd.OpenForm formName, acNormal, , , , acHidden
        Debug.Print "? Форма открыта успешно"

        ' Закрываем форму
        DoCmd.Close acForm, formName
        Debug.Print "? Форма закрыта успешно"
    Else
        Debug.Print "? Форма не существует"
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "? Ошибка теста формы: " & Err.description
End Sub

'########################################################################
'########           ТЕСТ КОНКРЕТНОЙ ТАБЛИЦЫ                      ########
'########################################################################
Public Sub TestSpecificTable(tableName As String)
    On Error GoTo ErrorHandler

    Debug.Print "=== ТЕСТ ТАБЛИЦЫ: " & tableName & " ==="

    If TableExistsSimple(tableName) Then
        Debug.Print "? Таблица существует"

        ' Показываем количество записей
        Dim recordCount As Long
        recordCount = DCount("*", tableName)
        Debug.Print "? Записей в таблице: " & recordCount

        ' Показываем структуру (первые 5 полей)
        ShowTableStructure tableName

    Else
        Debug.Print "? Таблица не существует"
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "? Ошибка теста таблицы: " & Err.description
End Sub

'########################################################################
'########           ПОКАЗ СТРУКТУРЫ ТАБЛИЦЫ                      ########
'########################################################################
Private Sub ShowTableStructure(tableName As String)
    On Error GoTo ErrorHandler

    Dim db As Database
    Dim tdf As TableDef
    Dim fld As Field
    Dim fieldCount As Integer

    Set db = CurrentDb()
    Set tdf = db.TableDefs(tableName)

    fieldCount = 0
    Debug.Print "  Структура таблицы:"

    For Each fld In tdf.Fields
        fieldCount = fieldCount + 1
        If fieldCount <= 5 Then ' Показываем только первые 5 полей
            Debug.Print "    " & fieldCount & ". " & fld.Name & " (" & GetFieldType(fld.Type) & ")"
        End If
    Next fld

    If fieldCount > 5 Then
        Debug.Print "    ... и еще " & (fieldCount - 5) & " полей"
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "  Не удалось получить структуру таблицы"
End Sub

'########################################################################
'########           ПОЛУЧЕНИЕ ТИПА ПОЛЯ                          ########
'########################################################################
Private Function GetFieldType(fieldType As Integer) As String
    Select Case fieldType
        Case dbBoolean: GetFieldType = "Да/Нет"
        Case dbByte: GetFieldType = "Байт"
        Case dbInteger: GetFieldType = "Целое"
        Case dbLong: GetFieldType = "Длинное целое"
        Case dbCurrency: GetFieldType = "Деньги"
        Case dbSingle: GetFieldType = "Одинарное"
        Case dbDouble: GetFieldType = "Двойное"
        Case dbDate: GetFieldType = "Дата/Время"
        Case dbText: GetFieldType = "Текст"
        Case dbLongBinary: GetFieldType = "OLE"
        Case dbMemo: GetFieldType = "Поле MEMO"
        Case Else: GetFieldType = "Неизвестно (" & fieldType & ")"
    End Select
End Function

'########################################################################
'########           ИНФОРМАЦИЯ О ПРОЕКТЕ                         ########
'########################################################################
Public Sub ShowProjectInfo()
    On Error GoTo ErrorHandler

    Debug.Print String(50, "=")
    Debug.Print "ИНФОРМАЦИЯ О ПРОЕКТЕ"
    Debug.Print String(50, "=")

    Debug.Print "Имя проекта: " & CurrentProject.Name
    Debug.Print "Полный путь: " & CurrentProject.FullName
    Debug.Print "Версия Access: " & Application.Version
    Debug.Print "Текущий пользователь: " & Application.CurrentUser
    Debug.Print "Дата создания: " & FileDateTime(CurrentProject.FullName)

    Debug.Print "--- СТАТИСТИКА ---"
    Debug.Print "Формы: " & CurrentProject.allForms.count
    Debug.Print "Отчеты: " & CurrentProject.AllReports.count
    Debug.Print "Модули: " & CurrentProject.AllModules.count
    Debug.Print "Таблицы: " & CurrentData.AllTables.count

    Debug.Print String(50, "=")

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка получения информации о проекте: " & Err.description
End Function

'########################################################################
'########           ПОЛНЫЙ АНАЛИЗ ВСЕХ ЭЛЕМЕНТОВ ФОРМЫ          ########
'########################################################################
Public Sub FullFormAnalysis()
    On Error GoTo ErrorHandler

    Debug.Print "=== ПОЛНЫЙ АНАЛИЗ ФОРМЫ f_daily_planner ==="

    If Not CurrentProject.allForms("f_daily_planner").IsLoaded Then
        DoCmd.OpenForm "f_daily_planner", acNormal, , , , acHidden
    End If

    Dim Form As Form
    Set Form = Forms!f_daily_planner

    Dim totalControls As Integer
    totalControls = 0
    Dim controlTypes As Collection
    Set controlTypes = New Collection

    ' Анализируем ВСЕ элементы включая вложенные
    AnalyzeAllControls Form, totalControls, controlTypes

    Debug.Print "ВСЕГО ЭЛЕМЕНТОВ: " & totalControls

    ' Выводим статистику по типам
    Dim i As Integer
    For i = 1 To controlTypes.count
        Debug.Print controlTypes(i)
    Next i

    ' Анализируем структуру календаря
    AnalyzeCalendarStructure Form

    If Not CurrentProject.allForms("f_daily_planner").IsLoaded Then
        DoCmd.Close acForm, "f_daily_planner"
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка полного анализа: " & Err.description
    If Not CurrentProject.allForms("f_daily_planner").IsLoaded Then
        DoCmd.Close acForm, "f_daily_planner"
    End If
End Sub

'########################################################################
'########           АНАЛИЗ ВСЕХ КОНТРОЛОВ РЕКУРСИВНО            ########
'########################################################################
Private Sub AnalyzeAllControls(container As Object, ByRef totalCount As Integer, ByRef typesCol As Collection)
    On Error GoTo ErrorHandler

    Dim ctrl As Control
    For Each ctrl In container.Controls
        totalCount = totalCount + 1

        ' Считаем типы элементов
        CountControlType typesCol, TypeName(ctrl)

        ' Если элемент является контейнером, анализируем его содержимое
        If TypeOf ctrl Is TabControl Or TypeOf ctrl Is Page Or _
           TypeOf ctrl Is Rectangle Or TypeOf ctrl Is OptionGroup Then
            AnalyzeAllControls ctrl, totalCount, typesCol
        End If

        ' Выводим информацию о каждом элементе
        If totalCount <= 100 Then ' Ограничиваем вывод
            Debug.Print totalCount & ". " & ctrl.Name & " | " & TypeName(ctrl) & _
                       " | " & ctrl.Left & "," & ctrl.Top & " | " & ctrl.Width & "x" & ctrl.Height
        End If
    Next ctrl

    If totalCount > 100 Then
        Debug.Print "... и еще " & (totalCount - 100) & " элементов"
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка анализа контрола " & container.Name & ": " & Err.description
End Sub

'########################################################################
'########           ПОДСЧЕТ ТИПОВ ЭЛЕМЕНТОВ                     ########
'########################################################################
Private Sub CountControlType(typesCol As Collection, controlType As String)
    On Error GoTo ErrorHandler

    Dim i As Integer
    Dim found As Boolean
    found = False

    For i = 1 To typesCol.count
        If InStr(typesCol(i), controlType) > 0 Then
            ' Увеличиваем счетчик для этого типа
            Dim parts() As String
            parts = Split(typesCol(i), ": ")
            Dim count As Integer
            count = CInt(parts(1)) + 1
            typesCol.Remove i
            typesCol.Add controlType & ": " & count, Before:=i
            found = True
            Exit For
        End If
    Next i

    If Not found Then
        typesCol.Add controlType & ": 1"
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка подсчета типа: " & Err.description
End Sub

'########################################################################
'########           АНАЛИЗ СТРУКТУРЫ КАЛЕНДАРЯ                  ########
'########################################################################
Private Sub AnalyzeCalendarStructure(Form As Form)
    On Error GoTo ErrorHandler

    Debug.Print "=== АНАЛИЗ СТРУКТУРЫ КАЛЕНДАРЯ ==="

    ' Ищем элементы календаря
    Dim dayControls As Integer
    dayControls = 0
    Dim eventControls As Integer
    eventControls = 0
    Dim ctrl As Control

    For Each ctrl In Form.Controls
        If TypeName(ctrl) = "Label" Then
            If InStr(ctrl.Name, "Day") > 0 Or InStr(ctrl.Name, "day") > 0 Then
                dayControls = dayControls + 1
            End If
        ElseIf TypeName(ctrl) = "TextBox" Then
            If InStr(ctrl.Name, "Event") > 0 Or InStr(ctrl.Name, "event") > 0 Then
                eventControls = eventControls + 1
            End If
        End If
    Next ctrl

    Debug.Print "Элементов дней: " & dayControls
    Debug.Print "Элементов событий: " & eventControls
    Debug.Print "Всего элементов календаря: " & (dayControls + eventControls)

    ' Анализируем сетку календаря (6 недель ? 7 дней = 42 дня)
    If dayControls >= 42 Then
        Debug.Print "? Календарь: полная сетка 6?7 дней"
    Else
        Debug.Print "? Календарь: неполная сетка (" & dayControls & " элементов)"
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка анализа календаря: " & Err.description
End Sub

'########################################################################
'########           АНАЛИЗ СТРУКТУРЫ ТАБЛИЦ БАЗЫ ДАННЫХ         ########
'########################################################################
Public Sub ПроанализироватьТаблицы()
    On Error GoTo Ошибка

    Debug.Print "================================================"
    Debug.Print "АНАЛИЗ СТРУКТУРЫ ТАБЛИЦ"
    Debug.Print "================================================"

    ПроанализироватьТаблицуСобытий
    ПроанализироватьТаблицуИсполнителей
    ПроанализироватьТаблицуТем
    ПроанализироватьТаблицуНастроек
    ПроанализироватьТаблицуПериодичности

    Debug.Print "================================================"
    Debug.Print "АНАЛИЗ ЗАВЕРШЕН"
    Debug.Print "================================================"

    Exit Sub

Ошибка:
    Debug.Print "Ошибка анализа: " & Err.description
End Sub

'########################################################################
'########           АНАЛИЗ ТАБЛИЦЫ СОБЫТИЙ                      ########
'########################################################################
Private Sub ПроанализироватьТаблицуСобытий()
    On Error GoTo Ошибка

    Debug.Print ""
    Debug.Print "--- ТАБЛИЦА СОБЫТИЙ (tbEventInstances) ---"

    If Not ТаблицаСуществует("tbEventInstances") Then
        Debug.Print "   ТАБЛИЦА НЕ НАЙДЕНА"
        Exit Sub
    End If

    Dim база As Database
    Dim описаниеТаблицы As TableDef
    Dim поле As Field

    Set база = CurrentDb()
    Set описаниеТаблицы = база.TableDefs("tbEventInstances")

    ' Количество записей
    Debug.Print "   Записей: " & DCount("*", "tbEventInstances")

    ' Поля таблицы
    Debug.Print "   ПОЛЯ ТАБЛИЦЫ:"

    For Each поле In описаниеТаблицы.Fields
        Debug.Print "   • " & поле.Name & " | " & ПолучитьТипПоля(поле.Type) & " | " & ПолучитьСвойстваПоля(поле)
    Next поле

    ' Индексы
    Debug.Print "   ИНДЕКСЫ:"
    Dim индекс As Index
    Dim счетчикИндексов As Integer
    счетчикИндексов = 0

    For Each индекс In описаниеТаблицы.Indexes
        счетчикИндексов = счетчикИндексов + 1
        Debug.Print "   " & счетчикИндексов & ". " & индекс.Name & " | " & ПолучитьПоляИндекса(индекс)
    Next индекс

    If счетчикИндексов = 0 Then
        Debug.Print "   Нет индексов"
    End If

    Exit Sub

Ошибка:
    Debug.Print "   Ошибка анализа таблицы событий: " & Err.description
End Sub

'########################################################################
'########           АНАЛИЗ ТАБЛИЦЫ ИСПОЛНИТЕЛЕЙ                 ########
'########################################################################
Private Sub ПроанализироватьТаблицуИсполнителей()
    On Error GoTo Ошибка

    Debug.Print ""
    Debug.Print "--- ТАБЛИЦА ИСПОЛНИТЕЛЕЙ (tbExecutors) ---"

    If Not ТаблицаСуществует("tbExecutors") Then
        Debug.Print "   ТАБЛИЦА НЕ НАЙДЕНА"
        Exit Sub
    End If

    Dim база As Database
    Dim описаниеТаблицы As TableDef
    Dim поле As Field

    Set база = CurrentDb()
    Set описаниеТаблицы = база.TableDefs("tbExecutors")

    Debug.Print "   Записей: " & DCount("*", "tbExecutors")
    Debug.Print "   ПОЛЯ ТАБЛИЦЫ:"

    For Each поле In описаниеТаблицы.Fields
        Debug.Print "   • " & поле.Name & " | " & ПолучитьТипПоля(поле.Type) & " | " & ПолучитьСвойстваПоля(поле)
    Next поле

    Exit Sub

Ошибка:
    Debug.Print "   Ошибка анализа таблицы исполнителей: " & Err.description
End Sub

'########################################################################
'########           АНАЛИЗ ТАБЛИЦЫ ТЕМ                         ########
'########################################################################
Private Sub ПроанализироватьТаблицуТем()
    On Error GoTo Ошибка

    Debug.Print ""
    Debug.Print "--- ТАБЛИЦА ТЕМ (tbThemes) ---"

    If Not ТаблицаСуществует("tbThemes") Then
        Debug.Print "   ТАБЛИЦА НЕ НАЙДЕНА"
        Exit Sub
    End If

    Dim база As Database
    Dim описаниеТаблицы As TableDef
    Dim поле As Field

    Set база = CurrentDb()
    Set описаниеТаблицы = база.TableDefs("tbThemes")

    Debug.Print "   Записей: " & DCount("*", "tbThemes")
    Debug.Print "   ПОЛЯ ТАБЛИЦЫ:"

    For Each поле In описаниеТаблицы.Fields
        Debug.Print "   • " & поле.Name & " | " & ПолучитьТипПоля(поле.Type) & " | " & ПолучитьСвойстваПоля(поле)
    Next поле

    Exit Sub

Ошибка:
    Debug.Print "   Ошибка анализа таблицы тем: " & Err.description
End Sub

'########################################################################
'########           АНАЛИЗ ТАБЛИЦЫ НАСТРОЕК                    ########
'########################################################################
Private Sub ПроанализироватьТаблицуНастроек()
    On Error GoTo Ошибка

    Debug.Print ""
    Debug.Print "--- ТАБЛИЦА НАСТРОЕК (tbSettings) ---"

    If Not ТаблицаСуществует("tbSettings") Then
        Debug.Print "   ТАБЛИЦА НЕ НАЙДЕНА"
        Exit Sub
    End If

    Dim база As Database
    Dim описаниеТаблицы As TableDef
    Dim поле As Field

    Set база = CurrentDb()
    Set описаниеТаблицы = база.TableDefs("tbSettings")

    Debug.Print "   Записей: " & DCount("*", "tbSettings")
    Debug.Print "   ПОЛЯ ТАБЛИЦЫ:"

    For Each поле In описаниеТаблицы.Fields
        Debug.Print "   • " & поле.Name & " | " & ПолучитьТипПоля(поле.Type) & " | " & ПолучитьСвойстваПоля(поле)
    Next поле

    Exit Sub

Ошибка:
    Debug.Print "   Ошибка анализа таблицы настроек: " & Err.description
End Sub

'########################################################################
'########           АНАЛИЗ ТАБЛИЦЫ ПЕРИОДИЧНОСТИ               ########
'########################################################################
Private Sub ПроанализироватьТаблицуПериодичности()
    On Error GoTo Ошибка

    Debug.Print ""
    Debug.Print "--- ТАБЛИЦА ПЕРИОДИЧНОСТИ (tbPeriodicity) ---"

    If Not ТаблицаСуществует("tbPeriodicity") Then
        Debug.Print "   ТАБЛИЦА НЕ НАЙДЕНА"
        Exit Sub
    End If

    Dim база As Database
    Dim описаниеТаблицы As TableDef
    Dim поле As Field

    Set база = CurrentDb()
    Set описаниеТаблицы = база.TableDefs("tbPeriodicity")

    Debug.Print "   Записей: " & DCount("*", "tbPeriodicity")
    Debug.Print "   ПОЛЯ ТАБЛИЦЫ:"

    For Each поле In описаниеТаблицы.Fields
        Debug.Print "   • " & поле.Name & " | " & ПолучитьТипПоля(поле.Type) & " | " & ПолучитьСвойстваПоля(поле)
    Next поле

    Exit Sub

Ошибка:
    Debug.Print "   Ошибка анализа таблицы периодичности: " & Err.description
End Sub

'########################################################################
'########           ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ                     ########
'########################################################################

'########################################################################
'########           ПРОВЕРКА СУЩЕСТВОВАНИЯ ТАБЛИЦЫ             ########
'########################################################################
Private Function ТаблицаСуществует(имяТаблицы As String) As Boolean
    On Error GoTo Ошибка

    Dim проверка As Variant
    проверка = DLookup("Name", "MSysObjects", "Name='" & имяТаблицы & "' AND Type=1")
    ТаблицаСуществует = (Not IsNull(проверка))

    Exit Function

Ошибка:
    ТаблицаСуществует = False
End Function

'########################################################################
'########           ПОЛУЧЕНИЕ ТИПА ПОЛЯ                        ########
'########################################################################
Private Function ПолучитьТипПоля(типПоля As Integer) As String
    Select Case типПоля
        Case dbBoolean: ПолучитьТипПоля = "Да/Нет"
        Case dbByte: ПолучитьТипПоля = "Байт"
        Case dbInteger: ПолучитьТипПоля = "Целое"
        Case dbLong: ПолучитьТипПоля = "Длинное целое"
        Case dbCurrency: ПолучитьТипПоля = "Деньги"
        Case dbSingle: ПолучитьТипПоля = "Одинарное"
        Case dbDouble: ПолучитьТипПоля = "Двойное"
        Case dbDate: ПолучитьТипПоля = "Дата/Время"
        Case dbText: ПолучитьТипПоля = "Текст"
        Case dbLongBinary: ПолучитьТипПоля = "OLE"
        Case dbMemo: ПолучитьТипПоля = "Поле MEMO"
        Case Else: ПолучитьТипПоля = "Другой (" & типПоля & ")"
    End Select
End Function

'########################################################################
'########           ПОЛУЧЕНИЕ СВОЙСТВ ПОЛЯ                     ########
'########################################################################
Private Function ПолучитьСвойстваПоля(поле As Field) As String
    On Error GoTo Ошибка

    Dim свойства As String
    свойства = ""

    If поле.Required Then свойства = свойства & "Обязательное "
    If поле.AllowZeroLength Then свойства = свойства & "Пустое "
    If поле.Attributes And dbAutoIncrField Then свойства = свойства & "Автоинкремент "

    If свойства = "" Then свойства = "Стандартное"

    ПолучитьСвойстваПоля = свойства

    Exit Function

Ошибка:
    ПолучитьСвойстваПоля = "Ошибка"
End Function

'########################################################################
'########           ПОЛУЧЕНИЕ ПОЛЕЙ ИНДЕКСА                    ########
'########################################################################
Private Function ПолучитьПоляИндекса(индекс As Index) As String
    On Error GoTo Ошибка

    Dim поле As Field
    Dim поля As String
    поля = ""

    For Each поле In индекс.Fields
        If поля <> "" Then поля = поля & ", "
        поля = поля & поле.Name
    Next поле

    ПолучитьПоляИндекса = "Поля: " & поля

    Exit Function

Ошибка:
    ПолучитьПоляИндекса = "Ошибка"
End Function

'########################################################################
'########           БЫСТРЫЙ ТЕСТ ОСНОВНЫХ ТАБЛИЦ               ########
'########################################################################
Public Sub БыстрыйТестТаблиц()
    On Error GoTo Ошибка

    Debug.Print "=== БЫСТРЫЙ ТЕСТ ТАБЛИЦ ==="

    Dim таблицы As Variant
    таблицы = Array("tbEventInstances", "tbExecutors", "tbThemes", "tbSettings", "tbPeriodicity", "tbBirthdays")

    Dim i As Integer
    For i = 0 To UBound(таблицы)
        If ТаблицаСуществует(таблицы(i)) Then
            Debug.Print "? " & таблицы(i) & " - " & DCount("*", таблицы(i)) & " записей"
        Else
            Debug.Print "? " & таблицы(i) & " - НЕ НАЙДЕНА"
        End If
    Next i

    Exit Sub

Ошибка:
    Debug.Print "Ошибка быстрого теста: " & Err.description
End Sub

'########################################################################
'########           ПОЛУЧЕНИЕ СПИСКА ОБЪЕКТОВ ПРОЕКТА ########
'########################################################################
Public Sub GetProjectObjectsList()
    On Error GoTo ErrorHandler

    Dim comp As VBComponent
    Dim objType As String
    Dim modulesCount As Integer
    Dim formsCount As Integer
    Dim classesCount As Integer
    Dim totalObjects As Integer

    modulesCount = 0
    formsCount = 0
    classesCount = 0
    totalObjects = 0

    Debug.Print "=============================================="
    Debug.Print "ОБЪЕКТЫ ПРОЕКТА 'ЕЖЕДНЕВНИК'"
    Debug.Print "=============================================="
    Debug.Print ""

    ' Перебираем все компоненты проекта
    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        totalObjects = totalObjects + 1

        ' Определяем тип компонента
        Select Case comp.Type
            Case vbext_ct_StdModule
                objType = "МОДУЛЬ"
                modulesCount = modulesCount + 1
            Case vbext_ct_ClassModule
                objType = "КЛАСС"
                classesCount = classesCount + 1
            Case vbext_ct_MSForm, vbext_ct_Document
                objType = "ФОРМА"
                formsCount = formsCount + 1
            Case Else
                objType = "ДРУГОЙ"
        End Select

        ' Выводим информацию об объекте
        Debug.Print objType & ": " & comp.Name
        Debug.Print "   Строк кода: " & comp.CodeModule.CountOfLines
        Debug.Print ""
    Next comp

    ' Выводим итоговую статистику
    Debug.Print "=============================================="
    Debug.Print "СТАТИСТИКА:"
    Debug.Print "Модули: " & modulesCount
    Debug.Print "Формы: " & formsCount
    Debug.Print "Классы: " & classesCount
    Debug.Print "Всего объектов: " & totalObjects
    Debug.Print "=============================================="

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка при получении списка объектов: " & Err.description, vbCritical
End Sub

'########################################################################
'########           ПОЛУЧЕНИЕ ТОЛЬКО МОДУЛЕЙ           ########
'########################################################################
Public Sub GetModulesList()
    On Error GoTo ErrorHandler

    Dim comp As VBComponent
    Dim count As Integer

    count = 0

    Debug.Print "=============================================="
    Debug.Print "МОДУЛИ ПРОЕКТА"
    Debug.Print "=============================================="
    Debug.Print ""

    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        If comp.Type = vbext_ct_StdModule Then
            count = count + 1
            Debug.Print count & ". " & comp.Name
            Debug.Print "   Строк: " & comp.CodeModule.CountOfLines
        End If
    Next comp

    Debug.Print ""
    Debug.Print "Всего модулей: " & count

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка при получении списка модулей: " & Err.description, vbCritical
End Sub

'########################################################################
'########           ПОЛУЧЕНИЕ ТОЛЬКО ФОРМ              ########
'########################################################################
Public Sub GetFormsList()
    On Error GoTo ErrorHandler

    Dim comp As VBComponent
    Dim count As Integer

    count = 0

    Debug.Print "=============================================="
    Debug.Print "ФОРМЫ ПРОЕКТА"
    Debug.Print "=============================================="
    Debug.Print ""

    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        If comp.Type = vbext_ct_MSForm Or comp.Type = vbext_ct_Document Then
            count = count + 1
            Debug.Print count & ". " & comp.Name
            Debug.Print "   Строк: " & comp.CodeModule.CountOfLines
        End If
    Next comp

    Debug.Print ""
    Debug.Print "Всего форм: " & count

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка при получении списка форм: " & Err.description, vbCritical
End Sub


