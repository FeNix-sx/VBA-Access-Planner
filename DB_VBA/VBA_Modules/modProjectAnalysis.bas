Option Compare Database

'########################################################################
'########           МОДУЛЬ АНАЛИЗА ПРОЕКТА "ЕЖЕДНЕВНИК"         ########
'########################################################################
' Вынесено в модули: modProjectAnalysisCore (проверки объектов),
' modProjectAnalysisDeep (анализ f_daily_planner),
' modProjectAnalysisTables (обход схемы таблиц),
' modProjectAnalysisVbe (списки VBE). Импортировать все в проект БД.
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

    TestBasicComponents

    TestAllForms

    TestAllTables

    TestAllModules

    ProjAn_AnalyzeMainFormStructure

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

    If ProjAn_FormExists("f_daily_planner") Then
        Debug.Print "? Главная форма: f_daily_planner"
        If CurrentProject.allForms("f_daily_planner").IsLoaded Then
            Debug.Print "  - Форма загружена"
        Else
            Debug.Print "  - Форма не загружена"
        End If
    Else
        Debug.Print "? Главная форма отсутствует: f_daily_planner"
    End If

    CheckEssentialModules

    CheckEssentialTables

    ProjAn_CheckBirthdaysArtifacts

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
        If ProjAn_ModuleExists(CStr(essentialModules(i))) Then
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
        If ProjAn_TableExistsSimple(CStr(essentialTables(i))) Then
            Debug.Print "? Таблица: " & essentialTables(i)

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

    TestFunctionAvailability "DCount", "Доступ к данным"
    TestFunctionAvailability "CurrentProject", "Объект CurrentProject"
    TestFunctionAvailability "DoCmd.OpenForm", "Команды DoCmd"

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
        If ProjAn_FormExists(CStr(allForms(i))) Then
            Debug.Print "? Форма: " & allForms(i)
        Else
            Debug.Print "? Форма отсутствует: " & allForms(i)
        End If
    Next i

    Debug.Print "Всего форм в проекте: " & CurrentProject.allForms.count

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка теста форм: " & Err.description
End Sub

'########################################################################
'########            ТЕСТ ВСЕХ ТАБЛИЦ БАЗЫ ДАННЫХ                ########
'########################################################################
Private Sub TestAllTables()
    On Error GoTo ErrorHandler

    Debug.Print "=== ВСЕ ТАБЛИЦЫ БАЗЫ ДАННЫХ ==="

    Dim tblObj As AccessObject
    Dim tableCount As Integer
    tableCount = 0

    For Each tblObj In CurrentData.AllTables
        If Left(tblObj.Name, 4) <> "MSys" Then
            tableCount = tableCount + 1
            Debug.Print tableCount & ". " & tblObj.Name
        End If
    Next tblObj

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

    Dim modObj As AccessObject
    Dim moduleCount As Integer
    moduleCount = 0

    For Each modObj In CurrentProject.AllModules
        moduleCount = moduleCount + 1
        Debug.Print moduleCount & ". " & modObj.Name
    Next modObj

    Debug.Print "Всего модулей: " & moduleCount

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка теста модулей: " & Err.description
End Sub

'########################################################################
'########           ПРОСТОЙ ТЕСТ СИСТЕМЫ                         ########
'########################################################################
Public Sub SimpleTest()
    On Error GoTo ErrorHandler

    Debug.Print "=== ПРОСТОЙ ТЕСТ СИСТЕМЫ ==="

    TestFormsExistence

    TestTablesExistence

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
        If ProjAn_FormExists(CStr(formsToTest(i))) Then
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
        If ProjAn_TableExistsSimple(CStr(tablesToTest(i))) Then
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
        If ProjAn_ModuleExists(CStr(modulesToTest(i))) Then
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
'########           ТЕСТ КОНКРЕТНОЙ ФОРМЫ                        ########
'########################################################################
Public Sub TestSpecificForm(formName As String)
    On Error GoTo ErrorHandler

    Debug.Print "=== ТЕСТ ФОРМЫ: " & formName & " ==="

    If ProjAn_FormExists(formName) Then
        Debug.Print "? Форма существует"

        DoCmd.OpenForm formName, acNormal, , , , acHidden
        Debug.Print "? Форма открыта успешно"

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

    If ProjAn_TableExistsSimple(tableName) Then
        Debug.Print "? Таблица существует"

        Dim recordCount As Long
        recordCount = DCount("*", tableName)
        Debug.Print "? Записей в таблице: " & recordCount

        ProjAn_ShowTableStructure tableName

    Else
        Debug.Print "? Таблица не существует"
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "? Ошибка теста таблицы: " & Err.description
End Sub

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
End Sub

'########################################################################
'########            ПОЛНЫЙ АНАЛИЗ ВСЕХ ЭЛЕМЕНТОВ ФОРМЫ          ########
'########################################################################
Public Sub FullFormAnalysis()
    ProjAn_FullFormAnalysis
End Sub

'########################################################################
'########            АНАЛИЗ СТРУКТУРЫ ТАБЛИЦ (делегирование)     ########
'########################################################################
Public Sub ПроанализироватьТаблицы()
    ProjAnTables_ПроанализироватьТаблицы
End Sub

'########################################################################
'########            БЫСТРЫЙ ТЕСТ ТАБЛИЦ (делегирование)         ########
'########################################################################
Public Sub БыстрыйТестТаблиц()
    ProjAnTables_БыстрыйТестТаблиц
End Sub

'########################################################################
'########             СПИСКИ VBE (делегирование)                 ########
'########################################################################
Public Sub GetProjectObjectsList()
    ProjAnVbe_GetProjectObjectsList
End Sub

Public Sub GetModulesList()
    ProjAnVbe_GetModulesList
End Sub

Public Sub GetFormsList()
    ProjAnVbe_GetFormsList
End Sub


