Option Compare Database
Option Explicit

'########################################################################
'########     АНАЛИЗ СТРУКТУРЫ ТАБЛИЦ (DAO, ВЫВОД В Immediate)   ########
'########################################################################
' Назначение: Русскоязычные процедуры обхода схемы ключевых таблиц.
' Принцип:    Локальные проверки через MSysObjects и TableDefs.
'########################################################################

'########################################################################
'########           АНАЛИЗ СТРУКТУРЫ ТАБЛИЦ БАЗЫ ДАННЫХ          ########
'########################################################################
Public Sub ProjAnTables_ПроанализироватьТаблицы()
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
'########            АНАЛИЗ ТАБЛИЦЫ СОБЫТИЙ                      ########
'########################################################################
Private Sub ПроанализироватьТаблицуСобытий()
    On Error GoTo Ошибка

    Debug.Print ""
    Debug.Print "--- ТАБЛИЦА СОБЫТИЙ (tbEventInstances) ---"

    If Not ProjAn_TableExistsSimple("tbEventInstances") Then
        Debug.Print "   ТАБЛИЦА НЕ НАЙДЕНА"
        Exit Sub
    End If

    Dim база As Database
    Dim описаниеТаблицы As TableDef
    Dim поле As Field

    Set база = CurrentDb()
    Set описаниеТаблицы = база.TableDefs("tbEventInstances")

    Debug.Print "   Записей: " & DCount("*", "tbEventInstances")

    Debug.Print "   ПОЛЯ ТАБЛИЦЫ:"

    For Each поле In описаниеТаблицы.Fields
        Debug.Print "   • " & поле.Name & " | " & ProjAn_GetFieldType(поле.Type) & " | " & ПолучитьСвойстваПоля(поле)
    Next поле

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
'########            АНАЛИЗ ТАБЛИЦЫ ИСПОЛНИТЕЛЕЙ                 ########
'########################################################################
Private Sub ПроанализироватьТаблицуИсполнителей()
    On Error GoTo Ошибка

    Debug.Print ""
    Debug.Print "--- ТАБЛИЦА ИСПОЛНИТЕЛЕЙ (tbExecutors) ---"

    If Not ProjAn_TableExistsSimple("tbExecutors") Then
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
        Debug.Print "   • " & поле.Name & " | " & ProjAn_GetFieldType(поле.Type) & " | " & ПолучитьСвойстваПоля(поле)
    Next поле

    Exit Sub

Ошибка:
    Debug.Print "   Ошибка анализа таблицы исполнителей: " & Err.description
End Sub

'########################################################################
'########             АНАЛИЗ ТАБЛИЦЫ ТЕМ                         ########
'########################################################################
Private Sub ПроанализироватьТаблицуТем()
    On Error GoTo Ошибка

    Debug.Print ""
    Debug.Print "--- ТАБЛИЦА ТЕМ (tbThemes) ---"

    If Not ProjAn_TableExistsSimple("tbThemes") Then
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
        Debug.Print "   • " & поле.Name & " | " & ProjAn_GetFieldType(поле.Type) & " | " & ПолучитьСвойстваПоля(поле)
    Next поле

    Exit Sub

Ошибка:
    Debug.Print "   Ошибка анализа таблицы тем: " & Err.description
End Sub

'########################################################################
'########             АНАЛИЗ ТАБЛИЦЫ НАСТРОЕК                    ########
'########################################################################
Private Sub ПроанализироватьТаблицуНастроек()
    On Error GoTo Ошибка

    Debug.Print ""
    Debug.Print "--- ТАБЛИЦА НАСТРОЕК (tbSettings) ---"

    If Not ProjAn_TableExistsSimple("tbSettings") Then
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
        Debug.Print "   • " & поле.Name & " | " & ProjAn_GetFieldType(поле.Type) & " | " & ПолучитьСвойстваПоля(поле)
    Next поле

    Exit Sub

Ошибка:
    Debug.Print "   Ошибка анализа таблицы настроек: " & Err.description
End Sub

'########################################################################
'########             АНАЛИЗ ТАБЛИЦЫ ПЕРИОДИЧНОСТИ               ########
'########################################################################
Private Sub ПроанализироватьТаблицуПериодичности()
    On Error GoTo Ошибка

    Debug.Print ""
    Debug.Print "--- ТАБЛИЦА ПЕРИОДИЧНОСТИ (tbPeriodicity) ---"

    If Not ProjAn_TableExistsSimple("tbPeriodicity") Then
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
        Debug.Print "   • " & поле.Name & " | " & ProjAn_GetFieldType(поле.Type) & " | " & ПолучитьСвойстваПоля(поле)
    Next поле

    Exit Sub

Ошибка:
    Debug.Print "   Ошибка анализа таблицы периодичности: " & Err.description
End Sub

'########################################################################
'########             ПОЛУЧЕНИЕ СВОЙСТВ ПОЛЯ                     ########
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
'########             ПОЛУЧЕНИЕ ПОЛЕЙ ИНДЕКСА                    ########
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
'########             БЫСТРЫЙ ТЕСТ ОСНОВНЫХ ТАБЛИЦ               ########
'########################################################################
Public Sub ProjAnTables_БыстрыйТестТаблиц()
    On Error GoTo Ошибка

    Debug.Print "=== БЫСТРЫЙ ТЕСТ ТАБЛИЦ ==="

    Dim таблицы As Variant
    таблицы = Array("tbEventInstances", "tbExecutors", "tbThemes", "tbSettings", "tbPeriodicity", "tbBirthdays")

    Dim i As Integer
    For i = 0 To UBound(таблицы)
        If ProjAn_TableExistsSimple(таблицы(i)) Then
            Debug.Print "? " & таблицы(i) & " - " & DCount("*", таблицы(i)) & " записей"
        Else
            Debug.Print "? " & таблицы(i) & " - НЕ НАЙДЕНА"
        End If
    Next i

    Exit Sub

Ошибка:
    Debug.Print "Ошибка быстрого теста: " & Err.description
End Sub

