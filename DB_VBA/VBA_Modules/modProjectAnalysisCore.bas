Option Compare Database
Option Explicit

'########################################################################
'########     ОБЩИЕ ПРОВЕРКИ И УТИЛИТЫ ДЛЯ modProjectAnalysis     ########
'########################################################################
' Назначение: Вынесенные из modProjectAnalysis проверки существования
'             объектов и вспомогательный вывод структуры таблицы.
' Принцип:    Публичные процедуры с префиксом ProjAn_ — только для
'             связанных модулей анализа проекта.
'########################################################################

'########################################################################
'########           ПРОВЕРКА СУЩЕСТВОВАНИЯ ФОРМЫ                 ########
'########################################################################
Public Function ProjAn_FormExists(ByVal formName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim frm As AccessObject
    For Each frm In CurrentProject.allForms
        If frm.Name = formName Then
            ProjAn_FormExists = True
            Exit Function
        End If
    Next frm

ErrorHandler:
    ProjAn_FormExists = False
End Function

'########################################################################
'########           ПРОВЕРКА СУЩЕСТВОВАНИЯ МОДУЛЯ                ########
'########################################################################
Public Function ProjAn_ModuleExists(ByVal moduleName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim comp As AccessObject
    For Each comp In CurrentProject.AllModules
        If comp.Name = moduleName Then
            ProjAn_ModuleExists = True
            Exit Function
        End If
    Next comp

ErrorHandler:
    ProjAn_ModuleExists = False
End Function

'########################################################################
'########           ПРОВЕРКА СУЩЕСТВОВАНИЯ ТАБЛИЦЫ (DAO)          ########
'########################################################################
Public Function ProjAn_TableExistsSimple(ByVal tableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim td As DAO.TableDef

    Set db = CurrentDb
    For Each td In db.TableDefs
        If StrComp(td.Name, tableName, vbTextCompare) = 0 Then
            ProjAn_TableExistsSimple = True
            Exit Function
        End If
    Next td

ErrorHandler:
    ProjAn_TableExistsSimple = False
End Function

'########################################################################
'########           ПРОВЕРКА СУЩЕСТВОВАНИЯ СОХРАНЁННОГО ЗАПРОСА ########
'########################################################################
Public Function ProjAn_QueryExistsSimple(ByVal queryName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim qd As DAO.QueryDef
    Set qd = CurrentDb.QueryDefs(queryName)
    ProjAn_QueryExistsSimple = True
    Exit Function

ErrorHandler:
    ProjAn_QueryExistsSimple = False
End Function

'########################################################################
'########           ПРОВЕРКА СУЩЕСТВОВАНИЯ ОТЧЁТА               ########
'########################################################################
Public Function ProjAn_ReportExistsSimple(ByVal reportName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim r As AccessObject
    For Each r In CurrentProject.AllReports
        If StrComp(r.Name, reportName, vbTextCompare) = 0 Then
            ProjAn_ReportExistsSimple = True
            Exit Function
        End If
    Next r

ErrorHandler:
    ProjAn_ReportExistsSimple = False
End Function

'########################################################################
'########           ПРОВЕРКА СУЩЕСТВОВАНИЯ ПРОЦЕДУРЫ             ########
'########################################################################
Public Function ProjAn_ProcedureExistsSimple(ByVal formName As String, ByVal procName As String) As Boolean
    On Error GoTo ErrorHandler

    If CurrentProject.allForms(formName).IsLoaded Then
        Dim frm As Form
        Set frm = Forms(formName)
        ProjAn_ProcedureExistsSimple = True
    Else
        ProjAn_ProcedureExistsSimple = True
    End If

    Exit Function

ErrorHandler:
    ProjAn_ProcedureExistsSimple = False
End Function

'########################################################################
'########           ОБЪЕКТЫ МОДУЛЯ «ДНИ РОЖДЕНИЯ»                ########
'########################################################################
Public Sub ProjAn_CheckBirthdaysArtifacts()
    On Error GoTo ErrorHandler

    Debug.Print "--- ОБЪЕКТЫ МОДУЛЯ «ДНИ РОЖДЕНИЯ» ---"

    If ProjAn_ModuleExists("modBirthdays") Then
        Debug.Print "? Модуль: modBirthdays"
    Else
        Debug.Print "? Модуль отсутствует: modBirthdays"
    End If

    If ProjAn_TableExistsSimple("tbBirthdays") Then
        Debug.Print "? Таблица: tbBirthdays"
    Else
        Debug.Print "? Таблица отсутствует: tbBirthdays"
    End If

    If ProjAn_QueryExistsSimple("qryBirthdaysForPanel") Then
        Debug.Print "? Запрос: qryBirthdaysForPanel"
    Else
        Debug.Print "? Запрос отсутствует: qryBirthdaysForPanel (см. EnsureQryBirthdaysForPanel в modBirthdays)"
    End If

    If ProjAn_ReportExistsSimple("rptBirthdays") Then
        Debug.Print "? Отчёт: rptBirthdays"
    Else
        Debug.Print "? Отчёт отсутствует: rptBirthdays (см. EnsureRptBirthdaysFromExportFile в modBirthdays)"
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка проверки объектов ДР: " & Err.description
End Sub

'########################################################################
'########           ПОКАЗ СТРУКТУРЫ ТАБЛИЦЫ                      ########
'########################################################################
Public Sub ProjAn_ShowTableStructure(ByVal tableName As String)
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
        If fieldCount <= 5 Then
            Debug.Print "    " & fieldCount & ". " & fld.Name & " (" & ProjAn_GetFieldType(fld.Type) & ")"
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
'########           ПОЛУЧЕНИЕ ТИПА ПОЛЯ (DAO)                    ########
'########################################################################
Public Function ProjAn_GetFieldType(ByVal fieldType As Integer) As String
    Select Case fieldType
        Case dbBoolean: ProjAn_GetFieldType = "Да/Нет"
        Case dbByte: ProjAn_GetFieldType = "Байт"
        Case dbInteger: ProjAn_GetFieldType = "Целое"
        Case dbLong: ProjAn_GetFieldType = "Длинное целое"
        Case dbCurrency: ProjAn_GetFieldType = "Деньги"
        Case dbSingle: ProjAn_GetFieldType = "Одинарное"
        Case dbDouble: ProjAn_GetFieldType = "Двойное"
        Case dbDate: ProjAn_GetFieldType = "Дата/Время"
        Case dbText: ProjAn_GetFieldType = "Текст"
        Case dbLongBinary: ProjAn_GetFieldType = "OLE"
        Case dbMemo: ProjAn_GetFieldType = "Поле MEMO"
        Case Else: ProjAn_GetFieldType = "Неизвестно (" & fieldType & ")"
    End Select
End Function

