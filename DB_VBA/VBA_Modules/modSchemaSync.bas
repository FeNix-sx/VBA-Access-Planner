Option Compare Database
Option Explicit

'################################################################
'########     МОДУЛЬ СИНХРОНИЗАЦИИ СТРУКТУРЫ БД (v2.0)   ########
'################################################################

'================================================================
' БАЗОВЫЕ ФУНКЦИИ ПРОВЕРКИ
'================================================================

' Проверка существования таблицы в указанной БД
Public Function DbHasTable(db As DAO.Database, tableName As String) As Boolean
    Dim tdf As DAO.TableDef
    On Error Resume Next
    Set tdf = db.TableDefs(tableName)
    DbHasTable = (Err.Number = 0)
    On Error GoTo 0
End Function

' Проверка существования поля в таблице
Public Function TableHasField(db As DAO.Database, tableName As String, fieldName As String) As Boolean
    Dim fld As DAO.Field
    On Error Resume Next
    Set fld = db.TableDefs(tableName).Fields(fieldName)
    TableHasField = (Err.Number = 0)
    On Error GoTo 0
End Function

' DEPRECATED: процедура сохранена для совместимости.
' Актуальная точка добавления структуры — EnsureField.
Public Sub EnsureTable(db As DAO.Database, tableName As String)
    Dim tdf As DAO.TableDef
    If Not DbHasTable(db, tableName) Then
        Set tdf = db.CreateTableDef(tableName)
        ' Для создания таблицы в DAO нужно хотя бы одно поле,
        ' поэтому мы просто создаем TableDef, а поля добавим в EnsureField.
        ' Однако DAO не позволяет добавить пустую таблицу.
        ' Поэтому EnsureTable просто проверяет. Реальное добавление произойдет в EnsureField.
    End If
End Sub

' Добавление поля, если оно не существует (с созданием таблицы при необходимости)
Public Sub EnsureField(db As DAO.Database, tableName As String, fieldName As String, fieldType As Integer, fieldSize As Integer, Optional isAutoIncr As Boolean = False)
    Dim tdf As DAO.TableDef
    Dim fld As DAO.Field
    Dim isNewTable As Boolean
    
    isNewTable = False
    
    If Not DbHasTable(db, tableName) Then
        Set tdf = db.CreateTableDef(tableName)
        isNewTable = True
    Else
        Set tdf = db.TableDefs(tableName)
    End If
    
    If isNewTable Or Not TableHasField(db, tableName, fieldName) Then
        Set fld = tdf.CreateField(fieldName, fieldType)
        If fieldSize > 0 And fieldType = dbText Then
            fld.Size = fieldSize
        End If
        If isAutoIncr Then
            fld.Attributes = dbAutoIncrField
        End If
        
        tdf.Fields.Append fld
        
        If isNewTable Then
            db.TableDefs.Append tdf
            db.TableDefs.Refresh
        End If
    End If
End Sub

'================================================================
' ОПИСАНИЕ СТРУКТУРЫ ВСЕХ ТАБЛИЦ
'================================================================

Public Sub DefineSchema(beDb As DAO.Database)
    ' tbBirthdays
    EnsureField beDb, "tbBirthdays", "ID", dbLong, 0, True
    EnsureField beDb, "tbBirthdays", "LastName", dbText, 100
    EnsureField beDb, "tbBirthdays", "FirstName", dbText, 100
    EnsureField beDb, "tbBirthdays", "MiddleName", dbText, 100
    EnsureField beDb, "tbBirthdays", "BirthDate", dbDate, 0
    EnsureField beDb, "tbBirthdays", "Notes", dbText, 255
    
    ' tbEventInstances
    EnsureField beDb, "tbEventInstances", "InstanceID", dbLong, 0, True
    EnsureField beDb, "tbEventInstances", "EventDate", dbDate, 0
    EnsureField beDb, "tbEventInstances", "EventNote", dbText, 255
    EnsureField beDb, "tbEventInstances", "Basis", dbText, 255
    EnsureField beDb, "tbEventInstances", "BasisAttachment", dbText, 255
    EnsureField beDb, "tbEventInstances", "CompletionDate", dbDate, 0
    EnsureField beDb, "tbEventInstances", "CompletionMark", dbText, 255
    EnsureField beDb, "tbEventInstances", "LastModified", dbDate, 0
    EnsureField beDb, "tbEventInstances", "AttachmentPath", dbText, 255
    EnsureField beDb, "tbEventInstances", "ExecutorID", dbLong, 0
    EnsureField beDb, "tbEventInstances", "Notes", dbText, 255
    
    ' tbExecutors
    EnsureField beDb, "tbExecutors", "ID", dbLong, 0, True
    EnsureField beDb, "tbExecutors", "LastName", dbText, 100
    EnsureField beDb, "tbExecutors", "FirstName", dbText, 100
    EnsureField beDb, "tbExecutors", "MiddleName", dbText, 100
    EnsureField beDb, "tbExecutors", "Position", dbText, 255
    EnsureField beDb, "tbExecutors", "SortOrder", dbLong, 0
    EnsureField beDb, "tbExecutors", "Notes", dbText, 255
    
    ' tbPeriodicity
    EnsureField beDb, "tbPeriodicity", "PeriodicityID", dbLong, 0, True
    EnsureField beDb, "tbPeriodicity", "PeriodicityName", dbText, 50
    EnsureField beDb, "tbPeriodicity", "Description", dbText, 255
    
    ' tbRules
    EnsureField beDb, "tbRules", "RuleID", dbLong, 0, True
    EnsureField beDb, "tbRules", "RuleName", dbText, 100
    EnsureField beDb, "tbRules", "RuleDescription", dbText, 255
    EnsureField beDb, "tbRules", "CalculationMethod", dbText, 50
    
    ' tbTempEvents
    EnsureField beDb, "tbTempEvents", "TempID", dbLong, 0
    EnsureField beDb, "tbTempEvents", "EventDate", dbDate, 0
    EnsureField beDb, "tbTempEvents", "EventNote", dbText, 255
    EnsureField beDb, "tbTempEvents", "OriginalDay", dbInteger, 0
    EnsureField beDb, "tbTempEvents", "ExecutorID", dbLong, 0
    EnsureField beDb, "tbTempEvents", "AttachmentPath", dbText, 255
    EnsureField beDb, "tbTempEvents", "BasisAttachment", dbText, 255
End Sub

'================================================================
' ГЛАВНАЯ ПРОЦЕДУРА СИНХРОНИЗАЦИИ
'================================================================

Public Sub SyncDatabaseSchema(backendPath As String)
    Dim beDb As DAO.Database
    
    On Error GoTo ErrorHandler
    
    ' Открываем Backend напрямую для изменения структуры
    Set beDb = DBEngine.OpenDatabase(backendPath)
    
    ' 1. Применяем схему (создаем таблицы и поля)
    Call DefineSchema(beDb)
    
    beDb.Close
    Set beDb = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка при синхронизации структуры БД: " & Err.description, vbCritical
    If Not beDb Is Nothing Then beDb.Close
End Sub

