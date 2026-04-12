Option Compare Database

'################################################################
'########              ОСНОВНЫЕ ПРОЦЕДУРЫ                ########
'################################################################

'################################################################
'########      СОЗДАНИЕ ТАБЛИЦЫ НАСТРОЕК ПОДКЛЮЧЕНИЯ     ########
'################################################################
Public Sub CreateTableConnections()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim td As DAO.TableDef
    Dim fld As DAO.Field

    Set db = CurrentDb

    ' СОЗДАЕМ ТАБЛИЦУ ЕСЛИ ЕЕ НЕТ
    On Error Resume Next
    db.TableDefs.Delete "tbTableConnections"
    On Error GoTo ErrorHandler

    Set td = db.CreateTableDef("tbTableConnections")

    ' ПОЛЕ "TableName" - ИМЯ ТАБЛИЦЫ
    Set fld = td.CreateField("TableName", dbText, 50)
    td.Fields.Append fld

    ' ПОЛЕ "TablePath" - ПУТЬ К ТАБЛИЦЕ
    Set fld = td.CreateField("TablePath", dbText, 255)
    td.Fields.Append fld

    ' ПОЛЕ "Description" - ОПИСАНИЕ ТАБЛИЦЫ
    Set fld = td.CreateField("Description", dbText, 100)
    td.Fields.Append fld

    ' ДОБАВЛЯЕМ ТАБЛИЦУ В БАЗУ
    db.TableDefs.Append td

    ' ЗАПОЛНЯЕМ ТАБЛИЦУ ДАННЫМИ
    Call FillTableConnections

    MsgBox "Таблица подключений создана и заполнена!", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка создания таблицы подключений: " & Err.description, vbCritical
End Sub

'################################################################
'########          ЗАПОЛНЕНИЕ ТАБЛИЦЫ ПОДКЛЮЧЕНИЙ        ########
'################################################################
Private Sub FillTableConnections()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Set db = CurrentDb

    ' ПУТЬ ПО УМОЛЧАНИЮ К BACKEND
    Dim defaultPath As String
    defaultPath = CurrentProject.Path & "\BE\Planner_BE.accdb"

    ' ОЧИЩАЕМ ТАБЛИЦУ
    db.Execute "DELETE FROM tbTableConnections"

    ' ДОБАВЛЯЕМ ВСЕ ТАБЛИЦЫ С ПУТЕМ ПО УМОЛЧАНИЮ
    db.Execute "INSERT INTO tbTableConnections (TableName, TablePath, Description) " & _
               "VALUES ('tbEventInstances', '" & defaultPath & "', 'Основные события календаря')"

    db.Execute "INSERT INTO tbTableConnections (TableName, TablePath, Description) " & _
               "VALUES ('tbExecutors', '" & defaultPath & "', 'Список исполнителей')"

    db.Execute "INSERT INTO tbTableConnections (TableName, TablePath, Description) " & _
               "VALUES ('tbThemes', '" & defaultPath & "', 'Темы оформления')"

    db.Execute "INSERT INTO tbTableConnections (TableName, TablePath, Description) " & _
               "VALUES ('tbTempEvents', '" & defaultPath & "', 'Временные события для генератора')"

    db.Execute "INSERT INTO tbTableConnections (TableName, TablePath, Description) " & _
               "VALUES ('tbPeriodicity', '" & defaultPath & "', 'Типы периодичности событий')"

    db.Execute "INSERT INTO tbTableConnections (TableName, TablePath, Description) " & _
               "VALUES ('tbRules', '" & defaultPath & "', 'Правила генерации событий')"

    db.Execute "INSERT INTO tbTableConnections (TableName, TablePath, Description) " & _
               "VALUES ('tbBirthdays', '" & defaultPath & "', 'Справочник дней рождения')"

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка заполнения таблицы подключений: " & Err.description, vbCritical
End Sub

'################################################################
'########             ОТКЛЮЧЕНИЕ ВСЕХ ТАБЛИЦ             ########
'################################################################
Public Sub DisconnectAllTables()
    On Error GoTo ErrorHandler

    Dim td As DAO.TableDef
    Dim tablesFound As Boolean

    Do
        tablesFound = False
        For Each td In CurrentDb.TableDefs
            If td.Connect <> "" Then
                tablesFound = True
                CurrentDb.TableDefs.Delete td.Name
                Exit For
            End If
        Next td
    Loop While tablesFound

    Exit Sub

ErrorHandler:
    ' ИГНОРИРУЕМ ОШИБКИ ОТКЛЮЧЕНИЯ ТАБЛИЦ
    Resume Next
End Sub

'################################################################
'########            ПОДКЛЮЧЕНИЕ ВСЕХ ТАБЛИЦ             ########
'################################################################
Public Sub ConnectAllTables()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim backendPath As String

    Set db = CurrentDb

    ' 1. ОТКЛЮЧАЕМ СТАРЫЕ ТАБЛИЦЫ
    Call DisconnectAllTables

    ' 2. ПРОВЕРЯЕМ/СОЗДАЕМ ТАБЛИЦУ ПОДКЛЮЧЕНИЙ
    If Not TableExists("tbTableConnections") Then
        Call CreateTableConnections
    End If

    ' 3. ПОЛУЧАЕМ ПУТЬ ИЗ ТАБЛИЦЫ
    Set rs = db.OpenRecordset("SELECT TOP 1 TablePath FROM tbTableConnections")
    backendPath = rs!TablePath
    rs.Close

    ' 4. ПРОВЕРЯЕМ СУЩЕСТВУЕТ ЛИ ФАЙЛ
    If Dir(backendPath) = "" Then
        ' ФАЙЛА НЕТ - ЗАПРАШИВАЕМ НОВЫЙ ПУТЬ
        backendPath = BrowseForBackendFile()
        If backendPath = "" Then Exit Sub

        ' СОХРАНЯЕМ НОВЫЙ ПУТЬ
        db.Execute "UPDATE tbTableConnections SET TablePath = '" & Replace(backendPath, "'", "''") & "'"
    End If

    ' 5. ПОДКЛЮЧАЕМ ВСЕ ТАБЛИЦЫ
    Set rs = db.OpenRecordset("SELECT TableName FROM tbTableConnections")
    Do While Not rs.EOF
        Call LinkTable(rs!tableName, backendPath)
        rs.MoveNext
    Loop

    rs.Close

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка подключения таблиц: " & Err.description, vbCritical
    If Not rs Is Nothing Then rs.Close
End Sub

'################################################################
'########           ПРОВЕРКА СУЩЕСТВОВАНИЯ ТАБЛИЦЫ       ########
'################################################################
Public Function TableExists(tableName As String) As Boolean
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim td As DAO.TableDef

    Set db = CurrentDb

    TableExists = False

    ' ПРОВЕРЯЕМ ВСЕ ТАБЛИЦЫ В БАЗЕ
    For Each td In db.TableDefs
        If td.Name = tableName Then
            TableExists = True
            Exit For
        End If
    Next td

    Exit Function

ErrorHandler:
    TableExists = False
End Function

'################################################################
'########        ПОДКЛЮЧЕНИЕ ОДНОЙ ТАБЛИЦЫ         ########
'################################################################
Public Sub LinkTable(tableName As String, backendPath As String)
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim td As DAO.TableDef

    Set db = CurrentDb

    ' СОЗДАЕМ ССЫЛКУ НА ТАБЛИЦУ В BACKEND
    Set td = db.CreateTableDef(tableName)
    td.Connect = ";DATABASE=" & backendPath
    td.SourceTableName = tableName

    ' ДОБАВЛЯЕМ ССЫЛКУ В БАЗУф
    db.TableDefs.Append td

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка подключения таблицы " & tableName & ": " & Err.description, vbCritical
End Sub

'################################################################
'########   МИГРАЦИЯ FE: СТРОКА tbBirthdays В ПОДКЛЮЧЕНИЯХ ########
'################################################################
' Однократно для уже развёрнутого frontend: не вызывать CreateTableConnections
' целиком. Путь BE берётся из любой существующей строки tbTableConnections.
Public Sub MigrateAddBirthdaysConnectionIfMissing()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim backendPath As String

    Set db = CurrentDb

    Set rs = db.OpenRecordset("SELECT COUNT(*) AS C FROM tbTableConnections WHERE TableName = 'tbBirthdays'")
    If rs!C > 0 Then
        rs.Close
        MsgBox "Запись tbBirthdays в tbTableConnections уже есть.", vbInformation
        Exit Sub
    End If
    rs.Close

    Set rs = db.OpenRecordset("SELECT TOP 1 TablePath FROM tbTableConnections")
    If rs.EOF Then
        rs.Close
        MsgBox "tbTableConnections пуста — сначала настройте подключения к BE.", vbExclamation
        Exit Sub
    End If
    backendPath = rs!TablePath
    rs.Close

    db.Execute "INSERT INTO tbTableConnections (TableName, TablePath, Description) VALUES (" & _
               "'tbBirthdays', '" & Replace(backendPath, "'", "''") & "', 'Справочник дней рождения')"

    MsgBox "Добавлена строка tbBirthdays. Выполните ConnectAllTables или перезапустите приложение.", vbInformation
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка миграции tbTableConnections: " & Err.description, vbCritical
End Sub

'################################################################
'########          ВЫБОР ФАЙЛА BACKEND С СООБЩЕНИЕМ      ########
'################################################################
Private Function BrowseForBackendFile() As String
    On Error GoTo ErrorHandler

    Dim fileDialog As Object
    Dim selectedFile As Variant
    Dim defaultPath As String

    ' ПОЛУЧАЕМ ПУТЬ ПО УМОЛЧАНИЮ ДЛЯ СООБЩЕНИЯ
    defaultPath = CurrentProject.Path & "\BE\Planner_BE.accdb"

    ' ПОКАЗЫВАЕМ СООБЩЕНИЕ О ТОМ ЧТО ФАЙЛ НЕ НАЙДЕН
    MsgBox "Файл базы данных не найден по пути:" & vbCrLf & _
           defaultPath & vbCrLf & vbCrLf & _
           "Пожалуйста, укажите расположение файла Planner_BE.accdb", _
           vbExclamation, "Файл не найден"

    Set fileDialog = Application.fileDialog(1) ' msoFileDialogOpen

    With fileDialog
        .title = "Выберите файл Backend (Planner_BE.accdb)"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Базы данных Access", "*.accdb"

        If .Show Then
            BrowseForBackendFile = .SelectedItems(1)
        Else
            BrowseForBackendFile = ""
        End If
    End With

    Exit Function

ErrorHandler:
    MsgBox "Ошибка выбора файла: " & Err.description, vbCritical
    BrowseForBackendFile = ""
End Function

'################################################################
'########           АВТОПОДКЛЮЧЕНИЕ ПРИ ЗАПУСКЕ          ########
'################################################################
Public Sub AutoConnectOnStartup()
    On Error GoTo ErrorHandler

    ' ПРОВЕРЯЕМ И ПОДКЛЮЧАЕМ ТАБЛИЦЫ ПРИ ЗАПУСКЕ
    Call ConnectAllTables

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка автоматического подключения: " & Err.description, vbCritical
End Sub


