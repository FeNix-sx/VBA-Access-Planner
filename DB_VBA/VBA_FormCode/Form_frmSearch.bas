Option Compare Database

'################################################################
'########      БАЗОВЫЙ SQL СПИСКА РЕЗУЛЬТАТОВ ПОИСКА     ########
'################################################################
Private Function BuildSearchFormBaseSql() As String
' Назначение: Единый SELECT+JOIN для источника записей формы поиска.
' Возврат:    SQL без WHERE и ORDER BY (строка заканчивается пробелом).
'================================================================
    BuildSearchFormBaseSql = "SELECT ei.EventDate, ei.EventNote, " & _
        "e.LastName & ' ' & Left(e.FirstName,1) & '.' & Left(e.MiddleName,1) & '.' AS ExecutorName, " & _
        "ei.CompletionMark, ei.AttachmentPath " & _
        "FROM tbEventInstances ei " & _
        "LEFT JOIN tbExecutors e ON ei.ExecutorID = e.ID "
End Function

'################################################################
'########            ЗАГРУЗКА ФОРМЫ ПОИСКА               ########
'################################################################
Private Sub Form_Load()
    ' ЗАГРУЖАЕМ ИСПОЛНИТЕЛЕЙ В КОМБОБОКС
    LoadExecutorsCombo

    ' ЗАГРУЖАЕМ СТАТУСЫ ВЫПОЛНЕНИЯ
    LoadStatusCombo

    ' ЗАГРУЖАЕМ ФИЛЬТР ПО ВЛОЖЕНИЯМ
    LoadAttachmentCombo

    ' УСТАНАВЛИВАЕМ НАЧАЛЬНЫЙ SQL БЕЗ ID
    Me.RecordSource = BuildSearchFormBaseSql() & _
                     "ORDER BY ei.EventDate DESC"

    Me.Requery

    On Error GoTo ErrorHandler

    ' УСТАНАВЛИВАЕМ РАЗМЕР И ПОЛОЖЕНИЕ ФОРМЫ
    ' DoCmd.MoveSize Left, Top, Width, Height
    ' Left   - отступ слева в твипах (1440 твипов = 1 дюйм = 2.54 см)
    ' Top    - отступ сверху в твипах
    ' Width  - ширина формы в твипах
    ' Height - высота формы в твипах

    DoCmd.MoveSize 5000, 1500, 14800, 14000
    Exit Sub

ErrorHandler:


End Sub

'################################################################
'########         ЗАГРУЗКА КОМБОБОКСА ИСПОЛНИТЕЛЕЙ       ########
'################################################################
Private Sub LoadExecutorsCombo()
    On Error GoTo ErrorHandler

    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim executorList As String

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT LastName & ' ' & Left(FirstName,1) & '.' & Left(MiddleName,1) & '.' AS FullName " & _
                             "FROM tbExecutors WHERE ID IS NOT NULL ORDER BY LastName, FirstName")

    ' СОЗДАЕМ СПИСОК ЗНАЧЕНИЙ
    executorList = "Все исполнители"

    Do While Not rs.EOF
        executorList = executorList & ";" & rs!FullName
        rs.MoveNext
    Loop

    ' УСТАНАВЛИВАЕМ ИСТОЧНИК ДАННЫХ
    Me.cboSearchExecutor.RowSourceType = "Value List"
    Me.cboSearchExecutor.rowSource = executorList
    Me.cboSearchExecutor.value = "Все исполнители"

    rs.Close
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка загрузки исполнителей: " & Err.description, vbExclamation
    If Not rs Is Nothing Then rs.Close
End Sub

'################################################################
'########          ЗАГРУЗКА КОМБОБОКСА СТАТУСОВ          ########
'################################################################
Private Sub LoadStatusCombo()
    Me.cboCompletionStatus.RowSourceType = "Value List"
    Me.cboCompletionStatus.rowSource = "Все статусы;Выполнено;Не выполнено"
    Me.cboCompletionStatus.value = "Все статусы"
End Sub

'################################################################
'########         ЗАГРУЗКА КОМБОБОКСА "ЕСТЬ ВЛОЖЕНИЯ"    ########
'################################################################
Private Sub LoadAttachmentCombo()
    Me.cboHasAttachment.RowSourceType = "Value List"
    Me.cboHasAttachment.rowSource = "Не важно;Есть вложения;Нет вложений"
    Me.cboHasAttachment.value = "Не важно"
End Sub

'################################################################
'########          КНОПКА "СБРОСИТЬ" - СБРОС ФИЛЬТРОВ    ########
'################################################################
Private Sub cmdReset_Click()
    ' СБРАСЫВАЕМ ВСЕ ПОЛЯ ПОИСКА
    Me.txtSearchText = ""
    Me.dtStartDate = Null
    Me.dtEndDate = Null
    Me.cboSearchExecutor = "Все исполнители"
    Me.cboCompletionStatus = "Все статусы"
    Me.cboHasAttachment = "Не важно"

    ' УСТАНАВЛИВАЕМ SQL БЕЗ ПОЛЯ ID
    Me.RecordSource = BuildSearchFormBaseSql() & _
                     "ORDER BY ei.EventDate DESC"

    Me.Requery
End Sub

'################################################################
'########          КНОПКА "НАЙТИ" - ПОИСК СОБЫТИЙ        ########
'################################################################
Private Sub cmdSearch_Click()
    On Error GoTo ErrorHandler

    Dim sqlWhere As String
    Dim sql As String

    ' ФОРМИРУЕМ УСЛОВИЯ WHERE
    sqlWhere = BuildSearchConditions

    ' ФОРМИРУЕМ ПОЛНЫЙ SQL ЗАПРОС БЕЗ ID
    sql = BuildSearchFormBaseSql()

    If sqlWhere <> "" Then
        sql = sql & " WHERE " & sqlWhere
    End If

    sql = sql & " ORDER BY ei.EventDate DESC"

    ' УСТАНАВЛИВАЕМ ИСТОЧНИК ДАННЫХ ДЛЯ ФОРМЫ
    Me.RecordSource = sql
    Me.Requery

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка выполнения поиска: " & Err.description, vbCritical
End Sub

'################################################################
'########          ПОСТРОЕНИЕ УСЛОВИЙ ПОИСКА             ########
'################################################################
Private Function BuildSearchConditions() As String
    Dim conditions As String
    conditions = ""

    ' УСЛОВИЕ ПО ТЕКСТУ СОБЫТИЯ
    If Not IsNull(Me.txtSearchText) And Me.txtSearchText <> "" Then
        If conditions <> "" Then conditions = conditions & " AND "
        conditions = conditions & "ei.EventNote LIKE '*" & Replace(Me.txtSearchText, "'", "''") & "*'"
    End If

    ' УСЛОВИЕ ПО ПЕРИОДУ ДАТ
    If Not IsNull(Me.dtStartDate) And Me.dtStartDate <> "" Then
        If conditions <> "" Then conditions = conditions & " AND "
        conditions = conditions & "ei.EventDate >= #" & Format(Me.dtStartDate, "yyyy-mm-dd") & "#"
    End If

    If Not IsNull(Me.dtEndDate) And Me.dtEndDate <> "" Then
        If conditions <> "" Then conditions = conditions & " AND "
        conditions = conditions & "ei.EventDate <= #" & Format(Me.dtEndDate, "yyyy-mm-dd") & "#"
    End If

    ' УСЛОВИЕ ПО ИСПОЛНИТЕЛЮ
    If Me.cboSearchExecutor <> "Все исполнители" Then
        If conditions <> "" Then conditions = conditions & " AND "
        conditions = conditions & "e.LastName & ' ' & Left(e.FirstName,1) & '.' & Left(e.MiddleName,1) & '.' = '" & Replace(Me.cboSearchExecutor, "'", "''") & "'"
    End If

    ' УСЛОВИЕ ПО СТАТУСУ ВЫПОЛНЕНИЯ
    If Me.cboCompletionStatus <> "Все статусы" Then
        If conditions <> "" Then conditions = conditions & " AND "
        If Me.cboCompletionStatus = "Выполнено" Then
            conditions = conditions & "(ei.CompletionMark IS NOT NULL AND ei.CompletionMark <> '')"
        Else
            conditions = conditions & "(ei.CompletionMark IS NULL OR ei.CompletionMark = '')"
        End If
    End If

    ' УСЛОВИЕ ПО НАЛИЧИЮ ВЛОЖЕНИЙ
    If Me.cboHasAttachment <> "Не важно" Then
        If conditions <> "" Then conditions = conditions & " AND "
        If Me.cboHasAttachment = "Есть вложения" Then
            conditions = conditions & "(ei.AttachmentPath IS NOT NULL AND ei.AttachmentPath <> '')"
        Else
            conditions = conditions & "(ei.AttachmentPath IS NULL OR ei.AttachmentPath = '')"
        End If
    End If

    BuildSearchConditions = conditions
End Function

'################################################################
'########          ДВОЙНОЙ КЛИК ПО ДАТЕ СОБЫТИЯ          ########
'################################################################
Private Sub resEventDate_DblClick(Cancel As Integer)
    OpenDayFromSearch
End Sub

'################################################################
'########         ДВОЙНОЙ КЛИК ПО ТЕКСТУ СОБЫТИЯ         ########
'################################################################
Private Sub resEventNote_DblClick(Cancel As Integer)
    OpenDayFromSearch
End Sub

'################################################################
'########         ДВОЙНОЙ КЛИК ПО ИСПОЛНИТЕЛЮ            ########
'################################################################
Private Sub resExecutor_DblClick(Cancel As Integer)
    OpenDayFromSearch
End Sub

'################################################################
'########          ДВОЙНОЙ КЛИК ПО СТАТУСУ               ########
'################################################################
Private Sub resStatus_DblClick(Cancel As Integer)
    OpenDayFromSearch
End Sub

'################################################################
'########       ОТКРЫТИЕ ФОРМЫ ДНЯ ИЗ РЕЗУЛЬТАТОВ        ########
'################################################################
Private Sub OpenDayFromSearch()
    On Error GoTo ErrorHandler

    If Not IsNull(Me.EventDate) Then
        ' ОТКРЫВАЕМ ФОРМУ ДНЯ С НУЖНОЙ ДАТОЙ
        DoCmd.OpenForm "frmDayEvents"

        ' УСТАНАВЛИВАЕМ ФИЛЬТР НА ДАТУ СОБЫТИЯ
        Forms!frmDayEvents.RecordSource = "SELECT * FROM tbEventInstances WHERE EventDate = #" & _
                                          Format(Me.EventDate, "yyyy-mm-dd") & "#"
        Forms!frmDayEvents.lblDate.Caption = Format(Me.EventDate, "d mmmm yyyy ""г.""")
    Else
        MsgBox "Не удалось определить дату события", vbExclamation
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка открытия события: " & Err.description, vbExclamation
End Sub

'################################################################
'########          ?????? "???????"                   ########
'################################################################
Private Sub cmdClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

'################################################################
'########          ПУБЛИЧНЫЕ МЕТОДЫ ДЛЯ ДЕМО-РЕЖИМА      ########
'################################################################

Public Sub ExecuteSearch()
    Call cmdSearch_Click
End Sub

Public Sub ResetSearch()
    Call cmdReset_Click
End Sub

Public Sub CloseSearchForm()
    Call cmdClose_Click
End Sub
