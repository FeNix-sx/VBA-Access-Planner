Option Compare Database

' Размер и позиция окна frmDayEvents (twips)
Private Const TWIPS_FRMDAYEVENTS_LEFT As Long = 6000
Private Const TWIPS_FRMDAYEVENTS_TOP As Long = 2000
Private Const TWIPS_FRMDAYEVENTS_WIDTH As Long = 18500
Private Const TWIPS_FRMDAYEVENTS_HEIGHT As Long = 11000

'################################################################
'########      РАЗМЕР И ПОЛОЖЕНИЕ ОКНА frmDayEvents       ########
'################################################################
Private Sub ApplyFrmDayEventsWindow()
' Назначение: Задаёт позицию и габариты формы на экране.
' Принцип:    Один вызов DoCmd.MoveSize с константами twips.
'================================================================
    DoCmd.MoveSize TWIPS_FRMDAYEVENTS_LEFT, TWIPS_FRMDAYEVENTS_TOP, _
                   TWIPS_FRMDAYEVENTS_WIDTH, TWIPS_FRMDAYEVENTS_HEIGHT
End Sub

'################################################################
'########        ОБНОВЛЕНИЕ ФОРМЫ ПРИ ЗАГРУЗКЕ          ########
'################################################################
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    ' Загружаем список исполнителей
    LoadExecutorsCombo
    
    ApplyFrmDayEventsWindow
    
    ' ВКЛЮЧАЕМ РЕЖИМ ПРОСМОТРА ПРИ ЗАГРУЗКЕ (один раз)
    SwitchToViewMode
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка загрузки формы: " & Err.description, vbCritical
End Sub

'################################################################
'########        ОТКРЫТИЕ ФОРМЫ (без дубля с Form_Load)  ########
'################################################################
Private Sub Form_Open(Cancel As Integer)
    ' Размер, позиция и SwitchToViewMode — в Form_Load после загрузки контролов.
End Sub

'################################################################
'########              НАВИГАЦИЯ ПО ДНЯМ                 ########
'################################################################
Private Sub cmdPrevDay_Click()
    On Error GoTo ErrorHandler
    Dim labelText As String
    labelText = Replace(Me.lblDate.Caption, " г.", "")
    Me.lblDate.Caption = Format(DateAdd("d", -1, CDate(labelText)), "d mmmm yyyy ""г.""")
    LoadDayEvents
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка перехода к предыдущему дню: " & Err.description, vbCritical
End Sub

Private Sub cmdNextDay_Click()
    On Error GoTo ErrorHandler
    Dim labelText As String
    labelText = Replace(Me.lblDate.Caption, " г.", "")
    Me.lblDate.Caption = Format(DateAdd("d", 1, CDate(labelText)), "d mmmm yyyy ""г.""")
    LoadDayEvents
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка перехода к следующему дню: " & Err.description, vbCritical
End Sub

'################################################################
'########           ЗАГРУЗКА СОБЫТИЙ ДНЯ               ########
'################################################################
Private Sub LoadDayEvents()
    On Error GoTo ErrorHandler
    Dim currentDate As Date
    currentDate = CDate(Replace(Me.lblDate.Caption, " г.", ""))
    
    Me.RecordSource = "SELECT * FROM tbEventInstances WHERE EventDate = " & _
                      Format(currentDate, "\#mm\/dd\/yyyy\#")
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка загрузки событий дня: " & Err.description, vbCritical
End Sub

'################################################################
'########           КНОПКА ЗАКРЫТЬ ФОРМУ                 ########
'################################################################
Private Sub cmdClose_Click()
    On Error GoTo ErrorHandler
    DoCmd.Close acForm, "frmDayEvents"
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка закрытия формы: " & Err.description, vbCritical
End Sub

'################################################################
'########        ПЕРЕКЛЮЧЕНИЕ РЕЖИМОВ (ОБНОВЛЕНО)       ########
'################################################################
Private Sub SwitchToViewMode()
    ' Режим просмотра
    Me.AllowEdits = False
    Me.AllowAdditions = False
    Me.AllowDeletions = False
    Me.txtCompletionMark.Locked = True
    Me.cboExecutor.Locked = True
    Me.txtBasic.Locked = True
    Me.txtBasisAttachment.Locked = True
    Me.cmdEdit.Visible = True
    Me.cmdEdit.SetFocus
    Me.cmdSave.Visible = False
    Me.cmdBrowseFile.Visible = False
    Me.cmdBrowseBasisAttachment.Visible = False
    Me.lblBrowse.Visible = False
End Sub

Private Sub SwitchToEditMode()
    ' Режим редактирования
    Me.AllowEdits = True
    Me.AllowAdditions = True
    Me.AllowDeletions = True
    Me.txtCompletionMark.Locked = False
    Me.cboExecutor.Locked = False
    Me.txtBasic.Locked = False
    Me.txtBasisAttachment.Locked = False
    Me.cmdSave.Visible = True
    Me.cmdSave.SetFocus
    Me.cmdEdit.Visible = False
    Me.cmdBrowseFile.Visible = True
    Me.cmdBrowseBasisAttachment.Visible = True
    Me.lblBrowse.Visible = True
End Sub

'################################################################
'########           СТИЛЬ РЕЖИМА ПРОСМОТРА               ########
'################################################################
Private Sub ApplyViewStyle()
    On Error GoTo ErrorHandler
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is TextBox Or TypeOf ctrl Is ComboBox Then
            ctrl.FontItalic = False
            ctrl.FontBold = False
            ctrl.ForeColor = RGB(0, 0, 0) ' Черный цвет
        End If
    Next ctrl
    Exit Sub
    
ErrorHandler:
    ' Пропускаем ошибки стилей - не критично
End Sub

'################################################################
'########           СТИЛЬ РЕЖИМА РЕДАКТИРОВАНИЯ          ########
'################################################################
Private Sub ApplyEditStyle()
    On Error GoTo ErrorHandler
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is TextBox Or TypeOf ctrl Is ComboBox Then
            ctrl.FontItalic = True
            ctrl.FontBold = True
            ctrl.ForeColor = RGB(0, 0, 128) ' Темно-синий цвет
        End If
    Next ctrl
    Exit Sub
    
ErrorHandler:
    ' Пропускаем ошибки стилей - не критично
End Sub

'################################################################
'########           КНОПКА РЕДАКТИРОВАТЬ                 ########
'################################################################
Private Sub cmdEdit_Click()
    On Error GoTo ErrorHandler
    SwitchToEditMode
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка перехода в режим редактирования: " & Err.description, vbCritical
End Sub

'################################################################
'########            КНОПКА СОХРАНИТЬ                    ########
'################################################################
Private Sub cmdSave_Click()
    On Error GoTo ErrorHandler
    ' Сохраняем изменения
    If Me.Dirty Then
        Me.Dirty = False
    End If
    
    ' Возвращаем в режим просмотра
    SwitchToViewMode
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка сохранения данных: " & Err.description, vbCritical
End Sub

'################################################################
'########           ОБНОВЛЕНИЕ КАЛЕНДАРЯ                 ########
'################################################################
Private Sub Form_Close()
    On Error GoTo ErrorHandler
    ' Обновляем основную форму календаря
    If CurrentProject.allForms("f_daily_planner").IsLoaded Then
        Form_f_daily_planner.BuildCalendar
    End If
    Exit Sub
    
ErrorHandler:
    ' Пропускаем ошибку обновления календаря - не критично
End Sub

'################################################################
'########            АВТОМАТИЧЕСКАЯ ДАТА                 ########
'################################################################
Private Sub Form_BeforeInsert(Cancel As Integer)
    On Error GoTo ErrorHandler
    ' Автоматически заполняем дату события при добавлении новой записи
    If Not IsNull(Me.OpenArgs) Then
        Me.EventDate = CDate(Me.OpenArgs)
    Else
        ' Если дата не передана, используем дату из заголовка
        Me.EventDate = CDate(Replace(Me.lblDate.Caption, " г.", ""))
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка установки даты события: " & Err.description, vbCritical
    Cancel = True
End Sub

'################################################################
'########     ОТКРЫТИЕ ВЛОЖЕНИЯ ПО ПУТИ (файл или папка) ########
'################################################################
Private Sub OpenAttachmentHyperlink(ByVal strPath As String)
' Назначение: Открывает файл или папку по пути вложения (как в проводнике).
' Принцип:    Dir + FollowHyperlink; песочные часы на время вызова.
'================================================================
    On Error GoTo ErrorHandler
    
    strPath = Trim$(strPath)
    If strPath = "" Then Exit Sub
    
    If Dir(strPath, vbDirectory) = "" Then
        MsgBox "Файл или папка не найдены: " & strPath, vbExclamation
        Exit Sub
    End If
    
    DoCmd.Hourglass True
    Me.Repaint
    FollowHyperlink strPath
    DoCmd.Hourglass False
    Exit Sub
    
ErrorHandler:
    DoCmd.Hourglass False
    MsgBox "Ошибка открытия: " & Err.description, vbCritical
End Sub

'################################################################
'########          ОБНОВЛЕНИЕ СОСТОЯНИЯ ФОРМЫ           ########
'################################################################
Private Sub Form_Current()
    On Error GoTo ErrorHandler
    ' Обновляем стили в зависимости от режима
    If Me.AllowEdits Then
        ApplyEditStyle
    Else
        ApplyViewStyle
    End If
    Exit Sub
    
ErrorHandler:
    ' Пропускаем ошибки обновления стилей - не критично
End Sub

'################################################################
'########          ОБРАБОТКА ОТМЕТКИ ВЫПОЛНЕНИЯ         ########
'################################################################
Private Sub txtCompletionMark_AfterUpdate()
    On Error GoTo ErrorHandler
    ' Если отметка заполнена впервые - ставим CompletionDate
    If Not IsNull(Me.txtCompletionMark) And Me.txtCompletionMark <> "" Then
        If IsNull(Me.txtCompletionDate) Then
            Me.txtCompletionDate = Date
        End If
        ' Всегда обновляем LastModified при изменении отметки
        Me.txtLastModified = Now()
    Else
        ' Если отметку очистили - сбрасываем даты
        Me.txtCompletionDate = Null
        Me.txtLastModified = Now()
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при обновлении отметки: " & Err.description, vbCritical
End Sub

'################################################################
'########           ОТКРЫТИЕ ФАЙЛА ИЛИ ПАПКИ             ########
'################################################################
Private Sub txtAttachmentPath_DblClick(Cancel As Integer)
    On Error GoTo ErrorHandler
    
    If Not IsNull(Me.txtAttachmentPath) And Me.txtAttachmentPath <> "" Then
        OpenAttachmentHyperlink CStr(Me.txtAttachmentPath)
    End If
    Exit Sub
    
ErrorHandler:
    Me.txtAttachmentPath.backColor = RGB(255, 255, 255)
    DoCmd.Hourglass False
    MsgBox "Ошибка открытия файла/папки: " & Err.description, vbCritical
End Sub

'################################################################
'########          ВЫБОР ФАЙЛА ИЛИ ПАПКИ               ########
'################################################################
Private Sub cmdBrowseFile_Click()
    On Error GoTo ErrorHandler
    ' Открываем форму выбора с параметром "Main"
    DoCmd.OpenForm "frmFileFolderSelector", , , , , acDialog, "Main"
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка открытия формы выбора файла: " & Err.description, vbCritical
End Sub

'################################################################
'########           ЗАГРУЗКА ИСПОЛНИТЕЛЕЙ В КОМБОБОКС   ########
'################################################################
Private Sub LoadExecutorsCombo()
    On Error GoTo ErrorHandler
    
    Me.cboExecutor.rowSource = "SELECT ID, LastName & ' ' & Left(FirstName,1) & '.' & Left(MiddleName,1) & '.' AS FullName " & _
                              "FROM tbExecutors WHERE ID IS NOT NULL ORDER BY SortOrder, LastName, FirstName"
    Me.cboExecutor.ColumnCount = 2
    Me.cboExecutor.BoundColumn = 1
    Me.cboExecutor.ColumnWidths = "0;4см"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка загрузки списка исполнителей: " & Err.description, vbExclamation
End Sub

'################################################################
'########     ОТКРЫТИЕ ФАЙЛА ОСНОВАНИЯ ПО ДВОЙНОМУ КЛИКУ ########
'################################################################
Private Sub txtBasisAttachment_DblClick(Cancel As Integer)
    On Error GoTo ErrorHandler
    
    If Not IsNull(Me.txtBasisAttachment) And Me.txtBasisAttachment <> "" Then
        ' Проверяем существует ли файл или папка
        If Dir(Me.txtBasisAttachment, vbDirectory) <> "" Then
            ' Показываем уведомление о начале открытия
            DoCmd.Hourglass True
            Me.Repaint
            
            ' Открываем файл или папку
            FollowHyperlink Me.txtBasisAttachment
            
            DoCmd.Hourglass False
        Else
            MsgBox "Файл или папка не найдены: " & Me.txtBasisAttachment, vbExclamation
        End If
    End If
    Exit Sub
    
ErrorHandler:
    ' Восстанавливаем цвет при ошибке
    Me.txtBasisAttachment.backColor = RGB(255, 255, 255)
    DoCmd.Hourglass False
    MsgBox "Ошибка открытия файла/папки основания: " & Err.description, vbCritical
End Sub

'################################################################
'########        ВЫБОР ФАЙЛА ОСНОВАНИЯ                 ########
'################################################################
Private Sub cmdBrowseBasisAttachment_Click()
    On Error GoTo ErrorHandler
    ' Открываем форму выбора с параметром "Basis"
    DoCmd.OpenForm "frmFileFolderSelector", , , , , acDialog, "Basis"
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка выбора файла основания: " & Err.description, vbCritical
End Sub

'################################################################
'########      ОТМЕТИТЬ ВСЕ КАК ВЫПОЛНЕННЫЕ           ########
'################################################################
Private Sub cmdMarkAllComplete_Click()
    On Error GoTo ErrorHandler
    
    If MsgBox("Отметить все события этого дня как выполненные?", vbQuestion + vbYesNo) = vbYes Then
        Dim currentDate As Date
        currentDate = CDate(Replace(Me.lblDate.Caption, " г.", ""))
        
        CurrentDb.Execute "UPDATE tbEventInstances SET " & _
                         "CompletionMark = 'Выполнено', " & _
                         "CompletionDate = Date(), " & _
                         "LastModified = Now() " & _
                         "WHERE EventDate = #" & Format(currentDate, "yyyy\/mm\/dd") & "#"
        
        Me.Requery
        MsgBox "Все события отмечены как выполненные!", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка массового обновления: " & Err.description, vbCritical
End Sub

'################################################################
'########          СНЯТЬ ВСЕ ОТМЕТКИ                  ########
'################################################################
Private Sub cmdUnmarkAll_Click()
    On Error GoTo ErrorHandler
    
    If MsgBox("Снять все отметки о выполнении за этот день?", vbQuestion + vbYesNo) = vbYes Then
        Dim currentDate As Date
        currentDate = CDate(Replace(Me.lblDate.Caption, " г.", ""))
        
        CurrentDb.Execute "UPDATE tbEventInstances SET " & _
                         "CompletionMark = Null, " & _
                         "CompletionDate = Null, " & _
                         "LastModified = Now() " & _
                         "WHERE EventDate = #" & Format(currentDate, "yyyy\/mm\/dd") & "#"
        
        Me.Requery
        MsgBox "Все отметки о выполнении сняты!", vbInformation
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка массового обновления: " & Err.description, vbCritical
End Sub

'################################################################
'########          ПУБЛИЧНЫЕ МЕТОДЫ ДЛЯ ДЕМО-РЕЖИМА      ########
'################################################################

Public Sub GoToNextDay()
    Call cmdNextDay_Click
End Sub

Public Sub GoToPreviousDay()
    Call cmdPrevDay_Click
End Sub

Public Sub StartEditMode()
    Call cmdEdit_Click
End Sub

Public Sub SaveChanges()
    Call cmdSave_Click
End Sub

Public Sub CloseForm()
    Call cmdClose_Click  ' < ЭТОТ МЕТОД ДОЛЖЕН БЫТЬ
End Sub


