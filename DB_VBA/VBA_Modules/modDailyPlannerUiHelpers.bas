Option Compare Database
Option Explicit

'################################################################
'########      ОБНОВЛЕНИЕ ГЛАВНОГО КАЛЕНДАРЯ             ########
'################################################################
Public Sub RefreshDailyPlannerIfLoaded()
    On Error GoTo ErrorHandler
    If CurrentProject.AllForms("f_daily_planner").IsLoaded Then
        Form_f_daily_planner.BuildCalendar
    End If
    Exit Sub
ErrorHandler:
    ' Пропускаем ошибку обновления календаря - не критично
End Sub

'################################################################
'########      ЗАГРУЗКА ИСПОЛНИТЕЛЕЙ В КОМБОБОКС         ########
'################################################################
Public Sub BindExecutorCombo(ByRef cbo As ComboBox)
    On Error GoTo ErrorHandler
    
    cbo.RowSource = "SELECT ID, LastName & ' ' & Left(FirstName,1) & '.' & Left(MiddleName,1) & '.' AS FullName " & _
                    "FROM tbExecutors WHERE ID IS NOT NULL ORDER BY SortOrder, LastName, FirstName"
    cbo.ColumnCount = 2
    cbo.BoundColumn = 1
    cbo.ColumnWidths = "0;4см"
    
    Exit Sub
ErrorHandler:
    MsgBox "Ошибка загрузки списка исполнителей: " & Err.description, vbExclamation
End Sub

'################################################################
'########   ОТКРЫТИЕ ВЛОЖЕНИЯ ПО ПУТИ (файл или папка)   ########
'################################################################
Public Sub OpenAttachmentHyperlink(ByVal strPath As String, Optional ByRef frm As Form = Nothing)
    On Error GoTo ErrorHandler

    strPath = Trim$(strPath)
    If strPath = "" Then Exit Sub

    If Dir(strPath, vbDirectory) = "" Then
        MsgBox "Файл или папка не найдены: " & strPath, vbExclamation
        Exit Sub
    End If

    DoCmd.Hourglass True
    If Not frm Is Nothing Then frm.Repaint
    
    FollowHyperlink strPath
    
    DoCmd.Hourglass False
    Exit Sub

ErrorHandler:
    DoCmd.Hourglass False
    MsgBox "Ошибка открытия: " & Err.description, vbCritical
End Sub
