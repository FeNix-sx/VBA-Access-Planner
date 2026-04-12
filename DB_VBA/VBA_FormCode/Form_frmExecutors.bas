Option Compare Database

'################################################################
'########           КНОПКА ЗАКРЫТЬ ФОРМУ                 ########
'################################################################
Private Sub cmdClose_Click()
    On Error GoTo ErrorHandler
    DoCmd.Close acForm, "frmExecutors"
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка закрытия формы: " & Err.description, vbCritical
End Sub

'################################################################
'########          ОБНОВЛЕНИЕ ФИЛЬТРА ИСПОЛНИТЕЛЕЙ       ########
'################################################################

Private Sub Form_Close()
    ' ОБНОВЛЯЕМ ФИЛЬТР ИСПОЛНИТЕЛЕЙ НА ГЛАВНОЙ ФОРМЕ
    If CurrentProject.allForms("f_daily_planner").IsLoaded Then
        Forms!f_daily_planner.InitializeExecutorFilter
    End If
End Sub

'################################################################
'########          НАСТРОЙКА РАЗМЕРА И ПОЛОЖЕНИЯ         ########
'################################################################

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    ' УСТАНАВЛИВАЕМ РАЗМЕР И ПОЛОЖЕНИЕ ФОРМЫ
    ' DoCmd.MoveSize Left, Top, Width, Height
    ' Left   - отступ слева в твипах (1440 твипов = 1 дюйм = 2.54 см)
    ' Top    - отступ сверху в твипах
    ' Width  - ширина формы в твипах
    ' Height - высота формы в твипах
    
    DoCmd.MoveSize 5000, 1500, 15500, 10000
    Exit Sub
    
ErrorHandler:

End Sub
