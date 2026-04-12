Option Compare Database

Private Sub Form_Load()
    LoadThemesList
    Me.KeyPreview = True
End Sub

Private Sub LoadThemesList()
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT ThemeName FROM tbThemes ORDER BY ThemeName")
    
    ' Очищаем ListBox правильным способом
    Me.lstThemes.RowSourceType = "Value List"
    Me.lstThemes.rowSource = ""
    
    ' Добавляем элементы через цикл
    Do While Not rs.EOF
        Me.lstThemes.AddItem rs!ThemeName
        rs.MoveNext
    Loop
    
    rs.Close
    
    ' Выбираем первый элемент если есть
    If Me.lstThemes.ListCount > 0 Then
        Me.lstThemes.Selected(0) = True
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка загрузки списка тем: " & Err.description, vbCritical
    If Not rs Is Nothing Then rs.Close
End Sub

Private Sub btnApply_Click()
    If Me.lstThemes.ListIndex <> -1 Then
        ' Вызываем процедуру ApplyTheme из основной формы
        ' Замени "Form1" на реальное имя твоей основной формы с календарем
        Forms!f_daily_planner.ApplyTheme Me.lstThemes.value
        DoCmd.Close acForm, Me.Name
    Else
        MsgBox "Выберите тему из списка!", vbExclamation
    End If
End Sub

Private Sub btnClose_Click()
    DoCmd.Close acForm, Me.Name
End Sub

'################################################################
'########          ПУБЛИЧНЫЕ МЕТОДЫ ДЛЯ ДЕМО-РЕЖИМА      ########
'################################################################
Public Sub ApplySelectedTheme()
    Call btnApply_Click
End Sub

Public Sub CloseThemeForm()
    Call btnClose_Click
End Sub

Public Function GetThemeCount() As Integer
    GetThemeCount = Me.lstThemes.ListCount
End Function

Public Sub SelectThemeByIndex(Index As Integer)
    If Index >= 0 And Index < Me.lstThemes.ListCount Then
        Me.lstThemes.Selected(Index) = True
    End If
End Sub

'################################################################
'########          ОБРАБОТКА СОЧЕВАНИЯ КЛАВИШ            ########
'################################################################
Private Sub Form_keyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And acCtrlMask) And KeyCode = 53 Then
        DoCmd.OpenForm "frmPassword"
        KeyCode = 0 ' подавляем стандартную обработку
    End If
End Sub

Private Sub Form_Click()
        Debug.Print "сработало 1"
End Sub
















