Option Compare Database

'################################################################
'########            ЗАГРУЗКА ФОРМЫ ПОИСКА               ########
'################################################################
Private Sub Form_Load()
    Me.fldPassword.InputMask = "Password"
    Me.fldPassword.SetFocus
End Sub

'################################################################
'########           ПРОВЕРКА ПАРОЛЯ ПРИ ENTER            ########
'################################################################
Private Sub fldPassword_KeyPress(KeyAscii As Integer)
    ' код клавиши Enter = 13
    If KeyAscii = 13 Then
        Call CheckPassword
        KeyCode = 0
    End If
End Sub

'################################################################
'########                   ПРОВЕРКА ПАРОЛЯ              ########
'################################################################
Private Sub CheckPassword()
    Dim inputPassword As String
    Dim storedHash As String
    Dim inputHash As String
    
    inputPassword = Nz(Me.fldPassword.value, "")
    
    If inputPassword = "" Then
        MsgBox "Введите пароль", vbExclamation
        Exit Sub
    End If
    
    ' ХЭШИРУЕМ ВВЕДЕРННЫЙ ПАРОЛЬ
    inputHash = AdvancedHash(inputPassword)
    
    ' получаем сохраненный ХЭШ из базы
    storedHash = LoadLicenseSetting("AdminPassword")
    
    ' сравниваем ХЭШИ
    If inputHash = storedHash Then
        Debug.Print "ПАРОЛЬ ВЕРНЫЙ"
        DoCmd.OpenForm "frmAdmin"
        DoCmd.Close acForm, "frmPassword"
    Else
        MsgBox "Неверный пароль", vbCritical
        DoCmd.Close acForm, "frmPassword"
    End If
End Sub
