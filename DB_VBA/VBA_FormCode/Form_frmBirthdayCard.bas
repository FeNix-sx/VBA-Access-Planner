Option Compare Database

'################################################################
'########   КАРТОЧКА ЗАПИСИ СПРАВОЧНИКА ДНЕЙ РОЖДЕНИЯ    ########
'################################################################
' Свойства формы (конструктор Access):
' - Record Source: tbBirthdays
' - Default View: Single Form
' - Allow Additions: Yes, Allow Edits: Yes, Allow Deletions: Yes
' - Поля: префикс fld_ (см. core-settings.mdc); кнопки: btn_save, btn_close
'################################################################

Private Sub Form_Load()
    On Error Resume Next
    DoCmd.MoveSize 10000, 5000, 5300, 3900
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    If Len(Trim(Nz(Me.fld_LastName, ""))) = 0 Then
        MsgBox "Укажите фамилию.", vbExclamation, "Дни рождения"
        Cancel = True
        Exit Sub
    End If
    If Len(Trim(Nz(Me.fld_FirstName, ""))) = 0 Then
        MsgBox "Укажите имя.", vbExclamation, "Дни рождения"
        Cancel = True
        Exit Sub
    End If
    If Not IsDate(Me.fld_BirthDate) Then
        MsgBox "Укажите дату рождения.", vbExclamation, "Дни рождения"
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub Form_AfterUpdate()
    Call RefreshBirthdaysUIAfterEdit
End Sub

Private Sub Form_AfterInsert()
    Call RefreshBirthdaysUIAfterEdit
End Sub

Private Sub Form_Close()
    Call RefreshBirthdaysUIAfterEdit
End Sub

Private Sub btn_save_Click()
    On Error GoTo ErrorHandler
    If Me.Dirty Then
        DoCmd.RunCommand acCmdSaveRecord
    End If
    Call RefreshBirthdaysUIAfterEdit
    Exit Sub
ErrorHandler:
    MsgBox "Не удалось сохранить: " & Err.description, vbExclamation, "Дни рождения"
End Sub

Private Sub btn_close_Click()
    On Error GoTo ErrorHandler
    DoCmd.Close acForm, Me.Name, acSavePrompt
    Exit Sub
ErrorHandler:
    MsgBox "Ошибка закрытия: " & Err.description, vbCritical
End Sub

