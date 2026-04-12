Option Compare Database

'################################################################
'########    СПИСОК СПРАВОЧНИКА ДНЕЙ РОЖДЕНИЯ (ЛЕНТА)    ########
'################################################################
' Свойства формы (конструктор Access):
' - Record Source: tbBirthdays
' - Default View: Continuous Forms
' - Allow Additions: No, Allow Edits: No, Allow Deletions: No
' - Order By: LastName, FirstName, MiddleName
' - Поля: fld_* (кроме ID — поле источника, обращение Me.ID без отдельного контрола)
' - Кнопки: btn_add, btn_edit, btn_delete, btn_close (core-settings.mdc)
' Редактирование — в frmBirthdayCard; удаление — кнопка с подтверждением.
'################################################################

Private Sub Form_Load()
    On Error Resume Next
    DoCmd.MoveSize 8000, 5000, 7700, 7000
End Sub

Private Sub btn_add_Click()
    On Error GoTo ErrorHandler
    DoCmd.OpenForm "frmBirthdayCard", , , , acFormAdd
    Exit Sub
ErrorHandler:
    MsgBox "Не удалось открыть карточку: " & Err.description, vbCritical
End Sub

Private Sub btn_edit_Click()
    Call OpenCurrentBirthdayCard
End Sub

Private Sub fld_LastName_DblClick(Cancel As Integer)
    Call OpenCurrentBirthdayCard
End Sub

Private Sub fld_FirstName_DblClick(Cancel As Integer)
    Call OpenCurrentBirthdayCard
End Sub

Private Sub fld_MiddleName_DblClick(Cancel As Integer)
    Call OpenCurrentBirthdayCard
End Sub

Private Sub fld_BirthDate_DblClick(Cancel As Integer)
    Call OpenCurrentBirthdayCard
End Sub

Private Sub OpenCurrentBirthdayCard()
    On Error GoTo ErrorHandler
    If IsNull(Me.ID) Then
        MsgBox "Выберите запись в списке.", vbInformation, "Дни рождения"
        Exit Sub
    End If
    DoCmd.OpenForm "frmBirthdayCard", , , "ID=" & CLng(Me.ID)
    Exit Sub
ErrorHandler:
    MsgBox "Не удалось открыть карточку: " & Err.description, vbCritical
End Sub

Private Sub btn_delete_Click()
    On Error GoTo ErrorHandler
    If IsNull(Me.ID) Then
        MsgBox "Выберите запись в списке.", vbInformation, "Дни рождения"
        Exit Sub
    End If

    Dim title As String
    title = Trim(Nz(Me.fld_LastName, "")) & " " & Trim(Nz(Me.fld_FirstName, ""))
    If Len(title) = 0 Then title = "запись"

    If MsgBox("Удалить " & title & "?", vbYesNo + vbQuestion + vbDefaultButton2, "Удаление") = vbNo Then
        Exit Sub
    End If

    CurrentDb.Execute "DELETE FROM tbBirthdays WHERE ID=" & CLng(Me.ID), dbFailOnError
    Me.Requery
    Call RefreshBirthdaysUIAfterEdit
    Exit Sub
ErrorHandler:
    MsgBox "Ошибка удаления: " & Err.description, vbCritical, "Дни рождения"
End Sub

Private Sub btn_close_Click()
    On Error GoTo ErrorHandler
    DoCmd.Close acForm, Me.Name
    Exit Sub
ErrorHandler:
    MsgBox "Ошибка закрытия: " & Err.description, vbCritical
End Sub

