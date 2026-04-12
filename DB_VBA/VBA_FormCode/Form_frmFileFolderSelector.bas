Option Compare Database

'################################################################
'########            ФОРМА ВЫБОРА ТИПА                   ########
'################################################################

'################################################################
'########         ОБРАБОТКА ПАРАМЕТРА ОТКРЫТИЯ           ########
'################################################################
Private Sub Form_Open(Cancel As Integer)
    ' Сохраняем параметр в Tag формы
    If Not IsNull(Me.OpenArgs) Then
        Me.Tag = Me.OpenArgs
    Else
        Me.Tag = "Main" ' По умолчанию
    End If
End Sub

'################################################################
'########            КНОПКА ВЫБРАТЬ ФАЙЛ                 ########
'################################################################
Private Sub cmdSelectFile_Click()
    BrowseForFile
    DoCmd.Close acForm, "frmFileFolderSelector"
End Sub

'################################################################
'########            КНОПКА ВЫБРАТЬ ПАПКУ                ########
'################################################################
Private Sub cmdSelectFolder_Click()
    BrowseForFolder
    DoCmd.Close acForm, "frmFileFolderSelector"
End Sub

'################################################################
'########            КНОПКА ОТМЕНА                       ########
'################################################################
Private Sub cmdCancel_Click()
    DoCmd.Close acForm, "frmFileFolderSelector"
End Sub

'################################################################
'########            ВЫБОР ФАЙЛА                         ########
'################################################################
Private Sub BrowseForFile()
    Dim fileDialog As Object
    Dim selectedFile As Variant
    
    Set fileDialog = Application.fileDialog(3)
    
    With fileDialog
        .title = "Выберите файл"
        .AllowMultiSelect = False
        If .Show Then
            ReturnResult .SelectedItems(1)
        End If
    End With
End Sub

'################################################################
'########            ВЫБОР ПАПКИ                         ########
'################################################################
Private Sub BrowseForFolder()
    Dim folderDialog As Object
    Dim selectedFolder As Variant
    
    Set folderDialog = Application.fileDialog(4)
    
    With folderDialog
        .title = "Выберите папку"
        .AllowMultiSelect = False
        If .Show Then
            ReturnResult .SelectedItems(1)
        End If
    End With
End Sub

'################################################################
'########         ВОЗВРАТ РЕЗУЛЬТАТА В ИСХОДНУЮ ФОРМУ    ########
'################################################################
Private Sub ReturnResult(SelectedPath As String)
    On Error GoTo ErrorHandler
    
    ' Определяем в какую форму возвращать результат
    If CurrentProject.allForms("frmEventGenerator").IsLoaded Then
        ' Возвращаем в форму генератора
        If Me.Tag = "Basis" Then
            Forms!frmEventGenerator!txtBasisAttachment = SelectedPath
        Else
            Forms!frmEventGenerator!txtAttachmentPath = SelectedPath
        End If
    Else
        ' Возвращаем в форму дня (по умолчанию)
        If Me.Tag = "Basis" Then
            Forms!frmDayEvents!txtBasisAttachment = SelectedPath
        Else
            Forms!frmDayEvents!txtAttachmentPath = SelectedPath
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка возврата результата: " & Err.description, vbCritical
End Sub
