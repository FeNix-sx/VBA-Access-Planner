Option Compare Database

'################################################################
'########      ЭКСПОРТ ВСЕГО КОДА В ОДИН ФАЙЛ         ########
'################################################################

Public Sub ExportAllCodeToSingleFile()
    On Error GoTo ExportAllCodeToSingleFile_Error
    
    Dim comp As Object
    Dim exportPath As String
    Dim fso As Object
    Dim txtFile As Object
    Dim i As Integer
    Dim lineCode As String
    
    ' ПАПКА И ФАЙЛ ДЛЯ ЭКСПОРТА
    exportPath = "C:\VBA_Export\"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' СОЗДАЕМ ПАПКУ
    If Not fso.FolderExists(exportPath) Then fso.CreateFolder exportPath
    
    ' СОЗДАЕМ ЕДИНЫЙ ФАЙЛ
    Set txtFile = fso.CreateTextFile(exportPath & "ALL_VBA_CODE.txt", True, True)
    
    ' ЗАГОЛОВОК ФАЙЛА
    txtFile.WriteLine "=========================================="
    txtFile.WriteLine "ВЕСЬ КОД VBA ИЗ ПРОЕКТА 'ЕЖЕДНЕВНИК'"
    txtFile.WriteLine "Сгенерировано: " & Now
    txtFile.WriteLine "==========================================" & vbCrLf
    
    ' ЭКСПОРТИРУЕМ ВСЕ КОМПОНЕНТЫ
    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        txtFile.WriteLine "------------------------------------------"
        txtFile.WriteLine "МОДУЛЬ: " & comp.Name
        txtFile.WriteLine "ТИП: " & GetComponentType(comp.Type)
        txtFile.WriteLine "------------------------------------------" & vbCrLf
        
        ' ЗАПИСЫВАЕМ ВЕСЬ КОД МОДУЛЯ
        With comp.CodeModule
            For i = 1 To .CountOfLines
                lineCode = .Lines(i, 1)
                txtFile.WriteLine lineCode
            Next i
        End With
        
        txtFile.WriteLine vbCrLf & vbCrLf
    Next comp
    
    txtFile.Close
    MsgBox "Весь код экспортирован в один файл: " & exportPath & "ALL_VBA_CODE.txt", vbInformation
    
    Exit Sub
    
ExportAllCodeToSingleFile_Error:
    MsgBox "Ошибка экспорта: " & Err.description, vbCritical
End Sub

'################################################################
'########          ОПРЕДЕЛЕНИЕ ТИПА МОДУЛЯ           ########
'################################################################

Private Function GetComponentType(compType As Integer) As String
    Select Case compType
        Case 1: GetComponentType = "Standard Module"
        Case 2: GetComponentType = "Class Module"
        Case 3: GetComponentType = "Form"
        Case 100: GetComponentType = "Document Module"
        Case Else: GetComponentType = "Unknown"
    End Select
End Function
