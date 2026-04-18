Option Compare Database
Option Explicit

'########################################################################
'########     СПИСКИ ОБЪЕКТОВ VBE (ТРЕБУЕТСЯ ДОСТУП К VBE)       ########
'########################################################################
' Назначение: Вывод состава VBProject в Immediate.
' Зависимости: Microsoft Visual Basic for Applications Extensibility 5.3.
'########################################################################

'########################################################################
'########           ПОЛУЧЕНИЕ СПИСКА ОБЪЕКТОВ ПРОЕКТА ########
'########################################################################
Public Sub ProjAnVbe_GetProjectObjectsList()
    On Error GoTo ErrorHandler

    Dim comp As VBComponent
    Dim objType As String
    Dim modulesCount As Integer
    Dim formsCount As Integer
    Dim classesCount As Integer
    Dim totalObjects As Integer

    modulesCount = 0
    formsCount = 0
    classesCount = 0
    totalObjects = 0

    Debug.Print "=============================================="
    Debug.Print "ОБЪЕКТЫ ПРОЕКТА 'ЕЖЕДНЕВНИК'"
    Debug.Print "=============================================="
    Debug.Print ""

    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        totalObjects = totalObjects + 1

        Select Case comp.Type
            Case vbext_ct_StdModule
                objType = "МОДУЛЬ"
                modulesCount = modulesCount + 1
            Case vbext_ct_ClassModule
                objType = "КЛАСС"
                classesCount = classesCount + 1
            Case vbext_ct_MSForm, vbext_ct_Document
                objType = "ФОРМА"
                formsCount = formsCount + 1
            Case Else
                objType = "ДРУГОЙ"
        End Select

        Debug.Print objType & ": " & comp.Name
        Debug.Print "   Строк кода: " & comp.CodeModule.CountOfLines
        Debug.Print ""
    Next comp

    Debug.Print "=============================================="
    Debug.Print "СТАТИСТИКА:"
    Debug.Print "Модули: " & modulesCount
    Debug.Print "Формы: " & formsCount
    Debug.Print "Классы: " & classesCount
    Debug.Print "Всего объектов: " & totalObjects
    Debug.Print "=============================================="

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка при получении списка объектов: " & Err.description, vbCritical
End Sub

'########################################################################
'########           ПОЛУЧЕНИЕ ТОЛЬКО МОДУЛЕЙ           ########
'########################################################################
Public Sub ProjAnVbe_GetModulesList()
    On Error GoTo ErrorHandler

    Dim comp As VBComponent
    Dim count As Integer

    count = 0

    Debug.Print "=============================================="
    Debug.Print "МОДУЛИ ПРОЕКТА"
    Debug.Print "=============================================="
    Debug.Print ""

    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        If comp.Type = vbext_ct_StdModule Then
            count = count + 1
            Debug.Print count & ". " & comp.Name
            Debug.Print "   Строк: " & comp.CodeModule.CountOfLines
        End If
    Next comp

    Debug.Print ""
    Debug.Print "Всего модулей: " & count

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка при получении списка модулей: " & Err.description, vbCritical
End Sub

'########################################################################
'########                  ПОЛУЧЕНИЕ ТОЛЬКО ФОРМ                 ########
'########################################################################
Public Sub ProjAnVbe_GetFormsList()
    On Error GoTo ErrorHandler

    Dim comp As VBComponent
    Dim count As Integer

    count = 0

    Debug.Print "=============================================="
    Debug.Print "ФОРМЫ ПРОЕКТА"
    Debug.Print "=============================================="
    Debug.Print ""

    For Each comp In Application.VBE.ActiveVBProject.VBComponents
        If comp.Type = vbext_ct_MSForm Or comp.Type = vbext_ct_Document Then
            count = count + 1
            Debug.Print count & ". " & comp.Name
            Debug.Print "   Строк: " & comp.CodeModule.CountOfLines
        End If
    Next comp

    Debug.Print ""
    Debug.Print "Всего форм: " & count

    Exit Sub

ErrorHandler:
    MsgBox "Ошибка при получении списка форм: " & Err.description, vbCritical
End Sub

