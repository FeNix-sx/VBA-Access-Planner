Option Compare Database
Option Explicit

'########################################################################
'########      АНАЛИЗ ГЛАВНОЙ ФОРМЫ И ГЛУБОКИЙ РАЗБОР UI         ########
'########################################################################
' Назначение: Логика анализа f_daily_planner и полного обхода контролов.
' Зависимости: ProjAn_* в modProjectAnalysisCore.
'########################################################################

'########################################################################
'########            АНАЛИЗ СТРУКТУРЫ ГЛАВНОЙ ФОРМЫ              ########
'########################################################################
Public Sub ProjAn_AnalyzeMainFormStructure()
    On Error GoTo ErrorHandler

    Debug.Print "=== АНАЛИЗ ГЛАВНОЙ ФОРМЫ f_daily_planner ==="

    If Not ProjAn_FormExists("f_daily_planner") Then
        Debug.Print "Форма f_daily_planner не найдена"
        Exit Sub
    End If

    Dim formWasLoaded As Boolean
    formWasLoaded = CurrentProject.allForms("f_daily_planner").IsLoaded

    If Not formWasLoaded Then
        DoCmd.OpenForm "f_daily_planner", acNormal, , , , acHidden
    End If

    ProjAn_AnalyzeFormControls

    ProjAn_AnalyzeFormProcedures

    If Not formWasLoaded Then
        DoCmd.Close acForm, "f_daily_planner"
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка анализа главной формы: " & Err.description
    If Not formWasLoaded And CurrentProject.allForms("f_daily_planner").IsLoaded Then
        DoCmd.Close acForm, "f_daily_planner"
    End If
End Sub

'########################################################################
'########            АНАЛИЗ ЭЛЕМЕНТОВ УПРАВЛЕНИЯ ФОРМЫ           ########
'########################################################################
Private Sub ProjAn_AnalyzeFormControls()
    On Error GoTo ErrorHandler

    Dim frm As Form
    Set frm = Forms!f_daily_planner

    Dim controlCount As Integer
    controlCount = 0
    Dim buttonCount As Integer
    buttonCount = 0
    Dim labelCount As Integer
    labelCount = 0
    Dim textBoxCount As Integer
    textBoxCount = 0

    Debug.Print "--- ЭЛЕМЕНТЫ УПРАВЛЕНИЯ ---"

    Dim ctrl As Control
    For Each ctrl In frm.Controls
        controlCount = controlCount + 1

        Select Case TypeName(ctrl)
            Case "CommandButton"
                buttonCount = buttonCount + 1
                Debug.Print "КНОПКА: " & ctrl.Name & " | '" & ctrl.Caption & _
                           "' | Pos: " & ctrl.Left & "," & ctrl.Top & _
                           " | Size: " & ctrl.Width & "x" & ctrl.Height

            Case "Label"
                labelCount = labelCount + 1

            Case "TextBox"
                textBoxCount = textBoxCount + 1

            Case Else
                Debug.Print "ДРУГОЙ: " & ctrl.Name & " | Тип: " & TypeName(ctrl)
        End Select
    Next ctrl

    Debug.Print "--- СТАТИСТИКА ---"
    Debug.Print "Всего элементов: " & controlCount
    Debug.Print "Кнопок: " & buttonCount
    Debug.Print "Надписей: " & labelCount
    Debug.Print "Текстовых полей: " & textBoxCount

    ProjAn_AnalyzeButtonLayout frm

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка анализа элементов управления: " & Err.description
End Sub

'########################################################################
'########            АНАЛИЗ РАСПОЛОЖЕНИЯ КНОПОК                  ########
'########################################################################
Private Sub ProjAn_AnalyzeButtonLayout(ByVal frm As Form)
    On Error GoTo ErrorHandler

    Debug.Print "--- РАСПОЛОЖЕНИЕ КНОПОК ---"

    Dim maxRight As Integer
    maxRight = 0
    Dim maxBottom As Integer
    maxBottom = 0
    Dim buttonList As String
    buttonList = ""

    Dim ctrl As Control
    For Each ctrl In frm.Controls
        If TypeName(ctrl) = "CommandButton" Then
            buttonList = buttonList & ctrl.Name & " ('" & ctrl.Caption & "'), "

            If ctrl.Left + ctrl.Width > maxRight Then
                maxRight = ctrl.Left + ctrl.Width
            End If

            If ctrl.Top + ctrl.Height > maxBottom Then
                maxBottom = ctrl.Top + ctrl.Height
            End If
        End If
    Next ctrl

    If buttonList <> "" Then
        buttonList = Left(buttonList, Len(buttonList) - 2)
    End If

    Debug.Print "Список кнопок: " & buttonList
    Debug.Print "Правая граница: " & maxRight
    Debug.Print "Нижняя граница: " & maxBottom
    Debug.Print "Рекомендуемая позиция для кнопки 'Демо':"
    Debug.Print "  - Left: " & maxRight + 100
    Debug.Print "  - Top: " & maxBottom - 100

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка анализа расположения кнопок: " & Err.description
End Sub

'########################################################################
'########            АНАЛИЗ ПРОЦЕДУР ФОРМЫ                       ########
'########################################################################
Private Sub ProjAn_AnalyzeFormProcedures()
    On Error GoTo ErrorHandler

    Debug.Print "--- ПРОЦЕДУРЫ ФОРМЫ ---"

    Dim criticalProcedures As Variant
    criticalProcedures = Array("BuildCalendar", "Form_Load", "Form_Open", _
                              "ApplyDayStyling", "LoadEventData", _
                              "cmdNextMonth_Click", "cmdPrevMonth_Click", _
                              "cmdToday_Click", "cmdExecutors_Click", _
                              "cmdThemes_Click", "cmdSearch_Click")

    Dim i As Integer
    For i = 0 To UBound(criticalProcedures)
        If ProjAn_ProcedureExistsSimple("f_daily_planner", CStr(criticalProcedures(i))) Then
            Debug.Print "? Процедура: " & criticalProcedures(i)
        Else
            Debug.Print "? Процедура отсутствует: " & criticalProcedures(i)
        End If
    Next i

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка анализа процедур: " & Err.description
End Sub

'########################################################################
'########            ПОЛНЫЙ АНАЛИЗ ВСЕХ ЭЛЕМЕНТОВ ФОРМЫ          ########
'########################################################################
Public Sub ProjAn_FullFormAnalysis()
    On Error GoTo ErrorHandler

    Debug.Print "=== ПОЛНЫЙ АНАЛИЗ ФОРМЫ f_daily_planner ==="

    If Not CurrentProject.allForms("f_daily_planner").IsLoaded Then
        DoCmd.OpenForm "f_daily_planner", acNormal, , , , acHidden
    End If

    Dim frm As Form
    Set frm = Forms!f_daily_planner

    Dim totalControls As Integer
    totalControls = 0
    Dim controlTypes As Collection
    Set controlTypes = New Collection

    ProjAn_AnalyzeAllControls frm, totalControls, controlTypes

    Debug.Print "ВСЕГО ЭЛЕМЕНТОВ: " & totalControls

    Dim i As Integer
    For i = 1 To controlTypes.count
        Debug.Print controlTypes(i)
    Next i

    ProjAn_AnalyzeCalendarStructure frm

    If Not CurrentProject.allForms("f_daily_planner").IsLoaded Then
        DoCmd.Close acForm, "f_daily_planner"
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка полного анализа: " & Err.description
    If Not CurrentProject.allForms("f_daily_planner").IsLoaded Then
        DoCmd.Close acForm, "f_daily_planner"
    End If
End Sub

'########################################################################
'########            АНАЛИЗ ВСЕХ КОНТРОЛОВ РЕКУРСИВНО            ########
'########################################################################
Private Sub ProjAn_AnalyzeAllControls(ByVal container As Object, ByRef totalCount As Integer, ByRef typesCol As Collection)
    On Error GoTo ErrorHandler

    Dim ctrl As Control
    For Each ctrl In container.Controls
        totalCount = totalCount + 1

        ProjAn_CountControlType typesCol, TypeName(ctrl)

        If TypeOf ctrl Is TabControl Or TypeOf ctrl Is Page Or _
           TypeOf ctrl Is Rectangle Or TypeOf ctrl Is OptionGroup Then
            ProjAn_AnalyzeAllControls ctrl, totalCount, typesCol
        End If

        If totalCount <= 100 Then
            Debug.Print totalCount & ". " & ctrl.Name & " | " & TypeName(ctrl) & _
                       " | " & ctrl.Left & "," & ctrl.Top & " | " & ctrl.Width & "x" & ctrl.Height
        End If
    Next ctrl

    If totalCount > 100 Then
        Debug.Print "... и еще " & (totalCount - 100) & " элементов"
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка анализа контрола " & container.Name & ": " & Err.description
End Sub

'########################################################################
'########            ПОДСЧЕТ ТИПОВ ЭЛЕМЕНТОВ                     ########
'########################################################################
Private Sub ProjAn_CountControlType(ByVal typesCol As Collection, ByVal controlType As String)
    On Error GoTo ErrorHandler

    Dim i As Integer
    Dim found As Boolean
    found = False

    For i = 1 To typesCol.count
        If InStr(typesCol(i), controlType) > 0 Then
            Dim parts() As String
            parts = Split(typesCol(i), ": ")
            Dim count As Integer
            count = CInt(parts(1)) + 1
            typesCol.Remove i
            typesCol.Add controlType & ": " & count, Before:=i
            found = True
            Exit For
        End If
    Next i

    If Not found Then
        typesCol.Add controlType & ": 1"
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка подсчета типа: " & Err.description
End Sub

'########################################################################
'########            АНАЛИЗ СТРУКТУРЫ КАЛЕНДАРЯ                  ########
'########################################################################
Private Sub ProjAn_AnalyzeCalendarStructure(ByVal frm As Form)
    On Error GoTo ErrorHandler

    Debug.Print "=== АНАЛИЗ СТРУКТУРЫ КАЛЕНДАРЯ ==="

    Dim dayControls As Integer
    dayControls = 0
    Dim eventControls As Integer
    eventControls = 0
    Dim ctrl As Control

    For Each ctrl In frm.Controls
        If TypeName(ctrl) = "Label" Then
            If InStr(ctrl.Name, "Day") > 0 Or InStr(ctrl.Name, "day") > 0 Then
                dayControls = dayControls + 1
            End If
        ElseIf TypeName(ctrl) = "TextBox" Then
            If InStr(ctrl.Name, "Event") > 0 Or InStr(ctrl.Name, "event") > 0 Then
                eventControls = eventControls + 1
            End If
        End If
    Next ctrl

    Debug.Print "Элементов дней: " & dayControls
    Debug.Print "Элементов событий: " & eventControls
    Debug.Print "Всего элементов календаря: " & (dayControls + eventControls)

    If dayControls >= 42 Then
        Debug.Print "? Календарь: полная сетка 6?7 дней"
    Else
        Debug.Print "? Календарь: неполная сетка (" & dayControls & " элементов)"
    End If

    Exit Sub

ErrorHandler:
    Debug.Print "Ошибка анализа календаря: " & Err.description
End Sub
