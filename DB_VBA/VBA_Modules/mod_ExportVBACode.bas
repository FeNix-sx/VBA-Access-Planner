Option Compare Database
'################################################################
'########          ЭКСПОРТ КОДА ФОРМ И МОДУЛЕЙ           ########
'################################################################
' Примечание: перебор идёт по CurrentProject.allForms, AllReports и по TableDefs текущей БД —
' новые формы, отчёты и таблицы (в т.ч. tbBirthdays) попадают в экспорт без правок списков.
' Сохранённые запросы (QueryDefs) этим модулем не выгружаются; например qryBirthdaysForPanel
' создаётся в базе из modBirthdays (EnsureQryBirthdaysForPanel / EnsureBirthdaysPanelQueryAndReport).
'################################################################
Public Sub ExportCodeToFolder()
' Назначение: Запрашивает режим экспорта (1–5) и вызывает RunExport.
'             1 — всё (в т.ч. разметка .txt); 2 — контролы; 3 — только VBA форм/отчётов;
'             4 — стандартные модули; 5 — таблицы (JSON). Отмена/пусто — выход.
' Принцип:    InputBox в цикле до числа 1–5 или отмены; тело — в RunExport.
' Зависимости: Для режимов 1, 3, 4 при работе с модулями — VBE
'              (Microsoft Visual Basic for Applications Extensibility 5.3).
'################################################################
    Const PROC_NAME As String = "ExportCodeToFolder"
    On Error GoTo Err_Handler

    Dim sInput As String
    Dim exportMode As Long
    Dim v As Double

    Do
        sInput = InputBox( _
            "Режим экспорта в папку DB_VBA (введите число 1–5):" & vbCrLf & vbCrLf & _
            "1 — всё полностью (разметка форм/отчётов .txt + VBA + контролы + таблицы)" & vbCrLf & _
            "2 — только контролы (формы и отчёты)" & vbCrLf & _
            "3 — только VBA форм и отчётов (процедуры и события), без .txt разметки" & vbCrLf & _
            "4 — только стандартные модули" & vbCrLf & _
            "5 — только таблицы (JSON)" & vbCrLf & vbCrLf & _
            "Отмена или пустой ввод — выход без экспорта.", _
            "Экспорт VBA и метаданных", "")

        If Len(Trim$(sInput)) = 0 Then
            Debug.Print "[ExportCodeToFolder] Отмена или пустой ввод — экспорт не выполнен."
            GoTo Exit_Procedure
        End If

        If IsNumeric(sInput) Then
            v = Val(sInput)
            If v >= 1# And v <= 5# And Fix(v) = v Then
                exportMode = CLng(v)
                Exit Do
            End If
        End If

        MsgBox "Введите целое число от 1 до 5 либо нажмите «Отмена».", vbExclamation, "Экспорт"
    Loop

    RunExport exportMode

Exit_Procedure:
    Exit Sub

Err_Handler:
    Debug.Print "=== Ошибка в процедуре: " & PROC_NAME & " ==="
    Debug.Print "   Описание: " & Err.description
    Debug.Print "   Номер: " & Err.Number
    Debug.Print String(60, "-")
    MsgBox "Ошибка в процедуре " & PROC_NAME & ":" & vbCrLf & _
           Err.description & vbCrLf & "(номер: " & Err.Number & ")", vbCritical
    Resume Exit_Procedure
End Sub

'################################################################
'########     Выполнение экспорта по режиму (1–5)        ########
'################################################################
Private Sub RunExport(ByVal exportMode As Long)
' Назначение: Создаёт только нужные подкаталоги и выполняет блоки экспорта
'             в соответствии с режимом (см. ExportCodeToFolder).
' Принцип:    Режим 1 — в т.ч. SaveAsText разметки форм и отчётов в .txt (VBA_Forms,
'             VBA_Reports). Режим 3 — только сценарии VBE форм/отчётов (.bas, события
'             и процедуры), без SaveAsText и без стандартных модулей/классов.
'             Режим 4 — только стандартные модули (Type=1).
'################################################################
    Const PROC_NAME As String = "RunExport"
    On Error GoTo Err_Handler

    Const FORMS_SUBFOLDER As String = "VBA_Forms"
    Const REPORTS_SUBFOLDER As String = "VBA_Reports"
    Const FORMCONTROLS_SUBFOLDER As String = "VBA_FormControls"
    Const REPORTCONTROLS_SUBFOLDER As String = "VBA_ReportControls"
    Const FORMCODE_SUBFOLDER As String = "VBA_FormCode"
    Const REPORTCODE_SUBFOLDER As String = "VBA_ReportCode"
    Const MODULES_SUBFOLDER As String = "VBA_Modules"
    Const TABLES_SUBFOLDER As String = "VBA_Tables"

    Dim BASE_PATH As String
    Dim vbProj As Object
    Dim vbComp As Object
    Dim frm As AccessObject
    Dim rptObj As AccessObject
    Dim db As DAO.Database
    Dim tdef As DAO.TableDef
    Dim strFullPath As String
    Dim lngExportedCount As Long
    Dim lngFormsExported As Long
    Dim lngReportsExported As Long
    Dim lngFormControlsExported As Long
    Dim lngReportControlsExported As Long
    Dim lngFormCodeExported As Long
    Dim lngReportCodeExported As Long
    Dim lngModulesExported As Long
    Dim lngTablesExported As Long
    Dim needVbe As Boolean
    Dim sSummary As String
    Dim sLine As String

    Set db = CurrentDb
    BASE_PATH = CurrentProject.Path & "\DB_VBA"

    lngFormsExported = 0
    lngReportsExported = 0
    lngFormControlsExported = 0
    lngReportControlsExported = 0
    lngFormCodeExported = 0
    lngReportCodeExported = 0
    lngModulesExported = 0
    lngTablesExported = 0

    needVbe = (exportMode = 1 Or exportMode = 3 Or exportMode = 4)
    Set vbProj = Nothing
    If needVbe Then
        On Error Resume Next
        Set vbProj = Application.VBE.ActiveVBProject
        On Error GoTo Err_Handler
        If vbProj Is Nothing Then
            If exportMode = 1 Then
                MsgBox "Не удалось получить доступ к проекту VBA." & vbCrLf & _
                       "Модули не будут экспортированы; остальные шаги режима 1 выполняются.", vbExclamation
            Else
                MsgBox "Нет доступа к проекту VBA (режим " & CStr(exportMode) & ")." & vbCrLf & _
                       "Включите доверие к проекту VBA в параметрах Access." & vbCrLf & _
                       "Экспорт модулей для этого режима пропущен.", vbExclamation
            End If
        End If
    End If

    Debug.Print String(60, "=")
    Debug.Print "[RunExport] НАЧАЛО ЭКСПОРТА, режим: " & CStr(exportMode)
    Debug.Print "[RunExport] База данных: " & CurrentProject.Name
    Debug.Print "[RunExport] Путь экспорта: " & BASE_PATH
    Debug.Print String(60, "-")

    ' --- Базовая папка (всегда) ---
    If Not EnsureFolderExists(BASE_PATH) Then
        MsgBox "Не удалось создать папку: " & BASE_PATH, vbCritical
        GoTo Exit_Procedure
    End If

    ' =========================================================
    '  РАЗМЕТКА ОБЪЕКТОВ В .txt (SaveAsText) — только режим 1.
    '  Режим 3 сюда не заходит: там только модули VBA (.bas), не описание форм/отчётов.
    ' =========================================================
    If exportMode = 1 Then
        If Not EnsureFolderExists(BASE_PATH & "\" & FORMS_SUBFOLDER) Then
            MsgBox "Не удалось создать папку для форм", vbCritical
            GoTo Exit_Procedure
        End If

        Debug.Print "[RunExport] РАЗМЕТКА ФОРМ (.txt, SaveAsText):"
        Debug.Print String(40, "-")

        For Each frm In CurrentProject.allForms
            strFullPath = BASE_PATH & "\" & FORMS_SUBFOLDER & "\" & frm.Name & ".txt"
            Application.SaveAsText acForm, frm.Name, strFullPath
            ConvertFileToUtf8 strFullPath
            lngFormsExported = lngFormsExported + 1
            Debug.Print "[RunExport] Форма [" & lngFormsExported & "]: " & frm.Name
            Debug.Print "[RunExport]   Путь: " & strFullPath
            Debug.Print String(30, "-")
        Next frm

        If Not EnsureFolderExists(BASE_PATH & "\" & REPORTS_SUBFOLDER) Then
            MsgBox "Не удалось создать папку для отчётов", vbCritical
            GoTo Exit_Procedure
        End If

        Debug.Print "[RunExport] РАЗМЕТКА ОТЧЁТОВ (.txt, SaveAsText):"
        Debug.Print String(40, "-")

        For Each rptObj In CurrentProject.AllReports
            strFullPath = BASE_PATH & "\" & REPORTS_SUBFOLDER & "\" & rptObj.Name & ".txt"
            Application.SaveAsText acReport, rptObj.Name, strFullPath
            ConvertFileToUtf8 strFullPath
            lngReportsExported = lngReportsExported + 1
            Debug.Print "[RunExport] Отчёт [" & lngReportsExported & "]: " & rptObj.Name
            Debug.Print "[RunExport]   Путь: " & strFullPath
            Debug.Print String(30, "-")
        Next rptObj
    End If

    ' =========================================================
    '  КОНТРОЛЫ ФОРМ — режимы 1 и 2
    ' =========================================================
    If exportMode = 1 Or exportMode = 2 Then
        If Not EnsureFolderExists(BASE_PATH & "\" & FORMCONTROLS_SUBFOLDER) Then
            MsgBox "Не удалось создать папку для контролов форм", vbCritical
            GoTo Exit_Procedure
        End If

        Debug.Print "[RunExport] ЭКСПОРТ КОНТРОЛОВ ФОРМ:"
        Debug.Print String(40, "-")

        For Each frm In CurrentProject.allForms
            strFullPath = BASE_PATH & "\" & FORMCONTROLS_SUBFOLDER & "\" & frm.Name & "_controls.json"
            ExportFormControlsToJson frm.Name, strFullPath
            ConvertFileToUtf8 strFullPath
            lngFormControlsExported = lngFormControlsExported + 1
            Debug.Print "[RunExport] Контролы формы [" & lngFormControlsExported & "]: " & frm.Name
            Debug.Print "[RunExport]   Путь: " & strFullPath
            Debug.Print String(30, "-")
        Next frm
    End If

    ' =========================================================
    '  КОНТРОЛЫ ОТЧЁТОВ — режимы 1 и 2
    ' =========================================================
    If exportMode = 1 Or exportMode = 2 Then
        If Not EnsureFolderExists(BASE_PATH & "\" & REPORTCONTROLS_SUBFOLDER) Then
            MsgBox "Не удалось создать папку для контролов отчётов", vbCritical
            GoTo Exit_Procedure
        End If

        Debug.Print "[RunExport] ЭКСПОРТ КОНТРОЛОВ ОТЧЁТОВ:"
        Debug.Print String(40, "-")

        For Each rptObj In CurrentProject.AllReports
            strFullPath = BASE_PATH & "\" & REPORTCONTROLS_SUBFOLDER & "\" & rptObj.Name & "_controls.json"
            ExportReportControlsToJson rptObj.Name, strFullPath
            ConvertFileToUtf8 strFullPath
            lngReportControlsExported = lngReportControlsExported + 1
            Debug.Print "[RunExport] Контролы отчёта [" & lngReportControlsExported & "]: " & rptObj.Name
            Debug.Print "[RunExport]   Путь: " & strFullPath
            Debug.Print String(30, "-")
        Next rptObj
    End If

    ' =========================================================
    '  VBE: режим 1 — полный обход; 3 — только документы форм/отчётов (без классов);
    '       4 — StdModule
    ' =========================================================
    If Not vbProj Is Nothing Then
        If exportMode = 1 Then
            If Not EnsureFolderExists(BASE_PATH & "\" & MODULES_SUBFOLDER) Then
                MsgBox "Не удалось создать папку для модулей", vbCritical
                GoTo Exit_Procedure
            End If
            If Not EnsureFolderExists(BASE_PATH & "\" & FORMCODE_SUBFOLDER) Then
                MsgBox "Не удалось создать папку для кода форм", vbCritical
                GoTo Exit_Procedure
            End If
            If Not EnsureFolderExists(BASE_PATH & "\" & REPORTCODE_SUBFOLDER) Then
                MsgBox "Не удалось создать папку для кода отчётов", vbCritical
                GoTo Exit_Procedure
            End If

            Debug.Print "[RunExport] Компоненты VBE (Name, Type):"
            For Each vbComp In vbProj.VBComponents
                Debug.Print "  " & vbComp.Name & " ; Type=" & vbComp.Type
            Next vbComp
            Debug.Print String(40, "-")
            Debug.Print "[RunExport] ЭКСПОРТ МОДУЛЕЙ, КОДА ФОРМ И ОТЧЁТОВ:"
            Debug.Print String(40, "-")

            For Each vbComp In vbProj.VBComponents
                If vbComp.Type = 1 Then
                    strFullPath = BASE_PATH & "\" & MODULES_SUBFOLDER & "\" & vbComp.Name & ".bas"
                    vbComp.Export strFullPath
                    StripVBEHeaderFromFile strFullPath
                    ConvertFileToUtf8 strFullPath
                    lngModulesExported = lngModulesExported + 1
                    Debug.Print "[RunExport] Модуль [" & lngModulesExported & "]: " & vbComp.Name
                    Debug.Print "[RunExport]   Путь: " & strFullPath
                    Debug.Print String(30, "-")
                ElseIf vbComp.Type = 100 Then
                    If IsNameInAllForms(vbComp.Name) Then
                        strFullPath = BASE_PATH & "\" & FORMCODE_SUBFOLDER & "\" & vbComp.Name & ".bas"
                        vbComp.Export strFullPath
                        StripVBEHeaderFromFile strFullPath
                        ConvertFileToUtf8 strFullPath
                        lngFormCodeExported = lngFormCodeExported + 1
                        Debug.Print "[RunExport] Код формы [" & lngFormCodeExported & "]: " & vbComp.Name
                        Debug.Print "[RunExport]   Путь: " & strFullPath
                        Debug.Print String(30, "-")
                    ElseIf IsNameInAllReports(vbComp.Name) Then
                        strFullPath = BASE_PATH & "\" & REPORTCODE_SUBFOLDER & "\" & vbComp.Name & ".bas"
                        vbComp.Export strFullPath
                        StripVBEHeaderFromFile strFullPath
                        ConvertFileToUtf8 strFullPath
                        lngReportCodeExported = lngReportCodeExported + 1
                        Debug.Print "[RunExport] Код отчёта [" & lngReportCodeExported & "]: " & vbComp.Name
                        Debug.Print "[RunExport]   Путь: " & strFullPath
                        Debug.Print String(30, "-")
                    Else
                        strFullPath = BASE_PATH & "\" & FORMCODE_SUBFOLDER & "\" & vbComp.Name & ".bas"
                        vbComp.Export strFullPath
                        StripVBEHeaderFromFile strFullPath
                        ConvertFileToUtf8 strFullPath
                        lngFormCodeExported = lngFormCodeExported + 1
                        Debug.Print "[RunExport] Код документа (прочее) [" & lngFormCodeExported & "]: " & vbComp.Name & " (Type=100)"
                        Debug.Print "[RunExport]   Путь: " & strFullPath
                        Debug.Print String(30, "-")
                    End If
                Else
                    strFullPath = BASE_PATH & "\" & FORMCODE_SUBFOLDER & "\" & vbComp.Name & ".bas"
                    vbComp.Export strFullPath
                    StripVBEHeaderFromFile strFullPath
                    ConvertFileToUtf8 strFullPath
                    lngFormCodeExported = lngFormCodeExported + 1
                    Debug.Print "[RunExport] Код класса/прочего [" & lngFormCodeExported & "]: " & vbComp.Name & " (Type=" & vbComp.Type & ")"
                    Debug.Print "[RunExport]   Путь: " & strFullPath
                    Debug.Print String(30, "-")
                End If
            Next vbComp

        ElseIf exportMode = 3 Then
            If Not EnsureFolderExists(BASE_PATH & "\" & FORMCODE_SUBFOLDER) Then
                MsgBox "Не удалось создать папку для кода форм", vbCritical
                GoTo Exit_Procedure
            End If
            If Not EnsureFolderExists(BASE_PATH & "\" & REPORTCODE_SUBFOLDER) Then
                MsgBox "Не удалось создать папку для кода отчётов", vbCritical
                GoTo Exit_Procedure
            End If

            Debug.Print "[RunExport] ТОЛЬКО VBA ФОРМ И ОТЧЁТОВ (.bas: события и процедуры; без SaveAsText .txt; Type=100 + AllForms/AllReports; без классов):"
            Debug.Print String(40, "-")

            For Each vbComp In vbProj.VBComponents
                Select Case vbComp.Type
                    Case 100
                        If IsNameInAllForms(vbComp.Name) Then
                            strFullPath = BASE_PATH & "\" & FORMCODE_SUBFOLDER & "\" & vbComp.Name & ".bas"
                            vbComp.Export strFullPath
                            StripVBEHeaderFromFile strFullPath
                            ConvertFileToUtf8 strFullPath
                            lngFormCodeExported = lngFormCodeExported + 1
                            Debug.Print "[RunExport] Код формы [" & lngFormCodeExported & "]: " & vbComp.Name
                            Debug.Print "[RunExport]   Путь: " & strFullPath
                            Debug.Print String(30, "-")
                        ElseIf IsNameInAllReports(vbComp.Name) Then
                            strFullPath = BASE_PATH & "\" & REPORTCODE_SUBFOLDER & "\" & vbComp.Name & ".bas"
                            vbComp.Export strFullPath
                            StripVBEHeaderFromFile strFullPath
                            ConvertFileToUtf8 strFullPath
                            lngReportCodeExported = lngReportCodeExported + 1
                            Debug.Print "[RunExport] Код отчёта [" & lngReportCodeExported & "]: " & vbComp.Name
                            Debug.Print "[RunExport]   Путь: " & strFullPath
                            Debug.Print String(30, "-")
                        End If
                    Case Else
                        ' Режим 3: не экспортируем стандартные модули (1), классы (2), MS Forms и т.д.
                End Select
            Next vbComp

        ElseIf exportMode = 4 Then
            If Not EnsureFolderExists(BASE_PATH & "\" & MODULES_SUBFOLDER) Then
                MsgBox "Не удалось создать папку для модулей", vbCritical
                GoTo Exit_Procedure
            End If

            Debug.Print "[RunExport] ЭКСПОРТ СТАНДАРТНЫХ МОДУЛЕЙ:"
            Debug.Print String(40, "-")

            For Each vbComp In vbProj.VBComponents
                If vbComp.Type = 1 Then
                    strFullPath = BASE_PATH & "\" & MODULES_SUBFOLDER & "\" & vbComp.Name & ".bas"
                    vbComp.Export strFullPath
                    StripVBEHeaderFromFile strFullPath
                    ConvertFileToUtf8 strFullPath
                    lngModulesExported = lngModulesExported + 1
                    Debug.Print "[RunExport] Модуль [" & lngModulesExported & "]: " & vbComp.Name
                    Debug.Print "[RunExport]   Путь: " & strFullPath
                    Debug.Print String(30, "-")
                End If
            Next vbComp
        End If
    ElseIf needVbe Then
        Debug.Print "[RunExport] Проект VBA недоступен — блок экспорта модулей пропущен (режим " & CStr(exportMode) & ")"
    End If

    ' =========================================================
    '  ТАБЛИЦЫ JSON — режимы 1 и 5
    ' =========================================================
    If exportMode = 1 Or exportMode = 5 Then
        If Not EnsureFolderExists(BASE_PATH & "\" & TABLES_SUBFOLDER) Then
            MsgBox "Не удалось создать папку для таблиц", vbCritical
            GoTo Exit_Procedure
        End If

        Debug.Print "[RunExport] ЭКСПОРТ ТАБЛИЦ:"
        Debug.Print String(40, "-")

        For Each tdef In db.TableDefs
            If Left$(tdef.Name, 4) <> "MSys" And Left$(tdef.Name, 1) <> "~" Then
                strFullPath = BASE_PATH & "\" & TABLES_SUBFOLDER & "\" & tdef.Name & ".json"
                ExportTableDefToJson tdef, strFullPath
                ConvertFileToUtf8 strFullPath
                lngTablesExported = lngTablesExported + 1
                Debug.Print "[RunExport] Таблица [" & lngTablesExported & "]: " & tdef.Name
                Debug.Print "[RunExport]   Путь: " & strFullPath
                Debug.Print String(30, "-")
            End If
        Next tdef
    End If

    ' =========================================================
    '  Итог (счётчики и MsgBox только по выполненным блокам режима)
    ' =========================================================
    lngExportedCount = lngFormsExported + lngReportsExported + lngFormControlsExported + lngReportControlsExported + lngFormCodeExported + lngReportCodeExported + lngModulesExported + lngTablesExported

    sSummary = "Экспорт завершён (режим " & CStr(exportMode) & ")." & vbCrLf & vbCrLf

    Debug.Print String(60, "=")
    Debug.Print "[RunExport] ИТОГ, всего единиц: " & CStr(lngExportedCount)

    Select Case exportMode
        Case 1
            sLine = "Форм (разметка .txt): " & CStr(lngFormsExported)
            sSummary = sSummary & sLine & vbCrLf
            Debug.Print "[RunExport]   " & sLine
            sLine = "Отчётов (разметка .txt): " & CStr(lngReportsExported)
            sSummary = sSummary & sLine & vbCrLf
            Debug.Print "[RunExport]   " & sLine
            sLine = "Файлов контролов форм: " & CStr(lngFormControlsExported)
            sSummary = sSummary & sLine & vbCrLf
            Debug.Print "[RunExport]   " & sLine
            sLine = "Файлов контролов отчётов: " & CStr(lngReportControlsExported)
            sSummary = sSummary & sLine & vbCrLf
            Debug.Print "[RunExport]   " & sLine
            sLine = "VBA форм (события, .bas): " & CStr(lngFormCodeExported)
            sSummary = sSummary & sLine & vbCrLf
            Debug.Print "[RunExport]   " & sLine
            sLine = "VBA отчётов (события, .bas): " & CStr(lngReportCodeExported)
            sSummary = sSummary & sLine & vbCrLf
            Debug.Print "[RunExport]   " & sLine
            sLine = "Стандартных модулей: " & CStr(lngModulesExported)
            sSummary = sSummary & sLine & vbCrLf
            Debug.Print "[RunExport]   " & sLine
            sLine = "Таблиц (JSON): " & CStr(lngTablesExported)
            sSummary = sSummary & sLine & vbCrLf
            Debug.Print "[RunExport]   " & sLine
        Case 2
            sLine = "Файлов контролов форм: " & CStr(lngFormControlsExported)
            sSummary = sSummary & sLine & vbCrLf
            Debug.Print "[RunExport]   " & sLine
            sLine = "Файлов контролов отчётов: " & CStr(lngReportControlsExported)
            sSummary = sSummary & sLine & vbCrLf
            Debug.Print "[RunExport]   " & sLine
        Case 3
            sLine = "VBA форм (только скрипты, .bas): " & CStr(lngFormCodeExported)
            sSummary = sSummary & sLine & vbCrLf
            Debug.Print "[RunExport]   " & sLine
            sLine = "VBA отчётов (только скрипты, .bas): " & CStr(lngReportCodeExported)
            sSummary = sSummary & sLine & vbCrLf
            Debug.Print "[RunExport]   " & sLine
        Case 4
            sLine = "Стандартных модулей: " & CStr(lngModulesExported)
            sSummary = sSummary & sLine & vbCrLf
            Debug.Print "[RunExport]   " & sLine
        Case 5
            sLine = "Таблиц (JSON): " & CStr(lngTablesExported)
            sSummary = sSummary & sLine & vbCrLf
            Debug.Print "[RunExport]   " & sLine
    End Select

    Debug.Print String(60, "=")

    MsgBox sSummary, vbInformation

Exit_Procedure:
    Set vbComp = Nothing
    Set vbProj = Nothing
    Set frm = Nothing
    Set rptObj = Nothing
    Set db = Nothing
    Exit Sub

Err_Handler:
    Debug.Print "=== Ошибка в процедуре: " & PROC_NAME & " ==="
    Debug.Print "   Описание: " & Err.description
    Debug.Print "   Номер: " & Err.Number
    Debug.Print String(60, "-")
    MsgBox "Ошибка в процедуре " & PROC_NAME & ":" & vbCrLf & _
           Err.description & vbCrLf & "(номер: " & Err.Number & ")", vbCritical
    Resume Exit_Procedure
End Sub

'################################################################
'########         ПРОВЕРКА СУЩЕСТВОВАНИЯ ПАПКИ           ########
'################################################################
Private Function EnsureFolderExists(ByVal strFolderPath As String) As Boolean
' Назначение: Проверяет наличие папки на диске и создает её
'             (включая все вложенные), если она отсутствует.
' Принцип:    Использует объект FileSystemObject.
' Возврат:    True, если папка существует или успешно создана,
'             False в случае ошибки.
'################################################################
    Const PROC_NAME As String = "EnsureFolderExists"
    '################################################################
    On Error GoTo Err_Handler

    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(strFolderPath) Then
        fso.CreateFolder strFolderPath
        Debug.Print "[EnsureFolderExists]   Создана папка: " & strFolderPath
    End If

    EnsureFolderExists = True

Exit_Procedure:
    Set fso = Nothing
    Exit Function

Err_Handler:
    Debug.Print "=== Ошибка в процедуре: " & PROC_NAME & " ==="
    Debug.Print "  Папка: " & strFolderPath
    Debug.Print "  Описание: " & Err.description
    Debug.Print "  Номер: " & Err.Number
    Debug.Print String(60, "-")
    EnsureFolderExists = False
    Resume Exit_Procedure
End Function

'################################################################
'########      Имя компонента среди объектов AllForms    ########
'################################################################
Private Function IsNameInAllForms(ByVal compName As String) As Boolean
' Назначение: Соответствует ли VBComponent документа какой-либо форме в AllForms.
' Принцип:    В VBE модуль формы обычно называется «Form_» + имя объекта формы;
'             в AllForms — только имя объекта. Учитываем оба варианта (без учёта регистра).
' Возврат:    True, если найдено соответствие.
'################################################################
    Dim ao As AccessObject

    For Each ao In CurrentProject.allForms
        If StrComp(ao.Name, compName, vbTextCompare) = 0 Then
            IsNameInAllForms = True
            Exit Function
        End If
        If StrComp("Form_" & ao.Name, compName, vbTextCompare) = 0 Then
            IsNameInAllForms = True
            Exit Function
        End If
    Next ao
End Function

'################################################################
'########     Имя компонента среди объектов AllReports  ########
'################################################################
Private Function IsNameInAllReports(ByVal compName As String) As Boolean
' Назначение: Соответствует ли VBComponent документа какому-либо отчёту в AllReports.
' Принцип:    В VBE модуль отчёта обычно «Report_» + имя объекта отчёта;
'             в AllReports — только имя объекта. Учитываем оба варианта.
' Возврат:    True, если найдено соответствие.
'################################################################
    Dim ao As AccessObject

    For Each ao In CurrentProject.AllReports
        If StrComp(ao.Name, compName, vbTextCompare) = 0 Then
            IsNameInAllReports = True
            Exit Function
        End If
        If StrComp("Report_" & ao.Name, compName, vbTextCompare) = 0 Then
            IsNameInAllReports = True
            Exit Function
        End If
    Next ao
End Function

'################################################################
'########          ВЫВОД ВСЕХ ФОРМ ИЗ ТЕКУЩЕЙ БД         ########
'################################################################
Public Sub ShowAllAccessForms()
' Назначение: Выводит в Immediate окно список ВСЕХ форм,
'             существующих в текущей базе данных Access
' Принцип:    Использует CurrentProject.AllForms для получения
'             списка форм независимо от наличия кода VBA
'################################################################
    Const PROC_NAME As String = "ShowAllAccessForms"
    '################################################################
    On Error GoTo Err_Handler

    ' =========================================================
    '  Переменные
    ' =========================================================
    Dim frm As AccessObject
    Dim lngCount As Long
    Dim lngLoaded As Long
    Dim lngTabsCount As Long
    Dim i As Integer

    ' =========================================================
    '  Заголовок
    ' =========================================================
    Debug.Print String(60, "=")
    Debug.Print "ВСЕ ФОРМЫ В БАЗЕ ДАННЫХ: " & CurrentProject.Name
    Debug.Print String(60, "=")

    ' =========================================================
    '  Получаем количество форм
    ' =========================================================
    lngCount = CurrentProject.allForms.count

    If lngCount = 0 Then
        Debug.Print "ВНИМАНИЕ: В текущей базе данных нет ни одной формы!"
        Debug.Print String(60, "=")
        MsgBox "В базе нет форм!", vbExclamation
        GoTo Exit_Procedure
    End If

    Debug.Print "Всего форм в базе: " & lngCount
    Debug.Print String(60, "-")

    ' =========================================================
    '  Перебираем все формы
    ' =========================================================
    For i = 0 To lngCount - 1
        Set frm = CurrentProject.allForms(i)

        ' Основная информация о форме
        Debug.Print "Форма #" & (i + 1) & ": " & frm.Name

        ' Статус загрузки
        If frm.IsLoaded Then
            Debug.Print "  [ОТКРЫТА]"
            lngLoaded = lngLoaded + 1
        Else
            Debug.Print "  [ЗАКРЫТА]"
        End If

        ' Проверяем наличие "tabs" в имени (регистронезависимо)
        If InStr(1, LCase(frm.Name), "tabs") > 0 Then
            Debug.Print "  --> СОДЕРЖИТ 'tabs' ?"
            lngTabsCount = lngTabsCount + 1
        End If

        ' Дата создания/изменения (если доступна)
        On Error Resume Next
        Debug.Print "  Дата создания: " & frm.DateCreated
        Debug.Print "  Дата изменения: " & frm.DateModified
        On Error GoTo Err_Handler

        Debug.Print String(40, "-")
    Next i

    ' =========================================================
    '  Итоговая статистика
    ' =========================================================
    Debug.Print String(60, "=")
    Debug.Print "СТАТИСТИКА:"
    Debug.Print "  Всего форм: " & lngCount
    Debug.Print "  Открыто форм: " & lngLoaded
    Debug.Print "  Закрыто форм: " & (lngCount - lngLoaded)
    Debug.Print "  Форм с 'tabs' в имени: " & lngTabsCount

    If lngTabsCount > 0 Then
        Debug.Print String(60, "=")
        Debug.Print "НАЙДЕНЫ ФОРМЫ С 'tabs':"

        ' Второй проход только для форм с tabs
        For i = 0 To lngCount - 1
            Set frm = CurrentProject.allForms(i)
            If InStr(1, LCase(frm.Name), "tabs") > 0 Then
                Debug.Print "  - " & frm.Name
            End If
        Next i
    End If

    Debug.Print String(60, "=")

    ' =========================================================
    '  Сообщение пользователю
    ' =========================================================
    MsgBox "Найдено форм: " & lngCount & vbCrLf & _
           "Форм с 'tabs': " & lngTabsCount, vbInformation, "Результат"

Exit_Procedure:
    Set frm = Nothing
    Exit Sub

Err_Handler:
    Debug.Print "=== Ошибка в процедуре: " & PROC_NAME & " ==="
    Debug.Print "   Описание: " & Err.description
    Debug.Print "   Номер: " & Err.Number
    Debug.Print String(60, "-")
    Resume Exit_Procedure
End Sub

'################################################################
'########      Перекодировка файла в UTF-8 (из Win-1251) ########
'################################################################
Private Sub ConvertFileToUtf8(ByVal sPath As String)
' Назначение: Перекодирует существующий текстовый файл из Windows-1251
'             в UTF-8 по тому же пути (перезаписывая исходный файл).
' Принцип:    Использует два объекта ADODB.Stream: первый читает текст
'             в указанной исходной кодировке, второй записывает тот же
'             текст в целевой кодировке UTF-8.
' Зависимости: Требуется доступ к библиотеке ADODB (ранний или поздний binding).
'################################################################
    Const PROC_NAME As String = "ConvertFileToUtf8"
    On Error GoTo Err_Handler

    Dim stmIn As Object  ' ADODB.Stream
    Dim stmOut As Object ' ADODB.Stream
    Dim sText As String

    ' Читаем файл как Windows-1251
    Set stmIn = CreateObject("ADODB.Stream")
    With stmIn
        .Type = 2                  ' текст
        .Charset = "windows-1251"  ' текущая кодировка
        .Open
        .LoadFromFile sPath
        sText = .ReadText
        .Close
    End With

    ' Пишем файл как UTF-8 (перезаписываем тот же путь)
    Set stmOut = CreateObject("ADODB.Stream")
    With stmOut
        .Type = 2                  ' текст
        .Charset = "utf-8"         ' целевая кодировка
        .Open
        .WriteText sText
        .Position = 0
        .SaveToFile sPath, 2       ' adSaveCreateOverWrite
        .Close
    End With

Exit_Sub:
    On Error Resume Next
    Set stmIn = Nothing
    Set stmOut = Nothing
    Exit Sub

Err_Handler:
    Debug.Print "=== Ошибка в процедуре: " & PROC_NAME & " ==="
    Debug.Print "  Файл: " & sPath
    Debug.Print "  Описание: " & Err.description
    Debug.Print "  Номер: " & Err.Number
    Debug.Print String(60, "-")
    Resume Exit_Sub
End Sub

'################################################################
'########      Экспорт контролов формы в JSON-файл       ########
'################################################################
Private Sub ExportFormControlsToJson(ByVal formName As String, ByVal sPath As String)
    Const PROC_NAME As String = "ExportFormControlsToJson"
    On Error GoTo Err_Handler

    Dim f As Integer
    Dim ctl As Control
    Dim frm As Form
    Dim firstControl As Boolean

    ' Открываем форму в режиме конструктора скрытно
    DoCmd.OpenForm formName, acDesign, , , , acHidden
    Set frm = Forms(formName)

    f = FreeFile
    Open sPath For Output As #f

    ' Пишем JSON вручную
    Print #f, "{"
    Print #f, "  ""formName"": """ & JsonEscape(formName) & ""","
    On Error Resume Next
    Dim recSrc As String
    recSrc = Nz(frm.RecordSource, "")
    On Error GoTo Err_Handler
    Print #f, "  ""recordSource"": """ & JsonEscape(recSrc) & ""","
    Print #f, "  ""controls"": ["

    Dim ctrlSource As String
    Dim rowSource As String
    Dim ctrlTypeName As String
    Dim tabIdx As Long

    firstControl = True
    For Each ctl In frm.Controls
        ctrlSource = ""
        rowSource = ""
        tabIdx = 0

        ' Безопасно читаем свойства
        On Error Resume Next
        ctrlSource = Nz(ctl.ControlSource, "")
        rowSource = Nz(ctl.rowSource, "")
        tabIdx = ctl.TabIndex
        On Error GoTo Err_Handler

        ctrlTypeName = ControlTypeToName(ctl.controlType)

        If Not firstControl Then
            Print #f, ","
        End If
        firstControl = False

        Print #f, "    {"
        Print #f, "      ""name"": """ & JsonEscape(ctl.Name) & ""","
        Print #f, "      ""type"": " & ctl.controlType & ","
        Print #f, "      ""controlTypeName"": """ & JsonEscape(ctrlTypeName) & ""","
        Print #f, "      ""controlSource"": """ & JsonEscape(ctrlSource) & ""","
        Print #f, "      ""rowSource"": """ & JsonEscape(rowSource) & ""","
        Print #f, "      ""tabIndex"": " & tabIdx
        Print #f, "    }"
    Next ctl

    Print #f, ""
    Print #f, "  ]"
    Print #f, "}"

Exit_Sub_Controls:
    On Error Resume Next
    If f <> 0 Then Close #f
    If Not frm Is Nothing Then
        DoCmd.Close acForm, formName, acSaveNo
        Set frm = Nothing
    End If
    Set ctl = Nothing
    Exit Sub

Err_Handler:
    Debug.Print "=== Ошибка в процедуре: " & PROC_NAME & " ==="
    Debug.Print "  Форма: " & formName
    Debug.Print "  Файл: " & sPath
    Debug.Print "  Описание: " & Err.description
    Debug.Print "  Номер: " & Err.Number
    Debug.Print String(60, "-")
    Resume Exit_Sub_Controls
End Sub

'################################################################
'########     Экспорт контролов отчёта в JSON-файл      ########
'################################################################
Private Sub ExportReportControlsToJson(ByVal reportName As String, ByVal sPath As String)
' Назначение: Выгружает метаданные контролов отчёта в JSON (как у форм).
' Принцип:    Открытие в конструкторе скрытно, обход Controls, ручная сборка JSON.
'################################################################
    Const PROC_NAME As String = "ExportReportControlsToJson"
    On Error GoTo Err_Handler

    Dim f As Integer
    Dim ctl As Control
    Dim rpt As Report
    Dim firstControl As Boolean

    DoCmd.OpenReport reportName, acDesign, , , acHidden
    Set rpt = Reports(reportName)

    f = FreeFile
    Open sPath For Output As #f

    Print #f, "{"
    Print #f, "  ""reportName"": """ & JsonEscape(reportName) & ""","
    On Error Resume Next
    Dim recSrc As String
    recSrc = Nz(rpt.RecordSource, "")
    On Error GoTo Err_Handler
    Print #f, "  ""recordSource"": """ & JsonEscape(recSrc) & ""","
    Print #f, "  ""controls"": ["

    Dim ctrlSource As String
    Dim rowSource As String
    Dim ctrlTypeName As String
    Dim tabIdx As Long
    Dim sOnClick As String

    firstControl = True
    For Each ctl In rpt.Controls
        ctrlSource = ""
        rowSource = ""
        tabIdx = 0
        sOnClick = ""

        On Error Resume Next
        ctrlSource = Nz(ctl.ControlSource, "")
        rowSource = Nz(ctl.rowSource, "")
        tabIdx = ctl.TabIndex
        sOnClick = Nz(ctl.OnClick, "")
        On Error GoTo Err_Handler

        ctrlTypeName = ControlTypeToName(ctl.controlType)

        If Not firstControl Then
            Print #f, ","
        End If
        firstControl = False

        Print #f, "    {"
        Print #f, "      ""name"": """ & JsonEscape(ctl.Name) & ""","
        Print #f, "      ""type"": " & ctl.controlType & ","
        Print #f, "      ""controlTypeName"": """ & JsonEscape(ctrlTypeName) & ""","
        Print #f, "      ""controlSource"": """ & JsonEscape(ctrlSource) & ""","
        Print #f, "      ""rowSource"": """ & JsonEscape(rowSource) & ""","
        Print #f, "      ""tabIndex"": " & tabIdx
        If Len(sOnClick) > 0 Then
            Print #f, ","
            Print #f, "      ""onClick"": """ & JsonEscape(sOnClick) & """"
        End If
        Print #f, "    }"
    Next ctl

    Print #f, ""
    Print #f, "  ]"
    Print #f, "}"

Exit_Sub_ReportControls:
    On Error Resume Next
    If f <> 0 Then Close #f
    If Not rpt Is Nothing Then
        DoCmd.Close acReport, reportName, acSaveNo
        Set rpt = Nothing
    End If
    Set ctl = Nothing
    Exit Sub

Err_Handler:
    Debug.Print "=== Ошибка в процедуре: " & PROC_NAME & " ==="
    Debug.Print "  Отчёт: " & reportName
    Debug.Print "  Файл: " & sPath
    Debug.Print "  Описание: " & Err.description
    Debug.Print "  Номер: " & Err.Number
    Debug.Print String(60, "-")
    Resume Exit_Sub_ReportControls
End Sub

'################################################################
'########      Экспорт структуры таблицы в JSON-файл     ########
'################################################################
Private Sub ExportTableDefToJson(ByVal tdef As DAO.TableDef, ByVal sPath As String)
' Назначение: Выгружает структуру таблицы DAO.TableDef (поля и индексы)
'             в JSON-файл для анализа и версионирования.
' Принцип:    Ручная сборка JSON (аналогично ExportFormControlsToJson).
' Зависимости: DAO.TableDef, DAO.Field, DAO.Index.
'################################################################
    Const PROC_NAME As String = "ExportTableDefToJson"
    On Error GoTo Err_Handler

    Dim f As Integer
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim idxFld As DAO.Field
    Dim firstField As Boolean
    Dim firstIndex As Boolean
    Dim firstIndexField As Boolean

    f = FreeFile
    Open sPath For Output As #f

    Print #f, "{"
    Print #f, "  ""tableName"": """ & JsonEscape(tdef.Name) & ""","
    Print #f, "  ""fields"": ["

    firstField = True
    For Each fld In tdef.Fields
        If Not firstField Then
            Print #f, ","
        End If
        firstField = False

        Print #f, "    {"
        Print #f, "      ""name"": """ & JsonEscape(fld.Name) & ""","
        Print #f, "      ""type"": " & fld.Type & ","
        Print #f, "      ""size"": " & fld.Size & ","
        Print #f, "      ""required"": " & JsonBool(fld.Required) & ","
        Print #f, "      ""attributes"": " & fld.Attributes
        Print #f, "    }"
    Next fld

    Print #f, ""
    Print #f, "  ],"
    Print #f, "  ""indexes"": ["

    firstIndex = True
    For Each idx In tdef.Indexes
        If Not firstIndex Then
            Print #f, ","
        End If
        firstIndex = False

        Print #f, "    {"
        Print #f, "      ""name"": """ & JsonEscape(idx.Name) & ""","
        Print #f, "      ""primary"": " & JsonBool(idx.Primary) & ","
        Print #f, "      ""unique"": " & JsonBool(idx.Unique) & ","
        Print #f, "      ""fields"": ["

        firstIndexField = True
        For Each idxFld In idx.Fields
            If Not firstIndexField Then
                Print #f, ","
            End If
            firstIndexField = False
            Print #f, "        """ & JsonEscape(idxFld.Name) & """"
        Next idxFld

        Print #f, ""
        Print #f, "      ]"
        Print #f, "    }"
    Next idx

    Print #f, ""
    Print #f, "  ]"
    Print #f, "}"

Exit_Sub:
    On Error Resume Next
    If f <> 0 Then Close #f
    Set fld = Nothing
    Set idx = Nothing
    Set idxFld = Nothing
    Exit Sub

Err_Handler:
    Debug.Print "=== Ошибка в процедуре: " & PROC_NAME & " ==="
    Debug.Print "  Таблица: " & tdef.Name
    Debug.Print "  Файл: " & sPath
    Debug.Print "  Описание: " & Err.description
    Debug.Print "  Номер: " & Err.Number
    Debug.Print String(60, "-")
    Resume Exit_Sub
End Sub

'################################################################
'########        Удаление служебной шапки VBE из .bas    ########
'################################################################
Private Sub StripVBEHeaderFromFile(ByVal sPath As String)
    Const PROC_NAME As String = "StripVBEHeaderFromFile"
    On Error GoTo Err_Handler

    Dim fIn As Integer
    Dim fOut As Integer
    Dim tempPath As String
    Dim lineText As String
    Dim inVersionBlock As Boolean

    tempPath = sPath & ".tmp"

    fIn = FreeFile
    Open sPath For Input As #fIn
    fOut = FreeFile + 1
    Open tempPath For Output As #fOut

    inVersionBlock = False

    Do While Not EOF(fIn)
        Line Input #fIn, lineText

        ' Блок VERSION ... END в начале класса
        If Not inVersionBlock And Left$(Trim$(lineText), 7) = "VERSION" Then
            inVersionBlock = True
            ' пропускаем строку VERSION и всё до END
            GoTo ContinueLoop
        End If

        If inVersionBlock Then
            If Trim$(lineText) = "END" Then
                inVersionBlock = False
            End If
            GoTo ContinueLoop
        End If

        ' Строки Attribute ... тоже убираем
        If Left$(Trim$(lineText), 9) = "Attribute" Then
            GoTo ContinueLoop
        End If

        Print #fOut, lineText

ContinueLoop:
    Loop

    Close #fIn
    Close #fOut

    ' Заменяем оригинал очищенным файлом
    Kill sPath
    Name tempPath As sPath

Exit_Sub_Strip:
    On Error Resume Next
    If fIn <> 0 Then Close #fIn
    If fOut <> 0 Then Close #fOut
    Exit Sub

Err_Handler:
    Debug.Print "=== Ошибка в процедуре: " & PROC_NAME & " ==="
    Debug.Print "  Файл: " & sPath
    Debug.Print "  Описание: " & Err.description
    Debug.Print "  Номер: " & Err.Number
    Debug.Print String(60, "-")
    Resume Exit_Sub_Strip
End Sub

'################################################################
'########    Вспомогательная функция экранирования JSON  ########
'################################################################
Private Function JsonEscape(ByVal value As String) As String
    Dim s As String
    s = value
    s = Replace$(s, "\", "\\")
    s = Replace$(s, """", "\""")
    s = Replace$(s, vbCrLf, "\n")
    s = Replace$(s, vbCr, "\n")
    s = Replace$(s, vbLf, "\n")
    JsonEscape = s
End Function

'################################################################
'########        Литералы true/false для JSON            ########
'################################################################
Private Function JsonBool(ByVal b As Boolean) As String
    If b Then
        JsonBool = "true"
    Else
        JsonBool = "false"
    End If
End Function

'################################################################
'########   Читаемое имя типа контрола по ControlType    ########
'################################################################
Private Function ControlTypeToName(ByVal controlType As Integer) As String
    Select Case controlType
        Case 106: ControlTypeToName = "Label"
        Case 109: ControlTypeToName = "TextBox"
        Case 111: ControlTypeToName = "ComboBox"
        Case 112: ControlTypeToName = "ListBox"
        Case 122: ControlTypeToName = "CommandButton"
        Case 113: ControlTypeToName = "CheckBox"
        Case 114: ControlTypeToName = "OptionButton"
        Case 115: ControlTypeToName = "ToggleButton"
        Case 118: ControlTypeToName = "Subform"
        Case Else
            ControlTypeToName = "Type" & CStr(controlType)
    End Select
End Function




