Option Compare Database
Option Explicit

'################################################################
'########     Оформление и действия rptEventInstances    ########
'################################################################

'################################################################
'########       Загрузка палитры в глобальный модуль     ########
'################################################################
Private Sub EnsureThemePaletteLoaded()
    On Error GoTo ExitSub

    If CurrentTheme_Back <> 0 Or HeaderTheme_Back <> 0 Then Exit Sub

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tbThemes WHERE IsActive = True", dbOpenSnapshot)

    If rs.EOF Then
        rs.Close
        Set rs = db.OpenRecordset("SELECT * FROM tbThemes ORDER BY ThemeID", dbOpenSnapshot)
    End If

    If Not rs.EOF Then
        Theme_WritePalette _
            rs!CurrentMonth_Text, rs!CurrentMonth_Back, rs!CurrentMonth_Border, _
            rs!OtherMonth_Text, rs!OtherMonth_Back, rs!OtherMonth_Border, _
            rs!Today_Back, rs!Today_Border, _
            rs!Header_Text, rs!Header_Back, rs!Header_Border, _
            rs!Form_Back
    End If

ExitSub:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub

'################################################################
'########   Секция: сброс темы Office, сплошной фон       ########
'########   (как в rptBirthdays — иначе «ломается» палитра) ########
'################################################################
Private Sub ApplySolidSectionBack(ByVal sec As Section, ByVal bg As Long)
    On Error Resume Next
    sec.BackThemeColorIndex = 0
    sec.AlternateBackThemeColorIndex = 0
    sec.backColor = bg
    sec.AlternateBackColor = bg
End Sub

'################################################################
'########   Вторая полоса зебры: слегка смешиваем с фоном формы
'################################################################
Private Function DetailZebraAlternateBack() As Long
    Dim r1 As Long, g1 As Long, b1 As Long
    Dim r2 As Long, g2 As Long, b2 As Long
    Dim c1 As Long, c2 As Long
    c1 = CurrentTheme_Back
    c2 = FormTheme_Back
    r1 = c1 And &HFF&: g1 = (c1 \ &H100&) And &HFF&: b1 = (c1 \ &H10000) And &HFF&
    r2 = c2 And &HFF&: g2 = (c2 \ &H100&) And &HFF&: b2 = (c2 \ &H10000) And &HFF&
    DetailZebraAlternateBack = RGB( _
        CLng(r1 * 0.72 + r2 * 0.28), _
        CLng(g1 * 0.72 + g2 * 0.28), _
        CLng(b1 * 0.72 + b2 * 0.28))
End Function

'################################################################
'########   Detail: чередование строк по текущей теме      ########
'########   Поля — прозрачные, чтобы был виден фон секции. ########
'################################################################
Private Sub ApplyDetailZebraFromTheme()
    On Error Resume Next

    Dim sec As Section
    Set sec = Me.Section(acDetail)
    sec.BackThemeColorIndex = 0
    sec.AlternateBackThemeColorIndex = 0
    sec.backColor = CurrentTheme_Back
    sec.AlternateBackColor = DetailZebraAlternateBack

    Me.fld_Number.BackThemeColorIndex = 0
    Me.fld_Number.BackStyle = 0
    Me.fld_Number.ForeThemeColorIndex = 0
    Me.fld_Number.ForeColor = CurrentTheme_Text
    Me.fld_Number.BorderThemeColorIndex = 0
    Me.fld_Number.borderColor = CurrentTheme_Border

    Me.fld_EventNote.BackThemeColorIndex = 0
    Me.fld_EventNote.BackStyle = 0
    Me.fld_EventNote.ForeThemeColorIndex = 0
    Me.fld_EventNote.ForeColor = CurrentTheme_Text
    Me.fld_EventNote.BorderThemeColorIndex = 0
    Me.fld_EventNote.borderColor = CurrentTheme_Border
End Sub

'################################################################
'########        Применение цветов к элементам отчета    ########
'################################################################
Private Sub ApplyReportTheme()
    On Error GoTo ExitSub

    EnsureThemePaletteLoaded

    ApplySolidSectionBack Me.Section(acHeader), HeaderTheme_Back
    ApplySolidSectionBack Me.Section(acPageHeader), HeaderTheme_Back
    ApplySolidSectionBack Me.Section(acGroupLevel1Header), FormTheme_Back

    ApplyDetailZebraFromTheme

    Me.lbl_Day.ForeColor = HeaderTheme_Text
    Me.lbl_Number.ForeColor = HeaderTheme_Text
    Me.lbl_EventNote.ForeColor = HeaderTheme_Text
    Me.fld_Completed.ForeColor = HeaderTheme_Text

ExitSub:
End Sub

'################################################################
'########   Публичный вход для применения темы отчета    ########
'################################################################
Public Sub ApplyThemeFromHost(Optional ByVal forceDataRefresh As Boolean = False)
    On Error GoTo ExitSub
    ApplyReportTheme
    If forceDataRefresh Then
        Me.Requery
    End If
ExitSub:
End Sub

'################################################################
'########         Показывать заголовок только для        ########
'########               выполненной группы               ########
'################################################################
Private Sub GroupHeading0_Format(Cancel As Integer, FormatCount As Integer)
    On Error Resume Next
    Me.Section("GroupHeading0").Visible = (Nz(Me!IsCompleted, 0) = 1)
End Sub

'################################################################
'########         DataArea (Detail): зебра видна         ########
'########  только если поля  остаются прозрачными        ########
'########     (как в rptBirthdays Format/Paint)          ########
'################################################################
Private Sub DataArea_Format(Cancel As Integer, FormatCount As Integer)
    On Error Resume Next
    Me.fld_Number.BackStyle = 0
    Me.fld_EventNote.BackStyle = 0
End Sub

Private Sub DataArea_Paint()
    On Error Resume Next
    Me.fld_Number.BackStyle = 0
    Me.fld_EventNote.BackStyle = 0
End Sub

'################################################################
'########          Открыть редактирование дня            ########
'########       по двойному клику в отчете панели        ########
'################################################################
Private Sub Detail_DblClick(Cancel As Integer)
    On Error GoTo ErrorHandler

    Dim clickDate As Date
    Dim executorFilter As String

    If Not IsDate(Me!EventDate) Then Exit Sub
    clickDate = DateValue(Me!EventDate)

    If CurrentProject.allForms("f_daily_planner").IsLoaded Then
        If Not IsNull(Forms!f_daily_planner!cboExecutorFilter.value) And Forms!f_daily_planner!cboExecutorFilter.value <> "" Then
            executorFilter = " AND ExecutorID = " & CLng(Forms!f_daily_planner!cboExecutorFilter.value)
        Else
            executorFilter = ""
        End If
    End If

    DoCmd.OpenForm "frmDayEvents"
    Forms!frmDayEvents.RecordSource = "SELECT * FROM tbEventInstances WHERE EventDate = " & _
                                      Format(clickDate, "\#mm\/dd\/yyyy\#") & executorFilter
    Forms!frmDayEvents.lblDate.Caption = Format(clickDate, "d mmmm yyyy ""г.""")
    Exit Sub

ErrorHandler:
    MsgBox "Ошибка открытия формы дня: " & Err.description, vbExclamation
End Sub

Private Sub fld_EventNote_DblClick(Cancel As Integer)
    Call Detail_DblClick(Cancel)
End Sub

Private Sub fld_Number_DblClick(Cancel As Integer)
    Call Detail_DblClick(Cancel)
End Sub

'################################################################
'########       Runtime RecordSource через TempVars      ########
'################################################################
Private Function BuildRuntimeRecordSource() As String
    BuildRuntimeRecordSource = _
        "SELECT InstanceID, EventDate, EventNote, CompletionDate, CompletionMark, " & _
        "IIf(IsNull([CompletionDate]),0,1) AS IsCompleted " & _
        "FROM tbEventInstances " & _
        "WHERE EventDate = TempVars!EventPanelDate " & _
        "AND (TempVars!EventPanelExecutorID Is Null OR ExecutorID = TempVars!EventPanelExecutorID) " & _
        "ORDER BY IIf(IsNull([CompletionDate]),0,1), InstanceID;"
End Function

Private Sub ApplyRuntimeRecordSource()
    On Error GoTo ExitSub
    Me.RecordSource = BuildRuntimeRecordSource()
ExitSub:
End Sub

'################################################################
'########          Жизненный цикл отчета                 ########
'################################################################
Private Sub Report_Open(Cancel As Integer)
    On Error Resume Next
    TempVars.Remove "EventPanelDate"
    TempVars.Remove "EventPanelExecutorID"
    TempVars.Add "EventPanelDate", Date
    TempVars.Add "EventPanelExecutorID", Null
    On Error GoTo 0

    ApplyRuntimeRecordSource
    ApplyReportTheme
End Sub

Private Sub Report_Activate()
    ApplyRuntimeRecordSource
    ApplyReportTheme
End Sub




