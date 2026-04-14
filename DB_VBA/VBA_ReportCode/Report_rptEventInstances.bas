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
'########        Применение цветов к элементам отчета    ########
'################################################################
Private Sub ApplyReportTheme()
    On Error GoTo ExitSub

    EnsureThemePaletteLoaded

    Me.Section(acHeader).backColor = HeaderTheme_Back
    Me.Section(acPageHeader).backColor = HeaderTheme_Back
    Me.Section(acDetail).backColor = FormTheme_Back
    Me.Section(acGroupLevel1Header).backColor = FormTheme_Back

    Me.lbl_Day.ForeColor = HeaderTheme_Text
    Me.lbl_Number.ForeColor = HeaderTheme_Text
    Me.lbl_EventNote.ForeColor = HeaderTheme_Text
    Me.lblCompleted.ForeColor = HeaderTheme_Text

    Me.fld_Number.backColor = CurrentTheme_Back
    Me.fld_Number.ForeColor = CurrentTheme_Text
    Me.fld_Number.borderColor = CurrentTheme_Border

    Me.fld_EventNote.backColor = CurrentTheme_Back
    Me.fld_EventNote.ForeColor = CurrentTheme_Text
    Me.fld_EventNote.borderColor = CurrentTheme_Border

ExitSub:
End Sub

'################################################################
'########   Публичный вход для применения темы отчета    ########
'################################################################
Public Sub ApplyThemeFromHost()
    On Error GoTo ExitSub
    ApplyReportTheme
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
'########          Открыть редактирование дня            ########
'########       по двойному клику в отчете панели        ########
'################################################################
Private Sub Detail_DblClick(Cancel As Integer)
    On Error GoTo ErrorHandler

    Dim clickDate As Date
    Dim executorFilter As String

    If Not IsDate(Me!EventDate) Then Exit Sub
    clickDate = DateValue(Me!EventDate)

    If CurrentProject.AllForms("f_daily_planner").IsLoaded Then
        If Not IsNull(Forms!f_daily_planner!cboExecutorFilter.Value) And Forms!f_daily_planner!cboExecutorFilter.Value <> "" Then
            executorFilter = " AND ExecutorID = " & CLng(Forms!f_daily_planner!cboExecutorFilter.Value)
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
    MsgBox "Ошибка открытия формы дня: " & Err.Description, vbExclamation
End Sub

Private Sub fld_EventNote_DblClick(Cancel As Integer)
    Call Detail_DblClick(Cancel)
End Sub

Private Sub fld_Number_DblClick(Cancel As Integer)
    Call Detail_DblClick(Cancel)
End Sub

'################################################################
'########        Runtime RecordSource через TempVars      ########
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
