Option Compare Database
Option Explicit

'################################################################
'########   rptBirthdays — оформление по палитре темы    ########
'################################################################
' Назначение: Подстановка цветов из tbThemes (modThemeColors) в секции
'             и поля отчёта в зависимости от даты события относительно сегодня.
' Принцип:    Format — предпросмотр/печать; Paint — отрисовка в режиме отчёта.
'             Секция Detail: в конструкторе «При окрашивании» = процедура события.
'################################################################

'################################################################
'########          ЗАГРУЗКА ПАЛИТРЫ ИЗ tbThemes            ########
'################################################################
Private Sub LoadPaletteFromDB()
' Назначение: Читает активную тему из tbThemes и записывает палитру через Theme_WritePalette.
' Принцип:    Сначала IsActive = True; если пусто — первая запись по ThemeID.
'================================================================
    On Error GoTo CleanUp

    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    Set rs = Nothing
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

CleanUp:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
End Sub

'################################################################
'########        ФОН СЕКЦИИ ОТЧЁТА (ТЕМА / RGB)          ########
'################################################################
Private Sub ApplyThemedSectionBack(ByVal sec As Section, ByVal bg As Long)
' Назначение: Задаёт однотонный фон секции, сбрасывая привязку к теме Office.
' Принцип:    BackThemeColorIndex = 0, затем backColor / AlternateBackColor.
'================================================================
    On Error Resume Next
    sec.BackThemeColorIndex = 0
    sec.AlternateBackThemeColorIndex = 0
    sec.backColor = bg
    sec.AlternateBackColor = bg
End Sub

'################################################################
'########           ФОН ЭЛЕМЕНТА УПРАВЛЕНИЯ (RGB)        ########
'################################################################
Private Sub ApplyThemedCtlBack(ByVal ctl As Control, ByVal bg As Long)
' Назначение: Непрозрачный фон поля по заданному RGB.
'================================================================
    On Error Resume Next
    ctl.BackThemeColorIndex = 0
    ctl.BackStyle = 1
    ctl.backColor = bg
End Sub

'################################################################
'########              ЦВЕТ ТЕКСТА ЭЛЕМЕНТА (RGB)        ########
'################################################################
Private Sub ApplyCtlForeRgb(ByVal ctl As Control, ByVal fc As Long)
' Назначение: Цвет переднего плана без темы Office (Tint 100%).
'================================================================
    On Error Resume Next
    ctl.ForeThemeColorIndex = 0
    ctl.ForeTint = 100
    ctl.ForeColor = fc
End Sub

'################################################################
'########   ЗАГОЛОВОК ОТЧЁТА (fld_Title) — ПАЛИТРА HEADER ########
'################################################################
Private Sub ApplyFldTitleHeaderTheme()
' Назначение: Оформляет fld_Title как кнопки шапки на f_daily_planner (Header).
' Принцип:    Фон/текст/рамка из HeaderTheme_*; не вызывается из ApplyGroupHeaderColors.
'================================================================
    If CurrentTheme_Back = 0 And OtherTheme_Back = 0 Then LoadPaletteFromDB
    ApplyThemedCtlBack Me.fld_Title, HeaderTheme_Back
    ApplyCtlForeRgb Me.fld_Title, HeaderTheme_Text
    Me.fld_Title.borderColor = HeaderTheme_Border
    Me.fld_Title.borderWidth = 1
End Sub

'################################################################
'########   ЗАГОЛОВОК ГРУППЫ ПО ДАТЕ (ЦВЕТА ПО dayDiff)  ########
'################################################################
Private Sub ApplyGroupHeaderColors(ByVal dayDiff As Long)
' Назначение: Оформляет GroupHeader0 и fld_OccurrenceDate: «другой месяц», «сегодня», текущий месяц.
' Принцип:    dayDiff = DateDiff("d", дата_события, Date): >0 прошлое, 0 сегодня, <0 будущее в месяце.
'================================================================
    If CurrentTheme_Back = 0 And OtherTheme_Back = 0 Then LoadPaletteFromDB
    Dim secBg As Long

    Me.fld_OccurrenceDate.BackStyle = 1

    ' ---------------------------------------------------------
    ' Дата раньше сегодняшней: секция и поле даты — Other, рамка 1 px, цвет Other_Text.
    ' ---------------------------------------------------------
    If dayDiff > 0 Then
        secBg = OtherTheme_Back
        ApplyThemedSectionBack Me.GroupHeader0, secBg
        ApplyThemedCtlBack Me.fld_OccurrenceDate, OtherTheme_Back
        Me.fld_OccurrenceDate.borderColor = OtherTheme_Border
        Me.fld_OccurrenceDate.borderWidth = 1
        ApplyCtlForeRgb Me.fld_OccurrenceDate, OtherTheme_Text
    ' ---------------------------------------------------------
    ' Сегодня: фон секции Today; поле даты — фон текущего месяца, рамка Today, толщина 2 px.
    ' ---------------------------------------------------------
    ElseIf dayDiff = 0 Then
        secBg = TodayTheme_Back
        ApplyThemedSectionBack Me.GroupHeader0, secBg
        ApplyThemedCtlBack Me.fld_OccurrenceDate, CurrentTheme_Back
        Me.fld_OccurrenceDate.borderColor = TodayTheme_Border
        Me.fld_OccurrenceDate.borderWidth = 2
        ApplyCtlForeRgb Me.fld_OccurrenceDate, CurrentTheme_Text
    ' ---------------------------------------------------------
    ' Дата позже сегодняшней: Current, рамка текущего месяца, 1 px.
    ' ---------------------------------------------------------
    Else
        secBg = CurrentTheme_Back
        ApplyThemedSectionBack Me.GroupHeader0, secBg
        ApplyThemedCtlBack Me.fld_OccurrenceDate, CurrentTheme_Back
        Me.fld_OccurrenceDate.borderColor = CurrentTheme_Border
        Me.fld_OccurrenceDate.borderWidth = 1
        ApplyCtlForeRgb Me.fld_OccurrenceDate, CurrentTheme_Text
    End If
End Sub

'################################################################
'########        СБРОС ТЕМЫ ШРИФТА В СЕКЦИИ DETAIL       ########
'################################################################
Private Sub ClearDetailFontTheme()
' Назначение: Убирает ThemeFontIndex у строки персоны и примечаний для предсказуемого вида.
'================================================================
    On Error Resume Next
    Me.fld_PersonLine.ThemeFontIndex = 0
    Me.fld_Notes.ThemeFontIndex = 0
End Sub

'################################################################
'########       СЕКЦИЯ DETAIL — ЦВЕТА ПО dayDiff         ########
'################################################################
Private Sub ApplyDetailColors(ByVal dayDiff As Long)
' Назначение: Фон секции Detail — как фон формы (FormTheme_Back); фон полей строки по dayDiff.
' Принцип:    Та же шкала dayDiff, что и в заголовке группы.
'================================================================
    If CurrentTheme_Back = 0 And OtherTheme_Back = 0 Then LoadPaletteFromDB
    ClearDetailFontTheme

    Me.fld_PersonLine.FontItalic = False
    Me.fld_Notes.FontItalic = False

    ' ---------------------------------------------------------
    ' Дата события раньше сегодняшней: «прошлое» в строке —
    ' палитра Other, серый текст, курсив.
    ' ---------------------------------------------------------
    If dayDiff > 0 Then
        ApplyThemedSectionBack Me.Detail, FormTheme_Back
        ApplyThemedCtlBack Me.fld_PersonLine, OtherTheme_Back
        ApplyThemedCtlBack Me.fld_Notes, OtherTheme_Back
        ApplyCtlForeRgb Me.fld_PersonLine, RGB(128, 128, 128)
        ApplyCtlForeRgb Me.fld_Notes, RGB(128, 128, 128)
        Me.fld_PersonLine.FontItalic = True
        Me.fld_Notes.FontItalic = True
        Me.fld_PersonLine.FontBold = False
        Me.fld_Notes.FontBold = False
    ' ---------------------------------------------------------
    ' Событие сегодня: фон Today, обычный цвет текста текущего месяца.
    ' ---------------------------------------------------------
    ElseIf dayDiff = 0 Then
        ApplyThemedSectionBack Me.Detail, FormTheme_Back
        ApplyThemedCtlBack Me.fld_PersonLine, TodayTheme_Back
        ApplyThemedCtlBack Me.fld_Notes, TodayTheme_Back
        ApplyCtlForeRgb Me.fld_PersonLine, CurrentTheme_Text
        ApplyCtlForeRgb Me.fld_Notes, CurrentTheme_Text
        Me.fld_PersonLine.FontBold = True
        Me.fld_Notes.FontBold = True
    ' ---------------------------------------------------------
    ' Дата позже сегодняшней: текущий месяц (Current), без курсива.
    ' ---------------------------------------------------------
    Else
        ApplyThemedSectionBack Me.Detail, FormTheme_Back
        ApplyThemedCtlBack Me.fld_PersonLine, CurrentTheme_Back
        ApplyThemedCtlBack Me.fld_Notes, CurrentTheme_Back
        ApplyCtlForeRgb Me.fld_PersonLine, CurrentTheme_Text
        ApplyCtlForeRgb Me.fld_Notes, CurrentTheme_Text
        Me.fld_PersonLine.FontBold = False
        Me.fld_Notes.FontBold = False
    End If

End Sub

'################################################################
'########             ОТЧЁТ — СОБЫТИЕ OPEN               ########
'################################################################
Private Sub Report_Open(Cancel As Integer)
' Назначение: Загрузка палитры при открытии отчёта.
'================================================================
    LoadPaletteFromDB
    ApplyFldTitleHeaderTheme
End Sub

'################################################################
'########            ОТЧЁТ — СОБЫТИЕ ACTIVATE            ########
'################################################################
Private Sub Report_Activate()
' Назначение: Повторная подгрузка палитры, если глобальные цвета ещё нулевые.
'================================================================
    If CurrentTheme_Back = 0 And OtherTheme_Back = 0 Then LoadPaletteFromDB
    ApplyFldTitleHeaderTheme
End Sub

'################################################################
'########      ПУБЛИЧНЫЙ ВХОД ДЛЯ ХОСТ-ФОРМЫ             ########
'################################################################
Public Sub ApplyThemeFromHost()
' Назначение: Немедленно применяет текущую палитру к статичному заголовку отчёта.
'================================================================
    On Error GoTo ExitSub
    If CurrentTheme_Back = 0 And OtherTheme_Back = 0 Then LoadPaletteFromDB
    ApplyFldTitleHeaderTheme
    Me.Requery
ExitSub:
End Sub

'################################################################
'########         GROUPHEADER0 — СОБЫТИЕ FORMAT          ########
'################################################################
Private Sub GroupHeader0_Format(Cancel As Integer, FormatCount As Integer)
' Назначение: Раскраска заголовка группы при форматировании (печать/предпросмотр).
'================================================================
    Dim occ As Variant
    occ = Me.OccurrenceDate
    If Not IsDate(occ) Then Exit Sub

    Dim dayDiff As Long
    dayDiff = DateDiff("d", DateValue(occ), Date)
    ApplyGroupHeaderColors dayDiff
End Sub

'################################################################
'########          GROUPHEADER0 — СОБЫТИЕ PAINT          ########
'################################################################
Private Sub GroupHeader0_Paint()
' Назначение: Раскраска заголовка группы при отрисовке в режиме отчёта.
'================================================================
    Dim occ As Variant
    Dim dayDiff As Long

    occ = Me.OccurrenceDate
    If Not IsDate(occ) Then Exit Sub
    dayDiff = DateDiff("d", DateValue(occ), Date)
    ApplyGroupHeaderColors dayDiff
End Sub

'################################################################
'########            DETAIL — СОБЫТИЕ FORMAT             ########
'################################################################
Private Sub Detail_Format(Cancel As Integer, FormatCount As Integer)
' Назначение: Цвета строки детали при форматировании.
'================================================================
    Dim occ As Variant
    occ = Me.OccurrenceDate
    If Not IsDate(occ) Then Exit Sub

    Dim dayDiff As Long
    dayDiff = DateDiff("d", DateValue(occ), Date)
    ApplyDetailColors dayDiff
End Sub

'################################################################
'########             DETAIL — СОБЫТИЕ PAINT             ########
'################################################################
Private Sub Detail_Paint()
' Назначение: Цвета строки детали при отрисовке; дата из fld_OccurrenceDate (как в конструкторе).
'================================================================
    Dim occ As Variant
    Dim dayDiff As Long

    occ = Me.fld_OccurrenceDate
    If Not IsDate(occ) Then Exit Sub
    dayDiff = DateDiff("d", DateValue(occ), Date)
    ApplyDetailColors dayDiff
End Sub





