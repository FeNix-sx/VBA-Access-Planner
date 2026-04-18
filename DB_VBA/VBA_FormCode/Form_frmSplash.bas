Option Compare Database
Option Explicit

'################################################################
'########     Заставка перед календарём (frmSplash)      ########
'################################################################
' В форме Access должны быть (имена обязательны):
'   lbl_Title         — заголовок (текст в Form_Load)
'   lbl_expand_1 … 7  — расшифровка по буквам; мельче шрифтом (FontSize = SPLASH_EXPAND_FONT_PT)
'   lbl_Version       — пустая в макете; набор текста «версия 2.0»
'   lbl_ProgressFill  — подложка-полоска (BackStyle=Обычный, узкая высота);
'                       ширина нарастает по таймеру до SPLASH_PROGRESS_MAX_TWIPS
'
' Рекомендуемые свойства формы: всплывающая, без области выделения,
' без кнопок навигации, автоматическое центрирование, тонкая рамка или «Область данных».
' Цвета: перед OpenForm вызывается Theme_LoadPaletteFromDatabase; здесь — только отображение.
'
' Настройки анимации (плавность / скорость):
'   SPLASH_TYPE_MS — интервал Form_Timer (мс); меньше = быстрее опрос, плавнее глазу.
'   SPLASH_PROGRESS_TOTAL_TICKS — сколько тиков длится полоса после набора версии;
'       больше значение = мельче шаг ширины, без «ступенек». Удобно кратно 8.
'   SPLASH_EXPAND_REVEAL_FIRST_TICK / SPLASH_EXPAND_REVEAL_TICK_STEP — на каких тиках
'       появляются строки расшифровки (по умолчанию 1,3,5,7,9,11,13; чётные тики только полоса).
'################################################################

Private Const SPLASH_TYPE_MS As Long = 40
Private Const SPLASH_VERSION_TEXT As String = "версия 2.0"
Private Const SPLASH_PROGRESS_MAX_TWIPS As Long = 5200
Private Const SPLASH_EXPAND_FONT_PT As Single = 10
Private Const SPLASH_PROGRESS_TOTAL_TICKS As Long = 32
Private Const SPLASH_EXPAND_REVEAL_FIRST_TICK As Long = 1
Private Const SPLASH_EXPAND_REVEAL_TICK_STEP As Long = 2

Private mVerPos As Long
Private mPhase As Integer
Private mProgressW As Long
Private mProgressTick As Long

'################################################################
'########        Полоска прогресса — конец шкалы        ########
'################################################################
Private Sub SplashProgressSnapToFull()
' Назначение: Доводит ширину полоски до SPLASH_PROGRESS_MAX_TWIPS перед закрытием.
' Принцип:    Последний кадр перед Close, чтобы не осталось «недобора» из-за округления twips.
'================================================================
    mProgressW = SPLASH_PROGRESS_MAX_TWIPS
    On Error Resume Next
    Me.lbl_ProgressFill.Width = mProgressW
    On Error GoTo 0
End Sub

'################################################################
'########    Подписи расшифровки по номеру тика прогресса ########
'################################################################
Private Sub UpdateExpandLabelsFromProgressTick(ByVal tick As Long)
' Назначение: Показывает lbl_expand_1..7 не каждый тик, а по сетке FIRST + (k-1)*STEP.
' Принцип:    На «пропущенных» тиках растёт только полоса — визуально мягче.
'================================================================
    Me.lbl_expand_1.Visible = (tick >= SPLASH_EXPAND_REVEAL_FIRST_TICK + 0 * SPLASH_EXPAND_REVEAL_TICK_STEP)
    Me.lbl_expand_2.Visible = (tick >= SPLASH_EXPAND_REVEAL_FIRST_TICK + 1 * SPLASH_EXPAND_REVEAL_TICK_STEP)
    Me.lbl_expand_3.Visible = (tick >= SPLASH_EXPAND_REVEAL_FIRST_TICK + 2 * SPLASH_EXPAND_REVEAL_TICK_STEP)
    Me.lbl_expand_4.Visible = (tick >= SPLASH_EXPAND_REVEAL_FIRST_TICK + 3 * SPLASH_EXPAND_REVEAL_TICK_STEP)
    Me.lbl_expand_5.Visible = (tick >= SPLASH_EXPAND_REVEAL_FIRST_TICK + 4 * SPLASH_EXPAND_REVEAL_TICK_STEP)
    Me.lbl_expand_6.Visible = (tick >= SPLASH_EXPAND_REVEAL_FIRST_TICK + 5 * SPLASH_EXPAND_REVEAL_TICK_STEP)
    Me.lbl_expand_7.Visible = (tick >= SPLASH_EXPAND_REVEAL_FIRST_TICK + 6 * SPLASH_EXPAND_REVEAL_TICK_STEP)
End Sub

'################################################################
'########           Один тик полосы (фаза 1)             ########
'################################################################
Private Sub AdvanceSplashProgressTick()
' Назначение: Увеличивает счётчик тиков, обновляет подписи и ширину полосы линейно.
' Принцип:    Ширина ∝ tick / SPLASH_PROGRESS_TOTAL_TICKS — много тиков ⇒ мелкий шаг.
'================================================================
    If mProgressTick < SPLASH_PROGRESS_TOTAL_TICKS Then
        mProgressTick = mProgressTick + 1
    End If

    Call UpdateExpandLabelsFromProgressTick(mProgressTick)

    mProgressW = CLng((CLng(SPLASH_PROGRESS_MAX_TWIPS) * mProgressTick) / SPLASH_PROGRESS_TOTAL_TICKS)
    If mProgressW > SPLASH_PROGRESS_MAX_TWIPS Then mProgressW = SPLASH_PROGRESS_MAX_TWIPS

    On Error Resume Next
    Me.lbl_ProgressFill.Width = mProgressW
    On Error GoTo 0
End Sub

'################################################################
'########      Цвета заставки из глобальной палитры      ########
'################################################################
Private Sub ApplySplashColorsFromGlobals()
' Назначение: Фон области данных (Detail) и подписи в цветах активной темы.
' Принцип:    Читает уже заполненные Theme_WritePalette до открытия календаря.
'================================================================
    On Error GoTo Err_Handler

    Const acTransparent As Long = 0

    Me.Section("Detail").backColor = FormTheme_Back

    Me.lbl_Title.BackStyle = acTransparent
    Me.lbl_Title.ForeColor = HeaderTheme_Text

    Me.lbl_Version.BackStyle = acTransparent
    Me.lbl_Version.ForeColor = OtherTheme_Text

    Me.lbl_expand_1.BackStyle = acTransparent
    Me.lbl_expand_1.ForeColor = OtherTheme_Text
    Me.lbl_expand_1.FontSize = SPLASH_EXPAND_FONT_PT
    Me.lbl_expand_2.BackStyle = acTransparent
    Me.lbl_expand_2.ForeColor = OtherTheme_Text
    Me.lbl_expand_2.FontSize = SPLASH_EXPAND_FONT_PT
    Me.lbl_expand_3.BackStyle = acTransparent
    Me.lbl_expand_3.ForeColor = OtherTheme_Text
    Me.lbl_expand_3.FontSize = SPLASH_EXPAND_FONT_PT
    Me.lbl_expand_4.BackStyle = acTransparent
    Me.lbl_expand_4.ForeColor = OtherTheme_Text
    Me.lbl_expand_4.FontSize = SPLASH_EXPAND_FONT_PT
    Me.lbl_expand_5.BackStyle = acTransparent
    Me.lbl_expand_5.ForeColor = OtherTheme_Text
    Me.lbl_expand_5.FontSize = SPLASH_EXPAND_FONT_PT
    Me.lbl_expand_6.BackStyle = acTransparent
    Me.lbl_expand_6.ForeColor = OtherTheme_Text
    Me.lbl_expand_6.FontSize = SPLASH_EXPAND_FONT_PT
    Me.lbl_expand_7.BackStyle = acTransparent
    Me.lbl_expand_7.ForeColor = OtherTheme_Text
    Me.lbl_expand_7.FontSize = SPLASH_EXPAND_FONT_PT

    Me.lbl_ProgressFill.BackStyle = 1
    Me.lbl_ProgressFill.backColor = HeaderTheme_Border
    Me.lbl_ProgressFill.borderColor = CurrentTheme_Border

    Exit Sub

Err_Handler:
    Debug.Print "[frmSplash][ApplySplashColorsFromGlobals] " & Err.Number & " " & Err.description
End Sub

'################################################################
'########     Подписи расшифровки заголовка (ДНЕВНИК)     ########
'################################################################
Private Sub SetSplashExpandCaptions()
' Назначение: Тексты строк расшифровки по буквам аббревиатуры.
' Принцип:    Соответствие порядка lbl_expand_1…7 буквам Д-Н-Е-В-Н-И-К.
'================================================================
    Me.lbl_expand_1.Caption = "Доступный"
    Me.lbl_expand_2.Caption = "Надежный"
    Me.lbl_expand_3.Caption = "Ежедневный"
    Me.lbl_expand_4.Caption = "Виртуальный"
    Me.lbl_expand_5.Caption = "Напоминальник"
    Me.lbl_expand_6.Caption = "Идеального"
    Me.lbl_expand_7.Caption = "Контроля"
End Sub

Private Sub Form_Load()
    Call ApplySplashColorsFromGlobals
    Call SetSplashExpandCaptions

    Me.lbl_Title.Caption = "ДНЕВНИК"
    Me.lbl_Version.Caption = ""
    mVerPos = 1
    mPhase = 0
    mProgressW = 8
    mProgressTick = 0
    Call UpdateExpandLabelsFromProgressTick(0)

    On Error Resume Next
    Me.lbl_ProgressFill.Width = mProgressW
    On Error GoTo 0

    Me.TimerInterval = SPLASH_TYPE_MS
End Sub

Private Sub Form_Timer()
    On Error GoTo ErrHandler

    Select Case mPhase
        Case 0
            If mVerPos <= Len(SPLASH_VERSION_TEXT) Then
                Me.lbl_Version.Caption = Me.lbl_Version.Caption & Mid$(SPLASH_VERSION_TEXT, mVerPos, 1)
                mVerPos = mVerPos + 1
            Else
                mPhase = 1
                mProgressTick = 0
                mProgressW = 8
                On Error Resume Next
                Me.lbl_ProgressFill.Width = mProgressW
                On Error GoTo ErrHandler
                Call UpdateExpandLabelsFromProgressTick(0)
            End If

        Case 1
            Call AdvanceSplashProgressTick
            If mProgressTick >= SPLASH_PROGRESS_TOTAL_TICKS Then
                mPhase = 2
                Call SplashProgressSnapToFull
            End If

        Case Else
            Me.TimerInterval = 0
            Call SplashProgressSnapToFull
            DoCmd.Close acForm, Me.Name
            Exit Sub
    End Select

    Exit Sub

ErrHandler:
    Me.TimerInterval = 0
    On Error Resume Next
    DoCmd.Close acForm, Me.Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.TimerInterval = 0
End Sub
