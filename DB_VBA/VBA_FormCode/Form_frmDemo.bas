Option Compare Database
Option Explicit

' Константы позиции и размеров формы (twips)
Private Const TWIPS_FRMDEMO_LEFT As Long = 6000
Private Const TWIPS_FRMDEMO_TOP As Long = 6000
Private Const TWIPS_FRMDEMO_WIDTH As Long = 10000
Private Const TWIPS_FRMDEMO_HEIGHT As Long = 6000

'################################################################
'########            ОБЪЯВЛЕНИЕ ПЕРЕМЕННЫХ               ########
'################################################################

Dim DemoStep As Integer
Dim DemoText As String
Dim DemoChar As Integer
Dim DemoInterval As Integer
Dim TimerMode As Integer ' 0-печать, 1-задержка

' Тексты шагов: modFrmDemoConstants (DEMO_STEP1_TITLE …)

'################################################################
'########          ЗАГРУЗКА ФОРМЫ ДЕМО-РЕЖИМА         ########
'################################################################

Private Sub Form_Load()
    On Error Resume Next
    DoCmd.MoveSize TWIPS_FRMDEMO_LEFT, TWIPS_FRMDEMO_TOP, TWIPS_FRMDEMO_WIDTH, TWIPS_FRMDEMO_HEIGHT
    On Error GoTo 0

    DemoStep = 1
    DemoInterval = 10
    Call UpdateDemoStep
End Sub

'################################################################
'########           ТАЙМЕР АНИМАЦИИ И ДЕЙСТВИЙ           ########
'################################################################

Private Sub Form_Timer()
    If TimerMode = 0 Then
        If Len(DemoText & "") = 0 Then
            Me.TimerInterval = 0
            Exit Sub
        End If

        If DemoChar <= Len(DemoText) Then
            Me.txtActionDescription.value = Me.txtActionDescription.value & Mid(DemoText, DemoChar, 1)
            DemoChar = DemoChar + 1
        Else
            TimerMode = 1
            Me.TimerInterval = 1000
        End If

    ElseIf TimerMode = 1 Then
        TimerMode = 0
        Call ExecuteDemoAction
        Me.TimerInterval = 0
    End If
End Sub

'################################################################
'########          ОБНОВЛЕНИЕ ШАГА ДЕМО-РЕЖИМА       ########
'################################################################

Private Sub UpdateDemoStep()
    DemoChar = 1
    Me.txtActionDescription.value = ""
    Me.txtCurrentAction.value = ""
    TimerMode = 0

    Select Case DemoStep
        Case 1
            Me.lblProcessName.Caption = DEMO_STEP1_TITLE
            DemoText = DEMO_STEP1_TEXT
        Case 2
            Me.lblProcessName.Caption = DEMO_STEP2_TITLE
            DemoText = DEMO_STEP2_TEXT
        Case 3
            Me.lblProcessName.Caption = DEMO_STEP3_TITLE
            DemoText = DEMO_STEP3_TEXT
        Case 4
            Me.lblProcessName.Caption = DEMO_STEP4_TITLE
            DemoText = DEMO_STEP4_TEXT
        Case 5
            Me.lblProcessName.Caption = DEMO_STEP5_TITLE
            DemoText = DEMO_STEP5_TEXT
        Case 6
            Me.lblProcessName.Caption = DEMO_STEP6_TITLE
            DemoText = DEMO_STEP6_TEXT
    End Select

    Me.TimerInterval = DemoInterval
End Sub

'################################################################
'########             КНОПКА "ДАЛЕЕ"                     ########
'################################################################

Private Sub cmdNext_Click()
    If DemoStep < 6 Then
        DemoStep = DemoStep + 1
        Call UpdateDemoStep
    Else
        DoCmd.Close acForm, "frmDemo"
    End If
End Sub

'################################################################
'########             КНОПКА "НАЗАД"                     ########
'################################################################

Private Sub cmdBack_Click()
    If DemoStep > 1 Then
        DemoStep = DemoStep - 1
        Call UpdateDemoStep
    End If
End Sub

'################################################################
'########          ВЫПОЛНЕНИЕ ДЕЙСТВИЙ ДЕМО-РЕЖИМА       ########
'################################################################

Private Sub ExecuteDemoAction()
    Select Case DemoStep
        Case 2
            FrmDemo_ExecuteNavigationDemo Me
        Case 3
            FrmDemo_ExecuteEventsDemo Me
        Case 4
            FrmDemo_ExecuteFilterDemo Me
        Case 5
            FrmDemo_ExecuteSearchDemo Me
        Case 6
            FrmDemo_ExecuteThemeDemo Me
    End Select
End Sub
