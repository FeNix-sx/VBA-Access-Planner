Option Compare Database
Option Explicit

'################################################################
'########       Переоткрыть форму по имени               ########
'################################################################
Public Sub ReopenPlannerFormGlobal(ByVal formName As String)
    On Error GoTo ErrHandler

    DoCmd.Close acForm, formName
    DoCmd.OpenForm formName
    Exit Sub

ErrHandler:
    Debug.Print "[modWindowMode][ERR][ReopenPlannerFormGlobal] " & Err.Number & " - " & Err.description
End Sub


