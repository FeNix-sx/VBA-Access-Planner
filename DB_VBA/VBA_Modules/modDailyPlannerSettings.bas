Option Compare Database
Option Explicit

' Ключ tbSettings: дата yyyy-mm-dd последнего показа заставки (не чаще раза в сутки).
Public Const PLN_SETTINGS_SPLASH_DATE As String = "PlannerSplashDate"

'################################################################
'########     НАСТРОЙКИ ПЛАНИРОВЩИКА (tbSettings)        ########
'################################################################
' Назначение: Чтение/запись строковых настроек и режима окна.
' Принцип:    Те же SQL-операции, что были в Form_f_daily_planner.
'################################################################

'################################################################
'########       Получить значение настройки по имени     ########
'################################################################
Public Function PlnSettings_GetValue(ByVal settingName As String, Optional ByVal defaultValue As String = "") As String
    On Error GoTo ExitFn
    Dim rs As DAO.Recordset

    Set rs = CurrentDb.OpenRecordset("SELECT SettingValue FROM tbSettings WHERE SettingName = '" & Replace(settingName, "'", "''") & "'")

    If Not rs.EOF Then
        PlnSettings_GetValue = Nz(rs!settingValue, defaultValue)
    Else
        PlnSettings_GetValue = defaultValue
    End If

ExitFn:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
End Function

'################################################################
'########      Сохранить значение настройки по имени     ########
'################################################################
Public Sub PlnSettings_SaveValue(ByVal settingName As String, ByVal settingValue As String)
    On Error GoTo Err_Handler
    Dim safeName As String
    Dim safeValue As String

    safeName = Replace(settingName, "'", "''")
    safeValue = Replace(settingValue, "'", "''")

    CurrentDb.Execute "DELETE FROM tbSettings WHERE SettingName = '" & safeName & "'"
    CurrentDb.Execute "INSERT INTO tbSettings (SettingName, SettingValue) VALUES ('" & safeName & "', '" & safeValue & "')"
    Exit Sub

Err_Handler:
    Debug.Print "[PlnSettings][ERR][SaveValue] " & Err.Number & " - " & Err.description & "; setting=" & settingName
End Sub

'################################################################
'########      Сохранить режим окна в tbSettings         ########
'################################################################
Public Sub PlnSettings_SaveWindowMode(ByVal modeValue As String)
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim sqlText As String

    Set db = CurrentDb
    db.Execute "DELETE FROM tbSettings WHERE SettingName = 'PlannerWindowMode'"

    sqlText = "INSERT INTO tbSettings (SettingName, SettingValue) VALUES ('PlannerWindowMode', '" & modeValue & "')"
    db.Execute sqlText
    Exit Sub

ErrHandler:
    Debug.Print "[PlnSettings][ERR][SaveWindowMode] " & Err.Number & " - " & Err.description
End Sub

'################################################################
'########      Загрузить режим окна из tbSettings        ########
'################################################################
Public Function PlnSettings_GetWindowMode() As String
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset

    PlnSettings_GetWindowMode = "windowed"

    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT SettingValue FROM tbSettings WHERE SettingName = 'PlannerWindowMode'")

    If Not rs.EOF Then
        PlnSettings_GetWindowMode = LCase$(Nz(rs!settingValue, "windowed"))
    End If

ExitFn:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Function

ErrHandler:
    Debug.Print "[PlnSettings][ERR][GetWindowMode] " & Err.Number & " - " & Err.description
    Resume ExitFn
End Function

