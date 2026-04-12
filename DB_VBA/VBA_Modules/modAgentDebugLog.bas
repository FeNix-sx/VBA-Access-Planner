Option Compare Database
Option Explicit

'################################################################
'########   Отладочный лог сессии (NDJSON в файл)        ########
'################################################################

Public Function AgentJsonEsc(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCrLf, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    AgentJsonEsc = t
End Function

Public Sub AgentDbgLog(ByVal hid As String, ByVal loc As String, ByVal msg As String, ByVal dataJson As String)
    Dim p As String
    Dim f As Integer
    Dim line As String
    Const MIRROR As String = "d:\Planner\debug-3cec1e.log"
    On Error Resume Next
    line = "{""sessionId"":""3cec1e"",""hid"":""" & AgentJsonEsc(hid) & """,""loc"":""" & AgentJsonEsc(loc) & """,""msg"":""" & AgentJsonEsc(msg) & """,""data"":" & dataJson & ",""ts"":" & CLng(Timer * 1000) & "}"
    p = CurrentProject.Path & "\debug-3cec1e.log"
    f = FreeFile
    Open p For Append As #f
    Print #f, line
    Close #f
    If LCase$(CurrentProject.Path) <> "d:\planner" Then
        f = FreeFile
        Open MIRROR For Append As #f
        Print #f, line
        Close #f
    End If
End Sub

Public Sub AgentDbgLogErr(ByVal hid As String, ByVal loc As String, ByVal errNum As Long, ByVal errDesc As String)
    AgentDbgLog hid, loc, "error", "{""errNum"":" & errNum & ",""errDesc"":""" & AgentJsonEsc(errDesc) & """}"
End Sub

