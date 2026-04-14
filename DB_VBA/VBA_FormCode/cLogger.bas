' cLogger.cls
Option Compare Database
Option Explicit

Public Enum eLogLevel
    lvlDEBUG = 10
    lvlINFO = 20
    lvlWARNING = 30
    lvlERROR = 40
    lvlCRITICAL = 50
End Enum

Private pName As String
Private pLevel As eLogLevel
Private pLogPath As String

Public Property Get LoggerName() As String: LoggerName = pName: End Property
Public Property Let LoggerName(value As String): pName = value: End Property

Public Property Get LogLevel() As eLogLevel: LogLevel = pLevel: End Property
Public Property Let LogLevel(value As eLogLevel): pLevel = value: End Property

Public Property Get LogPath() As String: LogPath = pLogPath: End Property
Public Property Let LogPath(value As String)
    pLogPath = value
    EnsureDir value
End Property

' Уровень DEBUG из вызывающего кода (снаружи класса нельзя писать cLogger.lvlDEBUG).
Public Sub SetLogLevelDebug()
    pLevel = lvlDEBUG
End Sub

Private Sub Class_Initialize()
    pName = "AccessApp"
    pLevel = lvlINFO
End Sub

Public Sub WriteLog(ByVal lvl As eLogLevel, ByVal msg As String, Optional ByVal source As String = "", Optional ByVal captureErr As Boolean = False)
    If lvl < pLevel Then Exit Sub

    Dim ts As String: ts = Format(Now, "yyyy-mm-dd hh:nn:ss")
    Dim line As String
    line = "[" & ts & "] [" & lvl & "] [" & pName & "] "
    If source <> "" Then line = line & "[" & source & "] "
    line = line & msg

    If captureErr And Err.Number <> 0 Then
        line = line & vbCrLf & "  -> Err " & Err.Number & ": " & Err.description
        If Erl <> 0 Then line = line & " (Line " & Erl & ")"
    End If

    Debug.Print line
    If pLogPath <> "" Then AppendToFile pLogPath, line
End Sub

Public Sub DebugLog(ByVal msg As String, Optional source As String = ""): WriteLog lvlDEBUG, msg, source: End Sub
Public Sub InfoLog(ByVal msg As String, Optional source As String = ""): WriteLog lvlINFO, msg, source: End Sub
Public Sub WarningLog(ByVal msg As String, Optional source As String = ""): WriteLog lvlWARNING, msg, source: End Sub
Public Sub ErrorLog(ByVal msg As String, Optional source As String = ""): WriteLog lvlERROR, msg, source, True: End Sub
Public Sub CriticalLog(ByVal msg As String, Optional source As String = ""): WriteLog lvlCRITICAL, msg, source, True: End Sub

Private Sub AppendToFile(path As String, text As String)
    Dim f As Integer: f = FreeFile
    Open path For Append As #f
    Print #f, text
    Close #f
End Sub

Private Sub EnsureDir(fullPath As String)
    If fullPath = "" Then Exit Sub
    Dim dirPath As String
    If InStrRev(fullPath, "\") > 0 Then
        dirPath = Left(fullPath, InStrRev(fullPath, "\") - 1)
        If Dir(dirPath, vbDirectory) = "" Then MkDir dirPath
    End If
End Sub
