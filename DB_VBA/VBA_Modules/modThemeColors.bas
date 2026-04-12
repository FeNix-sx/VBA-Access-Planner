Option Compare Database
Option Explicit

'################################################################
'########         ПАЛИТРА АКТИВНОЙ ТЕМЫ (ОБЩАЯ)          ########
'################################################################
' Пишется из Form_f_daily_planner.ApplyTheme; читается отчётами
' и другими модулями без обращения к Forms!f_daily_planner.

Public CurrentTheme_Text As Long
Public CurrentTheme_Back As Long
Public CurrentTheme_Border As Long
Public OtherTheme_Text As Long
Public OtherTheme_Back As Long
Public OtherTheme_Border As Long
Public TodayTheme_Back As Long
Public TodayTheme_Border As Long
Public HeaderTheme_Text As Long
Public HeaderTheme_Back As Long
Public HeaderTheme_Border As Long
Public FormTheme_Back As Long

'################################################################
'########             Theme_WritePalette                 ########
'################################################################
Public Sub Theme_WritePalette( _
    ByVal pCurrentTheme_Text As Long, _
    ByVal pCurrentTheme_Back As Long, _
    ByVal pCurrentTheme_Border As Long, _
    ByVal pOtherTheme_Text As Long, _
    ByVal pOtherTheme_Back As Long, _
    ByVal pOtherTheme_Border As Long, _
    ByVal pTodayTheme_Back As Long, _
    ByVal pTodayTheme_Border As Long, _
    ByVal pHeaderTheme_Text As Long, _
    ByVal pHeaderTheme_Back As Long, _
    ByVal pHeaderTheme_Border As Long, _
    ByVal pFormTheme_Back As Long)

    CurrentTheme_Text = pCurrentTheme_Text
    CurrentTheme_Back = pCurrentTheme_Back
    CurrentTheme_Border = pCurrentTheme_Border
    OtherTheme_Text = pOtherTheme_Text
    OtherTheme_Back = pOtherTheme_Back
    OtherTheme_Border = pOtherTheme_Border
    TodayTheme_Back = pTodayTheme_Back
    TodayTheme_Border = pTodayTheme_Border
    HeaderTheme_Text = pHeaderTheme_Text
    HeaderTheme_Back = pHeaderTheme_Back
    HeaderTheme_Border = pHeaderTheme_Border
    FormTheme_Back = pFormTheme_Back
End Sub


