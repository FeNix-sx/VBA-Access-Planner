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
'########            Записать  палитру                   ########
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

'################################################################
'########         Вывести Палитру ВсехТем                ########
'################################################################
Public Sub DumpAllThemesPalette()
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim fld As DAO.Field
    
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tbThemes ORDER BY ThemeID", dbOpenSnapshot)
    
    If rs.EOF Then
        Debug.Print "[THEME] Таблица tbThemes пуста."
        GoTo CleanExit
    End If
    
    Do Until rs.EOF
        Debug.Print "=== ТЕМА #" & rs!ThemeID & ": " & Nz(rs!ThemeName, "Без имени") & " ==="
        If rs!IsActive <> 0 Then Debug.Print "  [? ACTIVE]"
        
        For Each fld In rs.Fields
            Select Case fld.Name
                Case "ThemeID", "IsActive": ' skip
                Case Else
                    If IsNumeric(fld.value) And Not IsNull(fld.value) Then
                        If fld.Name Like "*_Back" Or fld.Name Like "*_Border" Or fld.Name Like "*_Text" Then
                            Debug.Print "  " & fld.Name & " => " & fld.value & "  (#" & LongToHexRGB(fld.value) & ")"
                        Else
                            Debug.Print "  " & fld.Name & " => " & fld.value
                        End If
                    Else
                        Debug.Print "  " & fld.Name & " => " & Nz(fld.value, "[NULL]")
                    End If
            End Select
        Next fld
        Debug.Print "----------------------------------------"
        rs.MoveNext
    Loop

CleanExit:
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing: Set db = Nothing
    Exit Sub

ErrHandler:
    Debug.Print "[ERR " & Err.Number & "] " & Err.description
    Resume CleanExit
End Sub

Private Function LongToHexRGB(ByVal clr As Long) As String
    Dim r As Long, g As Long, b As Long
    r = clr And &HFF&
    g = (clr \ &H100&) And &HFF&
    b = (clr \ &H10000) And &HFF&
    LongToHexRGB = Right("00" & Hex(r), 2) & Right("00" & Hex(g), 2) & Right("00" & Hex(b), 2)
End Function

'################################################################
'########   Инициализировать Финальные цветовые Темы     ########
'################################################################
' Удаляет старые/тестовые записи и вставляет финальные палитры.
Public Sub InitializeFinalThemes()
    On Error GoTo ErrHandler
    Dim db As DAO.Database
    Set db = CurrentDb

    db.Execute "DELETE FROM tbThemes WHERE ThemeName IN ('Тёмная','Бирюзовая','Коралловая','Морская','Пудровая','Кофейная','Персиковая','Медовая','Оливковая','Серая','Голубая','Зеленая','Фиолетовая')", dbFailOnError

    Dim sql As String
    sql = "INSERT INTO tbThemes (ThemeName, IsActive, CurrentMonth_Text, CurrentMonth_Back, CurrentMonth_Border, OtherMonth_Text, OtherMonth_Back, OtherMonth_Border, Today_Back, Today_Border, Header_Text, Header_Back, Header_Border, Form_Back) VALUES "
    
    ' 1. Персиковая
    db.Execute sql & "('Персиковая', 0, " & RGB(139, 69, 19) & ", " & RGB(255, 240, 230) & ", " & RGB(210, 105, 30) & ", " & RGB(160, 120, 90) & ", " & RGB(255, 250, 245) & ", " & RGB(220, 200, 180) & ", " & RGB(255, 228, 181) & ", " & RGB(255, 140, 0) & ", " & RGB(139, 69, 19) & ", " & RGB(255, 240, 230) & ", " & RGB(210, 105, 30) & ", " & RGB(255, 250, 245) & ")", dbFailOnError
    
    ' 2. Медовая
    db.Execute sql & "('Медовая', 0, " & RGB(101, 67, 33) & ", " & RGB(255, 236, 179) & ", " & RGB(255, 193, 7) & ", " & RGB(150, 120, 80) & ", " & RGB(255, 248, 225) & ", " & RGB(255, 224, 130) & ", " & RGB(255, 215, 0) & ", " & RGB(255, 140, 0) & ", " & RGB(101, 67, 33) & ", " & RGB(255, 236, 179) & ", " & RGB(255, 193, 7) & ", " & RGB(255, 248, 225) & ")", dbFailOnError
    
    ' 3. Оливковая
    db.Execute sql & "('Оливковая', 0, " & RGB(85, 107, 47) & ", " & RGB(240, 255, 240) & ", " & RGB(107, 142, 35) & ", " & RGB(120, 140, 80) & ", " & RGB(245, 255, 245) & ", " & RGB(180, 200, 150) & ", " & RGB(189, 236, 182) & ", " & RGB(85, 160, 70) & ", " & RGB(85, 107, 47) & ", " & RGB(240, 255, 240) & ", " & RGB(107, 142, 35) & ", " & RGB(245, 255, 245) & ")", dbFailOnError
    
    ' 4. Серая
    db.Execute sql & "('Серая', 0, " & RGB(68, 68, 68) & ", " & RGB(240, 240, 240) & ", " & RGB(191, 191, 191) & ", " & RGB(160, 160, 160) & ", " & RGB(255, 255, 255) & ", " & RGB(191, 191, 191) & ", " & RGB(204, 204, 204) & ", " & RGB(68, 68, 68) & ", " & RGB(68, 68, 68) & ", " & RGB(204, 204, 204) & ", " & RGB(191, 191, 191) & ", " & RGB(255, 255, 255) & ")", dbFailOnError
    
    ' 5. Голубая
    db.Execute sql & "('Голубая', 0, " & RGB(53, 53, 53) & ", " & RGB(197, 255, 255) & ", " & RGB(102, 204, 255) & ", " & RGB(160, 160, 160) & ", " & RGB(255, 255, 255) & ", " & RGB(102, 204, 255) & ", " & RGB(197, 255, 255) & ", " & RGB(0, 0, 255) & ", " & RGB(53, 53, 53) & ", " & RGB(197, 255, 255) & ", " & RGB(102, 204, 255) & ", " & RGB(231, 255, 255) & ")", dbFailOnError
    
    ' 6. Зеленая
    db.Execute sql & "('Зеленая', 0, " & RGB(53, 53, 53) & ", " & RGB(194, 254, 205) & ", " & RGB(128, 255, 128) & ", " & RGB(160, 160, 160) & ", " & RGB(233, 255, 221) & ", " & RGB(128, 255, 128) & ", " & RGB(194, 254, 205) & ", " & RGB(0, 128, 0) & ", " & RGB(53, 53, 53) & ", " & RGB(194, 254, 205) & ", " & RGB(128, 255, 128) & ", " & RGB(245, 255, 245) & ")", dbFailOnError
    
    ' 7. Фиолетовая
    db.Execute sql & "('Фиолетовая', 0, " & RGB(53, 53, 53) & ", " & RGB(226, 205, 255) & ", " & RGB(128, 0, 128) & ", " & RGB(160, 160, 160) & ", " & RGB(241, 231, 255) & ", " & RGB(128, 0, 128) & ", " & RGB(226, 205, 255) & ", " & RGB(128, 0, 128) & ", " & RGB(53, 53, 53) & ", " & RGB(226, 205, 255) & ", " & RGB(128, 0, 128) & ", " & RGB(246, 235, 255) & ")", dbFailOnError
    
    ' 8. Тёмная (IDE-style, тёплый серый)
    db.Execute sql & "('Тёмная', 0, " & RGB(225, 225, 228) & ", " & RGB(64, 60, 56) & ", " & RGB(104, 96, 88) & ", " & RGB(150, 145, 140) & ", " & RGB(50, 46, 42) & ", " & RGB(80, 74, 68) & ", " & RGB(74, 67, 60) & ", " & RGB(127, 110, 90) & ", " & RGB(228, 228, 230) & ", " & RGB(80, 74, 70) & ", " & RGB(114, 104, 94) & ", " & RGB(94, 90, 86) & ")", dbFailOnError
    
    ' 9. Бирюзовая (Material Teal accent)
    db.Execute sql & "('Бирюзовая', 0, " & RGB(0, 75, 70) & ", " & RGB(232, 245, 245) & ", " & RGB(38, 166, 154) & ", " & RGB(90, 110, 110) & ", " & RGB(242, 250, 250) & ", " & RGB(150, 200, 200) & ", " & RGB(190, 235, 230) & ", " & RGB(20, 140, 130) & ", " & RGB(0, 75, 70) & ", " & RGB(225, 240, 240) & ", " & RGB(38, 166, 154) & ", " & RGB(248, 252, 252) & ")", dbFailOnError
    
    ' 10. Коралловая (тёплый акцент, OtherMonth ярче, CurrentMonth мягче)
    db.Execute sql & "('Коралловая', 0, " & RGB(85, 45, 30) & ", " & RGB(255, 222, 210) & ", " & RGB(242, 128, 100) & ", " & RGB(145, 110, 92) & ", " & RGB(250, 236, 228) & ", " & RGB(232, 158, 135) & ", " & RGB(255, 170, 145) & ", " & RGB(210, 70, 35) & ", " & RGB(80, 40, 25) & ", " & RGB(255, 215, 195) & ", " & RGB(235, 95, 55) & ", " & RGB(255, 238, 228) & ")", dbFailOnError
    
    ' 11. Морская (деловой синий)
    db.Execute sql & "('Морская', 0, " & RGB(30, 58, 95) & ", " & RGB(227, 242, 253) & ", " & RGB(66, 165, 245) & ", " & RGB(120, 144, 156) & ", " & RGB(248, 250, 252) & ", " & RGB(176, 190, 197) & ", " & RGB(187, 222, 251) & ", " & RGB(30, 136, 229) & ", " & RGB(26, 54, 93) & ", " & RGB(226, 232, 240) & ", " & RGB(144, 164, 174) & ", " & RGB(244, 247, 250) & ")", dbFailOnError
    
    ' 12. Пудровая (мягкий розово-сиреневый)
    db.Execute sql & "('Пудровая', 0, " & RGB(92, 64, 80) & ", " & RGB(253, 242, 248) & ", " & RGB(224, 145, 167) & ", " & RGB(154, 136, 142) & ", " & RGB(252, 248, 249) & ", " & RGB(216, 196, 202) & ", " & RGB(252, 228, 236) & ", " & RGB(216, 27, 96) & ", " & RGB(109, 76, 90) & ", " & RGB(240, 228, 232) & ", " & RGB(196, 154, 168) & ", " & RGB(250, 245, 246) & ")", dbFailOnError
    
    ' 13. Кофейная (тёплый нейтральный, повышенный контраст)
    db.Execute sql & "('Кофейная', 0, " & RGB(40, 20, 15) & ", " & RGB(225, 215, 205) & ", " & RGB(130, 100, 85) & ", " & RGB(110, 90, 75) & ", " & RGB(238, 232, 225) & ", " & RGB(180, 165, 150) & ", " & RGB(205, 190, 175) & ", " & RGB(65, 35, 25) & ", " & RGB(45, 25, 18) & ", " & RGB(215, 200, 190) & ", " & RGB(115, 85, 70) & ", " & RGB(228, 220, 210) & ")", dbFailOnError

    Debug.Print ">>> Финальные темы добавлены (13 шт.)."
    Exit Sub

ErrHandler:
    Debug.Print "[ERR " & Err.Number & "] " & Err.description
End Sub
