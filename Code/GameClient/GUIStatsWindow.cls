VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GUIStatsWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Path to the skin
Private SkinPath As String

'Window location
Public cX As Integer
Public cY As Integer
Public cWidth As Integer
Public cHeight As Integer

'Font colors
Private cFontColor As Long

'Window text locations
Private cNameX As Integer
Private cNameY As Integer
Private cJobX As Integer
Private cJobY As Integer
Private cLevelX As Integer
Private cLevelY As Integer
Private cGuildX As Integer
Private cGuildY As Integer
Private cHPX As Integer
Private cHPY As Integer
Private cMPX As Integer
Private cMPY As Integer
Private cEXPX As Integer
Private cEXPY As Integer
Private cFameX As Integer
Private cFameY As Integer
Private cStrX As Integer
Private cStrY As Integer
Private cDexX As Integer
Private cDexY As Integer
Private cIntlX As Integer
Private cIntlY As Integer
Private cLukX As Integer
Private cLukY As Integer
Private cPointsX As Integer
Private cPointsY As Integer

'If the window is visible
Public cVisible As Boolean

'Graphic information
Private cGrh As tGrh

Public Property Get Visible() As Boolean
    Visible = cVisible
End Property

Public Property Let Visible(ByVal Value As Boolean)
    cVisible = Value
    If cVisible Then FocusWID = WID_StatsWindow
End Property

Public Property Get X() As Integer
    X = cX
End Property

Public Property Let X(ByVal Value As Integer)
    cX = Value
    If STAYINSCREEN Then
        If cX < 0 Then cX = 0
        If cX + Width > ScreenWidth Then cX = ScreenWidth - Width
    End If
End Property

Public Property Get Y() As Integer
    Y = cY
End Property

Public Property Let Y(ByVal Value As Integer)
    cY = Value
    If STAYINSCREEN Then
        If cY < 0 Then cY = 0
        If cY + Height > ScreenHeight Then cY = ScreenHeight - Height
    End If
End Property

Public Property Get Width() As Integer
    Width = cWidth
End Property

Public Property Get Height() As Integer
    Height = cHeight
End Property

Public Sub Load(Optional ByVal GUIName As String = "Default")
'*********************************************************************************
'Load the stats window
'*********************************************************************************

    'Set the skin path
    SkinPath = App.Path & "\Data\GUI\" & GUIName & ".ini"

    'Set the window information
    cX = Val(IO_INI_Read(SkinPath, "STATS", "X"))
    cY = Val(IO_INI_Read(SkinPath, "STATS", "Y"))
    cWidth = Val(IO_INI_Read(SkinPath, "STATS", "Width"))
    cHeight = Val(IO_INI_Read(SkinPath, "STATS", "Height"))
    
    'Font colors
    cFontColor = D3DColorARGB(255, Val(IO_INI_Read(SkinPath, "STATS", "R")), _
        Val(IO_INI_Read(SkinPath, "STATS", "G")), Val(IO_INI_Read(SkinPath, "STATS", "B")))
    
    'Window text locations
    cNameX = Val(IO_INI_Read(SkinPath, "STATS", "NameX"))
    cNameY = Val(IO_INI_Read(SkinPath, "STATS", "NameY"))
    cJobX = Val(IO_INI_Read(SkinPath, "STATS", "JobX"))
    cJobY = Val(IO_INI_Read(SkinPath, "STATS", "JobY"))
    cLevelX = Val(IO_INI_Read(SkinPath, "STATS", "LevelX"))
    cLevelY = Val(IO_INI_Read(SkinPath, "STATS", "LevelY"))
    cGuildX = Val(IO_INI_Read(SkinPath, "STATS", "GuildX"))
    cGuildY = Val(IO_INI_Read(SkinPath, "STATS", "GuildY"))
    cHPX = Val(IO_INI_Read(SkinPath, "STATS", "HPX"))
    cHPY = Val(IO_INI_Read(SkinPath, "STATS", "HPY"))
    cMPX = Val(IO_INI_Read(SkinPath, "STATS", "MPX"))
    cMPY = Val(IO_INI_Read(SkinPath, "STATS", "MPY"))
    cEXPX = Val(IO_INI_Read(SkinPath, "STATS", "EXPX"))
    cEXPY = Val(IO_INI_Read(SkinPath, "STATS", "EXPY"))
    cFameX = Val(IO_INI_Read(SkinPath, "STATS", "FameX"))
    cFameY = Val(IO_INI_Read(SkinPath, "STATS", "FameY"))
    cStrX = Val(IO_INI_Read(SkinPath, "STATS", "StrX"))
    cStrY = Val(IO_INI_Read(SkinPath, "STATS", "StrY"))
    cDexX = Val(IO_INI_Read(SkinPath, "STATS", "DexX"))
    cDexY = Val(IO_INI_Read(SkinPath, "STATS", "DexY"))
    cIntlX = Val(IO_INI_Read(SkinPath, "STATS", "IntlX"))
    cIntlY = Val(IO_INI_Read(SkinPath, "STATS", "IntlY"))
    cLukX = Val(IO_INI_Read(SkinPath, "STATS", "LukX"))
    cLukY = Val(IO_INI_Read(SkinPath, "STATS", "LukY"))
    cPointsX = Val(IO_INI_Read(SkinPath, "STATS", "PointsX"))
    cPointsY = Val(IO_INI_Read(SkinPath, "STATS", "PointsY"))
    
    'Init the graphic
    Graphics_SetGrh cGrh, Val(IO_INI_Read(SkinPath, "STATS", "Grh")), ANIMTYPE_LOOP

End Sub

Public Sub LeftClick(ByVal X As Integer, ByVal Y As Integer)
'*********************************************************************************
'Handles left-clicking the window
'*********************************************************************************
    
    'Nothing to click on this window, so just select it
    SelectedWID = WID_StatsWindow
    FocusWID = WID_StatsWindow

End Sub

Public Sub Draw()
'*********************************************************************************
'Draw the stats window
'*********************************************************************************

    'Confirm is visible
    If Not Visible Then Exit Sub

    'Window background
    Graphics_DrawGrh cGrh, cX, cY

    'Text
    FontDefault.Draw CharList(UserCharIndex).Name, cX + cNameX, cY + cNameY, cFontColor
    FontDefault.Draw "My Job", cX + cJobX, cY + cJobY, cFontColor
    FontDefault.Draw UserStats.Level, cX + cLevelX, cY + cLevelY, cFontColor
    FontDefault.Draw "My Guild", cX + cGuildX, cY + cGuildY, cFontColor
    FontDefault.Draw UserStats.HP & "/" & UserStats.MaxHP, cX + cHPX, cY + cHPY, cFontColor
    FontDefault.Draw UserStats.MP & "/" & UserStats.MaxMP, cX + cMPX, cY + cMPY, cFontColor
    FontDefault.Draw UserStats.EXP & "/" & User_ToNextLevel, cX + cEXPX, cY + cEXPY, cFontColor
    FontDefault.Draw "Fame", cX + cFameX, cY + cFameY, cFontColor
    FontDefault.Draw UserStats.Str & " (" & UserStats.ModStr & ")", cX + cStrX, cY + cStrY, cFontColor
    FontDefault.Draw UserStats.Dex & " (" & UserStats.ModDex & ")", cX + cDexX, cY + cDexY, cFontColor
    FontDefault.Draw UserStats.Intl & " ( " & UserStats.ModIntl & ")", cX + cIntlX, cY + cIntlY, cFontColor
    FontDefault.Draw UserStats.Luk & " (" & UserStats.ModLuk & ")", cX + cLukX, cY + cLukY, cFontColor
    FontDefault.Draw "Pts", cX + cPointsX, cY + cPointsY, cFontColor

End Sub
