VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GUIHUD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************************
'Heads-up display for the GUI
'*********************************************************************************

Option Explicit

'General information
Private SkinPath As String

'HUD
Private cX As Integer
Private cY As Integer
Private cWidth As Integer
Private cHeight As Integer
Private cGrh As tGrh

'Level label
Private cLevelX As Integer
Private cLevelY As Integer
Private cLevelXOffset As Integer
Private cLastLevel As Integer

'HP and MP bars
Private cHPX As Integer
Private cHPY As Integer
Private cHPWidth As Integer
Private cHPHeight As Integer
Private cMPX As Integer
Private cMPY As Integer
Private cMPWidth As Integer
Private cMPHeight As Integer

'Ryu
Private cRyuX As Integer
Private cRyuY As Integer

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

Public Sub LeftClick(ByVal X As Integer, ByVal Y As Integer)
'*********************************************************************************
'Handles leftclick events on the HUD
'*********************************************************************************

    'Drop an item from the inventory
    If InvSwapSlot > 0 Then
        sndBuf.Put_Byte PId.CS_DropItem
        sndBuf.Put_Byte InvSwapSlot
        InvSwapSlot = -1
    End If

End Sub

Public Sub Load(Optional ByVal GUIName As String = "Default")
'*********************************************************************************
'Load the HUD
'*********************************************************************************

    'Set the skin path
    SkinPath = App.Path & "\Data\GUI\" & GUIName & ".ini"

    'Set the HUD information
    cX = Val(IO_INI_Read(SkinPath, "HUD", "X"))
    cY = Val(IO_INI_Read(SkinPath, "HUD", "Y"))
    cWidth = Val(IO_INI_Read(SkinPath, "HUD", "Width"))
    cHeight = Val(IO_INI_Read(SkinPath, "HUD", "Height"))
    
    cHPX = Val(IO_INI_Read(SkinPath, "HUD", "HPX"))
    cHPY = Val(IO_INI_Read(SkinPath, "HUD", "HPY"))
    cHPWidth = Val(IO_INI_Read(SkinPath, "HUD", "HPWidth"))
    cHPHeight = Val(IO_INI_Read(SkinPath, "HUD", "HPHeight"))
    
    cMPX = Val(IO_INI_Read(SkinPath, "HUD", "MPX"))
    cMPY = Val(IO_INI_Read(SkinPath, "HUD", "MPY"))
    cMPWidth = Val(IO_INI_Read(SkinPath, "HUD", "MPWidth"))
    cMPHeight = Val(IO_INI_Read(SkinPath, "HUD", "MPHeight"))
    
    cRyuX = Val(IO_INI_Read(SkinPath, "HUD", "RyuX"))
    cRyuY = Val(IO_INI_Read(SkinPath, "HUD", "RyuY"))
    
    cLevelX = Val(IO_INI_Read(SkinPath, "HUD", "LevelX"))
    cLevelY = Val(IO_INI_Read(SkinPath, "HUD", "LevelY"))
    
    'Set the HUD Grh
    Graphics_SetGrh cGrh, Val(IO_INI_Read(SkinPath, "HUD", "Grh")), ANIMTYPE_LOOP
    
End Sub

Public Sub Draw()
'*********************************************************************************
'Draw the HUD
'*********************************************************************************
Const HPColor As Long = -65536
Const MPColor As Long = -16776961
Dim v(0 To 3) As TLVERTEX

    'Backdrop
    Graphics_DrawGrh cGrh, cX, cY
    
    'HP and MP bars
    Graphics_SetTexture 0
    If UserStats.MaxMP > 0 Then
        Graphics_DrawRect cHPX, cHPY, cHPWidth * (UserStats.HP / UserStats.MaxHP), cHPHeight, HPColor, HPColor, HPColor, HPColor
    End If
    If UserStats.MaxMP > 0 Then
        Graphics_DrawRect cMPX, cMPY, cMPWidth * (UserStats.MP / UserStats.MaxMP), cMPHeight, MPColor, MPColor, MPColor, MPColor
    End If
    
    'Ryu
    FontDefault.Draw UserStats.Ryu, cRyuX, cRyuY, -1
    
    'Level
    If cLastLevel <> UserStats.Level Then
        cLastLevel = UserStats.Level
        cLevelXOffset = -(FontDefault.Width(UserStats.Level) \ 2)
    End If
    FontDefault.Draw UserStats.Level, cX + cLevelX + cLevelXOffset, cY + cLevelY, -1

End Sub
