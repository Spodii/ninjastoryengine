VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GUIEquipWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Path to the skin
Private SkinPath As String

'Window location
Private cX As Integer
Private cY As Integer
Private cWidth As Integer
Private cHeight As Integer

'If the window is visible
Private cVisible As Boolean

'Slot locations
Private cCapX As Integer
Private cCapY As Integer
Private cForeheadX As Integer
Private cForeheadY As Integer
Private cRing1X As Integer
Private cRing1Y As Integer
Private cRing2X As Integer
Private cRing2Y As Integer
Private cRing3X As Integer
Private cRing3Y As Integer
Private cRing4X As Integer
Private cRing4Y As Integer
Private cEyeAccX As Integer
Private cEyeAccY As Integer
Private cEarAccX As Integer
Private cEarAccY As Integer
Private cMantleX As Integer
Private cMantleY As Integer
Private cClothesX As Integer
Private cClothesY As Integer
Private cPendantX As Integer
Private cPendantY As Integer
Private cWeaponX As Integer
Private cWeaponY As Integer
Private cShieldX As Integer
Private cShieldY As Integer
Private cGlovesX As Integer
Private cGlovesY As Integer
Private cPantsX As Integer
Private cPantsY As Integer
Private cShoesX As Integer
Private cShoesY As Integer

'Graphic information
Private cGrh As tGrh

Public Property Get Visible() As Boolean
    Visible = cVisible
End Property

Public Property Let Visible(ByVal Value As Boolean)
    cVisible = Value
    If cVisible Then
        FocusWID = WID_EquipWindow
    Else
        TooltipItemIndex = 0
    End If
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
'Load the equipment window
'*********************************************************************************

    'Set the skin path
    SkinPath = App.Path & "\Data\GUI\" & GUIName & ".ini"

    'Set the window information
    cX = Val(IO_INI_Read(SkinPath, "EQUIPPED", "X"))
    cY = Val(IO_INI_Read(SkinPath, "EQUIPPED", "Y"))
    cWidth = Val(IO_INI_Read(SkinPath, "EQUIPPED", "Width"))
    cHeight = Val(IO_INI_Read(SkinPath, "EQUIPPED", "Height"))
    
    'Slot locations
    cCapX = Val(IO_INI_Read(SkinPath, "EQUIPPED", "CapX"))
    cCapY = Val(IO_INI_Read(SkinPath, "EQUIPPED", "CapY"))
    cForeheadX = Val(IO_INI_Read(SkinPath, "EQUIPPED", "ForeheadX"))
    cForeheadY = Val(IO_INI_Read(SkinPath, "EQUIPPED", "ForeheadY"))
    cRing1X = Val(IO_INI_Read(SkinPath, "EQUIPPED", "Ring1X"))
    cRing1Y = Val(IO_INI_Read(SkinPath, "EQUIPPED", "Ring1Y"))
    cRing2X = Val(IO_INI_Read(SkinPath, "EQUIPPED", "Ring2X"))
    cRing2Y = Val(IO_INI_Read(SkinPath, "EQUIPPED", "Ring2Y"))
    cRing3X = Val(IO_INI_Read(SkinPath, "EQUIPPED", "Ring3X"))
    cRing3Y = Val(IO_INI_Read(SkinPath, "EQUIPPED", "Ring3Y"))
    cRing4X = Val(IO_INI_Read(SkinPath, "EQUIPPED", "Ring4X"))
    cRing4Y = Val(IO_INI_Read(SkinPath, "EQUIPPED", "Ring4Y"))
    cEyeAccX = Val(IO_INI_Read(SkinPath, "EQUIPPED", "EyeAccX"))
    cEyeAccY = Val(IO_INI_Read(SkinPath, "EQUIPPED", "EyeAccY"))
    cEarAccX = Val(IO_INI_Read(SkinPath, "EQUIPPED", "EarAccX"))
    cEarAccY = Val(IO_INI_Read(SkinPath, "EQUIPPED", "EarAccY"))
    cMantleX = Val(IO_INI_Read(SkinPath, "EQUIPPED", "MantleX"))
    cMantleY = Val(IO_INI_Read(SkinPath, "EQUIPPED", "MantleY"))
    cClothesX = Val(IO_INI_Read(SkinPath, "EQUIPPED", "ClothesX"))
    cClothesY = Val(IO_INI_Read(SkinPath, "EQUIPPED", "ClothesY"))
    cPendantX = Val(IO_INI_Read(SkinPath, "EQUIPPED", "PendantX"))
    cPendantY = Val(IO_INI_Read(SkinPath, "EQUIPPED", "PendantY"))
    cWeaponX = Val(IO_INI_Read(SkinPath, "EQUIPPED", "WeaponX"))
    cWeaponY = Val(IO_INI_Read(SkinPath, "EQUIPPED", "WeaponY"))
    cShieldX = Val(IO_INI_Read(SkinPath, "EQUIPPED", "ShieldX"))
    cShieldY = Val(IO_INI_Read(SkinPath, "EQUIPPED", "ShieldY"))
    cGlovesX = Val(IO_INI_Read(SkinPath, "EQUIPPED", "GlovesX"))
    cGlovesY = Val(IO_INI_Read(SkinPath, "EQUIPPED", "GlovesY"))
    cPantsX = Val(IO_INI_Read(SkinPath, "EQUIPPED", "PantsX"))
    cPantsY = Val(IO_INI_Read(SkinPath, "EQUIPPED", "PantsY"))
    cShoesX = Val(IO_INI_Read(SkinPath, "EQUIPPED", "ShoesX"))
    cShoesY = Val(IO_INI_Read(SkinPath, "EQUIPPED", "ShoesY"))

    'Init the graphic
    Graphics_SetGrh cGrh, Val(IO_INI_Read(SkinPath, "EQUIPPED", "Grh")), ANIMTYPE_LOOP

End Sub

Public Sub MoveCursor(ByVal X As Integer, ByVal Y As Integer)
'*********************************************************************************
'Handles the cursor moving over the window
'*********************************************************************************
Dim Slot As Byte
Dim Row As Byte
Dim Column As Byte

    'Set the window as the new target window
    FocusWID = WID_EquipWindow
    
    'Convert the X and Y values to relative to the window
    '(where 0,0 is the top-left of the window, not screen)
    X = X - cX
    Y = Y - cY
    
    'Find the slot the cursor is over
    If Math_Collision_PointRect(X, Y, cCapX, cCapY, 32, 32) Then
        TooltipItemIndex = CapItemIndex
    ElseIf Math_Collision_PointRect(X, Y, cForeheadX, cForeheadY, 32, 32) Then
        TooltipItemIndex = ForeheadItemIndex
    ElseIf Math_Collision_PointRect(X, Y, cRing1X, cRing1Y, 32, 32) Then
        TooltipItemIndex = Ring1ItemIndex
    ElseIf Math_Collision_PointRect(X, Y, cRing2X, cRing2Y, 32, 32) Then
        TooltipItemIndex = Ring2ItemIndex
    ElseIf Math_Collision_PointRect(X, Y, cRing3X, cRing3Y, 32, 32) Then
        TooltipItemIndex = Ring3ItemIndex
    ElseIf Math_Collision_PointRect(X, Y, cRing4X, cRing4Y, 32, 32) Then
        TooltipItemIndex = Ring4ItemIndex
    ElseIf Math_Collision_PointRect(X, Y, cEyeAccX, cEyeAccY, 32, 32) Then
        TooltipItemIndex = EyeAccItemIndex
    ElseIf Math_Collision_PointRect(X, Y, cEarAccX, cEarAccY, 32, 32) Then
        TooltipItemIndex = EarAccItemIndex
    ElseIf Math_Collision_PointRect(X, Y, cMantleX, cMantleY, 32, 32) Then
        TooltipItemIndex = MantleItemIndex
    ElseIf Math_Collision_PointRect(X, Y, cClothesX, cClothesY, 32, 32) Then
        TooltipItemIndex = ClothesItemIndex
    ElseIf Math_Collision_PointRect(X, Y, cPendantX, cPendantY, 32, 32) Then
        TooltipItemIndex = PendantItemIndex
    ElseIf Math_Collision_PointRect(X, Y, cWeaponX, cWeaponY, 32, 32) Then
        TooltipItemIndex = WeaponItemIndex
    ElseIf Math_Collision_PointRect(X, Y, cGlovesX, cGlovesY, 32, 32) Then
        TooltipItemIndex = GlovesItemIndex
    ElseIf Math_Collision_PointRect(X, Y, cPantsX, cPantsY, 32, 32) Then
        TooltipItemIndex = PantsItemIndex
    ElseIf Math_Collision_PointRect(X, Y, cShieldX, cShieldY, 32, 32) Then
        TooltipItemIndex = ShieldItemIndex
    ElseIf Math_Collision_PointRect(X, Y, cShoesX, cShoesY, 32, 32) Then
        TooltipItemIndex = ShoesItemIndex
    End If
    
End Sub

Public Sub LeftClick(ByVal X As Integer, ByVal Y As Integer)
'*********************************************************************************
'Handles left-clicking the window
'*********************************************************************************
    
    'Set the window as the new target window
    FocusWID = WID_EquipWindow
    
    'Convert the X and Y values to relative to the window
    '(where 0,0 is the top-left of the window, not screen)
    X = X - cX
    Y = Y - cY

    'Check for a valid inventory swap slot
    If InvSwapSlot > -1 Then
    
        'Check for a valid item type (something you can equip)
        If Items(UserInv(InvSwapSlot).ItemIndex).ItemType <> ITEMTYPE_USEONCE Then
        
            'Check if a slot was clicked
            If Math_Collision_PointRect(X, Y, cCapX, cCapY, 32, 32) Then
                sndBuf.Put_Byte PId.CS_UseInv
                sndBuf.Put_Byte InvSwapSlot
                InvSwapSlot = -1
            ElseIf Math_Collision_PointRect(X, Y, cForeheadX, cForeheadY, 32, 32) Then
                sndBuf.Put_Byte PId.CS_UseInv
                sndBuf.Put_Byte InvSwapSlot
                InvSwapSlot = -1
            ElseIf Math_Collision_PointRect(X, Y, cRing1X, cRing1Y, 32, 32) Then
                sndBuf.Put_Byte PId.CS_EquipRing
                sndBuf.Put_Byte InvSwapSlot
                sndBuf.Put_Byte 1
                InvSwapSlot = -1
            ElseIf Math_Collision_PointRect(X, Y, cRing2X, cRing2Y, 32, 32) Then
                sndBuf.Put_Byte PId.CS_EquipRing
                sndBuf.Put_Byte InvSwapSlot
                sndBuf.Put_Byte 2
                InvSwapSlot = -1
            ElseIf Math_Collision_PointRect(X, Y, cRing3X, cRing3Y, 32, 32) Then
                sndBuf.Put_Byte PId.CS_EquipRing
                sndBuf.Put_Byte InvSwapSlot
                sndBuf.Put_Byte 3
                InvSwapSlot = -1
            ElseIf Math_Collision_PointRect(X, Y, cRing4X, cRing4Y, 32, 32) Then
                sndBuf.Put_Byte PId.CS_EquipRing
                sndBuf.Put_Byte InvSwapSlot
                sndBuf.Put_Byte 4
                InvSwapSlot = -1
            ElseIf Math_Collision_PointRect(X, Y, cEyeAccX, cEyeAccY, 32, 32) Then
                sndBuf.Put_Byte PId.CS_UseInv
                sndBuf.Put_Byte InvSwapSlot
                InvSwapSlot = -1
            ElseIf Math_Collision_PointRect(X, Y, cEarAccX, cEarAccY, 32, 32) Then
                sndBuf.Put_Byte PId.CS_UseInv
                sndBuf.Put_Byte InvSwapSlot
                InvSwapSlot = -1
            ElseIf Math_Collision_PointRect(X, Y, cMantleX, cMantleY, 32, 32) Then
                sndBuf.Put_Byte PId.CS_UseInv
                sndBuf.Put_Byte InvSwapSlot
                InvSwapSlot = -1
            ElseIf Math_Collision_PointRect(X, Y, cClothesX, cClothesY, 32, 32) Then
                sndBuf.Put_Byte PId.CS_UseInv
                sndBuf.Put_Byte InvSwapSlot
                InvSwapSlot = -1
            ElseIf Math_Collision_PointRect(X, Y, cPendantX, cPendantY, 32, 32) Then
                sndBuf.Put_Byte PId.CS_UseInv
                sndBuf.Put_Byte InvSwapSlot
                InvSwapSlot = -1
            ElseIf Math_Collision_PointRect(X, Y, cWeaponX, cWeaponY, 32, 32) Then
                sndBuf.Put_Byte PId.CS_UseInv
                sndBuf.Put_Byte InvSwapSlot
                InvSwapSlot = -1
            ElseIf Math_Collision_PointRect(X, Y, cShieldX, cShieldY, 32, 32) Then
                sndBuf.Put_Byte PId.CS_UseInv
                sndBuf.Put_Byte InvSwapSlot
                InvSwapSlot = -1
            ElseIf Math_Collision_PointRect(X, Y, cGlovesX, cGlovesY, 32, 32) Then
                sndBuf.Put_Byte PId.CS_UseInv
                sndBuf.Put_Byte InvSwapSlot
                InvSwapSlot = -1
            ElseIf Math_Collision_PointRect(X, Y, cPantsX, cPantsY, 32, 32) Then
                sndBuf.Put_Byte PId.CS_UseInv
                sndBuf.Put_Byte InvSwapSlot
                InvSwapSlot = -1
            ElseIf Math_Collision_PointRect(X, Y, cShoesX, cShoesY, 32, 32) Then
                sndBuf.Put_Byte PId.CS_UseInv
                sndBuf.Put_Byte InvSwapSlot
                InvSwapSlot = -1
            Else
            
                'No slots were clicked, so select the window
                SelectedWID = WID_EquipWindow
                
            End If
        
        End If
        
    Else
     
        'No inventory swap slot, so select the window
        SelectedWID = WID_EquipWindow

    End If

End Sub

Public Sub RightClick(ByVal X As Integer, ByVal Y As Integer)
'*********************************************************************************
'Handles right-clicking on the window
'*********************************************************************************
Dim Slot As Byte

    'Set the window as the new target window
    FocusWID = WID_EquipWindow
    
    'Convert the X and Y values to relative to the window
    '(where 0,0 is the top-left of the window, not screen)
    X = X - cX
    Y = Y - cY
    
    'Check for collision against different slots
    If Math_Collision_PointRect(X, Y, cCapX, cCapY, 32, 32) Then
        If CapItemIndex > 0 Then
            sndBuf.Put_Byte PId.CS_UnEquip
            sndBuf.Put_Byte EQUIPSLOT_CAP
        End If
    ElseIf Math_Collision_PointRect(X, Y, cForeheadX, cForeheadY, 32, 32) Then
        If ForeheadItemIndex > 0 Then
            sndBuf.Put_Byte PId.CS_UnEquip
            sndBuf.Put_Byte EQUIPSLOT_FOREHEAD
        End If
    ElseIf Math_Collision_PointRect(X, Y, cRing1X, cRing1Y, 32, 32) Then
        If Ring1ItemIndex > 0 Then
            sndBuf.Put_Byte PId.CS_UnEquip
            sndBuf.Put_Byte EQUIPSLOT_RING1
        End If
    ElseIf Math_Collision_PointRect(X, Y, cRing2X, cRing2Y, 32, 32) Then
        If Ring2ItemIndex > 0 Then
            sndBuf.Put_Byte PId.CS_UnEquip
            sndBuf.Put_Byte EQUIPSLOT_RING2
        End If
    ElseIf Math_Collision_PointRect(X, Y, cRing3X, cRing3Y, 32, 32) Then
        If Ring3ItemIndex > 0 Then
            sndBuf.Put_Byte PId.CS_UnEquip
            sndBuf.Put_Byte EQUIPSLOT_RING3
        End If
    ElseIf Math_Collision_PointRect(X, Y, cRing4X, cRing4Y, 32, 32) Then
        If Ring4ItemIndex > 0 Then
            sndBuf.Put_Byte PId.CS_UnEquip
            sndBuf.Put_Byte EQUIPSLOT_RING4
        End If
    ElseIf Math_Collision_PointRect(X, Y, cEyeAccX, cEyeAccY, 32, 32) Then
        If EyeAccItemIndex > 0 Then
            sndBuf.Put_Byte PId.CS_UnEquip
            sndBuf.Put_Byte EQUIPSLOT_EYEACC
        End If
    ElseIf Math_Collision_PointRect(X, Y, cEarAccX, cEarAccY, 32, 32) Then
        If EarAccItemIndex > 0 Then
            sndBuf.Put_Byte PId.CS_UnEquip
            sndBuf.Put_Byte EQUIPSLOT_EARACC
        End If
    ElseIf Math_Collision_PointRect(X, Y, cMantleX, cMantleY, 32, 32) Then
        If MantleItemIndex > 0 Then
            sndBuf.Put_Byte PId.CS_UnEquip
            sndBuf.Put_Byte EQUIPSLOT_MANTLE
        End If
    ElseIf Math_Collision_PointRect(X, Y, cClothesX, cClothesY, 32, 32) Then
        If ClothesItemIndex > 0 Then
            sndBuf.Put_Byte PId.CS_UnEquip
            sndBuf.Put_Byte EQUIPSLOT_CLOTHES
        End If
    ElseIf Math_Collision_PointRect(X, Y, cPendantX, cPendantY, 32, 32) Then
        If PendantItemIndex > 0 Then
            sndBuf.Put_Byte PId.CS_UnEquip
            sndBuf.Put_Byte EQUIPSLOT_PENDANT
        End If
    ElseIf Math_Collision_PointRect(X, Y, cWeaponX, cWeaponY, 32, 32) Then
        If WeaponItemIndex > 0 Then
            sndBuf.Put_Byte PId.CS_UnEquip
            sndBuf.Put_Byte EQUIPSLOT_WEAPON
        End If
    ElseIf Math_Collision_PointRect(X, Y, cShieldX, cShieldY, 32, 32) Then
        If ShieldItemIndex > 0 Then
            sndBuf.Put_Byte PId.CS_UnEquip
            sndBuf.Put_Byte EQUIPSLOT_SHIELD
        End If
    ElseIf Math_Collision_PointRect(X, Y, cGlovesX, cGlovesY, 32, 32) Then
        If GlovesItemIndex > 0 Then
            sndBuf.Put_Byte PId.CS_UnEquip
            sndBuf.Put_Byte EQUIPSLOT_GLOVES
        End If
    ElseIf Math_Collision_PointRect(X, Y, cPantsX, cPantsY, 32, 32) Then
        If PantsItemIndex > 0 Then
            sndBuf.Put_Byte PId.CS_UnEquip
            sndBuf.Put_Byte EQUIPSLOT_PANTS
        End If
    ElseIf Math_Collision_PointRect(X, Y, cShoesX, cShoesY, 32, 32) Then
        If ShoesItemIndex > 0 Then
            sndBuf.Put_Byte PId.CS_UnEquip
            sndBuf.Put_Byte EQUIPSLOT_SHOES
        End If
    End If
    
End Sub

Public Sub Draw()
'*********************************************************************************
'Draw the equipment window
'*********************************************************************************

    'Confirm is visible
    If Not Visible Then Exit Sub

    'Window background
    Graphics_DrawGrh cGrh, cX, cY
    
    'Draw all the slots
    Graphics_DrawGrh CapGrh, cX + cCapX, cY + cCapY
    Graphics_DrawGrh ForeheadGrh, cX + cForeheadX, cY + cForeheadY
    Graphics_DrawGrh Ring1Grh, cX + cRing1X, cY + cRing1Y
    Graphics_DrawGrh Ring2Grh, cX + cRing2X, cY + cRing2Y
    Graphics_DrawGrh Ring3Grh, cX + cRing3X, cY + cRing3Y
    Graphics_DrawGrh Ring4Grh, cX + cRing4X, cY + cRing4Y
    Graphics_DrawGrh EyeAccGrh, cX + cEyeAccX, cY + cEyeAccY
    Graphics_DrawGrh EarAccGrh, cX + cEarAccX, cY + cEarAccY
    Graphics_DrawGrh MantleGrh, cX + cMantleX, cY + cMantleY
    Graphics_DrawGrh ClothesGrh, cX + cClothesX, cY + cClothesY
    Graphics_DrawGrh PendantGrh, cX + cPendantX, cY + cPendantY
    Graphics_DrawGrh WeaponGrh, cX + cWeaponX, cY + cWeaponY
    Graphics_DrawGrh ShieldGrh, cX + cShieldX, cY + cShieldY
    Graphics_DrawGrh GlovesGrh, cX + cGlovesX, cY + cGlovesY
    Graphics_DrawGrh PantsGrh, cX + cPantsX, cY + cPantsY
    Graphics_DrawGrh ShoesGrh, cX + cShoesX, cY + cShoesY
    
End Sub
