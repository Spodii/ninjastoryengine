VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GUIInvWindow"
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

'If the window is visible
Public cVisible As Boolean

'Slot positioning information
Private cFirstSlotX As Integer
Private cFirstSlotY As Integer
Private cSlotsPerRow As Byte
Private cGapX As Integer
Private cGapY As Integer

'Graphic information
Private cGrh As tGrh

Public Property Get Visible() As Boolean
    Visible = cVisible
End Property

Public Property Let Visible(ByVal Value As Boolean)
    cVisible = Value
    If cVisible Then
        FocusWID = WID_InvWindow
    Else
        InvSwapSlot = -1
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
'Load the inventory window
'*********************************************************************************

    'Set the skin path
    SkinPath = App.Path & "\Data\GUI\" & GUIName & ".ini"

    'Set the window information
    cX = Val(IO_INI_Read(SkinPath, "INVENTORY", "X"))
    cY = Val(IO_INI_Read(SkinPath, "INVENTORY", "Y"))
    cWidth = Val(IO_INI_Read(SkinPath, "INVENTORY", "Width"))
    cHeight = Val(IO_INI_Read(SkinPath, "INVENTORY", "Height"))
    
    'Set the slot positioning information
    cFirstSlotX = Val(IO_INI_Read(SkinPath, "INVENTORY", "FirstSlotX"))
    cFirstSlotY = Val(IO_INI_Read(SkinPath, "INVENTORY", "FirstSlotY"))
    cSlotsPerRow = Val(IO_INI_Read(SkinPath, "INVENTORY", "SlotsPerRow"))
    cGapX = Val(IO_INI_Read(SkinPath, "INVENTORY", "GapX"))
    cGapY = Val(IO_INI_Read(SkinPath, "INVENTORY", "GapY"))
    
    'Init the graphic
    Graphics_SetGrh cGrh, Val(IO_INI_Read(SkinPath, "INVENTORY", "Grh")), ANIMTYPE_LOOP

End Sub

Public Sub MoveCursor(ByVal X As Integer, ByVal Y As Integer)
'*********************************************************************************
'Handles the cursor moving over the window
'*********************************************************************************
Dim Slot As Byte
Dim Row As Byte
Dim Column As Byte

    'Set the window as the new target window
    FocusWID = WID_InvWindow
    
    'Convert the X and Y values to relative to the window
    '(where 0,0 is the top-left of the window, not screen)
    X = X - cX
    Y = Y - cY
    
    'Find the slot the cursor is over
    For Slot = 0 To USERINVSIZE
        Column = Slot \ cSlotsPerRow
        Row = Slot - (Column * cSlotsPerRow)
        If Math_Collision_PointRect(X, Y, cFirstSlotX + (Row * cGapX), cFirstSlotY + (Column * cGapY), 32, 32) Then
            TooltipItemIndex = UserInv(Slot).ItemIndex
            Exit Sub
        End If
    Next Slot
    
End Sub

Public Sub LeftClick(ByVal X As Integer, ByVal Y As Integer)
'*********************************************************************************
'Handles left-clicking the window
'*********************************************************************************
Dim Slot As Byte
Dim Row As Byte
Dim Column As Byte

    'Set the window as the new target window
    FocusWID = WID_InvWindow
    
    'Convert the X and Y values to relative to the window
    '(where 0,0 is the top-left of the window, not screen)
    X = X - cX
    Y = Y - cY
    
    'Loop through the slots to see if any were clicked
    For Slot = 0 To USERINVSIZE
        Column = Slot \ cSlotsPerRow
        Row = Slot - (Column * cSlotsPerRow)
        If Math_Collision_PointRect(X, Y, cFirstSlotX + (Row * cGapX), cFirstSlotY + (Column * cGapY), 32, 32) Then
            
            'Check if we are swapping or picking up
            If InvSwapSlot = -1 Then
                
                'Set the start swap slot
                If UserInv(Slot).ItemIndex > 0 Then
                    If UserInv(Slot).Amount > 0 Then
                        InvSwapSlot = Slot
                        Exit Sub
                    End If
                End If
                
            Else
            
                'Swap with the new slot
                If InvSwapSlot <> Slot Then
                    sndBuf.Put_Byte PId.CS_MoveInvSlot
                    sndBuf.Put_Byte InvSwapSlot
                    sndBuf.Put_Byte Slot
                    InvSwapSlot = -1
                Else
                    InvSwapSlot = -1
                End If
                
            End If
            
        End If
    Next Slot
    
    'No key parts of the window were clicked, so select the window
    SelectedWID = WID_InvWindow

End Sub

Public Sub RightClick(ByVal X As Integer, ByVal Y As Integer)
'*********************************************************************************
'Handles right-clicking on the window
'*********************************************************************************
Dim Slot As Byte
Dim Row As Byte
Dim Column As Byte

    'Set the window as the new target window
    FocusWID = WID_InvWindow
    
    'Convert the X and Y values to relative to the window
    '(where 0,0 is the top-left of the window, not screen)
    X = X - cX
    Y = Y - cY
    
    'Loop through the slots to see if any were clicked
    For Slot = 0 To USERINVSIZE
        Column = Slot \ cSlotsPerRow
        Row = Slot - (Column * cSlotsPerRow)
        If Math_Collision_PointRect(X, Y, cFirstSlotX + (Row * cGapX), cFirstSlotY + (Column * cGapY), 32, 32) Then
            
            'The slot was clicked, so if anything is in the slot, tell the server we want to use it
            If UserInv(Slot).ItemIndex > 0 Then
                If UserInv(Slot).Amount > 0 Then
                    sndBuf.Put_Byte PId.CS_UseInv
                    sndBuf.Put_Byte Slot
                    Exit Sub
                End If
            End If
            
        End If
    Next Slot

End Sub

Public Sub Draw()
'*********************************************************************************
'Draw the inventory window
'*********************************************************************************
Const SwapColor As Long = -1761607681
Dim Row As Byte
Dim Column As Byte
Dim Slot As Byte
Dim X As Integer
Dim Y As Integer

    'Confirm is visible
    If Not Visible Then Exit Sub
    
    'Draw the back
    Graphics_DrawGrh cGrh, cX, cY
    
    'Draw the inventory
    For Slot = 0 To USERINVSIZE
        If UserInv(Slot).ItemIndex > 0 Then
            Column = Slot \ cSlotsPerRow
            Row = Slot - (Column * cSlotsPerRow)
            X = cX + cFirstSlotX + (Row * cGapX)
            Y = cY + cFirstSlotY + (Column * cGapY)
            If InvSwapSlot = Slot Then
                Graphics_DrawGrh UserInv(Slot).Grh, X, Y, SwapColor, SwapColor, SwapColor, SwapColor
            Else
                Graphics_DrawGrh UserInv(Slot).Grh, X, Y
            End If
            FontDefault.Draw UserInv(Slot).Amount, X, Y, D3DColorARGB(255, 0, 255, 0)
        End If
    Next Slot

End Sub
