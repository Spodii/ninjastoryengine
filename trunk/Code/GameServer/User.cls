VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*********************************************************************************
'Handles individual users
'*********************************************************************************

Option Explicit

'How frequently the position is updated when the user is moving
Private Const UPDATEPOSRATE As Long = 3000

'Temporary/server-related information
Private cUserIndex As Long
Private cOnGround As Byte
Private cJump As Single
Private cLastUpdateTime As Long
Private cLastTileX As Integer
Private cLastTileY As Integer
Private cMoveDir As Byte
Private cPosUpdateTime As Long

'User's current action
Private cAction As eCharAction  'ID of the action
Private cActionTime As Long     'At what server tick the action will be over

'If the mod stats need to be updated (happens when one of the base stats change that
'can affect the mod stats or an item is equipped/removed)
Private cUpdateModStats As Boolean

'When the user last picked up an item off the ground
Private cLastPickedUpItem As Long

'Packet buffer information
Private Const BufferScaleSize As Integer = 128
Private BufferCopyPos As Integer
Private BufferSize As Integer
Private Buffer() As Byte

'User flags
Private Type tUserFlags
    Status As Long
End Type

'User's HP and MP percentages
Private cHPP As Byte
Private cMPP As Byte

'Last sent HP and MP percentages (send route is to the whole map)
Private cLastHPP As Byte
Private cLastMPP As Byte

'Inventory
Private Inv(0 To USERINVSIZE) As tServerInvSlot

'Character-related information
Private Type UserInfo
    Name As String          'The user's name
    CharIndex As Integer    'Character's index
    Body As Byte            'Body index
    Heading As Byte         'Current direction facing
    Map As Integer          'Map index the user is on
    X As Single             'X pixel co-ordinate
    Y As Single             'Y pixel co-ordinate
    WeaponItem As Integer   'Item index of the weapon the user has equipped
    CapItem As Integer      'Item indexes for the rest of the equipped stuff below
    ForeheadItem As Integer
    Ring1Item As Integer
    Ring2Item As Integer
    Ring3Item As Integer
    Ring4Item As Integer
    EyeAccItem As Integer
    EarAccItem As Integer
    GlovesItem As Integer
    PantsItem As Integer
    ShoesItem As Integer
    ShieldItem As Integer
    MantleItem As Integer
    ClothesItem As Integer
    PendantItem As Integer
    Flags As tUserFlags     'User flags
    Stats As tUserStats     'User stats
    LastSentStats As tUserStats 'The values of the user's stats that the client holds (DO NOT EDIT)
End Type
Dim u As UserInfo

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Sub DropInvItem(ByVal InvSlot As Byte)
'*********************************************************************************
'User drops an item from their inventory
'*********************************************************************************

    'Check for a valid slot
    If InvSlot > USERINVSIZE Then Exit Sub
    If Inv(InvSlot).ItemIndex < 1 Then Exit Sub
    
    'Add the item to the map
    Maps(u.Map).AddItem Inv(InvSlot).ItemIndex, Inv(InvSlot).Amount, X, Y + Height
    
    'Erase the item from the user's inventory
    Inv(InvSlot).ItemIndex = 0
    Inv(InvSlot).Amount = 0
    
    'Update the inventory slot on the client
    SendInvSlot InvSlot

End Sub

Public Sub Pickup(ByVal ItemSlot As Integer)
'*********************************************************************************
'Pick up an item off of the ground
'*********************************************************************************
Dim tmpItemIndex As Integer
Dim tmpItemAmount As Integer

    'Check that enough time has elapsed since the user last tried to pick something up
    If cLastPickedUpItem + PICKUPITEMTIME - 25 < timeGetTime Then
    
        'Set the pickup timer
        cLastPickedUpItem = timeGetTime
    
        'Check that the item is valid (this will also confirm the slot is valid)
        If Maps(u.Map).MapItemIndex(ItemSlot) > 0 Then
        
            'Check for a legal pickup distance
            If Math_Collision_PointRect(Maps(u.Map).MapItemX(ItemSlot), Maps(u.Map).MapItemY(ItemSlot), _
                X + Width \ 2 - (PICKUPITEMDISTSERVER \ 2), Y + Height \ 2 - (PICKUPITEMDISTSERVER \ 2), _
                PICKUPITEMDISTSERVER, PICKUPITEMDISTSERVER) Then

                'Give the item to the user (remember, AddToInv parameters are ByRef, so we can
                'see if anything was actually picked up, but we don't want to directly modify the
                'map values, so we do not use them directly)
                tmpItemIndex = Maps(u.Map).MapItemIndex(ItemSlot)
                tmpItemAmount = Maps(u.Map).MapItemAmount(ItemSlot)
                AddToInv tmpItemIndex, tmpItemAmount
                
                'Check that the user actually picked something up - if so, we take it all even if
                'not all of it fit in our inventory. If nothing was added to the inventory, we do not
                'pick it up, and instead return an "your inventory is full" message.
                If tmpItemAmount < Maps(u.Map).MapItemAmount(ItemSlot) Then
                    
                    'Erase the item off of the map
                    Maps(u.Map).RemoveItem ItemSlot, False
                    
                    'Send the pickup packet
                    conBuf.Clear
                    conBuf.Put_Byte PId.SC_Item_Pickup
                    conBuf.Put_Integer ItemSlot
                    conBuf.Put_Integer u.CharIndex
                    Data_Send ToMap, u.Map, conBuf.Get_Buffer()
                    
                Else
                
                    'Inventory is full
                    Data_Send ToIndex, cUserIndex, cMessage(9).Data()
                    
                End If
                
            End If
            
        End If
        
    End If
    
End Sub

Public Sub SwapInvSlots(ByVal SrcSlot As Byte, ByVal DestSlot As Byte)
'*********************************************************************************
'Swaps the contents of two inventory slots
'*********************************************************************************
Dim tmpSlot As tServerInvSlot

    'Check for valid values
    If SrcSlot > USERINVSIZE Then Exit Sub
    If DestSlot > USERINVSIZE Then Exit Sub
    
    'Swap
    tmpSlot = Inv(SrcSlot)
    Inv(SrcSlot) = Inv(DestSlot)
    Inv(DestSlot) = tmpSlot
    
    'Update
    SendInvSlot SrcSlot
    SendInvSlot DestSlot

End Sub

Public Sub UseInvItem(ByVal pSlot As Byte, Optional ByVal RingIndex As Byte = 1)
'*********************************************************************************
'Uses an inventory item as specified by the slot
'*********************************************************************************
Dim tItemIndex As Integer

    'Check for a valid item
    If pSlot > USERINVSIZE Then Exit Sub
    If Inv(pSlot).ItemIndex <= 0 Then Exit Sub
    If Inv(pSlot).Amount <= 0 Then Exit Sub

    With Items(Inv(pSlot).ItemIndex)
    
        'Find the item type
        Select Case .ItemType
        
        'Use a "use once" item (ie potions, food, etc)
        Case ITEMTYPE_USEONCE
            
            'Stat replenishing
            If .HP <> 0 Then HP = HP + .HP
            If .MP <> 0 Then MP = MP + .MP
            
            'Permanent stat raising
            If .Str <> 0 Then Str = Str + .Str
            If .Dex <> 0 Then Dex = Dex + .Dex
            If .Intl <> 0 Then Intl = Intl + .Intl
            If .Luk <> 0 Then Luk = Luk + .Luk
            If .MaxHP <> 0 Then MaxHP = MaxHP + .MaxHP
            If .MaxMP <> 0 Then MaxMP = MaxMP + .MaxMP
            
            'Decrease the item count
            Inv(pSlot).Amount = Inv(pSlot).Amount - 1
            If Inv(pSlot).Amount = 0 Then Inv(pSlot).ItemIndex = 0
            
            'Update the inventory slot
            SendInvSlot pSlot
        
        Case ITEMTYPE_WEAPON
        
            'First we store the weapon then remove it, then set the new weapon in case the chance that the user
            'is switching their weapon when the inventory is already full, this will allow us to have that one
            '"transitional" space for the swap
            tItemIndex = Inv(pSlot).ItemIndex
            
            'Erase the weapon slot
            Inv(pSlot).Amount = 0
            Inv(pSlot).ItemIndex = 0
            
            'Set the weapon
            SetWeapon tItemIndex
            
            'Update the changed slot
            SendInvSlot pSlot
        
        Case ITEMTYPE_CLOTHES

            'The rest of the equips work just like the weapon for the most part, so refer to the comments above
            tItemIndex = Inv(pSlot).ItemIndex
            Inv(pSlot).Amount = 0
            Inv(pSlot).ItemIndex = 0
            SetClothes tItemIndex
            SendInvSlot pSlot
            
        Case ITEMTYPE_CAP
        
            tItemIndex = Inv(pSlot).ItemIndex
            Inv(pSlot).Amount = 0
            Inv(pSlot).ItemIndex = 0
            SetCap tItemIndex
            SendInvSlot pSlot
            
        Case ITEMTYPE_EARACC
            
            tItemIndex = Inv(pSlot).ItemIndex
            Inv(pSlot).Amount = 0
            Inv(pSlot).ItemIndex = 0
            SetEarAcc tItemIndex
            SendInvSlot pSlot
        
        Case ITEMTYPE_EYEACC
        
            tItemIndex = Inv(pSlot).ItemIndex
            Inv(pSlot).Amount = 0
            Inv(pSlot).ItemIndex = 0
            SetEyeAcc tItemIndex
            SendInvSlot pSlot
        
        Case ITEMTYPE_FOREHEAD
        
            tItemIndex = Inv(pSlot).ItemIndex
            Inv(pSlot).Amount = 0
            Inv(pSlot).ItemIndex = 0
            SetForehead tItemIndex
            SendInvSlot pSlot
        
        Case ITEMTYPE_GLOVES
        
            tItemIndex = Inv(pSlot).ItemIndex
            Inv(pSlot).Amount = 0
            Inv(pSlot).ItemIndex = 0
            SetGloves tItemIndex
            SendInvSlot pSlot
        
        Case ITEMTYPE_MANTLE
        
            tItemIndex = Inv(pSlot).ItemIndex
            Inv(pSlot).Amount = 0
            Inv(pSlot).ItemIndex = 0
            SetMantle tItemIndex
            SendInvSlot pSlot

        Case ITEMTYPE_PANTS
        
            tItemIndex = Inv(pSlot).ItemIndex
            Inv(pSlot).Amount = 0
            Inv(pSlot).ItemIndex = 0
            SetPants tItemIndex
            SendInvSlot pSlot
        
        Case ITEMTYPE_PENDANT
            
            tItemIndex = Inv(pSlot).ItemIndex
            Inv(pSlot).Amount = 0
            Inv(pSlot).ItemIndex = 0
            SetPendant tItemIndex
            SendInvSlot pSlot
        
        Case ITEMTYPE_SHIELD
        
            tItemIndex = Inv(pSlot).ItemIndex
            Inv(pSlot).Amount = 0
            Inv(pSlot).ItemIndex = 0
            SetShield tItemIndex
            SendInvSlot pSlot
        
        Case ITEMTYPE_SHOES
        
            tItemIndex = Inv(pSlot).ItemIndex
            Inv(pSlot).Amount = 0
            Inv(pSlot).ItemIndex = 0
            SetShoes tItemIndex
            SendInvSlot pSlot
        
        Case ITEMTYPE_RING
        
            tItemIndex = Inv(pSlot).ItemIndex
            Inv(pSlot).Amount = 0
            Inv(pSlot).ItemIndex = 0
            Select Case RingIndex
            Case 2
                SetRing2 tItemIndex
            Case 3
                SetRing3 tItemIndex
            Case 4
                SetRing4 tItemIndex
            Case Else
                SetRing1 tItemIndex
            End Select
            SendInvSlot pSlot

        End Select
        
    End With

End Sub

Private Sub SendInvSlot(ByVal pSlot As Byte)
'*********************************************************************************
'Sends a single inventory slot to the client
'*********************************************************************************

    'Build the packet and send it
    conBuf.Clear
    conBuf.Put_Byte PId.SC_Inv_UpdateSlot
    conBuf.Put_Byte pSlot
    conBuf.Put_Integer Inv(pSlot).ItemIndex
    If Inv(pSlot).ItemIndex > 0 Then conBuf.Put_Integer Inv(pSlot).Amount
    Data_Send ToIndex, UserIndex, conBuf.Get_Buffer()

End Sub

Private Sub SendInv()
'*********************************************************************************
'Sends the complete inventory to the client
'*********************************************************************************
Dim Slot As Byte

    'Create the packet header
    conBuf.Clear
    conBuf.Put_Byte PId.SC_Inv_Update
    
    'Loop through the slots
    For Slot = 0 To USERINVSIZE
        
        'Send the index of the item for the slot
        conBuf.Put_Integer Inv(Slot).ItemIndex
        
        'Only send the amount if theres an item
        If Inv(Slot).ItemIndex > 0 Then
            conBuf.Put_Integer Inv(Slot).Amount
        End If
        
    Next Slot
    
    'Send the packet
    Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()

End Sub

Public Sub AddToInv(ByRef pItemIndex As Integer, ByRef pAmount As Integer)
'*********************************************************************************
'Adds an item to the user's inventory
'Be careful - this sub uses ByRef to let the parent sub know how much of the item
'was taken! If all the items were taken, pItemIndex and pAmount will return as 0.
'If only a partial amount is taken, pItemIndex will remain the same and pAmount
'will just be decreased by how many items are left that could not have been taken.
'*********************************************************************************
Dim Slot As Byte

    'Check for a valid item index
    If pItemIndex < 1 Then Exit Sub
    If pItemIndex > ItemsUBound Then Exit Sub

    'Check for the next slot that the item can go into
    For Slot = 0 To USERINVSIZE
    
        'Check for a used slot that we can stack on
        If Inv(Slot).ItemIndex = pItemIndex Then
            If Inv(Slot).Amount < Items(pItemIndex).Stacking Then
                
                'Add to the currently existing slot
                If Inv(Slot).Amount + pAmount > Items(pItemIndex).Stacking Then
                    
                    'Add as much as we can, then run the sub again
                    pAmount = pAmount - (Items(pItemIndex).Stacking - Inv(Slot).Amount)
                    Inv(Slot).Amount = Items(pItemIndex).Stacking
                    SendInvSlot Slot
                    AddToInv pItemIndex, pAmount
                    Exit Sub
                    
                Else
                
                    'Add it all, then exit
                    Inv(Slot).Amount = Inv(Slot).Amount + pAmount
                    SendInvSlot Slot
                    pAmount = 0
                    pItemIndex = 0
                    Exit Sub
                
                End If
                
            End If
        
        'Check for an unused slot
        ElseIf Inv(Slot).ItemIndex = 0 Then
        
            'Check if we surpass the stacking limit
            If pAmount > Items(pItemIndex).Stacking Then
            
                'Add as much as we can, then run the sub again
                pAmount = pAmount - (Items(pItemIndex).Stacking - Inv(Slot).Amount)
                Inv(Slot).ItemIndex = pItemIndex
                Inv(Slot).Amount = Items(pItemIndex).Stacking
                SendInvSlot Slot
                AddToInv pItemIndex, pAmount
                Exit Sub
            
            Else
                
                'Add the item to the slot, simple as that
                Inv(Slot).ItemIndex = pItemIndex
                Inv(Slot).Amount = pAmount
                SendInvSlot Slot
                
                'We added it all, so clear the parameters
                pAmount = 0
                pItemIndex = 0
                Exit Sub
        
            End If
        
        End If
        
    Next Slot
    
End Sub

Private Sub SendTNL()
'*********************************************************************************
'Sends the experience the user needs to the next level
'*********************************************************************************

    conBuf.Clear
    conBuf.Put_Byte PId.SC_User_ToNextLevel
    conBuf.Put_Long ToNextLevel
    Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()

End Sub

Public Sub SetWeapon(ByVal ItemIndex As Integer)
'*********************************************************************************
'Sets the weapon the user has equipped
'*********************************************************************************
Dim OldItemIndex As Integer

    'Check for a valid range
    If ItemIndex < 0 Then Exit Sub
    If ItemIndex > ItemsUBound Then Exit Sub
    
    'Store the old item index
    OldItemIndex = u.WeaponItem
    
    'Check for removing an item
    If ItemIndex = 0 Then
    
        'Remove the item
        If u.WeaponItem > 0 Then AddToInv u.WeaponItem, 1
        u.WeaponItem = 0
    
    Else
        
        'Check for a valid type - if valid, give the user the item they already have equipped back and set the new one
        If Items(ItemIndex).ItemType = ITEMTYPE_WEAPON Then
            If u.WeaponItem > 0 Then AddToInv u.WeaponItem, 1
            u.WeaponItem = ItemIndex
        End If

    End If
    
    'Check if the item index changed - if so, update the client and update the stats
    If OldItemIndex <> ItemIndex Then
        conBuf.Clear
        conBuf.Put_Byte PId.SC_SetEquipped
        conBuf.Put_Byte EQUIPSLOT_WEAPON
        conBuf.Put_Integer ItemIndex
        Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()
        UpdateModStats
    End If
    
End Sub

Public Sub SetCap(ByVal ItemIndex As Integer)
'*********************************************************************************
'Sets the cap the user has equipped - follows the same template as SetWeapon so look to that sub for comments
'*********************************************************************************
Dim OldItemIndex As Integer

    If ItemIndex < 0 Then Exit Sub
    If ItemIndex > ItemsUBound Then Exit Sub
    OldItemIndex = u.CapItem
    If ItemIndex = 0 Then
        If u.CapItem > 0 Then AddToInv u.CapItem, 1
        u.CapItem = 0
    Else
        If Items(ItemIndex).ItemType = ITEMTYPE_CAP Then
            If u.CapItem > 0 Then AddToInv u.CapItem, 1
            u.CapItem = ItemIndex
        End If
    End If
    If OldItemIndex <> ItemIndex Then
        conBuf.Clear
        conBuf.Put_Byte PId.SC_SetEquipped
        conBuf.Put_Byte EQUIPSLOT_CAP
        conBuf.Put_Integer ItemIndex
        Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()
        UpdateModStats
    End If

End Sub

Public Sub SetForehead(ByVal ItemIndex As Integer)
'*********************************************************************************
'Sets the forehead the user has equipped - follows the same template as SetWeapon so look to that sub for comments
'*********************************************************************************
Dim OldItemIndex As Integer

    If ItemIndex < 0 Then Exit Sub
    If ItemIndex > ItemsUBound Then Exit Sub
    OldItemIndex = u.ForeheadItem
    If ItemIndex = 0 Then
        If u.ForeheadItem > 0 Then AddToInv u.ForeheadItem, 1
        u.ForeheadItem = 0
    Else
        If Items(ItemIndex).ItemType = ITEMTYPE_FOREHEAD Then
            If u.ForeheadItem > 0 Then AddToInv u.ForeheadItem, 1
            u.ForeheadItem = ItemIndex
        End If
    End If
    If OldItemIndex <> ItemIndex Then
        conBuf.Clear
        conBuf.Put_Byte PId.SC_SetEquipped
        conBuf.Put_Byte EQUIPSLOT_FOREHEAD
        conBuf.Put_Integer ItemIndex
        Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()
        UpdateModStats
    End If

End Sub

Public Sub SetEyeAcc(ByVal ItemIndex As Integer)
'*********************************************************************************
'Sets the EyeAcc the user has equipped - follows the same template as SetWeapon so look to that sub for comments
'*********************************************************************************
Dim OldItemIndex As Integer

    If ItemIndex < 0 Then Exit Sub
    If ItemIndex > ItemsUBound Then Exit Sub
    OldItemIndex = u.EyeAccItem
    If ItemIndex = 0 Then
        If u.EyeAccItem > 0 Then AddToInv u.EyeAccItem, 1
        u.EyeAccItem = 0
    Else
        If Items(ItemIndex).ItemType = ITEMTYPE_EYEACC Then
            If u.EyeAccItem > 0 Then AddToInv u.EyeAccItem, 1
            u.EyeAccItem = ItemIndex
        End If
    End If
    If OldItemIndex <> ItemIndex Then
        conBuf.Clear
        conBuf.Put_Byte PId.SC_SetEquipped
        conBuf.Put_Byte EQUIPSLOT_EYEACC
        conBuf.Put_Integer ItemIndex
        Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()
        UpdateModStats
    End If

End Sub

Public Sub SetEarAcc(ByVal ItemIndex As Integer)
'*********************************************************************************
'Sets the EarAcc the user has equipped - follows the same template as SetWeapon so look to that sub for comments
'*********************************************************************************
Dim OldItemIndex As Integer

    If ItemIndex < 0 Then Exit Sub
    If ItemIndex > ItemsUBound Then Exit Sub
    OldItemIndex = u.EarAccItem
    If ItemIndex = 0 Then
        If u.EarAccItem > 0 Then AddToInv u.EarAccItem, 1
        u.EarAccItem = 0
    Else
        If Items(ItemIndex).ItemType = ITEMTYPE_EARACC Then
            If u.EarAccItem > 0 Then AddToInv u.EarAccItem, 1
            u.EarAccItem = ItemIndex
        End If
    End If
    If OldItemIndex <> ItemIndex Then
        conBuf.Clear
        conBuf.Put_Byte PId.SC_SetEquipped
        conBuf.Put_Byte EQUIPSLOT_EARACC
        conBuf.Put_Integer ItemIndex
        Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()
        UpdateModStats
    End If

End Sub

Public Sub SetGloves(ByVal ItemIndex As Integer)
'*********************************************************************************
'Sets the Gloves the user has equipped - follows the same template as SetWeapon so look to that sub for comments
'*********************************************************************************
Dim OldItemIndex As Integer

    If ItemIndex < 0 Then Exit Sub
    If ItemIndex > ItemsUBound Then Exit Sub
    OldItemIndex = u.GlovesItem
    If ItemIndex = 0 Then
        If u.GlovesItem > 0 Then AddToInv u.GlovesItem, 1
        u.GlovesItem = 0
    Else
        If Items(ItemIndex).ItemType = ITEMTYPE_GLOVES Then
            If u.GlovesItem > 0 Then AddToInv u.GlovesItem, 1
            u.GlovesItem = ItemIndex
        End If
    End If
    If OldItemIndex <> ItemIndex Then
        conBuf.Clear
        conBuf.Put_Byte PId.SC_SetEquipped
        conBuf.Put_Byte EQUIPSLOT_GLOVES
        conBuf.Put_Integer ItemIndex
        Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()
        UpdateModStats
    End If

End Sub

Public Sub SetPants(ByVal ItemIndex As Integer)
'*********************************************************************************
'Sets the Pants the user has equipped - follows the same template as SetWeapon so look to that sub for comments
'*********************************************************************************
Dim OldItemIndex As Integer

    If ItemIndex < 0 Then Exit Sub
    If ItemIndex > ItemsUBound Then Exit Sub
    OldItemIndex = u.PantsItem
    If ItemIndex = 0 Then
        If u.PantsItem > 0 Then AddToInv u.PantsItem, 1
        u.PantsItem = 0
    Else
        If Items(ItemIndex).ItemType = ITEMTYPE_PANTS Then
            If u.PantsItem > 0 Then AddToInv u.PantsItem, 1
            u.PantsItem = ItemIndex
        End If
    End If
    If OldItemIndex <> ItemIndex Then
        conBuf.Clear
        conBuf.Put_Byte PId.SC_SetEquipped
        conBuf.Put_Byte EQUIPSLOT_PANTS
        conBuf.Put_Integer ItemIndex
        Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()
        UpdateModStats
    End If

End Sub

Public Sub SetPendant(ByVal ItemIndex As Integer)
'*********************************************************************************
'Sets the Pendant the user has equipped - follows the same template as SetWeapon so look to that sub for comments
'*********************************************************************************
Dim OldItemIndex As Integer

    If ItemIndex < 0 Then Exit Sub
    If ItemIndex > ItemsUBound Then Exit Sub
    OldItemIndex = u.PendantItem
    If ItemIndex = 0 Then
        If u.PendantItem > 0 Then AddToInv u.PendantItem, 1
        u.PendantItem = 0
    Else
        If Items(ItemIndex).ItemType = ITEMTYPE_PENDANT Then
            If u.PendantItem > 0 Then AddToInv u.PendantItem, 1
            u.PendantItem = ItemIndex
        End If
    End If
    If OldItemIndex <> ItemIndex Then
        conBuf.Clear
        conBuf.Put_Byte PId.SC_SetEquipped
        conBuf.Put_Byte EQUIPSLOT_PENDANT
        conBuf.Put_Integer ItemIndex
        Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()
        UpdateModStats
    End If

End Sub

Public Sub SetRing1(ByVal ItemIndex As Integer)
'*********************************************************************************
'Sets the Ring1 the user has equipped - follows the same template as SetWeapon so look to that sub for comments
'*********************************************************************************
Dim OldItemIndex As Integer

    If ItemIndex < 0 Then Exit Sub
    If ItemIndex > ItemsUBound Then Exit Sub
    OldItemIndex = u.Ring1Item
    If ItemIndex = 0 Then
        If u.Ring1Item > 0 Then AddToInv u.Ring1Item, 1
        u.Ring1Item = 0
    Else
        If Items(ItemIndex).ItemType = ITEMTYPE_RING Then
            If u.Ring1Item > 0 Then AddToInv u.Ring1Item, 1
            u.Ring1Item = ItemIndex
        End If
    End If
    If OldItemIndex <> ItemIndex Then
        conBuf.Clear
        conBuf.Put_Byte PId.SC_SetEquipped
        conBuf.Put_Byte EQUIPSLOT_RING1
        conBuf.Put_Integer ItemIndex
        Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()
        UpdateModStats
    End If

End Sub

Public Sub SetRing2(ByVal ItemIndex As Integer)
'*********************************************************************************
'Sets the Ring2 the user has equipped - follows the same template as SetWeapon so look to that sub for comments
'*********************************************************************************
Dim OldItemIndex As Integer

    If ItemIndex < 0 Then Exit Sub
    If ItemIndex > ItemsUBound Then Exit Sub
    OldItemIndex = u.Ring2Item
    If ItemIndex = 0 Then
        If u.Ring2Item > 0 Then AddToInv u.Ring2Item, 1
        u.Ring2Item = 0
    Else
        If Items(ItemIndex).ItemType = ITEMTYPE_RING Then
            If u.Ring2Item > 0 Then AddToInv u.Ring2Item, 1
            u.Ring2Item = ItemIndex
        End If
    End If
    If OldItemIndex <> ItemIndex Then
        conBuf.Clear
        conBuf.Put_Byte PId.SC_SetEquipped
        conBuf.Put_Byte EQUIPSLOT_RING2
        conBuf.Put_Integer ItemIndex
        Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()
        UpdateModStats
    End If

End Sub

Public Sub SetRing3(ByVal ItemIndex As Integer)
'*********************************************************************************
'Sets the Ring3 the user has equipped - follows the same template as SetWeapon so look to that sub for comments
'*********************************************************************************
Dim OldItemIndex As Integer

    If ItemIndex < 0 Then Exit Sub
    If ItemIndex > ItemsUBound Then Exit Sub
    OldItemIndex = u.Ring3Item
    If ItemIndex = 0 Then
        If u.Ring3Item > 0 Then AddToInv u.Ring3Item, 1
        u.Ring3Item = 0
    Else
        If Items(ItemIndex).ItemType = ITEMTYPE_RING Then
            If u.Ring3Item > 0 Then AddToInv u.Ring3Item, 1
            u.Ring3Item = ItemIndex
        End If
    End If
    If OldItemIndex <> ItemIndex Then
        conBuf.Clear
        conBuf.Put_Byte PId.SC_SetEquipped
        conBuf.Put_Byte EQUIPSLOT_RING3
        conBuf.Put_Integer ItemIndex
        Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()
        UpdateModStats
    End If

End Sub

Public Sub SetRing4(ByVal ItemIndex As Integer)
'*********************************************************************************
'Sets the Ring4 the user has equipped - follows the same template as SetWeapon so look to that sub for comments
'*********************************************************************************
Dim OldItemIndex As Integer

    If ItemIndex < 0 Then Exit Sub
    If ItemIndex > ItemsUBound Then Exit Sub
    OldItemIndex = u.Ring4Item
    If ItemIndex = 0 Then
        If u.Ring4Item > 0 Then AddToInv u.Ring4Item, 1
        u.Ring4Item = 0
    Else
        If Items(ItemIndex).ItemType = ITEMTYPE_RING Then
            If u.Ring4Item > 0 Then AddToInv u.Ring4Item, 1
            u.Ring4Item = ItemIndex
        End If
    End If
    If OldItemIndex <> ItemIndex Then
        conBuf.Clear
        conBuf.Put_Byte PId.SC_SetEquipped
        conBuf.Put_Byte EQUIPSLOT_RING4
        conBuf.Put_Integer ItemIndex
        Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()
        UpdateModStats
    End If

End Sub

Public Sub SetShoes(ByVal ItemIndex As Integer)
'*********************************************************************************
'Sets the Shoes the user has equipped - follows the same template as SetWeapon so look to that sub for comments
'*********************************************************************************
Dim OldItemIndex As Integer

    If ItemIndex < 0 Then Exit Sub
    If ItemIndex > ItemsUBound Then Exit Sub
    OldItemIndex = u.ShoesItem
    If ItemIndex = 0 Then
        If u.ShoesItem > 0 Then AddToInv u.ShoesItem, 1
        u.ShoesItem = 0
    Else
        If Items(ItemIndex).ItemType = ITEMTYPE_SHOES Then
            If u.ShoesItem > 0 Then AddToInv u.ShoesItem, 1
            u.ShoesItem = ItemIndex
        End If
    End If
    If OldItemIndex <> ItemIndex Then
        conBuf.Clear
        conBuf.Put_Byte PId.SC_SetEquipped
        conBuf.Put_Byte EQUIPSLOT_SHOES
        conBuf.Put_Integer ItemIndex
        Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()
        UpdateModStats
    End If

End Sub

Public Sub SetShield(ByVal ItemIndex As Integer)
'*********************************************************************************
'Sets the Shield the user has equipped - follows the same template as SetWeapon so look to that sub for comments
'*********************************************************************************
Dim OldItemIndex As Integer

    If ItemIndex < 0 Then Exit Sub
    If ItemIndex > ItemsUBound Then Exit Sub
    OldItemIndex = u.ShieldItem
    If ItemIndex = 0 Then
        If u.ShieldItem > 0 Then AddToInv u.ShieldItem, 1
        u.ShieldItem = 0
    Else
        If Items(ItemIndex).ItemType = ITEMTYPE_SHIELD Then
            If u.ShieldItem > 0 Then AddToInv u.ShieldItem, 1
            u.ShieldItem = ItemIndex
        End If
    End If
    If OldItemIndex <> ItemIndex Then
        conBuf.Clear
        conBuf.Put_Byte PId.SC_SetEquipped
        conBuf.Put_Byte EQUIPSLOT_SHIELD
        conBuf.Put_Integer ItemIndex
        Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()
        UpdateModStats
    End If

End Sub

Public Sub SetMantle(ByVal ItemIndex As Integer)
'*********************************************************************************
'Sets the Mantle the user has equipped - follows the same template as SetWeapon so look to that sub for comments
'*********************************************************************************
Dim OldItemIndex As Integer

    If ItemIndex < 0 Then Exit Sub
    If ItemIndex > ItemsUBound Then Exit Sub
    OldItemIndex = u.MantleItem
    If ItemIndex = 0 Then
        If u.MantleItem > 0 Then AddToInv u.MantleItem, 1
        u.MantleItem = 0
    Else
        If Items(ItemIndex).ItemType = ITEMTYPE_MANTLE Then
            If u.MantleItem > 0 Then AddToInv u.MantleItem, 1
            u.MantleItem = ItemIndex
        End If
    End If
    If OldItemIndex <> ItemIndex Then
        conBuf.Clear
        conBuf.Put_Byte PId.SC_SetEquipped
        conBuf.Put_Byte EQUIPSLOT_MANTLE
        conBuf.Put_Integer ItemIndex
        Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()
        UpdateModStats
    End If

End Sub

Public Sub SetClothes(ByVal ItemIndex As Integer)
'*********************************************************************************
'Sets the Clothes the user has equipped - follows the same template as SetWeapon so look to that sub for comments
'*********************************************************************************
Dim OldItemIndex As Integer

    If ItemIndex < 0 Then Exit Sub
    If ItemIndex > ItemsUBound Then Exit Sub
    OldItemIndex = u.ClothesItem
    If ItemIndex = 0 Then
        If u.ClothesItem > 0 Then AddToInv u.ClothesItem, 1
        u.ClothesItem = 0
    Else
        If Items(ItemIndex).ItemType = ITEMTYPE_CLOTHES Then
            If u.ClothesItem > 0 Then AddToInv u.ClothesItem, 1
            u.ClothesItem = ItemIndex
        End If
    End If
    If OldItemIndex <> ItemIndex Then
        conBuf.Clear
        conBuf.Put_Byte PId.SC_SetEquipped
        conBuf.Put_Byte EQUIPSLOT_CLOTHES
        conBuf.Put_Integer ItemIndex
        Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()
        UpdateModStats
    End If

End Sub

Public Sub SendStats()
'*********************************************************************************
'Sends just EXP, Ryu, MP and HP changes
'This is updated much more frequently, and each stat has its own header, unlike the
'SendAllStatsToClient routine
'*********************************************************************************

    'Clear up the buffer
    conBuf.Clear

    'Check for EXP
    If u.LastSentStats.EXP <> u.Stats.EXP Then
        conBuf.Put_Byte PId.SC_User_EXP
        conBuf.Put_Long u.Stats.EXP
        u.LastSentStats.EXP = u.Stats.EXP
    End If
    
    'Check for Ryu
    If u.LastSentStats.Ryu <> u.Stats.Ryu Then
        conBuf.Put_Byte PId.SC_User_Ryu
        conBuf.Put_Long u.Stats.Ryu
        u.LastSentStats.Ryu = u.Stats.Ryu
    End If
    
    'Check for HP
    If u.LastSentStats.HP <> u.Stats.HP Then
        conBuf.Put_Byte PId.SC_User_HP
        conBuf.Put_Integer u.Stats.HP
        u.LastSentStats.HP = u.Stats.HP
    End If
    
    'Check for MP
    If u.LastSentStats.MP <> u.Stats.MP Then
        conBuf.Put_Byte PId.SC_User_MP
        conBuf.Put_Integer u.Stats.MP
        u.LastSentStats.MP = u.Stats.MP
    End If
    
    'If we made any changes, send them
    If conBuf.HasBuffer Then Data_Send ToIndex, UserIndex, conBuf.Get_Buffer()
    
    'Update the HP and MP percentages
    cHPP = CByte((u.Stats.HP / u.Stats.MaxHP) * 100)
    cMPP = CByte((u.Stats.MP / u.Stats.MaxMP) * 100)
    
    'Check if the values have changed
    If cHPP <> cLastHPP Or cMPP <> cLastMPP Then
        
        'Send the packet to the map
        conBuf.Clear
        conBuf.Put_Byte PId.SC_Char_HPMP
        conBuf.Put_Integer u.CharIndex
        conBuf.Put_Byte cHPP
        conBuf.Put_Byte cMPP
        Data_Send ToMapButIndex, u.CharIndex, conBuf.Get_Buffer()
        
        'Update the last sent values
        cLastHPP = cHPP
        cLastMPP = cMPP
        
    End If
    
End Sub

Public Sub SendAllStats()
'*********************************************************************************
'Sends all of the changed stats to the client
'Because of the rate EXP, Ryu, MP and HP changes, a separate routine is used
'*********************************************************************************
Dim Flags As Integer
Dim tHP As Integer
Dim tMP As Integer
Dim tRyu As Integer
Dim tEXP As Integer

    'Create the flags
    'Flags order:
    '0.  MaxHP
    '1.  MaxMP
    '2.  Level
    '3.  Str
    '4.  ModStr
    '5.  Dex
    '6.  ModDex
    '7.  Intl
    '8.  ModIntl
    '9.  Luk
    '10. ModLuk
    '11. MinHit
    '12. MaxHit
    Flags = 0
    
    With u.LastSentStats
    
        If .MaxHP <> u.Stats.MaxHP Then Flags = Flags Or 1
        If .MaxMP <> u.Stats.MaxMP Then Flags = Flags Or 2
        If .Level <> u.Stats.Level Then Flags = Flags Or 4
        If .Str <> u.Stats.Str Then Flags = Flags Or 8
        If .ModStr <> u.Stats.ModStr Then Flags = Flags Or 16
        If .Dex <> u.Stats.Dex Then Flags = Flags Or 32
        If .ModDex <> u.Stats.ModDex Then Flags = Flags Or 64
        If .Intl <> u.Stats.Intl Then Flags = Flags Or 128
        If .ModIntl <> u.Stats.ModIntl Then Flags = Flags Or 256
        If .Luk <> u.Stats.Luk Then Flags = Flags Or 512
        If .ModLuk <> u.Stats.ModLuk Then Flags = Flags Or 1024
        If .MinHit <> u.Stats.MinHit Then Flags = Flags Or 2048
        If .MaxHit <> u.Stats.MaxHit Then Flags = Flags Or 4096

    End With
    
    'Check if we have any stats to send
    If Flags = 0 Then Exit Sub
    
    'Start building the packet
    conBuf.Clear
    conBuf.Put_Byte PId.SC_User_Stats
    conBuf.Put_Integer Flags
    
    'Now we get to go back through and find out again what changed and put the values
    With u.Stats
        If Flags And 1 Then conBuf.Put_Integer .MaxHP
        If Flags And 2 Then conBuf.Put_Integer .MaxMP
        If Flags And 4 Then conBuf.Put_Integer .Level
        If Flags And 8 Then conBuf.Put_Integer .Str
        If Flags And 16 Then conBuf.Put_Integer .ModStr
        If Flags And 32 Then conBuf.Put_Integer .Dex
        If Flags And 64 Then conBuf.Put_Integer .ModDex
        If Flags And 128 Then conBuf.Put_Integer .Intl
        If Flags And 256 Then conBuf.Put_Integer .ModIntl
        If Flags And 512 Then conBuf.Put_Integer .Luk
        If Flags And 1024 Then conBuf.Put_Integer .ModLuk
        If Flags And 2048 Then conBuf.Put_Integer .MinHit
        If Flags And 4096 Then conBuf.Put_Integer .MaxHit
    End With
    
    'Send the packet
    Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()
    
    'Set the last sent stats to the current
    'We avoid updating the HP/MP/EXP/Ryu because we update that all in a separate routine
    tHP = u.LastSentStats.HP
    tMP = u.LastSentStats.MP
    tEXP = u.LastSentStats.EXP
    tRyu = u.LastSentStats.Ryu
    u.LastSentStats = u.Stats
    u.LastSentStats.HP = tHP
    u.LastSentStats.MP = tMP
    u.LastSentStats.EXP = tEXP
    u.LastSentStats.Ryu = tRyu

End Sub

Public Sub SetHitAction()
'*********************************************************************************
'Set the user's action to being hit
'*********************************************************************************

    'Set the action
    cAction = eHit
    cActionTime = timeGetTime + HITTIME
    
    'If the user is moving, stop them
    If MoveDir <> 0 Then MoveDir = 0

End Sub

Private Sub UpdateModStats()
'*********************************************************************************
'Updates the user's modified stats
'*********************************************************************************

    'Start by setting the mod values equal to the base value
    MaxHP = 10 + Items(u.WeaponItem).HP
    MaxMP = 10 + Items(u.WeaponItem).MP
    ModStr = Str + Items(u.WeaponItem).Str
    ModDex = Dex + Items(u.WeaponItem).Dex
    ModIntl = Intl + Items(u.WeaponItem).Intl
    ModLuk = Luk + Items(u.WeaponItem).Luk
    
    'Min/max hit
    MinHit = 1 + Str + Items(u.WeaponItem).MinHit
    MaxHit = 2 + Str + Items(u.WeaponItem).MaxHit
    
    'Defense
    Def = Items(u.WeaponItem).Def
    
    'All done updating
    cUpdateModStats = False

End Sub

Public Property Get StatusFlag(ByVal Flag As Long) As Boolean
'*********************************************************************************
'Return the status flag
'*********************************************************************************

    StatusFlag = (u.Flags.Status And Flag)

End Property

Public Property Let StatusFlag(ByVal Flag As Long, ByVal Value As Boolean)
'*********************************************************************************
'Set the status flag
'*********************************************************************************

    If Value Then
    
        'Enable
        If Not (u.Flags.Status And Flag) Then
            u.Flags.Status = u.Flags.Status Or Flag
        End If
        
    Else
    
        'Disable
        If (u.Flags.Status And Flag) Then
            u.Flags.Status = u.Flags.Status Xor Flag
        End If
        
    End If

End Property

Public Sub DamageByNPC(ByVal UserIndex As Integer, ByVal Damage As Long)
'*********************************************************************************
'NPC damages the user
'*********************************************************************************

    'Inflict the damage
    HP = HP - Damage
    
    'Check if the user's HP ran out
    If HP < 1 Then
        
        'Kill off the user
        Kill
        
    Else
    
        'Hit the user
        SetHitAction
    
    End If

End Sub

Public Sub Kill()
'*********************************************************************************
'Kill the user
'*********************************************************************************

    'Restore their HP and MP
    HP = MaxHP
    MP = MaxMP
    
    'Set the position
    SetPos 1, 50, 50

End Sub

Public Property Get Width() As Integer
'*********************************************************************************
'Returns the user's collision width
'*********************************************************************************

    Width = BodyInfo(u.Body).Width

End Property

Public Property Get Height() As Integer
'*********************************************************************************
'Returns the user's collision height
'*********************************************************************************

    Height = BodyInfo(u.Body).Height

End Property

Public Sub Punch(Optional ByVal HitCharIndex As Integer = 0)
'*********************************************************************************
'Makes the user punch
'*********************************************************************************
Dim HitDamage As Long
Dim HitIndex As Integer
Dim ValidHit As Boolean

    'Check if the user is already performing an action
    If cAction <> eNone Then Exit Sub
    
    'If the user is moving, stop them
    If MoveDir > 0 Then MoveDir = 0
    
    'Check for a valid hit index
    If HitCharIndex < 0 Then HitCharIndex = 0
    If HitCharIndex > LastChar Then HitCharIndex = 0
    If HitCharIndex > 0 Then
        If CharList(HitCharIndex).CharType = CHARTYPE_NPC Then
            HitIndex = CharList(HitCharIndex).Index
            If HitIndex <= LastNPC Then
                If HitIndex > 0 Then
                    If Not NPCList(HitIndex) Is Nothing Then
                        If NPCList(HitIndex).StatusFlag(NPCSTATUSFLAG_SPAWNED) Then
                            If NPCList(HitIndex).Map = u.Map Then
                                
                                'Check the collision area
                                If Math_Collision_Rect(NPCList(HitIndex).X, NPCList(HitIndex).Y, NPCList(HitIndex).Width, _
                                    NPCList(HitIndex).Height, u.X + Width - (BodyInfo(u.Body).PunchWidth * -(u.Heading = WEST)), _
                                    u.Y, BodyInfo(u.Body).PunchWidth, Height) Then
                                    ValidHit = True
                                    
                                    'Calculate the hit damage
                                    HitDamage = Math_Random(MinHit, MaxHit)
                                    If HitDamage > 32000 Then HitDamage = 32000
                                    
                                    'Damage the NPC
                                    NPCList(HitIndex).DamageByPC cUserIndex, HitDamage
                                    
                                End If
                            
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    'Set the user's action time
    cActionTime = timeGetTime + BodyInfo(u.Body).PunchTime
    
    'Set the user's action
    cAction = ePunch
    
    'Send the action packet
    If ValidHit Then
        conBuf.Clear
        conBuf.Put_Byte PId.SC_Punch_Hit
        conBuf.Put_Integer u.CharIndex
        conBuf.Put_Integer HitCharIndex
        conBuf.Put_Integer HitDamage
        Data_Send ToMap, u.Map, conBuf.Get_Buffer()
    Else
        conBuf.Clear
        conBuf.Put_Byte PId.SC_Punch
        conBuf.Put_Integer u.CharIndex
        Data_Send ToMap, u.Map, conBuf.Get_Buffer()
    End If
    
End Sub

Public Property Get Action() As eCharAction
'*********************************************************************************
'Get the user's current action
'*********************************************************************************

    Action = cAction

End Property

Public Sub Jump()
'*********************************************************************************
'Makes the user jump
'*********************************************************************************

    If cJump = 0 Then
        If Action <> eHit Then
            If cOnGround = 1 Then
                cJump = JUMPHEIGHT
                conBuf.Clear
                conBuf.Put_Byte PId.SC_Jump
                conBuf.Put_Integer u.CharIndex
                conBuf.Put_Integer u.X
                conBuf.Put_Integer u.Y
                Data_Send ToMap, u.Map, conBuf.Get_Buffer()
            End If
        End If
    End If

End Sub

Public Property Get MoveDir() As Byte
'*********************************************************************************
'Get the direction the user is moving
'*********************************************************************************

    MoveDir = cMoveDir
        
End Property

Public Property Let MoveDir(ByVal Value As Byte)
'*********************************************************************************
'Set the direction the user is moving
'*********************************************************************************

    'Check for a valid value
    If Value <> EAST Then
        If Value <> WEST Then
            
            'Make sure the value is 0 so we don't something funky
            Value = 0
            
        End If
    End If

    'Confirm the value has changed
    If cMoveDir <> Value Then
    
        'If starting to move, check if they are currently performing an action (if so, disallow it)
        If Value > 0 Then
            If Action <> eNone Then
                Exit Property
            End If
        End If
        
        'Reset the move update time
        cPosUpdateTime = timeGetTime + UPDATEPOSRATE

        'Clear the conversion buffer
        conBuf.Clear
        
        'Find the direction the user is going
        Select Case Value
            
            'Moving East
            Case EAST
                conBuf.Put_Byte PId.SC_Move_EastStart
            
            'Moving West
            Case WEST
                conBuf.Put_Byte PId.SC_Move_WestStart
                
            'Movement stopped
            Case Else
                
                'Check their old direction
                If cMoveDir = EAST Then
                    
                    'Was moving East
                    conBuf.Put_Byte PId.SC_Move_EastEnd
                    
                Else
                
                    'Was moving West
                    conBuf.Put_Byte PId.SC_Move_WestEnd
                    
                End If
                
        End Select

        'Attach the character index and the character's current
        'position then send it to everyone on the map
        conBuf.Put_Integer u.CharIndex
        conBuf.Put_Integer u.X
        conBuf.Put_Integer u.Y
        Data_Send ToMap, u.Map, conBuf.Get_Buffer()
            
        'Change the value
        cMoveDir = Value
        
        'Set the heading
        If Value <> 0 Then u.Heading = Value
    
    End If

End Property

Public Sub SetChar(Optional ByVal BodyIndex As Integer = -1)
'*********************************************************************************
'Sets the different visual paper-dolling values of the character
'*********************************************************************************
Dim UpdateBody As Boolean

    'Check for a new body
    If BodyIndex > -1 Then
        If u.Body <> BodyIndex Then
            u.Body = BodyIndex
            UpdateBody = True
        End If
    End If
    
    'Send the packets
    If UpdateBody Then
        conBuf.Clear
        conBuf.Put_Byte PId.SC_Char_SetPaperDoll
        conBuf.Put_Integer u.CharIndex
        conBuf.Put_Byte BodyIndex
        Data_Send ToMap, u.Map, conBuf.Get_Buffer()
    End If

End Sub

Public Sub Update()
'*********************************************************************************
'Updates the user
'*********************************************************************************

    'Check if the user's action time has run out
    If cAction <> eNone Then
        If cActionTime < timeGetTime Then
            cActionTime = 0
            cAction = eNone
        End If
    End If

    'Update mod stats
    If cUpdateModStats Then UpdateModStats

    'Update position
    UpdatePos

End Sub

Private Sub UpdatePos()
'*********************************************************************************
'Updates the user's position
'*********************************************************************************
Dim OldOnGround As Byte
Dim CheckTileX As Integer
Dim CheckTileY As Integer
Dim NewCheckTile As Integer
Dim UpdateX As Boolean
Dim UpdateY As Boolean
Dim TileChangeX As Integer
Dim i As Long
Dim ElapsedTime As Long
Dim RemainderX As Long

    'Get the elapsed time
    ElapsedTime = timeGetTime - cLastUpdateTime
    If ElapsedTime > 33 Then ElapsedTime = 33
    
    'Get the X co-ordinate remainder
    RemainderX = X Mod GRIDSIZE
    
    'Store the old OnGround
    OldOnGround = cOnGround

    'User is jumping
    If cJump > 0 Then
        Y = Y - (ElapsedTime * cJump * MOVESPEED)
        cJump = cJump - (ElapsedTime * JUMPDECAY)
        If cJump < 0 Then cJump = 0
        cOnGround = 0
    End If
    
    'Update the user's position
    Select Case MoveDir
        Case EAST
            X = X + ElapsedTime * MOVESPEED
        Case WEST
            X = X - ElapsedTime * MOVESPEED
    End Select
    
    'Check if the tile X has changed
    If cLastTileX <> X \ GRIDSIZE Then
        UpdateX = True
        TileChangeX = ((X \ GRIDSIZE) - cLastTileX)
        cLastTileX = X \ GRIDSIZE
    End If
    
    If UpdateX Then
    
        If Maps(u.Map).HasFloatingBlocks Then
        
            'Check for blocking to the right
            If TileChangeX > 0 Then
                CheckTileX = ((X + Width - 5) \ GRIDSIZE)
                If CheckTileX <= Maps(u.Map).TileWidth Then
                    For i = 0 To Height \ GRIDSIZE
                        CheckTileY = i + (Y \ GRIDSIZE)
                        If Maps(u.Map).TileInfo(CheckTileX, CheckTileY) = TILETYPE_BLOCKED Then
                            X = (CheckTileX * GRIDSIZE) - Width - 1
                            MoveDir = 0
                            cLastTileX = X \ GRIDSIZE
                            Exit For
                        End If
                    Next i
                End If
                
            'Check for blocking to the left
            Else
                CheckTileX = (X + 5) \ GRIDSIZE
                If CheckTileX <= Maps(u.Map).TileWidth Then
                    For i = 0 To Height \ GRIDSIZE
                        CheckTileY = i + (Y \ GRIDSIZE)
                        If Maps(u.Map).TileInfo(CheckTileX, CheckTileY) = TILETYPE_BLOCKED Then
                            X = ((CheckTileX + 1) * GRIDSIZE) + 1
                            MoveDir = 0
                            cLastTileX = X \ GRIDSIZE
                            Exit For
                        End If
                    Next i
                End If
            End If
            
        End If

        'Check if the user will be dropping
        CheckTileY = ((Y + Height) \ GRIDSIZE) + 1    'Get the tile below the user
        If CheckTileY <= Maps(u.Map).TileHeight Then
            For i = 0 To (Width + RemainderX) \ GRIDSIZE
                CheckTileX = i + (X \ GRIDSIZE)
                If Maps(u.Map).TileInfo(CheckTileX, CheckTileY) = TILETYPE_BLOCKED Or _
                    Maps(u.Map).TileInfo(CheckTileX, CheckTileY) = TILETYPE_PLATFORM Then
                    i = 0
                    Exit For
                End If
            Next i
            If i > 0 Then
                cOnGround = 0
            End If
        End If
        
    End If
    
    If cJump < 1 And cOnGround = 0 Then
        Y = Y + (ElapsedTime * MOVESPEED)
    End If
    
    'Check if the tile Y has changed
    If cLastTileY <> Y \ GRIDSIZE Then
        UpdateY = True
        cLastTileY = Y \ GRIDSIZE
    End If
    
    If cJump < 1 Then
    
        'Dropping handling
        If cJump = 0 Then
            NewCheckTile = ((Y + Height + 1) \ GRIDSIZE)
            If NewCheckTile <> CheckTileY Then
                For i = 0 To (Width + RemainderX) \ GRIDSIZE
                    CheckTileX = i + (X \ GRIDSIZE)
                    Select Case Maps(u.Map).TileInfo(CheckTileX, NewCheckTile)
                        Case TILETYPE_BLOCKED, TILETYPE_PLATFORM
                            Y = NewCheckTile * GRIDSIZE - Height - 1
                            cOnGround = 1
                            cJump = 0
                            Exit For
                    End Select
                Next i
            End If
        End If
        
    End If
    
    If cJump >= 1 Then
        If Maps(u.Map).HasFloatingBlocks Then
    
            'Head-hitting handling
            If UpdateY Then
                NewCheckTile = (Y \ GRIDSIZE)
                For i = 0 To (Width + RemainderX) \ GRIDSIZE
                    CheckTileX = i + (X \ GRIDSIZE)
                    If Maps(u.Map).TileInfo(CheckTileX, NewCheckTile) = TILETYPE_BLOCKED Then
                        Y = (NewCheckTile * GRIDSIZE) + GRIDSIZE - 1
                        cOnGround = 0
                        cJump = 0
                        Exit For
                    End If
                Next i
            End If
        
        End If
    End If
    
    'Check to update the user's position while they are moving
    If MoveDir <> 0 Then
        If cPosUpdateTime < timeGetTime Then
            cPosUpdateTime = timeGetTime + UPDATEPOSRATE
            conBuf.Clear
            conBuf.Put_Byte PId.SC_Char_UpdatePos
            conBuf.Put_Integer u.CharIndex
            conBuf.Put_Integer X
            conBuf.Put_Integer Y
            Data_Send ToPCArea, cUserIndex, conBuf.Get_Buffer()
        End If
    End If
    
End Sub

Public Sub AddMakeCharToBuffer()
'*********************************************************************************
'Adds the MakeChar (Char_MakePC) packet information to the conversion buffer
'*********************************************************************************
    
    Select Case cMoveDir
    Case WEST
        conBuf.Put_Byte PId.SC_Char_MakePC_MoveWest
    Case EAST
        conBuf.Put_Byte PId.SC_Char_MakePC_MoveEast
    Case Else
        conBuf.Put_Byte PId.SC_Char_MakePC
    End Select
    conBuf.Put_Integer u.CharIndex
    conBuf.Put_Integer u.X
    conBuf.Put_Integer u.Y
    conBuf.Put_String u.Name
    If cMoveDir = 0 Then conBuf.Put_Byte u.Heading
    conBuf.Put_Byte u.Body

End Sub

Public Sub SendMapInfo()
'*********************************************************************************
'Sends all the information about the map to the user (for when switching to a new map)
'This adds to the buffer, it does NOT actually send it
'*********************************************************************************
Dim MapChars() As Integer
Dim MapCharsUBound As Integer
Dim i As Long

    'Clear the buffer
    conBuf.Clear

    'Send the new map index
    conBuf.Put_Byte PId.SC_User_SetMap
    conBuf.Put_Integer u.Map

    'Send the users in the map
    Maps(u.Map).GetMapUsers MapChars, MapCharsUBound
    For i = 0 To MapCharsUBound
        If Not UserList(MapChars(i)) Is Nothing Then
            UserList(MapChars(i)).AddMakeCharToBuffer
        End If
    Next i
    
    'Send the NPCs in the map
    Maps(u.Map).GetMapNPCs MapChars, MapCharsUBound
    For i = 0 To MapCharsUBound
        If Not NPCList(MapChars(i)) Is Nothing Then
            NPCList(MapChars(i)).AddMakeCharToBuffer
        End If
    Next i
    
    'Items on the map
    For i = 0 To Maps(u.Map).NumMapItems
        If Maps(u.Map).MapItemIndex(i) > 0 Then
            conBuf.Put_Byte PId.SC_Item_Make
            conBuf.Put_Integer i
            conBuf.Put_Integer Maps(u.Map).MapItemIndex(i)
            conBuf.Put_Integer Maps(u.Map).MapItemAmount(i)
            conBuf.Put_Integer Maps(u.Map).MapItemX(i)
            conBuf.Put_Integer Maps(u.Map).MapItemY(i)
        End If
    Next i
    
    'Send the buffer
    If conBuf.HasBuffer Then Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()

End Sub

Public Property Get CharIndex() As Integer

    'Return the character index
    CharIndex = u.CharIndex

End Property

Private Sub Class_Initialize()

    'Create the initial buffer size
    BufferSize = BufferScaleSize - 1
    ReDim Buffer(0 To BufferSize)

End Sub

Private Sub IncreaseBufferSize()
'*********************************************************************************
'Increases the size of the packet buffer
'*********************************************************************************

    BufferSize = BufferSize + BufferScaleSize
    ReDim Preserve Buffer(0 To BufferSize)

End Sub

Public Sub SendBuffer()
'*********************************************************************************
'Sends the buffer to the user
'*********************************************************************************

    'Check that there is a buffer
    If BufferCopyPos = 0 Then Exit Sub
    
    'Resize the buffer
    If BufferCopyPos - 1 <> BufferSize Then
        BufferSize = BufferCopyPos - 1
        ReDim Preserve Buffer(0 To BufferSize)
    End If
    
    'Send the packet
    frmMain.GOREsock.SendData UserIndex, Buffer()
    
    'Reset the buffer
    ReDim Buffer(0 To BufferScaleSize - 1)
    BufferCopyPos = 0
    BufferSize = BufferScaleSize
    
    'Check if the user is disconnecting
    If StatusFlag(PCSTATUSFLAG_DISCONNECTING) Then frmMain.GOREsock.Shut cUserIndex

End Sub

Public Sub SendData(ByRef Data() As Byte, ByVal DataSize As Long)
'*********************************************************************************
'Places a packet into the user's outgoing packet buffer
'*********************************************************************************

    'Check if the buffer size needs to be increased
    Do While BufferCopyPos + DataSize > BufferSize
        IncreaseBufferSize
    Loop
    
    'Add to the buffer
    CopyMemory Buffer(BufferCopyPos), Data(0), DataSize
    
    'Increase the copy position
    BufferCopyPos = BufferCopyPos + DataSize

End Sub

Public Sub SetPos(ByVal pMap As Integer, ByVal pX As Integer, ByVal pY As Integer)
'*********************************************************************************
'Sets the user's position
'*********************************************************************************
    
    'Check if the map has changed
    If Map <> pMap Then
        
        'Erase the user from the map
        EraseFromMap

        'Set the new map
        u.Map = pMap
        Maps(u.Map).AddUser cUserIndex

        'Send the new map's information
        SendMapInfo
        
        'Tell everyone else already on the map the user joined
        conBuf.Clear
        AddMakeCharToBuffer
        Data_Send ToMapButIndex, cUserIndex, conBuf.Get_Buffer()
        
    End If
    
    'Check if the position has changed
    If X <> pX Or Y <> pY Then
        
        'Set the user's new position
        X = pX
        Y = pY

        'Send the new position
        conBuf.Clear
        conBuf.Put_Byte PId.SC_Char_SetPos
        conBuf.Put_Integer u.CharIndex
        conBuf.Put_Integer X
        conBuf.Put_Integer Y
        Data_Send ToMap, u.Map, conBuf.Get_Buffer()

    End If

End Sub

Public Sub EraseFromMap()

    'Chekc for a valid map
    If u.Map < 1 Then Exit Sub
    If u.Map > MapsUBound Then Exit Sub

    'Remove the user from the map's internal user list
    Maps(u.Map).RemoveUser cUserIndex

    'Send the packet to the clients
    conBuf.Clear
    conBuf.Put_Byte PId.SC_Char_Erase
    conBuf.Put_Integer u.CharIndex
    Data_Send ToMapButIndex, cUserIndex, conBuf.Get_Buffer()

End Sub

Public Property Get Map() As Integer

    'Return the user's Map position
    Map = u.Map

End Property

Public Property Get Y() As Single

    'Return the user's Y position
    Y = u.Y

End Property

Public Property Get X() As Single
    
    'Return the user's X position
    X = u.X
    
End Property

Private Property Let X(ByVal Value As Single)

    'Set the user's X co-ordinate
    If Value < 0 Then Value = 0
    If Value > Maps(u.Map).TileWidth * GRIDSIZE Then Value = (Maps(u.Map).TileWidth - 2) * GRIDSIZE
    u.X = Value

End Property

Private Property Let Y(ByVal Value As Single)

    'Set the user's Y co-ordinate
    If Value < 0 Then Value = 0
    If Value > Maps(u.Map).TileHeight * GRIDSIZE Then Value = (Maps(u.Map).TileHeight - 2) * GRIDSIZE
    u.Y = Value

End Property

Public Property Get UserIndex() As Long

    'Return the UserIndex
    UserIndex = cUserIndex

End Property

Public Property Get Name() As String
    
    'Return the name
    Name = u.Name

End Property

Public Property Let Name(ByVal Value As String)

    'Set the new name
    If u.Name <> Value Then
        u.Name = Value
    End If

End Property

Public Sub Save()
'*********************************************************************************
'Save the user's character
'*********************************************************************************
Dim InvStr As String
Dim i As Long

    'Make the query to the database
    DB_RS.Open "SELECT * FROM users WHERE `name`='" & u.Name & "'", DB_Conn, adOpenStatic, adLockOptimistic
    
    'Check that the user exists
    If DB_RS.EOF Then
        DB_RS.Close
        Exit Sub
    End If
    
    'Build the inventory string
    For i = 0 To USERINVSIZE
        InvStr = InvStr & Inv(i).ItemIndex & "," & Inv(i).Amount & vbNewLine
    Next i
    
    'Save the information to the database
    With DB_RS
    
        'Position
        !Pos_X = Int(u.X)
        !Pos_Y = Int(u.Y)
        !Pos_Map = u.Map
        
        'Character visuals
        !char_body = u.Body
        
        'Stats
        !stat_lvl = Level
        !stat_exp = EXP
        !stat_str = Str
        !stat_dex = Dex
        !stat_intl = Intl
        !stat_luk = Luk
        !stat_maxhp = MaxHP
        !stat_maxmp = MaxMP
        !stat_hp = HP
        !stat_mp = MP
        !stat_ryu = Ryu
        
        'Equipped
        !eq_weapon = u.WeaponItem
        !eq_cap = u.CapItem
        !eq_forehead = u.ForeheadItem
        !eq_ring1 = u.Ring1Item
        !eq_ring2 = u.Ring2Item
        !eq_ring3 = u.Ring3Item
        !eq_ring4 = u.Ring4Item
        !eq_eyeacc = u.EyeAccItem
        !eq_earacc = u.EarAccItem
        !eq_gloves = u.GlovesItem
        !eq_pendant = u.PendantItem
        !eq_pants = u.PantsItem
        !eq_shoes = u.ShoesItem
        !eq_shield = u.ShieldItem
        !eq_mantle = u.MantleItem
        !eq_clothes = u.ClothesItem
        
        'Misc
        !Inv = InvStr
        
    End With
    
    'Update
    DB_RS.Update
    
    'Close
    DB_RS.Close

End Sub

Public Function Load(ByVal pUserName As String, ByVal pUserIndex As Long) As Boolean
'*********************************************************************************
'Load the user's character
'*********************************************************************************
Dim TempX As Single
Dim TempY As Single
Dim TempMap As Integer
Dim InvStr As String
Dim i As Long
Dim s() As String
Dim s2() As String

    Log "Loading user " & pUserName & " from the database."
    
    'Set the user's index
    cUserIndex = pUserIndex

    'Make the query to the database
    DB_RS.Open "SELECT * FROM users WHERE `name`='" & pUserName & "'", DB_Conn, adOpenStatic, adLockOptimistic
    
    'Check that the user exists
    If DB_RS.EOF Then
        Log "Failed to find database entry for user " & pUserName & "!"
        DB_RS.Close
        Exit Function
    End If
    
    'Load the information
    With DB_RS
        
        'Position
        TempX = !Pos_X
        TempY = !Pos_Y
        TempMap = !Pos_Map
        
        'Character visuals
        u.Body = !char_body
        
        'Stats
        Level = !stat_lvl
        EXP = !stat_exp
        Str = !stat_str
        Dex = !stat_dex
        Intl = !stat_intl
        Luk = !stat_luk
        MaxHP = !stat_maxhp
        MaxMP = !stat_maxmp
        HP = !stat_hp
        MP = !stat_mp
        Ryu = !stat_ryu
        
        'Equipment
        SetWeapon !eq_weapon
        SetCap !eq_cap
        SetForehead !eq_forehead
        SetEyeAcc !eq_eyeacc
        SetEarAcc !eq_earacc
        SetGloves !eq_gloves
        SetShoes !eq_shoes
        SetShield !eq_shield
        SetPants !eq_pants
        SetMantle !eq_mantle
        SetClothes !eq_clothes
        SetRing1 !eq_ring1
        SetRing2 !eq_ring2
        SetRing3 !eq_ring3
        SetRing4 !eq_ring4
        SetPendant !eq_pendant
        
        'Misc
        InvStr = !Inv
        
    End With

    'Close the connect to the recordset
    DB_RS.Close
    
    'Store the user's name
    u.Name = pUserName
    
    'Get the CharIndex
    u.CharIndex = Char_GetIndex(CHARTYPE_PC, pUserIndex)
    
    'Create the inventory
    If LenB(InvStr) Then
        s() = Split(InvStr, vbNewLine)
        For i = 0 To USERINVSIZE
            s2() = Split(s(i), ",")
            Inv(i).ItemIndex = Val(s2(0))
            Inv(i).Amount = Val(s2(1))
        Next i
    End If
    
    'Update the user's position
    SetPos TempMap, TempX, TempY

    'Send the user their char index
    conBuf.Clear
    conBuf.Put_Byte PId.SC_User_SetIndex
    conBuf.Put_Integer u.CharIndex
    Data_Send ToIndex, UserIndex, conBuf.Get_Buffer()
    
    'Send the inventory
    SendInv
    
    'Set the last update time to now
    cLastUpdateTime = timeGetTime
    
End Function

Public Sub Unload()
'*********************************************************************************
'Shut down the user
'*********************************************************************************

    'Save the user
    Save
    
    'Delete the user from the map
    If u.Map > 0 Then
        EraseFromMap
    End If
    
    'Free the CharIndex slot
    Char_FreeIndex u.CharIndex

End Sub

Public Property Get ModLuk() As Integer
'*********************************************************************************
'Return the user's mod luck
'*********************************************************************************

    ModLuk = u.Stats.ModLuk
    
End Property

Public Property Let ModLuk(ByVal Value As Integer)
'*********************************************************************************
'Set the user's mod luck
'*********************************************************************************

    u.Stats.ModLuk = Value
    
End Property

Public Property Get ModIntl() As Integer
'*********************************************************************************
'Return the user's mod intelligence
'*********************************************************************************

    ModIntl = u.Stats.ModIntl
    
End Property

Public Property Let ModIntl(ByVal Value As Integer)
'*********************************************************************************
'Set the user's mod intelligence
'*********************************************************************************

    u.Stats.ModIntl = Value
    
End Property

Public Property Get ModDex() As Integer
'*********************************************************************************
'Return the user's mod dexterity
'*********************************************************************************

    ModDex = u.Stats.ModDex
    
End Property

Public Property Let ModDex(ByVal Value As Integer)
'*********************************************************************************
'Set the user's mod speed
'*********************************************************************************

    u.Stats.ModDex = Value
    
End Property

Public Property Get ModStr() As Integer
'*********************************************************************************
'Return the user's mod strength
'*********************************************************************************

    ModStr = u.Stats.ModStr
    
End Property

Public Property Let ModStr(ByVal Value As Integer)
'*********************************************************************************
'Set the user's mod strength
'*********************************************************************************

    u.Stats.ModStr = Value
    
End Property

Public Property Get MaxMP() As Integer
'*********************************************************************************
'Return the user's mod max mana
'*********************************************************************************

    MaxMP = u.Stats.MaxMP
    
End Property

Public Property Let MaxMP(ByVal Value As Integer)
'*********************************************************************************
'Set the user's mod mana
'*********************************************************************************

    u.Stats.MaxMP = Value
    
End Property

Public Property Get MaxHP() As Integer
'*********************************************************************************
'Return the user's max health
'*********************************************************************************

    MaxHP = u.Stats.MaxHP
    
End Property

Public Property Let MaxHP(ByVal Value As Integer)
'*********************************************************************************
'Set the user's max health
'*********************************************************************************

    u.Stats.MaxHP = Value
    
End Property

Public Property Get Luk() As Integer
'*********************************************************************************
'Return the user's luck
'*********************************************************************************

    Luk = u.Stats.Luk
    
End Property

Public Property Let Luk(ByVal Value As Integer)
'*********************************************************************************
'Set the user's luck
'*********************************************************************************

    u.Stats.Luk = Value
    cUpdateModStats = True
    
End Property

Public Property Get Intl() As Integer
'*********************************************************************************
'Return the user's intelligence
'*********************************************************************************

    Intl = u.Stats.Intl
    
End Property

Public Property Let Intl(ByVal Value As Integer)
'*********************************************************************************
'Set the user's speed
'*********************************************************************************

    u.Stats.Intl = Value
    cUpdateModStats = True
    
End Property

Public Property Get MinHit() As Integer
'*********************************************************************************
'Return the user's MinHit
'*********************************************************************************

    MinHit = u.Stats.MinHit
    
End Property

Public Property Let MinHit(ByVal Value As Integer)
'*********************************************************************************
'Set the user's MinHit
'*********************************************************************************

    u.Stats.MinHit = Value

End Property

Public Property Get MaxHit() As Integer
'*********************************************************************************
'Return the user's MaxHit
'*********************************************************************************

    MaxHit = u.Stats.MaxHit
    
End Property

Public Property Let MaxHit(ByVal Value As Integer)
'*********************************************************************************
'Set the user's MaxHit
'*********************************************************************************

    u.Stats.MaxHit = Value

End Property

Public Property Get Dex() As Integer
'*********************************************************************************
'Return the user's dexterity
'*********************************************************************************

    Dex = u.Stats.Dex
    
End Property

Public Property Let Dex(ByVal Value As Integer)
'*********************************************************************************
'Set the user's dexterity
'*********************************************************************************

    u.Stats.Dex = Value
    cUpdateModStats = True
    
End Property

Public Property Get Def() As Integer
'*********************************************************************************
'Return the user's Def
'*********************************************************************************

    Def = u.Stats.Def
    
End Property

Public Property Let Def(ByVal Value As Integer)
'*********************************************************************************
'Set the user's Def
'*********************************************************************************

    u.Stats.Def = Value
    cUpdateModStats = True
    
End Property

Public Property Get Str() As Integer
'*********************************************************************************
'Return the user's strength
'*********************************************************************************

    Str = u.Stats.Str
    
End Property

Public Property Let Str(ByVal Value As Integer)
'*********************************************************************************
'Set the user's strength
'*********************************************************************************

    u.Stats.Str = Value
    cUpdateModStats = True
    
End Property

Public Property Get EXP() As Long
'*********************************************************************************
'Return the user's experience
'*********************************************************************************

    EXP = u.Stats.EXP
    
End Property

Public Property Let EXP(ByVal Value As Long)
'*********************************************************************************
'Set the user's experience
'*********************************************************************************

    u.Stats.EXP = Value
    
    'Check to increase the user's level
    If u.Stats.EXP > ToNextLevel Then
        LevelUp
    End If
    
End Property

Private Sub LevelUp()
'*********************************************************************************
'Makes the user level up
'*********************************************************************************

    'Decrease the user's exp
    u.Stats.EXP = u.Stats.EXP - ToNextLevel
    
    'Increase the user's level
    Level = Level + 1
    
    'Tell the user they have leveled up
    Data_Send ToIndex, cUserIndex, cMessage(5).Data()

End Sub

Public Function ToNextLevel() As Long
'*********************************************************************************
'Returns the amount of experience needed for the next level
'*********************************************************************************

    ToNextLevel = 10 + u.Stats.Level * 5

End Function

Public Property Get Ryu() As Long
'*********************************************************************************
'Return the user's Ryu
'*********************************************************************************

    Ryu = u.Stats.Ryu
    
End Property

Public Property Let Ryu(ByVal Value As Long)
'*********************************************************************************
'Set the user's Ryu
'*********************************************************************************

    u.Stats.Ryu = Value
    
End Property

Public Property Get HP() As Integer
'*********************************************************************************
'Return the user's health
'*********************************************************************************

    HP = u.Stats.HP
    
End Property

Public Property Let HP(ByVal Value As Integer)
'*********************************************************************************
'Set the user's speed
'*********************************************************************************

    If Value > u.Stats.MaxHP Then Value = u.Stats.MaxHP
    u.Stats.HP = Value
    
End Property

Public Property Get MP() As Integer
'*********************************************************************************
'Return the user's mana
'*********************************************************************************

    MP = u.Stats.MP
    
End Property

Public Property Let MP(ByVal Value As Integer)
'*********************************************************************************
'Set the user's speed
'*********************************************************************************

    If Value > u.Stats.MaxMP Then Value = u.Stats.MaxMP
    u.Stats.MP = Value
    
End Property

Public Property Get Level() As Integer
'*********************************************************************************
'Return the user's level
'*********************************************************************************

    Level = u.Stats.Level
    
End Property

Public Property Let Level(ByVal Value As Integer)
'*********************************************************************************
'Set the user's level
'*********************************************************************************

    u.Stats.Level = Value
    cUpdateModStats = True
    
    'Tell the user the amount of EXP they need to level
    SendTNL
    
End Property

Public Sub GiveEXPandRyu(ByVal pEXP As Integer, ByVal pRyu As Integer)
'*********************************************************************************
'Raises the user's EXP and Ryu along with attaches the message
'*********************************************************************************

    'Check what to send in the message
    If pEXP <> 0 Then
        If pRyu <> 0 Then
            
            'Increase EXP and Ryu
            conBuf.Clear
            conBuf.Put_Byte PId.SC_Message
            conBuf.Put_Byte 3
            conBuf.Put_Integer pEXP
            conBuf.Put_Integer pRyu
            Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()
        
        Else
        
            'Increase EXP
            conBuf.Clear
            conBuf.Put_Byte PId.SC_Message
            conBuf.Put_Byte 1
            conBuf.Put_Integer pEXP
            Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()
            
        End If
    
    Else
        If pRyu <> 0 Then
            
            'Increase Ryu
            conBuf.Clear
            conBuf.Put_Byte PId.SC_Message
            conBuf.Put_Byte 2
            conBuf.Put_Integer pRyu
            Data_Send ToIndex, cUserIndex, conBuf.Get_Buffer()
            
        Else
        
            'Do nothing, nothing to change
            Exit Sub
            
        End If
    End If
    
    'Increase the values
    EXP = EXP + pEXP
    Ryu = Ryu + pRyu

End Sub
