Attribute VB_Name = "Packets"
Option Explicit

Public Sub Data_Handle(ByVal inSox As Long, ByRef Data() As Byte)
'*********************************************************************************
'Global routine for handling incoming packets. Packets are forwarded to another
'method corresponding to their PacketID.
'*********************************************************************************
Dim DataUBound As Long
Dim PacketID As Byte

    'Get the size of the data array
    DataUBound = UBound(Data)

    'Create the buffer
    rBuf.Set_Buffer Data()
    
    'Add to the bandwidth calculation
    'BufferSize (UBound + 1) + TCP (20) + IPv4 (20)
    BytesIn = BytesIn + DataUBound + 41

    'Start the loop through the packet
    Do
        
        'Get the packet ID
        PacketID = rBuf.Get_Byte
        
        'Forward to the corresponding packet handling method
        With PId
            Select Case PacketID
            Case PId.SC_Char_MakePC: Data_Char_MakePC
            Case PId.SC_Char_MakePC_MoveEast: Data_Char_MakePC_MoveEast
            Case PId.SC_Char_MakePC_MoveWest: Data_Char_MakePC_MoveWest
            Case PId.SC_Char_MakeNPC: Data_Char_MakeNPC
            Case PId.SC_Char_MakeNPC_MoveEast: Data_Char_MakeNPC_MoveEast
            Case PId.SC_Char_MakeNPC_MoveWest: Data_Char_MakeNPC_MoveWest
            Case PId.SC_Char_SetPos: Data_Char_SetPos
            Case PId.SC_Char_UpdatePos: Data_Char_UpdatePos
            Case PId.SC_Char_Erase: Data_Char_Erase
            Case PId.SC_Char_SetPaperDoll: Data_Char_SetPaperDoll
            Case PId.SC_Char_Kill: Data_Char_Kill
            Case PId.SC_Char_HPMP: Data_Char_HPMP
            
            Case PId.SC_Move_EastStart: Data_Move_EastStart
            Case PId.SC_Move_EastEnd: Data_Move_EastEnd
            Case PId.SC_Move_WestStart: Data_Move_WestStart
            Case PId.SC_Move_WestEnd: Data_Move_WestEnd
            Case PId.SC_Jump: Data_Jump
            Case PId.SC_Punch: Data_Punch
            Case PId.SC_Punch_Hit: Data_Punch_Hit
            Case PId.SC_Char_SetHeading: Data_Char_SetHeading
            
            Case PId.SC_User_SetMap: Data_User_SetMap
            Case PId.SC_User_SetIndex: Data_User_SetIndex
            Case PId.SC_User_Stats: Data_User_Stats
            Case PId.SC_User_HP: Data_User_HP
            Case PId.SC_User_MP: Data_User_MP
            Case PId.SC_User_EXP: Data_User_EXP
            Case PId.SC_User_Ryu: Data_User_Ryu
            Case PId.SC_User_ToNextLevel: Data_User_ToNextLevel
            
            Case PId.SC_Inv_Update: Data_Inv_Update
            Case PId.SC_Inv_UpdateSlot: Data_Inv_UpdateSlot
            
            Case PId.SC_Item_Make: Data_Item_Make
            Case PId.SC_Item_Erase: Data_Item_Erase
            Case PId.SC_Item_Pickup: Data_Item_Pickup
            Case PId.SC_Item_Drop: Data_Item_Drop
            
            Case PId.SC_Ping: Data_Ping
            Case PId.SC_Message_UserAlreadyOnline: Data_Message_UserAlreadyOnline
            Case PId.SC_Message_AccountInUse: Data_Message_AccountInUse
            Case PId.SC_Message: Data_Message
            Case PId.SC_Chat_Say: Data_Chat_Say
            Case PId.SC_SetEquipped: Data_SetEquipped
            
            Case 0: rBuf.Overflow       'An error occured
            Case Else: rBuf.Overflow    'An error occured
            End Select
        End With
        
        'Check if the buffer ran out
        If rBuf.Get_ReadPos > DataUBound Then Exit Do
        
    Loop

End Sub

Private Sub Data_Item_Pickup()
'*********************************************************************************
'An item was picked up off the ground
'<ItemSlot(I)><CharIndex(I)>
'*********************************************************************************
Dim ItemSlot As Integer
Dim CharIndex As Integer
Dim s() As String

    'Get the data
    ItemSlot = rBuf.Get_Integer
    CharIndex = rBuf.Get_Integer
    
    'Check for valid character
    If Not IsValidChar(CharIndex) Then
    
        'If the character is invalid, just erase the item
        EraseMapItem ItemSlot
        Exit Sub
        
    End If
    
    'Check if our user picked it up
    If CharIndex = UserCharIndex Then
        
        'Pickup message
        If MapItems(ItemSlot).Amount = 1 Then
            ReDim s(0)
            s(0) = Items(MapItems(ItemSlot).ItemIndex).Name
            InfoBox.Add Msg.GrabReplace(7, s)
        Else
            ReDim s(1)
            s(0) = MapItems(ItemSlot).Amount
            s(1) = Items(MapItems(ItemSlot).ItemIndex).Name
            InfoBox.Add Msg.GrabReplace(8, s)
        End If
        
    End If
    
    'Pick up the item
    AddPickedupItem MapItems(ItemSlot).X, MapItems(ItemSlot).Y, GrhData(Items(MapItems(ItemSlot).ItemIndex).GrhIndex).Width, _
        GrhData(Items(MapItems(ItemSlot).ItemIndex).GrhIndex).Height, CharIndex, MapItems(ItemSlot).Grh
    
    'Erase the item
    EraseMapItem ItemSlot

End Sub

Private Sub Data_Item_Erase()
'*********************************************************************************
'Erases an item from the map
'<ItemSlot(I)>
'*********************************************************************************
Dim ItemSlot As Integer

    'Get the slot
    ItemSlot = rBuf.Get_Integer

    'Create the fading item
    AddFadeItem MapItems(ItemSlot).X, MapItems(ItemSlot).Y, MapItems(ItemSlot).Grh

    'Erase the real item from the map
    EraseMapItem ItemSlot
    
End Sub

Private Sub Data_Item_Drop()
'*********************************************************************************
'Makes an item drop with a nice little effect instead of just appearing
'<ItemSlot(I)><ItemIndex(I)><Amount(I)><X(I)><Y(I)><DestY(I)>
'*********************************************************************************
Dim ItemSlot As Integer
Dim ItemIndex As Integer
Dim Amount As Integer
Dim X As Integer
Dim Y As Integer
Dim DestY As Integer

    'Get the data
    ItemSlot = rBuf.Get_Integer
    ItemIndex = rBuf.Get_Integer
    Amount = rBuf.Get_Integer
    X = rBuf.Get_Integer
    Y = rBuf.Get_Integer
    DestY = rBuf.Get_Integer
    
    'Check if the items array needs to be increased
    If ItemSlot > MapItemsUBound Then
        MapItemsUBound = ItemSlot
        ReDim Preserve MapItems(0 To MapItemsUBound)
    End If
    
    'Check for a new highest map item index
    If ItemSlot > LastMapItem Then LastMapItem = ItemSlot
    
    'Make sure the DestY > Y
    If DestY < Y Then Y = DestY - 1
    
    'Create the item
    With MapItems(ItemSlot)
        .ItemIndex = ItemIndex
        .Amount = Amount
        .X = X
        .Y = Y
        .DestY = DestY
        .DestX = X
        .Xv = 1 - (Rnd * 2)
        .Yv = -5 - (Rnd * 2)
        .Moving = True
        Graphics_SetGrh .Grh, Items(ItemIndex).GrhIndex, ANIMTYPE_LOOP
    End With

End Sub

Private Sub Data_Item_Make()
'*********************************************************************************
'Creates an item on the map (no fancy drop effect)
'<ItemSlot(I)><ItemIndex(I)><Amount(I)><X(I)><Y(I)>
'*********************************************************************************
Dim ItemSlot As Integer
Dim ItemIndex As Integer
Dim Amount As Integer
Dim X As Integer
Dim Y As Integer

    'Get the data
    ItemSlot = rBuf.Get_Integer
    ItemIndex = rBuf.Get_Integer
    Amount = rBuf.Get_Integer
    X = rBuf.Get_Integer
    Y = rBuf.Get_Integer
    
    'Check if the items array needs to be increased
    If ItemSlot > MapItemsUBound Then
        MapItemsUBound = ItemSlot
        ReDim Preserve MapItems(0 To MapItemsUBound)
    End If
    
    'Check for a new highest map item index
    If ItemSlot > LastMapItem Then LastMapItem = ItemSlot
    
    'Create the item
    With MapItems(ItemSlot)
        .ItemIndex = ItemIndex
        .Amount = Amount
        .X = X
        .Y = Y + 2
        Graphics_SetGrh .Grh, Items(ItemIndex).GrhIndex, ANIMTYPE_LOOP
    End With

End Sub

Private Sub Data_SetEquipped()
'*********************************************************************************
'Set the user's equipped slot
'<Slot(B)><ItemIndex(I)>
'*********************************************************************************
Dim Slot As Byte
Dim ItemIndex As Integer

    'Get the data
    Slot = rBuf.Get_Byte
    ItemIndex = rBuf.Get_Integer
    
    'Find the slot, and change it if it is different
    Select Case Slot
    
    Case EQUIPSLOT_CAP
        If CapItemIndex <> ItemIndex Then
            CapItemIndex = ItemIndex
            Graphics_SetGrh CapGrh, Items(ItemIndex).GrhIndex, ANIMTYPE_LOOP
        End If
    
    Case EQUIPSLOT_CLOTHES
        If ClothesItemIndex <> ItemIndex Then
            ClothesItemIndex = ItemIndex
            Graphics_SetGrh ClothesGrh, Items(ItemIndex).GrhIndex, ANIMTYPE_LOOP
        End If
    
    Case EQUIPSLOT_EARACC
        If EarAccItemIndex <> ItemIndex Then
            EarAccItemIndex = ItemIndex
            Graphics_SetGrh EarAccGrh, Items(ItemIndex).GrhIndex, ANIMTYPE_LOOP
        End If
    
    Case EQUIPSLOT_EYEACC
        If EyeAccItemIndex <> ItemIndex Then
            EyeAccItemIndex = ItemIndex
            Graphics_SetGrh EyeAccGrh, Items(ItemIndex).GrhIndex, ANIMTYPE_LOOP
        End If
    
    Case EQUIPSLOT_FOREHEAD
        If ForeheadItemIndex <> ItemIndex Then
            ForeheadItemIndex = ItemIndex
            Graphics_SetGrh ForeheadGrh, Items(ItemIndex).GrhIndex, ANIMTYPE_LOOP
        End If
    
    Case EQUIPSLOT_GLOVES
        If GlovesItemIndex <> ItemIndex Then
            GlovesItemIndex = ItemIndex
            Graphics_SetGrh GlovesGrh, Items(ItemIndex).GrhIndex, ANIMTYPE_LOOP
        End If
    
    Case EQUIPSLOT_MANTLE
        If MantleItemIndex <> ItemIndex Then
            MantleItemIndex = ItemIndex
            Graphics_SetGrh MantleGrh, Items(ItemIndex).GrhIndex, ANIMTYPE_LOOP
        End If
    
    Case EQUIPSLOT_PANTS
        If PantsItemIndex <> ItemIndex Then
            PantsItemIndex = ItemIndex
            Graphics_SetGrh PantsGrh, Items(ItemIndex).GrhIndex, ANIMTYPE_LOOP
        End If
    
    Case EQUIPSLOT_PENDANT
        If PendantItemIndex <> ItemIndex Then
            PendantItemIndex = ItemIndex
            Graphics_SetGrh PendantGrh, Items(ItemIndex).GrhIndex, ANIMTYPE_LOOP
        End If
    
    Case EQUIPSLOT_RING1
        If Ring1ItemIndex <> ItemIndex Then
            Ring1ItemIndex = ItemIndex
            Graphics_SetGrh Ring1Grh, Items(ItemIndex).GrhIndex, ANIMTYPE_LOOP
        End If
    
    Case EQUIPSLOT_RING2
        If Ring2ItemIndex <> ItemIndex Then
            Ring2ItemIndex = ItemIndex
            Graphics_SetGrh Ring2Grh, Items(ItemIndex).GrhIndex, ANIMTYPE_LOOP
        End If
    
    Case EQUIPSLOT_RING3
        If Ring3ItemIndex <> ItemIndex Then
            Ring3ItemIndex = ItemIndex
            Graphics_SetGrh Ring3Grh, Items(ItemIndex).GrhIndex, ANIMTYPE_LOOP
        End If
    
    Case EQUIPSLOT_RING4
        If Ring4ItemIndex <> ItemIndex Then
            Ring4ItemIndex = ItemIndex
            Graphics_SetGrh Ring4Grh, Items(ItemIndex).GrhIndex, ANIMTYPE_LOOP
        End If
    
    Case EQUIPSLOT_SHIELD
        If ShieldItemIndex <> ItemIndex Then
            ShieldItemIndex = ItemIndex
            Graphics_SetGrh ShieldGrh, Items(ItemIndex).GrhIndex, ANIMTYPE_LOOP
        End If
    
    Case EQUIPSLOT_SHOES
        If ShoesItemIndex <> ItemIndex Then
            ShoesItemIndex = ItemIndex
            Graphics_SetGrh ShoesGrh, Items(ItemIndex).GrhIndex, ANIMTYPE_LOOP
        End If
    
    Case EQUIPSLOT_WEAPON
        If WeaponItemIndex <> ItemIndex Then
            WeaponItemIndex = ItemIndex
            Graphics_SetGrh WeaponGrh, Items(ItemIndex).GrhIndex, ANIMTYPE_LOOP
        End If
    
    End Select

End Sub

Private Sub Data_Inv_Update()
'*********************************************************************************
'Updates the user's whole inventory
'Loop: <ItemIndex(I)>(<Amount(I)>)
'*********************************************************************************
Dim Slot As Byte
Dim ItemIndex As Integer
Dim Amount As Integer

    'Loop through the user's complete inventory
    For Slot = 0 To USERINVSIZE

        'Get the information
        ItemIndex = rBuf.Get_Integer
        If ItemIndex > 0 Then Amount = rBuf.Get_Integer
        
        'Set the slot information
        If UserInv(Slot).ItemIndex <> ItemIndex Then
            UserInv(Slot).ItemIndex = ItemIndex
            Graphics_SetGrh UserInv(Slot).Grh, Items(UserInv(Slot).ItemIndex).GrhIndex, ANIMTYPE_LOOP
        End If
        UserInv(Slot).Amount = Amount

    Next Slot

End Sub

Private Sub Data_Inv_UpdateSlot()
'*********************************************************************************
'Updates just one slot of the user's inventory
'<Slot(B)><ItemIndex(I)>(<Amount(I)>)
'*********************************************************************************
Dim Slot As Byte
Dim ItemIndex As Integer
Dim Amount As Integer

    'Get the data
    Slot = rBuf.Get_Byte
    ItemIndex = rBuf.Get_Integer
    If ItemIndex > 0 Then Amount = rBuf.Get_Integer
    
    'Set the slot information
    If UserInv(Slot).ItemIndex <> ItemIndex Then
        UserInv(Slot).ItemIndex = ItemIndex
        Graphics_SetGrh UserInv(Slot).Grh, Items(UserInv(Slot).ItemIndex).GrhIndex, ANIMTYPE_LOOP
    End If
    UserInv(Slot).Amount = Amount

End Sub

Private Sub Data_User_ToNextLevel()
'*********************************************************************************
'The amount of experience the user needs to level
'<ToNextLevel(L)>
'*********************************************************************************

    User_ToNextLevel = rBuf.Get_Long

End Sub

Private Sub Data_Char_SetHeading()
'*********************************************************************************
'Changes the character's heading
'<CharIndex(I)><Heading(B)>
'*********************************************************************************
Dim CharIndex As Integer
Dim Heading As Byte

    'Get the data
    CharIndex = rBuf.Get_Integer
    Heading = rBuf.Get_Byte
    
    'Check for a valid character
    If Not IsValidChar(CharIndex) Then Exit Sub
    
    'Set the heading
    CharList(CharIndex).Heading = Heading

End Sub

Private Sub Data_Char_HPMP()
'*********************************************************************************
'Updates a character's HP and MP percentages
'<CharIndex(I)><HPP(B)><MPP(B)>
'*********************************************************************************
Dim CharIndex As Integer
Dim HPP As Byte
Dim MPP As Byte

    'Get the data
    CharIndex = rBuf.Get_Integer
    HPP = rBuf.Get_Byte
    MPP = rBuf.Get_Byte
    
    'Check for a valid character
    If Not IsValidChar(CharIndex) Then Exit Sub
    
    'Set the character's HP and MP percents
    CharList(CharIndex).HPP = HPP
    CharList(CharIndex).MPP = MPP

End Sub

Private Sub Data_User_HP()
'*********************************************************************************
'Updates the user's HP
'<HP(I)>
'*********************************************************************************

    UserStats.HP = rBuf.Get_Integer
    If UserStats.MaxHP > 0 Then CharList(UserCharIndex).HPP = (UserStats.HP / UserStats.MaxHP) * 100

End Sub

Private Sub Data_User_MP()
'*********************************************************************************
'Updates a user's MP
'<MP(I)>
'*********************************************************************************

    UserStats.MP = rBuf.Get_Integer
    If UserStats.MaxMP > 0 Then CharList(UserCharIndex).MPP = (UserStats.MP / UserStats.MaxMP) * 100

End Sub

Private Sub Data_User_EXP()
'*********************************************************************************
'Updates the user's EXP
'<EXP(L)>
'*********************************************************************************

    UserStats.EXP = rBuf.Get_Long

End Sub

Private Sub Data_User_Ryu()
'*********************************************************************************
'Updates the user's Ryu
'<Ryu(L)>
'*********************************************************************************

    UserStats.Ryu = rBuf.Get_Long

End Sub

Private Sub Data_User_Stats()
'*********************************************************************************
'Updates the user's stats with the exception of EXP, Ryu, MP and HP
'<Flags(I)><???>
'*********************************************************************************
Dim Flags As Integer

    'Get the flags
    Flags = rBuf.Get_Integer

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
    If Flags And 1 Then UserStats.MaxHP = rBuf.Get_Integer
    If Flags And 2 Then UserStats.MaxMP = rBuf.Get_Integer
    If Flags And 4 Then UserStats.Level = rBuf.Get_Integer
    If Flags And 8 Then UserStats.Str = rBuf.Get_Integer
    If Flags And 16 Then UserStats.ModStr = rBuf.Get_Integer
    If Flags And 32 Then UserStats.Dex = rBuf.Get_Integer
    If Flags And 64 Then UserStats.ModDex = rBuf.Get_Integer
    If Flags And 128 Then UserStats.Intl = rBuf.Get_Integer
    If Flags And 256 Then UserStats.ModIntl = rBuf.Get_Integer
    If Flags And 512 Then UserStats.Luk = rBuf.Get_Integer
    If Flags And 1024 Then UserStats.ModLuk = rBuf.Get_Integer
    If Flags And 2048 Then UserStats.MinHit = rBuf.Get_Integer
    If Flags And 4096 Then UserStats.MaxHit = rBuf.Get_Integer

End Sub

Private Sub Data_Chat_Say()
'*********************************************************************************
'A character says something
'<CharIndex(I)><Text(S)>
'*********************************************************************************
Dim CharIndex As Integer
Dim Text As String
Dim s(0 To 1) As String

    'Get the data
    CharIndex = rBuf.Get_Integer
    Text = rBuf.Get_String
    
    'Check for a valid character
    If Not IsValidChar(CharIndex) Then Exit Sub
    
    'Add to the chat buffer
    s(0) = CharList(CharIndex).Name
    s(1) = Text
    ChatBox.AddHistory Msg.GrabReplace(4, s())

End Sub

Private Sub Data_Char_Kill()
'*********************************************************************************
'Erases a character that has died
'<CharIndex(I)>
'*********************************************************************************
Dim CharIndex As Integer

    'Get the data
    CharIndex = rBuf.Get_Integer
    
    'Check if a valid character
    If Not IsValidChar(CharIndex) Then Exit Sub
    
    'Set the character as unused
    CharList(CharIndex).Used = False
    
    'Set the action
    CharList(CharIndex).Action = eDeath
    
    'Set the death animaiton
    If CharList(CharIndex).IsNPC Then
        Graphics_SetGrh CharList(CharIndex).Body, PDSprite(CharList(CharIndex).BodyIndex).Death, ANIMTYPE_LOOPONCE
    Else
        Graphics_SetGrh CharList(CharIndex).Body, PDBody(CharList(CharIndex).BodyIndex).Death, ANIMTYPE_LOOPONCE
    End If
    
End Sub

Private Sub Data_Char_Erase()
'*********************************************************************************
'Erases a character (went invisible, changed maps, etc)
'<CharIndex(I)>
'*********************************************************************************
Dim CharIndex As Integer

    'Get the data
    CharIndex = rBuf.Get_Integer
    
    'Check if a valid character
    If Not IsValidChar(CharIndex) Then Exit Sub
    
    'Set the character as unused
    CharList(CharIndex).Used = False

End Sub

Private Sub Data_Message_AccountInUse()
'*********************************************************************************
'The user tried to log on, but a character from the account was already online
'<>
'*********************************************************************************

    MsgBox "Another character from that account is already online." & vbNewLine & _
        "Please disconnect that character first.", vbOKOnly
    EngineRunning = False

End Sub

Private Sub Data_Message_UserAlreadyOnline()
'*********************************************************************************
'The user tried to log on, but the character was already online
'<>
'*********************************************************************************

    MsgBox "That user is already connected. Please try again.", vbOKOnly
    EngineRunning = False

End Sub

Private Sub Data_Ping()
'*********************************************************************************
'Ping came back from the server
'<>
'*********************************************************************************
Dim l As Long
Dim i As Long

    'Get the time it took for the ping
    l = timeGetTime - PingSentTime

    'Check if to increase the number of pings we have
    If NumPings < TOTALPINGS Then
        NumPings = NumPings + 1
    End If
        
    'Add the ping to the list and remove the oldest
    For i = 0 To TOTALPINGS - 2
        PingTime(i + 1) = PingTime(i)
    Next i
    PingTime(0) = l

    'Get the ping average
    Ping = 0
    For i = 0 To NumPings - 1
        Ping = Ping + PingTime(i)
    Next i
    Ping = Ping / NumPings

End Sub

Private Sub Data_Punch_Hit()
'*********************************************************************************
'Makes a character punch, and the punch hit someone
'<CharIndex(I)><HitIndex(I)><Damage(I)>
'*********************************************************************************
Dim CharIndex As Integer
Dim HitIndex As Integer
Dim Damage As Integer

    'Get the data
    CharIndex = rBuf.Get_Integer
    HitIndex = rBuf.Get_Integer
    Damage = rBuf.Get_Integer
    
    'Check for a valid character
    If Not IsValidChar(CharIndex) Then Exit Sub
    If HitIndex < 1 Then Exit Sub
    If HitIndex > CharListUBound Then Exit Sub
    
    'Set the character's action
    CharList(CharIndex).Action = ePunch
    If CharList(CharIndex).IsNPC Then
        Graphics_SetGrh CharList(CharIndex).Body, PDSprite(CharList(CharIndex).BodyIndex).Punch, ANIMTYPE_LOOPONCE
    Else
        Graphics_SetGrh CharList(CharIndex).Body, PDBody(CharList(CharIndex).BodyIndex).Punch, ANIMTYPE_LOOPONCE
    End If
    
    'Display the damage
    AddDamage CharList(HitIndex).X, CharList(HitIndex).Y, Damage
    
    'Set the being hit action
    If CharList(HitIndex).Action <> eDeath Then
        CharList(HitIndex).Action = eHit
        If CharList(CharIndex).IsNPC Then
            Graphics_SetGrh CharList(HitIndex).Body, PDSprite(CharList(HitIndex).BodyIndex).Hit, ANIMTYPE_LOOPONCE
        Else
            Graphics_SetGrh CharList(HitIndex).Body, PDBody(CharList(HitIndex).BodyIndex).Hit, ANIMTYPE_LOOPONCE
        End If
    End If
    
    'Play the hit sound effect
    Sfx_Play 2
    
End Sub

Private Sub Data_Punch()
'*********************************************************************************
'Makes a character punch
'<CharIndex(I)>
'*********************************************************************************
Dim CharIndex As Integer

    'Get the data
    CharIndex = rBuf.Get_Integer
    
    'Check for a valid character
    If Not IsValidChar(CharIndex) Then Exit Sub
    
    'Set the character's action
    CharList(CharIndex).Action = ePunch
    If CharList(CharIndex).IsNPC Then
        Graphics_SetGrh CharList(CharIndex).Body, PDSprite(CharList(CharIndex).BodyIndex).Punch, ANIMTYPE_LOOPONCE
    Else
        Graphics_SetGrh CharList(CharIndex).Body, PDBody(CharList(CharIndex).BodyIndex).Punch, ANIMTYPE_LOOPONCE
    End If
    
    'Miss sound effect
    Sfx_Play 1
    
End Sub

Private Sub Data_Jump()
'*********************************************************************************
'Makes a character jump
'<CharIndex(I)><X(I)><Y(I)>
'*********************************************************************************
Dim CharIndex As Integer
Dim X As Integer
Dim Y As Integer

    'Get the data
    CharIndex = rBuf.Get_Integer
    X = rBuf.Get_Integer
    Y = rBuf.Get_Integer
    
    'Check for a valid character
    If Not IsValidChar(CharIndex) Then Exit Sub
    
    'Set the values
    CharList(CharIndex).X = X
    CharList(CharIndex).Y = Y
    CharList(CharIndex).Jump = JUMPHEIGHT

End Sub

Private Sub Data_Move_EastEnd()
'*********************************************************************************
'Ends the character moving to the East
'<CharIndex(I)><X(I)><Y(I)>
'*********************************************************************************
Dim CharIndex As Integer
Dim X As Integer
Dim Y As Integer
    
    'Get the data
    CharIndex = rBuf.Get_Integer
    X = rBuf.Get_Integer
    Y = rBuf.Get_Integer
    
    'Check for a valid character
    If Not IsValidChar(CharIndex) Then Exit Sub
    
    'Set the values
    CharList(CharIndex).X = X
    CharList(CharIndex).Y = Y
    CharList(CharIndex).MoveDir = 0
    CharList(CharIndex).Heading = EAST
    
    'Update LastSentMoveDir if this is the user's character
    If CharIndex = UserCharIndex Then LastSentMoveDir = 0

End Sub

Private Sub Data_Move_WestEnd()
'*********************************************************************************
'Ends the character moving to the West
'<CharIndex(I)><X(I)><Y(I)>
'*********************************************************************************
Dim CharIndex As Integer
Dim X As Integer
Dim Y As Integer
    
    'Get the data
    CharIndex = rBuf.Get_Integer
    X = rBuf.Get_Integer
    Y = rBuf.Get_Integer
    
    'Check for a valid character
    If Not IsValidChar(CharIndex) Then Exit Sub

    'Set the values
    CharList(CharIndex).X = X
    CharList(CharIndex).Y = Y
    CharList(CharIndex).MoveDir = 0
    CharList(CharIndex).Heading = WEST
    
    'Update LastSentMoveDir if this is the user's character
    If CharIndex = UserCharIndex Then LastSentMoveDir = 0

End Sub

Private Sub Data_Message()
'*********************************************************************************
'Handles various messages from the server
'<MessageID(B)><Parameters(?)>
'*********************************************************************************
Dim MessageID As Byte
Dim s() As String

    'Get the ID
    MessageID = rBuf.Get_Byte
    
    Select Case MessageID
        
        'You got <0> experience.
        Case 1
            ReDim s(0)
            s(0) = rBuf.Get_Integer
            InfoBox.Add Msg.GrabReplace(MessageID, s)
        
        'You got <0> Ryu.
        Case 2
            ReDim s(0)
            s(0) = rBuf.Get_Integer
            InfoBox.Add Msg.GrabReplace(MessageID, s)
            
        'You got <0> experience and <1> Ryu.
        Case 3
            ReDim s(1)
            s(0) = rBuf.Get_Integer
            s(1) = rBuf.Get_Integer
            InfoBox.Add Msg.GrabReplace(MessageID, s)
            
        'Death
        Case 6
            ChatBox.AddHistory Msg.Grab(MessageID)
            
        'Pick up a single item
        Case 7
            ReDim s(0)
            s(0) = Items(rBuf.Get_Integer).Name
            InfoBox.Add Msg.GrabReplace(MessageID, s)
            
        'Pick up multiple items
        Case 8
            ReDim s(1)
            s(0) = rBuf.Get_Integer
            s(1) = Items(rBuf.Get_Integer).Name
            InfoBox.Add Msg.GrabReplace(MessageID, s)
            
        'Inventory is full
        Case 9
            InfoBox.Add Msg.Grab(MessageID)
        
        'Everything else, throw in the box as-is
        Case Else
            ChatBox.AddHistory Msg.Grab(MessageID)
            
    End Select

End Sub

Private Sub Data_Move_WestStart()
'*********************************************************************************
'Start the character moving to the West
'<CharIndex(I)><X(I)><Y(I)>
'*********************************************************************************
Dim CharIndex As Integer
Dim X As Integer
Dim Y As Integer
    
    'Get the data
    CharIndex = rBuf.Get_Integer
    X = rBuf.Get_Integer
    Y = rBuf.Get_Integer
    
    'Check for a valid character
    If Not IsValidChar(CharIndex) Then Exit Sub
    
    'Modify the X co-ordinate with the ping taken into consideration
    Engine_ModifyXForLag X, WEST, MOVESPEED

    'Set the values
    CharList(CharIndex).X = X
    CharList(CharIndex).Y = Y
    CharList(CharIndex).MoveDir = WEST
    CharList(CharIndex).Heading = WEST
    
    'Update LastSentMoveDir if this is the user's character
    If CharIndex = UserCharIndex Then LastSentMoveDir = WEST

End Sub

Private Sub Data_Move_EastStart()
'*********************************************************************************
'Start the character moving to the East
'<CharIndex(I)><X(I)><Y(I)>
'*********************************************************************************
Dim CharIndex As Integer
Dim X As Integer
Dim Y As Integer
    
    'Get the data
    CharIndex = rBuf.Get_Integer
    X = rBuf.Get_Integer
    Y = rBuf.Get_Integer
    
    'Check for a valid character
    If Not IsValidChar(CharIndex) Then Exit Sub
    
    'Modify the X co-ordinate with the ping taken into consideration
    Engine_ModifyXForLag X, EAST, MOVESPEED

    'Set the values
    CharList(CharIndex).X = X
    CharList(CharIndex).Y = Y
    CharList(CharIndex).MoveDir = EAST
    CharList(CharIndex).Heading = EAST
    
    'Update LastSentMoveDir if this is the user's character
    If CharIndex = UserCharIndex Then LastSentMoveDir = EAST

End Sub

Private Sub Data_Char_SetPaperDoll()
'*********************************************************************************
'Set the user's paperdolling information
'<CharIndex(I)><Body(B)>
'*********************************************************************************
Dim CharIndex As Integer
Dim Body As Byte

    'Get the data
    CharIndex = rBuf.Get_Integer
    Body = rBuf.Get_Byte
    
    'Check for a valid character
    If Not IsValidChar(CharIndex) Then Exit Sub

    'Set the paperdoll information
    If CharList(CharIndex).BodyIndex <> Body Then
        CharList(CharIndex).BodyIndex = Body
        CharList(CharIndex).Width = PDBody(Body).Width
        CharList(CharIndex).Height = PDBody(Body).Height
    End If

End Sub

Private Sub Data_User_SetIndex()
'*********************************************************************************
'Sets the user's character index
'<UserCharIndex(I)>
'*********************************************************************************

    'Get the data
    UserCharIndex = rBuf.Get_Integer

End Sub

Private Sub Data_User_SetMap()
'*********************************************************************************
'Make the client load a map
'<MapIndex(I)>
'*********************************************************************************
Dim MapIndex As Integer

    'Get the data
    MapIndex = rBuf.Get_Integer

    'Stop all the sound effects
    Sfx_Stop

    'Load the map information
    IO_Map_LoadInfo MapIndex, MapInfo
    IO_Map_LoadGrhs MapIndex, MapGrh(), NumMapGrhs, BG(), BGSizeX, BGSizeY
    CurrMapIndex = MapIndex
    UpdateMapGrhPtr = True
    Music_Play MapInfo.MusicID

End Sub

Private Sub Data_Char_UpdatePos()
'*********************************************************************************
'Updates the character's position while they're moving
'<CharIndex(I)><X(I)><Y(I)>
'*********************************************************************************
Dim CharIndex As Integer
Dim X As Integer
Dim Y As Integer

    'Get the data
    CharIndex = rBuf.Get_Integer
    X = rBuf.Get_Integer
    Y = rBuf.Get_Integer
    
    'Check that the character is valid
    If Not IsValidChar(CharIndex) Then Exit Sub
    
    'Set the character's position
    Engine_ModifyXForLag X, CharList(CharIndex).Heading, MOVESPEED
    CharList(CharIndex).X = X
    CharList(CharIndex).Y = Y
    
End Sub

Private Sub Data_Char_SetPos()
'*********************************************************************************
'Sets the character's position
'<CharIndex(I)><X(I)><Y(I)>
'*********************************************************************************
Dim CharIndex As Integer
Dim X As Integer
Dim Y As Integer

    'Get the data
    CharIndex = rBuf.Get_Integer
    X = rBuf.Get_Integer
    Y = rBuf.Get_Integer
    
    'Check that the character is valid
    If Not IsValidChar(CharIndex) Then Exit Sub
    
    'Set the character's position
    CharList(CharIndex).X = X
    CharList(CharIndex).Y = Y
    CharList(CharIndex).DrawX = X
    CharList(CharIndex).DrawY = Y
    
End Sub

Private Sub Data_Char_MakePC()
'*********************************************************************************
'Make a player character on the client
'<CharIndex(I)><X(I)><Y(I)><Name(S)><Heading(B)><Body(B)>
'*********************************************************************************
Dim CharIndex As Integer
Dim X As Integer
Dim Y As Integer
Dim Name As String
Dim Body As Byte
Dim Heading As Byte

    'Get the data
    CharIndex = rBuf.Get_Integer
    X = rBuf.Get_Integer
    Y = rBuf.Get_Integer
    Name = rBuf.Get_String
    Heading = rBuf.Get_Byte
    Body = rBuf.Get_Byte

    'Make the PC
    MakePC CharIndex, X, Y, Name, Heading, Body

End Sub

Private Sub Data_Char_MakePC_MoveEast()
'*********************************************************************************
'Make a player character on the client moving to the East
'<CharIndex(I)><X(I)><Y(I)><Name(S)><Body(B)>
'*********************************************************************************
Dim CharIndex As Integer
Dim X As Integer
Dim Y As Integer
Dim Name As String
Dim Body As Byte

    'Get the data
    CharIndex = rBuf.Get_Integer
    X = rBuf.Get_Integer
    Y = rBuf.Get_Integer
    Name = rBuf.Get_String
    Body = rBuf.Get_Byte

    'Make the PC
    MakePC CharIndex, X, Y, Name, EAST, Body, EAST

End Sub

Private Sub Data_Char_MakePC_MoveWest()
'*********************************************************************************
'Make a player character on the client moving to the West
'<CharIndex(I)><X(I)><Y(I)><Name(S)><Body(B)>
'*********************************************************************************
Dim CharIndex As Integer
Dim X As Integer
Dim Y As Integer
Dim Name As String
Dim Body As Byte

    'Get the data
    CharIndex = rBuf.Get_Integer
    X = rBuf.Get_Integer
    Y = rBuf.Get_Integer
    Name = rBuf.Get_String
    Body = rBuf.Get_Byte

    'Make the PC
    MakePC CharIndex, X, Y, Name, WEST, Body, WEST

End Sub

Private Sub Data_Char_MakeNPC()
'*********************************************************************************
'Make a non-player character on the client
'<CharIndex(I)><X(I)><Y(I)><TemplateID(I)>
'*********************************************************************************
Dim TemplateID As Integer
Dim CharIndex As Integer
Dim X As Integer
Dim Y As Integer

    'Get the data
    CharIndex = rBuf.Get_Integer
    X = rBuf.Get_Integer
    Y = rBuf.Get_Integer
    TemplateID = rBuf.Get_Integer

    'Make the NPC
    MakeNPC CharIndex, X, Y, TemplateID

End Sub

Private Sub Data_Char_MakeNPC_MoveEast()
'*********************************************************************************
'Make a non-player character on the client moving to the East
'<CharIndex(I)><X(I)><Y(I)><TemplateID(I)>
'*********************************************************************************
Dim TemplateID As Integer
Dim CharIndex As Integer
Dim X As Integer
Dim Y As Integer

    'Get the data
    CharIndex = rBuf.Get_Integer
    X = rBuf.Get_Integer
    Y = rBuf.Get_Integer
    TemplateID = rBuf.Get_Integer

    'Make the NPC
    MakeNPC CharIndex, X, Y, TemplateID, EAST
    
End Sub

Private Sub Data_Char_MakeNPC_MoveWest()
'*********************************************************************************
'Make a non-player character on the client moving to the Wast
'<CharIndex(I)><X(I)><Y(I)><TemplateID(I)>
'*********************************************************************************
Dim TemplateID As Integer
Dim CharIndex As Integer
Dim X As Integer
Dim Y As Integer

    'Get the data
    CharIndex = rBuf.Get_Integer
    X = rBuf.Get_Integer
    Y = rBuf.Get_Integer
    TemplateID = rBuf.Get_Integer

    'Make the NPC
    MakeNPC CharIndex, X, Y, TemplateID, WEST
    
End Sub
