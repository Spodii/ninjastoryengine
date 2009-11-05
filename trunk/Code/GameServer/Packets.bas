Attribute VB_Name = "Packets"
Option Explicit

Public Enum SendRoute
    ToIndex = 1
    ToAll = 2
    ToMap = 3
    ToMapButIndex = 4
    ToPCArea = 5
    ToNPCArea = 6
End Enum

'Conversion buffer (for making packets)
Public conBuf As ByteBuffer

'Receive buffer (for reading packets)
Public rBuf As ByteBuffer

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

    'Start the loop through the packet
    Do
        
        'Get the packet ID
        PacketID = rBuf.Get_Byte
        
        'If the user does not exist, the only thing we will allow is PId.CS_Connect
        If UserList(inSox) Is Nothing Then
            If PacketID <> PId.CS_Connect Then Exit Sub
        End If
        
        'Forward to the corresponding packet handling method
        With PId
            Select Case PacketID
            
            Case PId.CS_Jump: Data_Jump inSox
            Case PId.CS_MoveEast: Data_MoveEast inSox
            Case PId.CS_MoveWest: Data_MoveWest inSox
            Case PId.CS_MoveStop: Data_MoveStop inSox
            Case PId.CS_Ping: Data_Ping inSox
            Case PId.CS_Punch: Data_Punch inSox
            Case PId.CS_Punch_Hit: Data_Punch_Hit inSox
            Case PId.CS_Pickup: Data_Pickup inSox
            Case PId.CS_DropItem: Data_DropItem inSox
            
            Case PId.CS_Connect: Data_Connect inSox
            Case PId.CS_UseInv: Data_UseInv inSox
            Case PId.CS_EquipRing: Data_EquipRing inSox
            Case PId.CS_UnEquip: Data_UnEquip inSox
            Case PId.CS_Chat_Say: Data_Chat_Say inSox
            Case PId.CS_MoveInvSlot: Data_MoveInvSlot inSox
            
            Case 0: rBuf.Overflow       'An error occured
            Case Else: rBuf.Overflow    'An error occured
            End Select
        End With
        
        'Check if the buffer ran out
        If rBuf.Get_ReadPos > DataUBound Then Exit Do
        
    Loop

End Sub

Private Sub Data_DropItem(ByVal inSox As Long)
'*********************************************************************************
'User drops an item from their inventory
'<InvSlot(B)>
'*********************************************************************************
Dim InvSlot As Byte

    InvSlot = rBuf.Get_Byte
    
    UserList(inSox).DropInvItem InvSlot

End Sub

Private Sub Data_Pickup(ByVal inSox As Long)
'*********************************************************************************
'User picks up an item off the ground of the map
'<ItemSlot(I)>
'*********************************************************************************
Dim ItemSlot As Integer

    ItemSlot = rBuf.Get_Integer
    
    'Invalid values handled by the sub
    UserList(inSox).Pickup ItemSlot

End Sub

Private Sub Data_MoveInvSlot(ByVal inSox As Long)
'*********************************************************************************
'User moves an inventory slot to another slot
'<SrcSlot(B)><DestSlot(B)>
'*********************************************************************************
Dim SrcSlot As Byte
Dim DestSlot As Byte

    SrcSlot = rBuf.Get_Byte
    DestSlot = rBuf.Get_Byte
    
    'Swap the slots (invalid values handled by the sub)
    UserList(inSox).SwapInvSlots SrcSlot, DestSlot

End Sub

Private Sub Data_UnEquip(ByVal inSox As Long)
'*********************************************************************************
'User unequips an equipped item
'<Slot(B)>
'*********************************************************************************
Dim Slot As Byte

    Slot = rBuf.Get_Byte
    
    Select Case Slot
    
    Case EQUIPSLOT_CAP
        UserList(inSox).SetCap 0
    
    Case EQUIPSLOT_CLOTHES
        UserList(inSox).SetClothes 0
    
    Case EQUIPSLOT_EARACC
        UserList(inSox).SetEarAcc 0
    
    Case EQUIPSLOT_EYEACC
        UserList(inSox).SetEyeAcc 0
    
    Case EQUIPSLOT_FOREHEAD
        UserList(inSox).SetForehead 0
    
    Case EQUIPSLOT_GLOVES
        UserList(inSox).SetGloves 0
    
    Case EQUIPSLOT_MANTLE
        UserList(inSox).SetMantle 0
    
    Case EQUIPSLOT_PANTS
        UserList(inSox).SetPants 0
    
    Case EQUIPSLOT_PENDANT
        UserList(inSox).SetPendant 0
    
    Case EQUIPSLOT_RING1
        UserList(inSox).SetRing1 0
    
    Case EQUIPSLOT_RING2
        UserList(inSox).SetRing2 0
    
    Case EQUIPSLOT_RING3
        UserList(inSox).SetRing3 0
    
    Case EQUIPSLOT_RING4
        UserList(inSox).SetRing4 0
    
    Case EQUIPSLOT_SHIELD
        UserList(inSox).SetShield 0
    
    Case EQUIPSLOT_SHOES
        UserList(inSox).SetShoes 0
    
    Case EQUIPSLOT_WEAPON
        UserList(inSox).SetWeapon 0
    
    End Select

End Sub

Private Sub Data_EquipRing(ByVal inSox As Long)
'*********************************************************************************
'User equips a ring in a specific slot
'<Slot(B)><RingIndex(B)>
'*********************************************************************************
Dim Slot As Byte
Dim RingIndex As Byte

    Slot = rBuf.Get_Byte
    RingIndex = rBuf.Get_Byte
    UserList(inSox).UseInvItem Slot, RingIndex

End Sub

Private Sub Data_UseInv(ByVal inSox As Long)
'*********************************************************************************
'User uses one of their inventory slots
'<Slot(B)>
'*********************************************************************************
Dim Slot As Byte

    Slot = rBuf.Get_Byte
    UserList(inSox).UseInvItem Slot

End Sub

Private Sub Data_Chat_Say(ByVal inSox As Long)
'*********************************************************************************
'User said something in the local chat
'<Text(S)>
'*********************************************************************************
Dim Text As String

    Text = rBuf.Get_String
    
    'Check if legal
    If Text = vbNullString Then Exit Sub
    If Len(Text) > MAXCHATLENGTH Then Exit Sub
    If Not IsLegalString(Text) Then Exit Sub
    
    'Forward the message to everyone else on the map
    conBuf.Clear
    conBuf.Put_Byte PId.SC_Chat_Say
    conBuf.Put_Integer UserList(inSox).CharIndex
    conBuf.Put_String Text
    Data_Send ToPCArea, inSox, conBuf.Get_Buffer()

End Sub

Private Sub Data_Punch_Hit(ByVal inSox As Long)
'*********************************************************************************
'User punched something / someone
'<HitIndex(I)>
'*********************************************************************************
Dim HitIndex As Integer

    HitIndex = rBuf.Get_Integer
    UserList(inSox).Punch HitIndex

End Sub

Private Sub Data_Punch(ByVal inSox As Long)
'*********************************************************************************
'User wants to punch
'<>
'*********************************************************************************

    UserList(inSox).Punch

End Sub

Private Sub Data_Ping(ByVal inSox As Long)
'*********************************************************************************
'User pings the server
'<>
'*********************************************************************************
Dim b(0) As Byte
    
   b(0) = PId.SC_Ping
   Data_Send ToIndex, inSox, b()

End Sub

Private Sub Data_Jump(ByVal inSox As Long)
'*********************************************************************************
'User jumps
'<>
'*********************************************************************************

    UserList(inSox).Jump

End Sub

Private Sub Data_MoveStop(ByVal inSox As Long)
'*********************************************************************************
'User stops moving
'<>
'*********************************************************************************

    UserList(inSox).MoveDir = 0

End Sub

Private Sub Data_MoveWest(ByVal inSox As Long)
'*********************************************************************************
'User starts moving to the West
'<>
'*********************************************************************************

    UserList(inSox).MoveDir = WEST

End Sub

Private Sub Data_MoveEast(ByVal inSox As Long)
'*********************************************************************************
'User starts moving to the East
'<>
'*********************************************************************************

    UserList(inSox).MoveDir = EAST

End Sub

Private Sub Data_Connect(ByVal inSox As Long)
'*********************************************************************************
'User connects to the server
'<Account(S)><CharName(S)><Password(S)>
'*********************************************************************************
Dim dbUser(1 To 5) As String
Dim Account As String
Dim CharName As String
Dim Password As String
Dim i As Long
Dim b(0) As Byte

    'Get the variables
    Account = rBuf.Get_String
    CharName = rBuf.Get_String
    Password = rBuf.Get_String
    Log "Attempting to log in [Account] " & Account & " [Char] " & CharName & " [Password] " & Password
    
    'Check if the user is already online
    i = Server_UserNameToIndex(CharName)
    If i > 0 Then
        b(0) = PId.SC_Message_UserAlreadyOnline
        frmMain.GOREsock.SendData inSox, b()
        'Close the user of the same name, that way the user doesn't get locked out of their account
        UserList(i).StatusFlag(PCSTATUSFLAG_DISCONNECTING) = True
        Log "User already online, disconnecting"
        Exit Sub
    End If
    
    'Confirm that the account is valid
    DB_RS.Open "SELECT * FROM accounts WHERE `name`='" & Account & "'", DB_Conn, adOpenStatic, adLockOptimistic
    If DB_RS.EOF Then
        DB_RS.Close
        frmMain.GOREsock.Shut inSox
        Exit Sub
    End If
    
    'Check if the password was valid
    If MD5_String(Password) <> DB_RS!Pass Then
        DB_RS.Close
        frmMain.GOREsock.Shut inSox
        Log "Invalid password supplied for account, disconnecting"
        Exit Sub
    End If
    
    'Store the information from the database
    dbUser(1) = DB_RS!user1
    dbUser(2) = DB_RS!user2
    dbUser(3) = DB_RS!user3
    dbUser(4) = DB_RS!user4
    dbUser(5) = DB_RS!user5
    DB_RS.Close
    
    'Check if the character is part of the user's account
    For i = 1 To 5
        If dbUser(i) = CharName Then    'No need for UCase$() since it should already be proper casing
            GoTo ValidUser
        End If
    Next i
    
    'If we made it here, the user is invalid
    Log "Invalid login attempt - character did not exist in account"
    Exit Sub
    
    'The user was valid
ValidUser:

    'Check if a user from the account is aleady online
    For i = 1 To 5
        If Server_UserNameToIndex(dbUser(i)) > 0 Then

            'User already online
            b(0) = PId.SC_Message_AccountInUse
            frmMain.GOREsock.SendData inSox, b()
            Log "Account already online, disconnecting"
            Exit Sub

        End If
    Next i
    
    'Create the user class
    If Not UserList(inSox) Is Nothing Then Set UserList(inSox) = Nothing
    Set UserList(inSox) = New User
    
    'Check for a new highest index
    If inSox > LastUser Then LastUser = inSox
    
    'Load the user
    UserList(inSox).Load CharName, inSox

    'Send the user's buffer
    UserList(inSox).SendBuffer
    
    Log "User logged in successfully"
    
End Sub

Public Sub Data_Send(ByVal Route As SendRoute, ByVal SendIndex As Integer, ByRef Data() As Byte)
'*********************************************************************************
'Sends data to a single or multiple users
'*********************************************************************************
Dim MapUsers() As Integer
Dim MapUsersUBound As Integer
Dim CopySize As Long
Dim i As Long

    'Get the copy size
    CopySize = UBound(Data) + 1
    
    'Check the send routes
    Select Case Route
        
    'Send to a single index
    Case SendRoute.ToIndex
        If UserList(SendIndex) Is Nothing Then Exit Sub
        UserList(SendIndex).SendData Data(), CopySize
    
    'Send to every user
    Case SendRoute.ToAll
        For i = 1 To LastUser
            If Not UserList(i) Is Nothing Then
                UserList(i).SendData Data(), CopySize
            End If
        Next i
    
    'Send to every user on a map
    Case SendRoute.ToMap
        Maps(SendIndex).GetMapUsers MapUsers, MapUsersUBound
        For i = 0 To MapUsersUBound
            If Not UserList(MapUsers(i)) Is Nothing Then
                UserList(MapUsers(i)).SendData Data(), CopySize
            End If
        Next i
    
    'Send to every user on the map but the SendIndex
    Case SendRoute.ToMapButIndex
        If UserList(SendIndex) Is Nothing Then Exit Sub
        Maps(UserList(SendIndex).Map).GetMapUsers MapUsers, MapUsersUBound
        For i = 0 To MapUsersUBound
            If MapUsers(i) <> SendIndex Then
                If Not UserList(MapUsers(i)) Is Nothing Then
                    UserList(MapUsers(i)).SendData Data(), CopySize
                End If
            End If
        Next i
        
    'Send to everyone in the user's view (plus a little more)
    Case SendRoute.ToPCArea
        If UserList(SendIndex) Is Nothing Then Exit Sub
        Maps(UserList(SendIndex).Map).GetMapUsers MapUsers, MapUsersUBound
        For i = 0 To MapUsersUBound
            If Not UserList(MapUsers(i)) Is Nothing Then
                If Math_Collision_PointRect(UserList(MapUsers(i)).X, UserList(MapUsers(i)).Y, _
                    UserList(SendIndex).X - ScreenWidth \ 2 - 100, UserList(SendIndex).Y - ScreenHeight \ 2 - 100, _
                    ScreenWidth + 200, ScreenHeight + 200) Then
                    UserList(MapUsers(i)).SendData Data(), CopySize
                End If
            End If
        Next i
        
    'Send to everyone in the NPC's view
    Case SendRoute.ToNPCArea
        If NPCList(SendIndex) Is Nothing Then Exit Sub
        Maps(NPCList(SendIndex).Map).GetMapUsers MapUsers, MapUsersUBound
        For i = 0 To MapUsersUBound
            If Not UserList(MapUsers(i)) Is Nothing Then
                If Math_Collision_PointRect(UserList(MapUsers(i)).X, UserList(MapUsers(i)).Y, _
                    NPCList(SendIndex).X - ScreenWidth \ 2 - 100, NPCList(SendIndex).Y - ScreenHeight \ 2 - 100, _
                    ScreenWidth + 200, ScreenHeight + 200) Then
                    UserList(MapUsers(i)).SendData Data(), CopySize
                End If
            End If
        Next i
    
    End Select
            
End Sub
