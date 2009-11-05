Attribute VB_Name = "CommonIDs"
'*********************************************************************************
'This module contains common IDs between the server and client. These IDs are often
'used to decrease the size of packets, and can include anything from stating what
'packet it is, different skills and spells, stats, etc. Every ID must be identical
'on the server and client both, which is why this is in the Common code.
'
'For the packets, theres two kinds of prefixes:
'  CS - Client to Server
'  SC - Server to Client
'*********************************************************************************

Option Explicit

'Packet IDs for the game server / client
Public Type PId
    CS_Connect As Byte
    CS_MoveWest As Byte
    CS_MoveEast As Byte
    CS_MoveStop As Byte
    CS_Jump As Byte
    CS_Ping As Byte
    CS_Punch As Byte
    CS_Punch_Hit As Byte
    CS_Chat_Say As Byte
    CS_UseInv As Byte
    CS_UnEquip As Byte
    CS_MoveInvSlot As Byte
    CS_Pickup As Byte
    CS_DropItem As Byte
    CS_EquipRing As Byte
    
    SC_Char_SetPos As Byte
    SC_Char_MakePC As Byte
    SC_Char_MakePC_MoveEast As Byte
    SC_Char_MakePC_MoveWest As Byte
    SC_Char_MakeNPC As Byte
    SC_Char_MakeNPC_MoveEast As Byte
    SC_Char_MakeNPC_MoveWest As Byte
    SC_Char_SetPaperDoll As Byte
    SC_Char_UpdatePos As Byte
    SC_Char_Erase As Byte
    SC_Char_Kill As Byte
    SC_User_SetMap As Byte
    SC_User_SetIndex As Byte
    SC_Move_WestStart As Byte
    SC_Move_WestEnd As Byte
    SC_Move_EastStart As Byte
    SC_Move_EastEnd As Byte
    SC_Message_UserAlreadyOnline As Byte
    SC_Message_AccountInUse As Byte
    SC_Ping As Byte
    SC_Jump As Byte
    SC_Punch As Byte
    SC_Punch_Hit As Byte
    SC_Message As Byte
    SC_Chat_Say As Byte
    SC_User_Stats As Byte
    SC_User_HP As Byte
    SC_User_MP As Byte
    SC_User_EXP As Byte
    SC_User_Ryu As Byte
    SC_User_ToNextLevel As Byte
    SC_Char_HPMP As Byte
    SC_Char_SetHeading As Byte
    SC_Inv_Update As Byte
    SC_Inv_UpdateSlot As Byte
    SC_SetEquipped As Byte
    SC_Item_Make As Byte
    SC_Item_Erase As Byte
    SC_Item_Pickup As Byte
    SC_Item_Drop As Byte
End Type
Public PId As PId

'Packet IDs for the account server / client
Public Type AccountPId
    CS_GetChars As Byte
    
    SC_SendChars As Byte
    SC_NoChars As Byte
    SC_BadPass As Byte
End Type
Public AccountPId As AccountPId

Public Sub InitCommonIDs()
'*********************************************************************************
'Sets the values for all the common IDs. This routine must be called before any
'of these values are used.
'*********************************************************************************

    With PId
        .CS_Connect = 1
        .CS_MoveWest = 2
        .CS_MoveEast = 3
        .CS_Jump = 4
        .CS_MoveStop = 5
        .CS_Ping = 6
        .CS_Punch = 7
        .CS_Punch_Hit = 8
        .CS_Chat_Say = 9
        .CS_UseInv = 10
        .CS_UnEquip = 11
        .CS_MoveInvSlot = 12
        .CS_Pickup = 13
        .CS_DropItem = 14
        .CS_EquipRing = 15
        
        .SC_Char_SetPos = 1
        .SC_User_SetMap = 2
        .SC_User_SetIndex = 3
        .SC_Char_MakePC = 4
        .SC_Move_EastStart = 5
        .SC_Move_EastEnd = 6
        .SC_Move_WestStart = 7
        .SC_Move_WestEnd = 8
        .SC_Char_SetPaperDoll = 9
        .SC_Ping = 10
        .SC_Jump = 11
        .SC_Char_UpdatePos = 12
        .SC_Message_UserAlreadyOnline = 13
        .SC_Message_AccountInUse = 14
        .SC_Char_Erase = 15
        .SC_Char_MakeNPC = 16
        .SC_Punch = 17
        .SC_Punch_Hit = 18
        .SC_Message = 19
        .SC_Char_Kill = 20
        .SC_Chat_Say = 21
        .SC_User_Stats = 22
        .SC_User_HP = 23
        .SC_User_MP = 24
        .SC_User_EXP = 25
        .SC_User_Ryu = 26
        .SC_User_ToNextLevel = 27
        .SC_Char_HPMP = 28
        .SC_Char_SetHeading = 29
        .SC_Inv_Update = 30
        .SC_Inv_UpdateSlot = 31
        .SC_SetEquipped = 32
        .SC_Item_Make = 33
        .SC_Item_Erase = 34
        .SC_Item_Pickup = 35
        .SC_Char_MakePC_MoveEast = 36
        .SC_Char_MakePC_MoveWest = 37
        .SC_Char_MakeNPC_MoveEast = 38
        .SC_Char_MakeNPC_MoveWest = 39
        .SC_Item_Drop = 40
    End With

    With AccountPId
        .CS_GetChars = 1
        
        .SC_SendChars = 1
        .SC_NoChars = 2
        .SC_BadPass = 3
    End With

End Sub

