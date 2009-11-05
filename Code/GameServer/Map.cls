VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Information for the map
Private MapInfo As tMapInfo

'If the map is loaded into memory
Private IsLoaded As Boolean

'The index of the map for this class
Private MapIndex As Integer

'The array of spawn tiles
Private SpawnTile() As tTilePos
Private SpawnTileUBound As Integer

'Users on the map
Private MapUsers() As Integer       'Indexes of the users on the map
Private MapUsersUBound As Integer   'Size of the MapUsers() array
Private LastMapUser As Integer      'Highest used index in the MapUsers() array
Private NumMapUsersFree As Integer  'Number of free slots in the MapUsers() array

'NPCs on the map
Private MapNPCs() As Integer        'Indexes of the npcs on the map
Private MapNPCsUBound As Integer    'Size of the MapNPCs() array
Private LastMapNPC As Integer       'Highest used index in the MapNPCs() array
Private NumMapNPCsFree As Integer   'Number of free slots in the MapNPCs() array

'Objects on the map
Private MapItems() As tServerMapItem    'List of objects on the map
Private MapItemsUBound As Integer       'UBound of the MapObjs() array
Private LastMapItem As Integer          'Highest in-use MapObjs() index
Private NumMapItemsFree As Integer      'Number of unused indicies in the MapObjs() array

Public Property Get NumMapItems() As Integer
'*********************************************************************************
'Returns the number of map items
'*********************************************************************************

    NumMapItems = LastMapItem

End Property

Public Property Get MapItemX(ByVal ItemSlot As Integer) As Integer
'*********************************************************************************
'Returns the X pixel co-ordinate of the item
'*********************************************************************************

    'Check for a valid range
    If ItemSlot > LastMapItem Then Exit Property
    If ItemSlot < 0 Then Exit Property
    
    MapItemX = MapItems(ItemSlot).X

End Property

Public Property Get MapItemY(ByVal ItemSlot As Integer) As Integer
'*********************************************************************************
'Returns the Y pixel co-ordinate of the item
'*********************************************************************************

    'Check for a valid range
    If ItemSlot > LastMapItem Then Exit Property
    If ItemSlot < 0 Then Exit Property
    
    MapItemY = MapItems(ItemSlot).Y

End Property

Public Property Get MapItemIndex(ByVal ItemSlot As Integer) As Integer
'*********************************************************************************
'Returns the index of a map item
'*********************************************************************************
    
    'Check for a valid range
    If ItemSlot > LastMapItem Then Exit Property
    If ItemSlot < 0 Then Exit Property
    
    MapItemIndex = MapItems(ItemSlot).ItemIndex
    
End Property

Public Property Get MapItemAmount(ByVal ItemSlot As Integer) As Integer
'*********************************************************************************
'Returns the amount of a map item
'*********************************************************************************
    
    'Check for a valid range
    If ItemSlot > LastMapItem Then Exit Property
    If ItemSlot < 0 Then Exit Property
    
    MapItemAmount = MapItems(ItemSlot).Amount
    
End Property

Public Sub AddItem(ByVal pItemIndex As Integer, ByVal pAmount As Integer, ByVal pX As Integer, ByVal pY As Integer)
'*********************************************************************************
'Creates an item on the map at the specified location
'*********************************************************************************
Dim ItemSlot As Integer
Dim TileX As Integer
Dim TileY As Integer
Dim OldpY As Integer

    'Get the index for the item in the MapItems() array
    ItemSlot = GetMapItemSlot
    
    'Pull the item down to the closest platform or blocked tile
    TileX = pX \ GRIDSIZE
    TileY = pY \ GRIDSIZE
    Do While TileInfo(TileX, TileY) <> TILETYPE_BLOCKED And TileInfo(TileX, TileY) <> TILETYPE_PLATFORM
        TileY = TileY + 1
        If TileY > TileHeight Then
            TileY = TileHeight
            Exit Do
        End If
    Loop
    OldpY = pY - GRIDSIZE
    pY = (TileY - 1) * GRIDSIZE
    
    'Fill up the information
    MapItems(ItemSlot).ItemIndex = pItemIndex
    MapItems(ItemSlot).Amount = pAmount
    MapItems(ItemSlot).X = pX
    MapItems(ItemSlot).Y = pY
    MapItems(ItemSlot).Life = timeGetTime + MAPITEMLIFE
    
    'Send a packet to everyone in the map telling them about the new item
    conBuf.Clear
    conBuf.Put_Byte PId.SC_Item_Drop
    conBuf.Put_Integer ItemSlot
    conBuf.Put_Integer pItemIndex
    conBuf.Put_Integer pAmount
    conBuf.Put_Integer pX
    conBuf.Put_Integer OldpY
    conBuf.Put_Integer pY
    Data_Send ToMap, MapIndex, conBuf.Get_Buffer

End Sub

Public Sub UpdateItemLife()
'*********************************************************************************
'Checks if any items on the map have expired, and removes them if they have
'*********************************************************************************
Dim i As Long

    'Loop through all the active items
    For i = 0 To LastMapItem
        
        'Confirm it is a valid item
        If MapItems(i).ItemIndex > 0 Then
        
            'Check the life
            If MapItems(i).Life < timeGetTime Then
            
                'The life has run out, so delete it
                RemoveItem i
                
            End If
            
        End If
        
    Next i

End Sub

Private Function GetMapItemSlot() As Integer
'*********************************************************************************
'Returns the next free slot for the map
'*********************************************************************************
Dim i As Integer

    'Check for any existing free slots
    If NumMapItemsFree > 0 Then
        For i = 0 To MapItemsUBound
        
            'Any unused slot will have their ObjIndex as 0
            If MapItems(i).ItemIndex = 0 Then
            
                'Return the unused slot
                GetMapItemSlot = i
                
                'Decrease the "free slots" count
                NumMapItemsFree = NumMapItemsFree - 1
                
                'Check if we have a new highest index
                If GetMapItemSlot > LastMapItem Then LastMapItem = GetMapItemSlot
                
                Exit Function
                
            End If
        Next i
    End If
    
    'It is safe to assume the first index of the new slots we will make is going to be available
    GetMapItemSlot = MapItemsUBound + 1
    
    'We have no free slots, so we must make more by resizing the array
    MapItemsUBound = MapItemsUBound + 25
    ReDim Preserve MapItems(0 To MapItemsUBound)
    
    'We made 25 more slots which are now empty, so we have to reflect that in the number of free items count
    'Remember, we are already using the first index that we just made
    NumMapItemsFree = NumMapItemsFree + 24
    
    'Check if we have a new highest index
    If GetMapItemSlot > LastMapItem Then LastMapItem = GetMapItemSlot

End Function

Public Sub RemoveItem(ByVal pItemSlot As Integer, Optional ByVal SendPacket As Boolean = True)
'*********************************************************************************
'Removes an item from the MapItems() array by its array index
'*********************************************************************************

    'Clear the item
    ZeroMemory MapItems(pItemSlot), LenB(MapItems(pItemSlot))
    
    'Increase the count on the number of free slots
    NumMapItemsFree = NumMapItemsFree + 1
    
    'Check if the highest used index has lowered
    If pItemSlot >= LastMapItem Then
    
        'Loop through all of the items, in reverse, until we find a used one
        Do While MapItems(LastMapItem).ItemIndex = 0
        
            'Decrease the index to look at
            LastMapItem = LastMapItem - 1
            
            'Check if we hit the end of the list
            If LastMapItem = -1 Then
                
                'No slots are in use
                Exit Do
                
            End If
            
        Loop
        
    End If
    
    'Send the packet to everyone on the map telling them to remove the item
    If SendPacket Then
        conBuf.Clear
        conBuf.Put_Byte PId.SC_Item_Erase
        conBuf.Put_Integer pItemSlot
        Data_Send ToMap, MapIndex, conBuf.Get_Buffer()
    End If

End Sub

Public Property Get HasFloatingBlocks() As Boolean
'*********************************************************************************
'Returns the HasFloatingBlocks value
'*********************************************************************************

    HasFloatingBlocks = MapInfo.HasFloatingBlocks

End Property

Public Function HasUsers() As Boolean
'*********************************************************************************
'If the map has any users on it
'*********************************************************************************

    HasUsers = (LastMapUser > -1)

End Function

Public Function HasNPCs() As Boolean
'*********************************************************************************
'If the map has any NPCs on it
'*********************************************************************************

    HasNPCs = (LastMapNPC > -1)
    
End Function

Public Sub GetMapUsers(ByRef pMapUsers() As Integer, ByRef pMapUsersUBound As Integer)
'*********************************************************************************
'Return the array of users on the map
'*********************************************************************************

    'Check if the map is loaded
    If Not IsLoaded Then
        pMapUsersUBound = -1
        Exit Sub
    End If

    'We return the UBound as LastMapUser, even though its not the "true" UBound, but
    'every index after it should be 0, so theres no point in checking them
    pMapUsers() = MapUsers()
    pMapUsersUBound = LastMapUser

End Sub

Public Sub GetMapNPCs(ByRef pMapNPCs() As Integer, ByRef pMapNPCsUBound As Integer)
'*********************************************************************************
'Return the array of NPCs on the map
'*********************************************************************************

    'Check if the map is loaded
    If Not IsLoaded Then
        pMapNPCsUBound = -1
        Exit Sub
    End If

    'Same concept as GetMapUsers
    pMapNPCs() = MapNPCs()
    pMapNPCsUBound = LastMapNPC

End Sub

Public Sub AddNPC(ByVal NPCIndex As Integer)
'*********************************************************************************
'Adds a NPC to the map's list of NPCs
'*********************************************************************************
Dim i As Integer

    'Check if the map is loaded
    If Not IsLoaded Then Load

    'Check if theres any free slots
    If NumMapNPCsFree > 0 Then
        For i = 0 To MapNPCsUBound
            If MapNPCs(i) = 0 Then
                MapNPCs(i) = NPCIndex
                NumMapNPCsFree = NumMapNPCsFree - 1
                If i > LastMapNPC Then LastMapNPC = i
                Exit Sub
            End If
        Next i
    End If
    
    'No free slots, so make some
    MapNPCsUBound = MapNPCsUBound + 20
    ReDim Preserve MapNPCs(0 To MapNPCsUBound)
    NumMapNPCsFree = NumMapNPCsFree + 20

    'Increase the LastMapNPC
    LastMapNPC = LastMapNPC + 1
    
    'Use the next free index, which is the LastMapNPC
    MapNPCs(LastMapNPC) = NPCIndex
    NumMapNPCsFree = NumMapNPCsFree + 1

End Sub

Public Sub AddUser(ByVal UserIndex As Integer)
'*********************************************************************************
'Adds a user to the map's list of users
'*********************************************************************************
Dim i As Integer

    'Check if the map is loaded
    If Not IsLoaded Then Load

    'Check if theres any free slots
    If NumMapUsersFree > 0 Then
        For i = 0 To MapUsersUBound
            If MapUsers(i) = 0 Then
                MapUsers(i) = UserIndex
                NumMapUsersFree = NumMapUsersFree - 1
                If i > LastMapUser Then LastMapUser = i
                Exit Sub
            End If
        Next i
    End If
    
    'No free slots, so make some
    MapUsersUBound = MapUsersUBound + 20
    ReDim Preserve MapUsers(0 To MapUsersUBound)
    NumMapUsersFree = NumMapUsersFree + 20
    
    'Increase the LastMapUser
    LastMapUser = LastMapUser + 1
    
    'Use the next free index, which is the LastMapUser
    MapUsers(LastMapUser) = UserIndex
    NumMapUsersFree = NumMapUsersFree - 1

End Sub

Public Sub RemoveNPC(ByVal NPCIndex As Integer)
'*********************************************************************************
'Removes a NPC from the map's list of NPCs
'*********************************************************************************
Dim i As Integer

    'Check if the map is loaded
    If Not IsLoaded Then Exit Sub

    'Loop through every NPC looking for the NPC's index
    For i = 0 To LastMapNPC
        If MapNPCs(i) = NPCIndex Then
            MapNPCs(i) = 0
            NumMapNPCsFree = NumMapNPCsFree + 1
            Exit For
        End If
    Next i
    
    'If the NPC was the LastMapNPC, find the new LastMapNPC
    If i = LastMapNPC Then
        Do While MapNPCs(LastMapNPC) = 0
            LastMapNPC = LastMapNPC - 1
            If LastMapNPC = -1 Then Exit Do
        Loop
    End If

End Sub

Public Sub RemoveUser(ByVal UserIndex As Integer)
'*********************************************************************************
'Removes a user from the map's list of users
'*********************************************************************************
Dim i As Integer

    'Check if the map is loaded
    If Not IsLoaded Then Exit Sub

    'Loop through every user looking for the user's index
    For i = 0 To LastMapUser
        If MapUsers(i) = UserIndex Then
            MapUsers(i) = 0
            NumMapUsersFree = NumMapUsersFree + 1
            Exit For
        End If
    Next i
    
    'If the user was the LastMapUser, find the new LastMapUser
    If i = LastMapUser Then
        Do While MapUsers(LastMapUser) = 0
            LastMapUser = LastMapUser - 1
            If LastMapUser = -1 Then Exit Do
        Loop
    End If
    
    'If the last user has left the map, stop all NPC movement on the map
    If LastMapUser = -1 Then
        For i = 0 To LastMapNPC
            If MapNPCs(i) > 0 Then
                NPCList(MapNPCs(i)).MoveDir = 0
            End If
        Next i
    End If

End Sub

Public Function GetSpawnTile(ByRef retTileX As Integer, ByRef retTileY As Integer) As Boolean
'*********************************************************************************
'Returns a random spawn tile
'*********************************************************************************
Dim i As Integer

    'Check if the map is loaded
    If Not IsLoaded Then Load

    'Check if we have any spawn tiles
    If SpawnTileUBound = -1 Then Exit Function

    'Get a random tile
    i = Int(Rnd * (SpawnTileUBound + 1))
    
    'Return the tile positions
    retTileX = SpawnTile(i).TileX
    retTileY = SpawnTile(i).TileY
    
    'All went well
    GetSpawnTile = True
    
End Function

Public Property Get TileInfo(ByVal TileX As Integer, ByVal TileY As Integer) As Byte
'*********************************************************************************
'Get the tile information for the specified tile
'*********************************************************************************

    'Make sure the map is loaded
    Load
    
    'Return the tile property
    If TileX > MapInfo.TileWidth Then
        TileInfo = TILETYPE_BLOCKED
        Exit Property
    End If
    If TileY > MapInfo.TileHeight Then
        TileInfo = TILETYPE_BLOCKED
        Exit Property
    End If
    If TileX < 0 Then
        TileInfo = TILETYPE_BLOCKED
        Exit Property
    End If
    If TileY < 0 Then
        TileInfo = TILETYPE_BLOCKED
        Exit Property
    End If
    TileInfo = MapInfo.TileInfo(TileX, TileY)
    
End Property

Public Property Get TileHeight() As Integer
'*********************************************************************************
'Width of the map in tiles
'*********************************************************************************

    'Make sure the map is loaded
    Load
    
    'Return the name
    TileHeight = MapInfo.TileHeight

End Property

Public Property Get TileWidth() As Integer
'*********************************************************************************
'Width of the map in tiles
'*********************************************************************************

    'Make sure the map is loaded
    Load
    
    'Return the name
    TileWidth = MapInfo.TileWidth

End Property

Public Property Get Name() As String
'*********************************************************************************
'Get the name
'*********************************************************************************

    'Make sure the map is loaded
    Load
    
    'Return the name
    Name = MapInfo.Name

End Property

Private Sub Load()
'*********************************************************************************
'Load the map information
'*********************************************************************************
Dim NPCSpawn() As tMapSpawn
Dim NPCSpawnUBound As Integer
Dim i As Long
Dim j As Long
Dim n As Integer

    'Confirm the map is not loaded
    If Not IsLoaded Then
    
        'Load the map information
        IO_Map_LoadInfo MapIndex, MapInfo
        IO_Map_LoadServerInfo MapIndex, SpawnTile(), SpawnTileUBound, NPCSpawn(), NPCSpawnUBound

        'Set as loaded
        IsLoaded = True
        
        'Empty out the internal user list
        MapUsersUBound = -1
        Erase MapUsers
        LastMapUser = -1
        NumMapUsersFree = 0
        
        'Empty out the internal npc list
        MapNPCsUBound = -1
        Erase MapNPCs
        LastMapNPC = -1
        NumMapNPCsFree = 0
        
        'Empty out the interal items list
        MapItemsUBound = -1
        Erase MapItems
        LastMapItem = -1
        NumMapItemsFree = 0
        
        'Spawn the NPCs
        For i = 0 To NPCSpawnUBound
            For j = 1 To NPCSpawn(i).Amount
                n = NPC_GetIndex
                NPCList(n).Load NPCSpawn(i).NPCID, n, MapIndex
                NPCList(n).Spawn
            Next j
        Next i
        
        '//!!
        For i = 0 To 300
            Maps(MapIndex).AddItem 1, 1, Rnd * Maps(MapIndex).TileWidth * GRIDSIZE, 10
        Next i
            
    End If
    
End Sub

Public Sub Unload()
'*********************************************************************************
'Unload the map information
'*********************************************************************************
Dim i As Integer

    'Confirm the map is loaded
    If IsLoaded Then
        
        'Set as unloaded
        IsLoaded = False
        
        'Remove the map information
        MapInfo.Name = vbNullString
        Erase MapInfo.TileInfo
        
        'Delete all the NPCs on the map
        For i = 0 To LastMapNPC
            If MapNPCs(i) > 0 Then
                NPC_Close MapNPCs(i)
            End If
        Next i
        
        'Remove the internal NPC and User arrays, along with the map items
        MapUsersUBound = -1
        Erase MapUsers
        LastMapUser = -1
        NumMapUsersFree = 0
        MapNPCsUBound = -1
        Erase MapNPCs
        LastMapNPC = -1
        NumMapNPCsFree = 0
        MapItemsUBound = -1
        Erase MapItems
        LastMapItem = -1
        NumMapItemsFree = 0
        
    End If

End Sub

Public Sub Create(ByVal pMapIndex As Integer)
'*********************************************************************************
'Set the map's index - MUST be done before anything else, and only called once!
'*********************************************************************************

    IsLoaded = False
    MapIndex = pMapIndex
    MapUsersUBound = -1
    LastMapUser = -1
    MapNPCsUBound = -1
    LastMapNPC = -1
    MapItemsUBound = -1
    LastMapItem = -1
    
End Sub

