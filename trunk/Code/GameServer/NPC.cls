VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NPC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'How frequently the position is updated when the NPC is moving
Private Const UPDATEPOSRATE As Long = 3000

'NPC's index in the NPCList() array
Private cNPCIndex As Integer

'Temporary/server-related information
Private cOnGround As Byte
Private cJump As Single
Private cLastUpdateTime As Long
Private cLastTileX As Integer
Private cLastTileY As Integer
Private cMoveDir As Byte
Private cPosUpdateTime As Long

'At what time the NPC will respawn
Private cRespawnTime As Long

'NPC's current action
Private cAction As eCharAction  'ID of the action
Private cActionTime As Long     'At what server tick the action will be over

'NPC type
Private n As tServerNPC

'If the mod stats need to be updated (happens when one of the base stats change that
'can affect the mod stats or an item is equipped/removed)
Private cUpdateModStats As Boolean

Public Sub Punch(Optional ByVal HitUserIndex As Integer = 0)
'*********************************************************************************
'Makes the user punch
'*********************************************************************************
Dim HitDamage As Long
Dim ValidHit As Boolean

    'Check if the user is already performing an action
    If cAction <> eNone Then Exit Sub
    
    'If the user is moving, stop them
    If MoveDir > 0 Then MoveDir = 0
    
    'Check for a valid hit index
    If HitUserIndex < 0 Then HitUserIndex = 0
    If HitUserIndex > LastUser Then HitUserIndex = 0
    If HitUserIndex > 0 Then
                                
        'Check the collision area
        If Math_Collision_Rect(UserList(HitUserIndex).X, UserList(HitUserIndex).Y, UserList(HitUserIndex).Width, _
            UserList(HitUserIndex).Height, n.X + Width - (SpriteInfo(n.Sprite).PunchWidth * -(n.Heading = WEST)), _
            n.Y, SpriteInfo(n.Sprite).PunchWidth, Height) Then
            ValidHit = True
            
            'Calculate the hit damage
            HitDamage = Math_Random(MinHit, MaxHit)
            If HitDamage > 32000 Then HitDamage = 32000
            
            'Damage the user
            UserList(HitUserIndex).DamageByNPC cNPCIndex, HitDamage
            
            'Confirm the heading is correct
            If UserList(HitUserIndex).X > n.X Then
                Heading = EAST
            Else
                Heading = WEST
            End If
            
        End If

    End If
    
    'Set the user's action time
    cActionTime = timeGetTime + SpriteInfo(n.Sprite).PunchTime
    
    'Set the user's action
    cAction = ePunch
    
    'Send the action packet
    If ValidHit Then
        conBuf.Clear
        conBuf.Put_Byte PId.SC_Punch_Hit
        conBuf.Put_Integer n.CharIndex
        conBuf.Put_Integer UserList(HitUserIndex).CharIndex
        conBuf.Put_Integer HitDamage
        Data_Send ToMap, n.Map, conBuf.Get_Buffer()
    Else
        conBuf.Clear
        conBuf.Put_Byte PId.SC_Punch
        conBuf.Put_Integer n.CharIndex
        Data_Send ToMap, n.Map, conBuf.Get_Buffer()
    End If
    
End Sub

Public Sub SetHitAction()
'*********************************************************************************
'Set the NPC's action to being hit
'*********************************************************************************

    'Set the action
    cAction = eHit
    cActionTime = timeGetTime + HITTIME

    'If the NPC is moving, stop them
    If MoveDir <> 0 Then MoveDir = 0

End Sub

Public Property Get Action() As eCharAction
'*********************************************************************************
'Get the NPC's current action
'*********************************************************************************

    Action = cAction

End Property

Public Property Get Heading() As Byte
'*********************************************************************************
'Get the NPC's heading
'*********************************************************************************

    Heading = n.Heading

End Property

Public Property Let Heading(ByVal Value As Byte)
'*********************************************************************************
'Set the NPC's heading
'*********************************************************************************
    
    'Check for a valid value
    If Value <> EAST Then
        If Value <> WEST Then
            Exit Property
        End If
    End If
    
    'Confirm the value has changed
    If Value <> n.Heading Then
    
        'Don't change the heading when being hit
        If Action = eHit Then
            If Value <> 0 Then Exit Property
        End If
        
        'Set the heading
        n.Heading = Value
        
        'Send the new heading
        conBuf.Clear
        conBuf.Put_Byte PId.SC_Char_SetHeading
        conBuf.Put_Integer n.CharIndex
        conBuf.Put_Byte n.Heading
        Data_Send ToMap, n.Map, conBuf.Get_Buffer()
        
    End If

End Property

Public Property Get MoveDir() As Byte
'*********************************************************************************
'Get the direction the NPC is moving
'*********************************************************************************

    MoveDir = cMoveDir
        
End Property

Public Property Let MoveDir(ByVal Value As Byte)
'*********************************************************************************
'Set the direction the NPC is moving
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
    
        'Don't allow movement to be set when being hit
        If Action = eHit Then
            If Value <> 0 Then Exit Property
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
        conBuf.Put_Integer n.CharIndex
        conBuf.Put_Integer n.X
        conBuf.Put_Integer n.Y
        Data_Send ToMap, n.Map, conBuf.Get_Buffer()
            
        'Change the value
        cMoveDir = Value
        
        'Set the heading
        If Value <> 0 Then n.Heading = Value
    
    End If

End Property

Public Sub Jump()
'*********************************************************************************
'Makes the NPC jump
'*********************************************************************************

    If cJump = 0 Then
        If cOnGround = 1 Then
            cJump = JUMPHEIGHT
            conBuf.Clear
            conBuf.Put_Byte PId.SC_Jump
            conBuf.Put_Integer n.CharIndex
            conBuf.Put_Integer n.X
            conBuf.Put_Integer n.Y
            Data_Send ToMap, n.Map, conBuf.Get_Buffer()
        End If
    End If

End Sub

Public Sub Update()
'*********************************************************************************
'Updates the NPC
'*********************************************************************************

    'Check if the NPC is spawned
    If StatusFlag(NPCSTATUSFLAG_SPAWNED) = False Then
    
        'Check if it is time to respawn
        If cRespawnTime < timeGetTime Then
            Spawn
        End If
    
    Else
    
        'Check if the NPC's action time has run out
        If cAction <> eNone Then
            If cActionTime < timeGetTime Then
                cActionTime = 0
                cAction = eNone
            End If
        End If
    
        'AI
        UpdateAI
        
        'Update mod stats
        If cUpdateModStats Then UpdateModStats
        
        'Position update
        UpdatePos
        
    End If

End Sub

Private Sub UpdateAI()
'*********************************************************************************
'Updates the NPC's AI
'*********************************************************************************
Dim MapUsers() As Integer
Dim MapUsersUBound As Integer
Dim i As Long

    'Test AI
    If Action = eNone Then
    
        'Attacking
        If Int(Rnd * 25) = 0 Then
            Maps(n.Map).GetMapUsers MapUsers(), MapUsersUBound
            For i = 0 To MapUsersUBound
                If Abs(UserList(MapUsers(i)).Y - n.Y) < Height Then
                    If Abs(UserList(MapUsers(i)).X - n.X) < SpriteInfo(n.Sprite).PunchWidth Then
                        '//!! Punch MapUsers(i)
                    End If
                End If
            Next i
        End If

        'Movement
        If MoveDir = 0 Then
            If Int(Rnd * 100) = 0 Then
                If Int(Rnd * 2) = 0 Then MoveDir = EAST Else MoveDir = WEST
            End If
        Else
            If Int(Rnd * 25) = 0 Then MoveDir = 0
        End If
        
        'Jumping
        If Int(Rnd * 100) = 0 Then Jump
        
    End If
    
End Sub

Public Property Get Width() As Integer
'*********************************************************************************
'Returns the NPC's collision width
'*********************************************************************************

    Width = SpriteInfo(n.Sprite).Width

End Property

Public Property Get Height() As Integer
'*********************************************************************************
'Returns the NPC's collision height
'*********************************************************************************

    Height = SpriteInfo(n.Sprite).Height

End Property

Private Sub UpdatePos()
'*********************************************************************************
'Updates the NPC's position
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

    'Make sure the NPC is spawned
    If Not StatusFlag(NPCSTATUSFLAG_SPAWNED) Then Exit Sub

    'Get the elapsed time
    ElapsedTime = timeGetTime - cLastUpdateTime
    If ElapsedTime > 33 Then ElapsedTime = 33
    
    'Get the X co-ordinate remainder
    RemainderX = X Mod GRIDSIZE
    
    'Store the old OnGround
    OldOnGround = cOnGround

    'NPC is jumping
    If cJump > 0 Then
        Y = Y - (ElapsedTime * cJump * MOVESPEED)
        cJump = cJump - (ElapsedTime * JUMPDECAY)
        If cJump < 0 Then cJump = 0
        cOnGround = 0
    End If
    
    'Update the NPC's position
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
    
        If Maps(n.Map).HasFloatingBlocks Then
        
            'Check for blocking to the right
            If TileChangeX > 0 Then
                CheckTileX = ((X + Width - 5) \ GRIDSIZE)
                If CheckTileX <= Maps(n.Map).TileWidth Then
                    For i = 0 To Height \ GRIDSIZE
                        CheckTileY = i + (Y \ GRIDSIZE)
                        If Maps(n.Map).TileInfo(CheckTileX, CheckTileY) = TILETYPE_BLOCKED Then
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
                If CheckTileX <= Maps(n.Map).TileWidth Then
                    For i = 0 To Height \ GRIDSIZE
                        CheckTileY = i + (Y \ GRIDSIZE)
                        If Maps(n.Map).TileInfo(CheckTileX, CheckTileY) = TILETYPE_BLOCKED Then
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
        If CheckTileY <= Maps(n.Map).TileHeight Then
            For i = 0 To (Width + RemainderX) \ GRIDSIZE
                CheckTileX = i + (X \ GRIDSIZE)
                If Maps(n.Map).TileInfo(CheckTileX, CheckTileY) = TILETYPE_BLOCKED Or _
                    Maps(n.Map).TileInfo(CheckTileX, CheckTileY) = TILETYPE_PLATFORM Then
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
                    Select Case Maps(n.Map).TileInfo(CheckTileX, NewCheckTile)
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
        If Maps(n.Map).HasFloatingBlocks Then

            'Head-hitting handling
            If UpdateY Then
                NewCheckTile = (Y \ GRIDSIZE)
                For i = 0 To (Width + RemainderX) \ GRIDSIZE
                    CheckTileX = i + (X \ GRIDSIZE)
                    If Maps(n.Map).TileInfo(CheckTileX, NewCheckTile) = TILETYPE_BLOCKED Then
                        Y = (NewCheckTile * GRIDSIZE) + GRIDSIZE - 1
                        cOnGround = 0
                        cJump = 0
                        Exit For
                    End If
                Next i
            End If
            
        End If
    End If
    
    'Check to update the NPC's position while they are moving
    If MoveDir <> 0 Then
        If cPosUpdateTime < timeGetTime Then
            cPosUpdateTime = timeGetTime + UPDATEPOSRATE
            conBuf.Clear
            conBuf.Put_Byte PId.SC_Char_UpdatePos
            conBuf.Put_Integer n.CharIndex
            conBuf.Put_Integer X
            conBuf.Put_Integer Y
            Data_Send ToNPCArea, cNPCIndex, conBuf.Get_Buffer()
        End If
    End If
    
End Sub

Public Function Load(ByVal pTemplateIndex As Integer, ByVal pNPCIndex As Integer, ByVal pMap As Integer) As Boolean
'*********************************************************************************
'Loads the NPC into the class
'*********************************************************************************

    'Load the NPC
    If Not IO_ServerNPC_Load(pTemplateIndex, n) Then
        
        'There was an error (probably the NPC doesn't exist)
        Exit Function
    
    End If
    
    'Add the NPC to the map
    Maps(pMap).AddNPC pNPCIndex
    
    'Set the misc variables
    cNPCIndex = pNPCIndex
    n.Map = pMap
    
    'Give the NPC a CharIndex
    n.CharIndex = Char_GetIndex(CHARTYPE_NPC, cNPCIndex)
    
    'All successful
    Load = True
    
End Function

Public Sub Spawn()
'*********************************************************************************
'Makes the NPC spawn on the map in a random location
'*********************************************************************************
Dim SpawnX As Integer
Dim SpawnY As Integer

    'Check if the NPC is already spawned
    If StatusFlag(NPCSTATUSFLAG_SPAWNED) Then Exit Sub

    'Get the spawn tile
    If Maps(n.Map).GetSpawnTile(SpawnX, SpawnY) = False Then
    
        'We couldn't get a spawn tile
        Exit Sub
        
    End If
    
    'Refill the NPC's HP
    HP = n.Stats.MaxHP
    
    'Set the NPC's position
    X = (CLng(SpawnX) * GRIDSIZE) - Width + Int(Rnd * 32)   'Randomly place them anywhere on the tile
    Y = (CLng(SpawnY + 1) * GRIDSIZE) - Height              'Add 1 to the height to put them at the bottom of the tile
    
    'Create the NPC on the map
    conBuf.Clear
    AddMakeCharToBuffer
    Data_Send ToMap, n.Map, conBuf.Get_Buffer()
    
    'Set the last update time to now
    cLastUpdateTime = timeGetTime
    
    'Set the NPC as spawned
    StatusFlag(NPCSTATUSFLAG_SPAWNED) = True

End Sub

Public Property Get StatusFlag(ByVal Flag As Long) As Boolean
'*********************************************************************************
'Return the status flag
'*********************************************************************************

    StatusFlag = (n.Flags.Status And Flag)

End Property

Public Property Let StatusFlag(ByVal Flag As Long, ByVal Value As Boolean)
'*********************************************************************************
'Set the status flag
'*********************************************************************************

    If Value Then
    
        'Enable
        If Not (n.Flags.Status And Flag) Then
            n.Flags.Status = n.Flags.Status Or Flag
        End If
        
    Else
    
        'Disable
        If (n.Flags.Status And Flag) Then
            n.Flags.Status = n.Flags.Status Xor Flag
        End If
        
    End If

End Property

Public Sub Unload()
'*********************************************************************************
'Shut down the NPC
'*********************************************************************************

    'Remove the NPC from the map
    Maps(n.Map).RemoveNPC cNPCIndex
    
    'Free the CharIndex slot
    Char_FreeIndex n.CharIndex

End Sub

Public Sub AddMakeCharToBuffer()
'*********************************************************************************
'Adds the MakeChar (Char_MakeNPC) packet information to the conversion buffer
'*********************************************************************************

    Select Case cMoveDir
    Case WEST
        conBuf.Put_Byte PId.SC_Char_MakeNPC_MoveWest
    Case EAST
        conBuf.Put_Byte PId.SC_Char_MakeNPC_MoveEast
    Case Else
        conBuf.Put_Byte PId.SC_Char_MakeNPC
    End Select
    conBuf.Put_Integer n.CharIndex
    conBuf.Put_Integer X
    conBuf.Put_Integer Y
    conBuf.Put_Integer n.TemplateID

End Sub

Public Property Get Y() As Single

    'Return the NPC's Y position
    Y = n.Y

End Property

Public Property Get X() As Single
    
    'Return the NPC's X position
    X = n.X
    
End Property

Private Property Let X(ByVal Value As Single)

    'Set the NPC's X co-ordinate
    If Value < 0 Then Value = 0
    If Value > Maps(n.Map).TileWidth * GRIDSIZE Then Value = (Maps(n.Map).TileWidth - 2) * GRIDSIZE
    n.X = Value

End Property

Public Property Get Map() As Integer

    'Return the NPC's Map position
    Map = n.Map

End Property

Private Property Let Y(ByVal Value As Single)

    'Set the NPC's Y co-ordinate
    If Value < 0 Then Value = 0
    If Value > Maps(n.Map).TileHeight * GRIDSIZE Then Value = (Maps(n.Map).TileHeight - 2) * GRIDSIZE
    n.Y = Value

End Property

Private Sub UpdateModStats()
'*********************************************************************************
'Updates the NPC's modified stats
'*********************************************************************************

    'Start by setting the mod values equal to the base value
    MaxHP = 10
    MaxMP = 10
    ModStr = Str
    ModDex = Dex
    ModIntl = Intl
    ModLuk = Luk
    
    'All done updating
    cUpdateModStats = False

End Sub

Public Property Get ModLuk() As Integer
'*********************************************************************************
'Return the NPC's mod luck
'*********************************************************************************

    ModLuk = n.Stats.ModLuk
    
End Property

Public Property Let ModLuk(ByVal Value As Integer)
'*********************************************************************************
'Set the NPC's mod luck
'*********************************************************************************

    n.Stats.ModLuk = Value
    
End Property

Public Property Get ModIntl() As Integer
'*********************************************************************************
'Return the NPC's mod intelligence
'*********************************************************************************

    ModIntl = n.Stats.ModIntl
    
End Property

Public Property Let ModIntl(ByVal Value As Integer)
'*********************************************************************************
'Set the NPC's mod intelligence
'*********************************************************************************

    n.Stats.ModIntl = Value
    
End Property

Public Property Get ModDex() As Integer
'*********************************************************************************
'Return the NPC's mod dexterity
'*********************************************************************************

    ModDex = n.Stats.ModDex
    
End Property

Public Property Let ModDex(ByVal Value As Integer)
'*********************************************************************************
'Set the NPC's mod speed
'*********************************************************************************

    n.Stats.ModDex = Value
    
End Property

Public Property Get ModStr() As Integer
'*********************************************************************************
'Return the NPC's mod strength
'*********************************************************************************

    ModStr = n.Stats.ModStr
    
End Property

Public Property Let ModStr(ByVal Value As Integer)
'*********************************************************************************
'Set the NPC's mod strength
'*********************************************************************************

    n.Stats.ModStr = Value
    
End Property

Public Property Get MaxMP() As Integer
'*********************************************************************************
'Return the NPC's mod max mana
'*********************************************************************************

    MaxMP = n.Stats.MaxMP
    
End Property

Public Property Let MaxMP(ByVal Value As Integer)
'*********************************************************************************
'Set the NPC's mod mana
'*********************************************************************************

    n.Stats.MaxMP = Value
    
End Property

Public Property Get BaseMaxMP() As Integer
'*********************************************************************************
'Return the NPC's base max mana
'*********************************************************************************

    BaseMaxMP = n.Stats.BaseMaxMP
    
End Property

Public Property Let BaseMaxMP(ByVal Value As Integer)
'*********************************************************************************
'Set the NPC's base mana
'*********************************************************************************

    n.Stats.BaseMaxMP = Value
    
End Property

Public Property Get MaxHP() As Integer
'*********************************************************************************
'Return the NPC's max health
'*********************************************************************************

    MaxHP = n.Stats.MaxHP
    
End Property

Public Property Let MaxHP(ByVal Value As Integer)
'*********************************************************************************
'Set the NPC's max health
'*********************************************************************************

    n.Stats.MaxHP = Value
    
End Property

Private Property Get BaseMaxHP() As Integer
'*********************************************************************************
'Return the NPC's base max health
'*********************************************************************************

    BaseMaxHP = n.Stats.BaseMaxHP
    
End Property

Private Property Let BaseMaxHP(ByVal Value As Integer)
'*********************************************************************************
'Set the NPC's base max health
'*********************************************************************************

    n.Stats.BaseMaxHP = Value
    
End Property

Public Property Get Luk() As Integer
'*********************************************************************************
'Return the NPC's luck
'*********************************************************************************

    Luk = n.Stats.Luk
    
End Property

Public Property Let Luk(ByVal Value As Integer)
'*********************************************************************************
'Set the NPC's luck
'*********************************************************************************

    n.Stats.Luk = Value
    cUpdateModStats = True
    
End Property

Public Property Get Intl() As Integer
'*********************************************************************************
'Return the NPC's intelligence
'*********************************************************************************

    Intl = n.Stats.Intl
    
End Property

Public Property Let Intl(ByVal Value As Integer)
'*********************************************************************************
'Set the NPC's intelligence
'*********************************************************************************

    n.Stats.Intl = Value
    cUpdateModStats = True
    
End Property

Public Property Get Dex() As Integer
'*********************************************************************************
'Return the NPC's dexterity
'*********************************************************************************

    Dex = n.Stats.Dex
    
End Property

Public Property Let Dex(ByVal Value As Integer)
'*********************************************************************************
'Set the NPC's speed
'*********************************************************************************

    n.Stats.Dex = Value
    cUpdateModStats = True
    
End Property

Public Property Get Str() As Integer
'*********************************************************************************
'Return the NPC's strength
'*********************************************************************************

    Str = n.Stats.Str
    
End Property

Public Property Let Str(ByVal Value As Integer)
'*********************************************************************************
'Set the NPC's strength
'*********************************************************************************

    n.Stats.Str = Value
    cUpdateModStats = True
    
End Property

Public Property Get HP() As Integer
'*********************************************************************************
'Return the NPC's health
'*********************************************************************************

    HP = n.Stats.HP
    
End Property

Public Property Let HP(ByVal Value As Integer)
'*********************************************************************************
'Set the NPC's speed
'*********************************************************************************

    If Value > n.Stats.MaxHP Then Value = n.Stats.MaxHP
    n.Stats.HP = Value
    
End Property

Public Sub DamageByPC(ByVal UserIndex As Integer, ByVal Damage As Long)
'*********************************************************************************
'A player inflicts damage upon the NPC
'*********************************************************************************

    'Inflict the damage
    HP = HP - Damage
    
    'Check if the NPC's HP ran out
    If HP < 1 Then
        
        'Kill off the NPC
        Kill
        
        'Give the user the EXP
        UserList(UserIndex).GiveEXPandRyu EXP, Ryu
        
    Else
    
        'Hit the NPC
        SetHitAction
    
    End If

End Sub

Public Sub Kill()
'*********************************************************************************
'Kills off the NPC
'*********************************************************************************
Dim i As Long

    'Set the respawn timer and flag
    StatusFlag(NPCSTATUSFLAG_SPAWNED) = False
    cRespawnTime = timeGetTime + n.Spawn
    
    'Erase the NPC from the map
    conBuf.Clear
    conBuf.Put_Byte PId.SC_Char_Kill
    conBuf.Put_Integer n.CharIndex
    Data_Send ToMap, n.Map, conBuf.Get_Buffer()
    
    'Item dropping
    If n.NumDrops > 0 Then
        For i = 0 To n.NumDrops - 1
            If Rnd * 100 <= n.Drops(i).Chance Then
                Maps(n.Map).AddItem n.Drops(i).ItemIndex, n.Drops(i).Amount, X, Y + Height
            End If
        Next i
    End If
    
End Sub

Public Property Get MP() As Integer
'*********************************************************************************
'Return the NPC's mana
'*********************************************************************************

    MP = n.Stats.MP
    
End Property

Public Property Let MP(ByVal Value As Integer)
'*********************************************************************************
'Set the NPC's mana
'*********************************************************************************

    If Value > n.Stats.MaxMP Then Value = n.Stats.MaxMP
    n.Stats.MP = Value
    
End Property

Public Property Get Level() As Integer
'*********************************************************************************
'Return the NPC's level
'*********************************************************************************

    Level = n.Stats.Level
    
End Property

Public Property Let Level(ByVal Value As Integer)
'*********************************************************************************
'Set the NPC's level
'*********************************************************************************

    n.Stats.Level = Value
    cUpdateModStats = True
    
End Property

Public Property Get Ryu() As Integer
'*********************************************************************************
'Return the NPC's Ryu
'*********************************************************************************

    Ryu = n.Stats.Ryu
    
End Property

Public Property Let Ryu(ByVal Value As Integer)
'*********************************************************************************
'Set the NPC's Ryu
'*********************************************************************************

    n.Stats.Ryu = Value

End Property

Public Property Get EXP() As Integer
'*********************************************************************************
'Return the NPC's EXP
'*********************************************************************************

    EXP = n.Stats.EXP
    
End Property

Public Property Let EXP(ByVal Value As Integer)
'*********************************************************************************
'Set the NPC's EXP
'*********************************************************************************

    n.Stats.EXP = Value

End Property

Public Property Get MinHit() As Integer
'*********************************************************************************
'Return the NPC's MinHit
'*********************************************************************************

    MinHit = n.Stats.MinHit
    
End Property

Public Property Let MinHit(ByVal Value As Integer)
'*********************************************************************************
'Set the NPC's MinHit
'*********************************************************************************

    n.Stats.MinHit = Value
    cUpdateModStats = True
    
End Property

Public Property Get MaxHit() As Integer
'*********************************************************************************
'Return the NPC's MaxHit
'*********************************************************************************

    MaxHit = n.Stats.MaxHit
    
End Property

Public Property Let MaxHit(ByVal Value As Integer)
'*********************************************************************************
'Set the NPC's MaxHit
'*********************************************************************************

    n.Stats.MaxHit = Value
    cUpdateModStats = True
    
End Property

Public Property Get Def() As Integer
'*********************************************************************************
'Return the NPC's defense
'*********************************************************************************

    Def = n.Stats.Def
    
End Property

Public Property Let Def(ByVal Value As Integer)
'*********************************************************************************
'Set the NPC's defense
'*********************************************************************************

    n.Stats.Def = Value
    cUpdateModStats = True
    
End Property
