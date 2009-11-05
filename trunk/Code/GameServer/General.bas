Attribute VB_Name = "General"
Option Explicit

'Server counters and loop controls (all times are in milliseconds)
Private Const SL_Frequency As Long = 10         'Server loop runs frequency
Private Const SL_UpdateUser As Long = 30        'Update user frequency
Private Const SL_UpdateNPC As Long = 30         'Update NPC frequency
Private Const SL_SendBuffer As Long = 100       'How frequent the user's buffer is sent
Private Const SL_SendAllStats As Long = 750     'Rate at which stats are checked to be sent (does not include HP/MP/EXP/Ryu)
Private Const SL_SendStats As Long = 300        'Rate at which HP, MP, EXP and Ryu are sent to the client
Private Const SL_KeepAlive As Long = 240000     'How frequently a query is made to the database to keep it alive
Private Const SL_UpdateMapItems As Long = 2000  'How often map items are checked to see if their life has expired
Private SLTime_UpdateUser As Long
Private SLTime_UpdateNPC As Long
Private SLTime_SendBuffer As Long
Private SLTime_SendAllStats As Long
Private SLTime_SendStats As Long
Private SLTime_KeepAlive As Long
Private SLTime_UpdateMapItems As Long

'Priority handling functions
Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Public Sub Server_Loop()
'*********************************************************************************
'Runs the primary server loop
'*********************************************************************************
Dim bSendStats As Boolean
Dim LoopStartTime As Long
Dim i As Integer

    'Begin the loop
    ServerRunning = True
    Do While ServerRunning
    
        'Store the starting time
        LoopStartTime = timeGetTime
        
        'Update the users
        If SLTime_UpdateUser <= timeGetTime Then
            SLTime_UpdateUser = timeGetTime + SL_UpdateUser
            For i = 1 To LastUser
                If Not UserList(i) Is Nothing Then UserList(i).Update
            Next i
        End If
        
        'Send all of the user's stats but Ryu, EXP, HP and MP
        If SLTime_SendAllStats <= timeGetTime Then
            SLTime_SendAllStats = timeGetTime + SL_SendAllStats
            For i = 1 To LastUser
                If Not UserList(i) Is Nothing Then UserList(i).SendAllStats
            Next i
        End If
        
        'Send Ryu, EXP, HP and MP
        If SLTime_SendStats <= timeGetTime Then
            SLTime_SendStats = timeGetTime + SL_SendStats
            For i = 1 To LastUser
                If Not UserList(i) Is Nothing Then UserList(i).SendStats
            Next i
        End If

        'Update the NPCs
        If SLTime_UpdateNPC <= timeGetTime Then
            SLTime_UpdateNPC = timeGetTime + SL_UpdateNPC
            For i = 1 To LastNPC
                If Not NPCList(i) Is Nothing Then
                    If NPCList(i).Map > 0 Then
                        If Maps(NPCList(i).Map).HasUsers Then
                            NPCList(i).Update
                        End If
                    End If
                End If
            Next i
        End If
        
        'Update map items
        If SLTime_UpdateMapItems <= timeGetTime Then
            SLTime_UpdateMapItems = timeGetTime + SL_UpdateMapItems
            For i = 1 To MapsUBound
                Maps(i).UpdateItemLife
            Next i
        End If
        
        'Send the buffers (make sure this is right before cool-down so all
        'of the game states have a change to update)
        If SLTime_SendBuffer <= timeGetTime Then
            SLTime_SendBuffer = timeGetTime + SL_SendBuffer
            For i = 1 To LastUser
                If Not UserList(i) Is Nothing Then UserList(i).SendBuffer
            Next i
        End If
        
        'Check if to send the "Keep-Alive" to the database to make sure the connection
        'with the database stays open and alive
        If SLTime_KeepAlive < timeGetTime Then
            SLTime_KeepAlive = timeGetTime + SL_KeepAlive
            
            'Do a small query to keep the connection alive
            DB_RS.Open "SELECT id FROM items WHERE `id`='1'", DB_Conn, adOpenStatic, adLockOptimistic
            DB_RS.Close
            
        End If
    
        'Cool-down
        DoEvents
        If timeGetTime - LoopStartTime < SL_Frequency Then
            Sleep SL_Frequency - (timeGetTime - LoopStartTime)
        End If
        
    Loop
    
    'Terminate the server
    frmMain.UnloadTmr.Enabled = True

End Sub

Public Sub Init_NPCs()
'*********************************************************************************
'Creates all of the temporary NPC files - this will convert the NPCs from the database
'to a binary file structured just like the tNPC type, which allows for very fast and
'easy access directly from the hard drive instead of the database
'*********************************************************************************
Dim HighestIndex As Integer 'Highest NPC index
Dim n() As tClientNPC   'Client NPC information
Dim t As tServerNPC     'Server NPC information
Dim DropTxt As String
Dim s1() As String
Dim s2() As String
Dim i As Long

    'Create the directory
    MakeFilePath App.Path & "\Server Data\temp\"

    'Get the highest NPC index
    DB_RS.Open "SELECT id FROM npcs ORDER BY id DESC LIMIT 1", DB_Conn, adOpenStatic, adLockOptimistic
    HighestIndex = Val(DB_RS(0))
    DB_RS.Close

    'Resize the array to fit all of the NPCs
    If MakeClientInfo Then ReDim n(1 To HighestIndex)

    'Grab the NPCs from the database
    DB_RS.Open "SELECT *,heading+0 FROM npcs", DB_Conn, adOpenStatic, adLockOptimistic
    
    'Loop through every recordset
    Do While Not DB_RS.EOF

        With t
        
            'Apply the values to the temporary server NPC
            .Name = Trim$(DB_RS!Name)
            .Sprite = DB_RS!Sprite
            .TemplateID = DB_RS!id
            .Spawn = DB_RS!Spawn
            .Stats.Ryu = DB_RS!stat_ryu
            .Stats.EXP = DB_RS!stat_exp
            .Stats.MaxHP = DB_RS!stat_hp
            .Stats.HP = DB_RS!stat_hp
            .Stats.BaseMaxHP = DB_RS!stat_hp
            .Stats.MaxMP = DB_RS!stat_mp
            .Stats.MP = DB_RS!stat_mp
            .Stats.BaseMaxMP = DB_RS!stat_mp
            .Stats.Str = DB_RS!stat_str
            .Stats.ModStr = DB_RS!stat_str
            .Stats.Dex = DB_RS!stat_dex
            .Stats.ModDex = DB_RS!stat_dex
            .Stats.Intl = DB_RS!stat_intl
            .Stats.ModIntl = DB_RS!stat_intl
            .Stats.Luk = DB_RS!stat_luk
            .Stats.ModLuk = DB_RS!stat_luk
            .Stats.MinHit = DB_RS!stat_minhit
            .Stats.MaxHit = DB_RS!stat_maxhit
            .Stats.Def = DB_RS!stat_def
            .Stats.Level = 1   '//!! Create level calculation
            If Trim$(DB_RS!Heading) = "WEST" Then .Heading = WEST Else .Heading = EAST

            'Create the item drop information
            DropTxt = Trim$(DB_RS!Drop)
            If LenB(DropTxt) > 0 Then
                s1() = Split(DropTxt, vbNewLine)
                .NumDrops = UBound(s1) + 1
                ReDim .Drops(0 To .NumDrops - 1)
                For i = 0 To UBound(s1)
                    s2() = Split(s1(i), ",")
                    .Drops(i).ItemIndex = Val(s2(0))
                    .Drops(i).Amount = Val(s2(1))
                    .Drops(i).Chance = Val(s2(2))
                Next i
            End If
            
            'Add the values to the client NPC, too
            If MakeClientInfo Then
                n(.TemplateID).Name = .Name
                n(.TemplateID).Sprite = .Sprite
                n(.TemplateID).Level = .Stats.Level
                n(.TemplateID).Width = SpriteInfo(.Sprite).Width
                n(.TemplateID).Height = SpriteInfo(.Sprite).Height
                n(.TemplateID).Heading = .Heading
            End If
            
        End With
        
        'Save the server NPC to a file
        IO_ServerNPC_Save DB_RS!id, t
        
        'Move to the next recordset
        DB_RS.MoveNext
    
    Loop
    
    'Close the recordset
    DB_RS.Close
    
    'Save the client NPC information
    If MakeClientInfo Then IO_ClientNPCs_Save HighestIndex, n()

End Sub

Public Function Server_UserNameToIndex(ByVal UserName As String) As Integer
'*********************************************************************************
'Takes a user name and returns the user's index (0 means they aren't online)
'*********************************************************************************
Dim i As Long

    'Turn the name to lowercase
    UserName = LCase$(UserName)

    'Loop through the users
    For i = 1 To UserListUBound
        
        'Check that the user exists
        If Not UserList(i) Is Nothing Then
            
            'Check the name
            If LCase$(UserList(i).Name) = UserName Then
                
                'Match found
                Server_UserNameToIndex = i
                Exit Function
                
            End If
        
        End If
        
    Next i

End Function

Sub Main()
'*********************************************************************************
'Entry point for the server - starts loading up all the server components
'*********************************************************************************

    'Make sure the server is not already running
    If App.PrevInstance Then
        MsgBox "You are already running an instance of the server!" & vbNewLine & _
            "Only one instance of the server per server ID may be run at a time.", vbOKOnly
        Unload frmMain
        End
    End If
    
    'Open the log file
    Log_Open App.Path & "\serverlog.txt"
    
    'Move the form to the tray
    Load frmMain
    frmMain.Visible = False
    frmMain.Hide
    TrayAdd frmMain, frmMain.Caption, MouseMove
    DoEvents
    
    'Create the byte buffers
    Set rBuf = New ByteBuffer
    Set conBuf = New ByteBuffer
    
    'Randomize the random values
    Randomize
    
    'Set the common IDs
    InitCommonIDs
    
    'Set the timer resolution
    timeBeginPeriod 1
    
    'Set the server priority
    Server_SetPriority (Val(IO_INI_Read(App.Path & "\Server Data\Settings.ini", "GAME", "HighPriority")) <> 0)

    'Check if we will be generating the client files
    MakeClientInfo = (Val(IO_INI_Read(App.Path & "\Server Data\Settings.ini", "GAME", "MakeClientInfo")) <> 0)

    'Make the maps
    Init_Maps
    
    'Load the body and sprite sizes
    Init_BodyInfo
    Init_SpriteInfo

    'Connect to the database
    If Not MySQL_Init Then
        Server_Unload
        End
        Exit Sub
    End If
    
    'Move the NPCs from the database to a binary file
    Init_NPCs
    
    'Create the items
    Init_Items
    
    'Create the cached server messages
    Init_CachedMessages

    'Create the socket
    Init_Socket
    
    'Start the server loop
    Server_Loop

End Sub

Private Sub Init_Items()
'*********************************************************************************
'Load up all the items
'*********************************************************************************
Dim ci() As tClientItem
Dim i As Long

    'Get the highest item index
    DB_RS.Open "SELECT id FROM items ORDER BY id DESC LIMIT 1", DB_Conn, adOpenStatic, adLockOptimistic
    ItemsUBound = Val(DB_RS(0))
    DB_RS.Close
    
    'Resize the items array
    ReDim Items(0 To ItemsUBound)
    If MakeClientInfo Then ReDim ci(0 To ItemsUBound)
    
    'Init the items
    For i = 0 To ItemsUBound
    
        'Create the item class
        Set Items(i) = New Item
        
        'Load the item information from the database
        Items(i).Init i
        
        'Store the item information in the array for the client
        If MakeClientInfo Then
            With Items(i)
                ci(i).Def = .Def
                ci(i).Dex = .Dex
                ci(i).HP = .HP
                ci(i).Intl = .Intl
                ci(i).ItemType = .ItemType
                ci(i).Luk = .Luk
                ci(i).MaxHit = .MaxHit
                ci(i).MinHit = .MinHit
                ci(i).MP = .MP
                ci(i).Str = .Str
                ci(i).Stacking = .Stacking
            End With
        End If
        
        'These values aren't stored in the item class, so we need to grab them directly
        'from the database instead of the class
        DB_RS.Open "SELECT name,`desc`,grhindex FROM items WHERE `id`='" & i & "'", DB_Conn, adOpenStatic, adLockOptimistic
        If Not DB_RS.EOF Then
            ci(i).Name = Trim$(DB_RS!Name)
            ci(i).Desc = DB_RS!Desc
            ci(i).GrhIndex = DB_RS!GrhIndex
        End If
        DB_RS.Close
        
    Next i
    
    'Save the client item information
    If MakeClientInfo Then IO_ClientItems_Save ItemsUBound, ci()

End Sub

Private Sub Init_BodyInfo()
'*********************************************************************************
'Load all of the body sizes and attack information
'*********************************************************************************
Dim i As Byte

    'Get the number of bodies
    BodyInfoUBound = Val(IO_INI_Read(App.Path & "\Data\Body.dat", "GENERAL", "NumBodies"))
    
    'Resize the array
    ReDim BodyInfo(0 To BodyInfoUBound)
    
    'Load the body sizes
    For i = 1 To BodyInfoUBound
        BodyInfo(i).Width = Val(IO_INI_Read(App.Path & "\Data\Body.dat", i, "Width"))
        BodyInfo(i).Height = Val(IO_INI_Read(App.Path & "\Data\Body.dat", i, "Height"))
        BodyInfo(i).PunchTime = Val(IO_INI_Read(App.Path & "\Data\Body.dat", i, "PunchTime"))
        BodyInfo(i).PunchWidth = Val(IO_INI_Read(App.Path & "\Data\Body.dat", i, "PunchWidth"))
    Next i
    
End Sub

Private Sub Init_CachedMessages()
'*********************************************************************************
'Loads up the cached server messages with no parameters
'*********************************************************************************
Dim i As Long

    For i = 1 To NumcMessages
        cMessage(i).Data(0) = PId.SC_Message
        cMessage(i).Data(1) = i
    Next i

End Sub

Private Sub Init_SpriteInfo()
'*********************************************************************************
'Load all of the sprite sizes and attack information
'*********************************************************************************
Dim i As Byte

    'Get the number of bodies
    SpriteInfoUBound = Val(IO_INI_Read(App.Path & "\Data\Sprite.dat", "GENERAL", "NumSprites"))
    
    'Resize the array
    ReDim SpriteInfo(0 To SpriteInfoUBound)
    
    'Load the sprite sizes
    For i = 1 To SpriteInfoUBound
        SpriteInfo(i).Width = Val(IO_INI_Read(App.Path & "\Data\Sprite.dat", i, "Width"))
        SpriteInfo(i).Height = Val(IO_INI_Read(App.Path & "\Data\Sprite.dat", i, "Height"))
        SpriteInfo(i).PunchTime = Val(IO_INI_Read(App.Path & "\Data\Sprite.dat", i, "PunchTime"))
        SpriteInfo(i).PunchWidth = Val(IO_INI_Read(App.Path & "\Data\Sprite.dat", i, "PunchWidth"))
    Next i
    
End Sub

Private Sub Init_Maps()
'*********************************************************************************
'Ready all of the maps
'*********************************************************************************
Dim i As Integer

    'Get the number of maps
    MapsUBound = Val(IO_INI_Read(App.Path & "\Maps\maps.ini", "GENERAL", "NumMaps"))
    
    'Resize the array
    ReDim Maps(1 To MapsUBound)
    
    'Create the objects
    For i = 1 To MapsUBound
        Set Maps(i) = New Map
        Maps(i).Create i
    Next i

End Sub

Public Sub Server_Unload()
'*********************************************************************************
'Unloads the server - this may end up being called a few times before it finally
'goes through since GOREsock can be picky on the shutdown
'*********************************************************************************

    'Terminate the socket
    If frmMain.GOREsock.ShutDown <> soxERROR Then
        frmMain.GOREsock.UnHook
        
        'Remove the icon from the system tray
        TrayDelete
        
        'Close the log file
        Log_Close
        
        'Unload the forms
        Unload frmSettings
        Unload frmMain

        'Finish up
        End
        
    End If

End Sub

Public Sub Server_SetPriority(ByVal High As Boolean)
'*********************************************************************************
'Sets the priority of the server between high and normal
'*********************************************************************************
    
    'Set the new priority
    If Not (IsHighPriority = High) Then
        If High Then
            SetThreadPriority GetCurrentThread, 2
            SetPriorityClass GetCurrentProcess, &H80
        Else
            SetThreadPriority GetCurrentThread, 0
            SetPriorityClass GetCurrentProcess, &H20
        End If
    End If
    
    'Store the priority
    IsHighPriority = High
    
End Sub

Private Sub Init_Socket()
'*********************************************************************************
'Create the socket and open it for connections
'*********************************************************************************
Dim PacketKeys() As String
Dim IsPublic As Boolean
Dim PublicIP As String
Dim Port As Integer

    'Get the settings
    Port = Val(IO_INI_Read(App.Path & "\Server Data\Settings.ini", "GAME", "Port"))
    IsPublic = (Val(IO_INI_Read(App.Path & "\Server Data\Settings.ini", "GENERAL", "Public")) <> 0)
    
    'Clear the socket image (cleans up a little RAM)
    frmMain.GOREsock.ClearPicture
    
    'Set the encryption
    GenerateEncryptionKeys PacketKeys
    frmMain.GOREsock.ClearPicture
    frmMain.GOREsock.SetEncryption PacketEncTypeServerIn, PacketEncTypeServerOut, PacketKeys()
    Erase PacketKeys
    
    'Load the socket
    If IsPublic Then
        PublicIP = Trim$(IO_INI_Read(App.Path & "\Server Data\Settings.ini", "GENERAL", "PublicIP"))
        LocalSocketID = frmMain.GOREsock.Listen(PublicIP, Port)
    Else
        LocalSocketID = frmMain.GOREsock.Listen("127.0.0.1", Port)
    End If
    
    'Check for a valid connection
    If frmMain.GOREsock.Address(LocalSocketID) = "-1" Then
        MsgBox "Error binding the socket to the IP address. Please make sure port " & Port & " is not in use." & vbNewLine & _
        "IsPublic is set to " & IsPublic & ". If this value is " & True & " and you are behind a router, make sure you " & vbNewLine & _
        "properly forward the port " & Port & " in your router's configuration.", vbOKOnly
        Unload frmMain
        End
    End If
    
End Sub

Public Sub Char_FreeIndex(ByVal CharIndex As Integer)
'*********************************************************************************
'Removes a CharIndex from usage
'*********************************************************************************
    
    'Confirm a valid CharIndex
    If CharIndex < 1 Then Exit Sub
    If CharIndex > CharListUBound Then Exit Sub

    'Clear the slot
    CharList(CharIndex).CharType = CHARTYPE_NONE
    CharList(CharIndex).Index = 0
    
    'Increase the unused slots count
    NumCharsFree = NumCharsFree + 1
    
    'If CharIndex was LastChar then find the new LastChar
    If CharIndex = LastChar Then
        Do While CharList(LastChar).CharType <> CHARTYPE_NONE
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If

End Sub

Public Sub NPC_Close(ByVal NPCIndex As Integer)
'*********************************************************************************
'Removes a NPC completely. This is not the same as removing a NPC from the map when
'you kill it, this will clear out all memory it uses in the server and no longer be
'able to reference it. Primarly for when unloading a map from memory or when a
'summoned NPC (one that does not respawn) dies.
'*********************************************************************************

    'Check that the NPC exists
    If NPCList(NPCIndex) Is Nothing Then Exit Sub

    'Unload the NPC from memory
    NPCList(NPCIndex).Unload
    Set NPCList(NPCIndex) = Nothing
    NumNPCsFree = NumNPCsFree + 1

End Sub

Public Function NPC_GetIndex() As Integer
'*********************************************************************************
'Creates an NPC object and returns the index used
'*********************************************************************************
Dim i As Long

    'Check for empty slots
    If NumNPCsFree > 0 Then
        For i = 1 To NPCListUBound
            If NPCList(i) Is Nothing Then
                
                'Return the free index
                NPC_GetIndex = i
                
                'Decrease the number of free slots
                NumNPCsFree = NumNPCsFree - 1
                
                'Create the object
                Set NPCList(NPC_GetIndex) = New NPC
                
                'Update the LastNPC
                If NPC_GetIndex > LastNPC Then LastNPC = NPC_GetIndex
                
                Exit Function
            End If
        Next i
    End If
    
    'Increase the LastNPC count since theres no free slots
    LastNPC = LastNPC + 1
    
    'Check to resize the NPCList() array
    If LastNPC > NPCListUBound Then
        NPCListUBound = NPCListUBound + 50  'Increase the array by 50
        ReDim Preserve NPCList(1 To NPCListUBound)
        NumNPCsFree = NumNPCsFree + 50
    End If
    
    'Return the lowest free index, which is the highest index, or NumNPCs
    NPC_GetIndex = LastNPC
    NumNPCsFree = NumNPCsFree - 1

    'Create the object
    Set NPCList(NPC_GetIndex) = New NPC

End Function

Public Function Char_GetIndex(ByVal CharType As Byte, ByVal Index As Integer) As Integer
'*********************************************************************************
'Returns the next free character index
'*********************************************************************************
Dim i As Long

    'Check for empty slots
    If NumCharsFree > 0 Then
        For i = 1 To CharListUBound
            If CharList(i).CharType = CHARTYPE_NONE Then
            
                'Return the free index
                Char_GetIndex = i
                
                'Set the variables in the CharList()
                CharList(Char_GetIndex).CharType = CharType
                CharList(Char_GetIndex).Index = Index
                
                'Decrease the number of free slots
                NumCharsFree = NumCharsFree - 1
                
                'Update the LastChar
                If Char_GetIndex > LastChar Then LastChar = Char_GetIndex
                
                Exit Function
            End If
        Next i
    End If
    
    'Increase the LastChar count since theres no free slots
    LastChar = LastChar + 1
    
    'Check to resize the CharList array
    If LastChar > CharListUBound Then
        CharListUBound = CharListUBound + 50    'Increase the array by 50
        ReDim Preserve CharList(1 To CharListUBound)
        NumCharsFree = NumCharsFree + 50
    End If

    'Return the lowest free index, which is the highest index, or NumChars
    Char_GetIndex = LastChar
    NumCharsFree = NumCharsFree - 1
    
    'Set the character type
    CharList(Char_GetIndex).CharType = CharType
    CharList(Char_GetIndex).Index = Index

End Function

Sub Graphics_SetGrh(ByRef pGrh As tGrh, ByVal GrhIndex As Long, Optional ByVal AnimType As Byte = 0)
'Dummy routine - just ignore it
End Sub
