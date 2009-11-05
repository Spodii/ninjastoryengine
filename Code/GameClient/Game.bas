Attribute VB_Name = "Game"
Option Explicit

'The local socket ID
Public LocalSocketID As Long

'Send packet buffer
Public sndBuf As ByteBuffer
Public rBuf As ByteBuffer

'The map information
Public MapInfo As tMapInfo

'If the socket is open
Public SocketOpen As Boolean

'If we are getting the account
Public GettingAccount As Boolean

'Last time the user sent an action to the server
Public LastInputMoveTime As Long
Public LastInputJumpTime As Long
Public LastInputAttackTime As Long
Public LastInputPickupTime As Long
Public LastSentMoveDir As Byte

'NPC template list
Public NPCTemplate() As tClientNPC
Public NPCTemplateUBound As Integer

'Character list
Public CharList() As tClientChar    'Holds information on every character
Public CharListUBound As Integer    'Highest character index (UBound of CharList)
Public UserCharIndex As Integer     'The index of the user's character

'Damage list
Private Type tDamage
    X As Integer
    Y As Integer
    Value As Integer    '0 = unused, >0 = damage value, -1 = miss
    Life As Single
End Type
Private DamageList() As tDamage             'Damage array
Private DamageListUBound As Integer         'UBound of the array
Private LastDamageListIndex As Integer      'Highest used damage index
Private DamageListFreeIndexes As Integer    'Number of free indexes in the damage array
Private Const DamageLife As Single = 3000

'User's inventory
Public UserInv(0 To USERINVSIZE) As tClientInvSlot
Public InvSwapSlot As Integer

'Picked-up items
'Maximum number of items that can be doing the "pickup animation" at once
Public Const MAXPICKEDUPITEMS As Integer = 50
Public PickedupItem(0 To MAXPICKEDUPITEMS) As tClientPickupItem

'Fading items
Public Const MAXFADEITEMS As Integer = 20
Public FadeItem(0 To MAXFADEITEMS) As tClientFadeItem

'Item tooltip ID
Public TooltipItemIndex As Integer

'Messages
Public Msg As Messages

'Stats for our user
Public UserStats As tUserStats
Public User_ToNextLevel As Long

'Equipped items
Public CapItemIndex As Integer
Public CapGrh As tGrh
Public ForeheadItemIndex As Integer
Public ForeheadGrh As tGrh
Public Ring1ItemIndex As Integer
Public Ring1Grh As tGrh
Public Ring2ItemIndex As Integer
Public Ring2Grh As tGrh
Public Ring3ItemIndex As Integer
Public Ring3Grh As tGrh
Public Ring4ItemIndex As Integer
Public Ring4Grh As tGrh
Public EyeAccItemIndex As Integer
Public EyeAccGrh As tGrh
Public EarAccItemIndex As Integer
Public EarAccGrh As tGrh
Public MantleItemIndex As Integer
Public MantleGrh As tGrh
Public ClothesItemIndex As Integer
Public ClothesGrh As tGrh
Public PendantItemIndex As Integer
Public PendantGrh As tGrh
Public WeaponItemIndex As Integer
Public WeaponGrh As tGrh
Public ShieldItemIndex As Integer
Public ShieldGrh As tGrh
Public GlovesItemIndex As Integer
Public GlovesGrh As tGrh
Public PantsItemIndex As Integer
Public PantsGrh As tGrh
Public ShoesItemIndex As Integer
Public ShoesGrh As tGrh

Public Sub UpdateFadeItems()
'*********************************************************************************
'Updates all of the fading items
'*********************************************************************************
Dim i As Long

    'Loop through all the fading items
    For i = 0 To MAXFADEITEMS
        
        'Check that the item is in use
        If FadeItem(i).Alpha > 0 Then
        
            'Decrease the alpha
            FadeItem(i).Alpha = FadeItem(i).Alpha - ElapsedTime * 0.2
        
        End If
        
    Next i
    
End Sub

Public Sub UpdatePickedupItems()
'*********************************************************************************
'Updates all of the picked up items
'*********************************************************************************
Dim i As Long
Dim a As Single
    
    'Loop through all the items
    For i = 0 To MAXPICKEDUPITEMS
    
        'Check that the item is in use
        If PickedupItem(i).Used Then
        
            With PickedupItem(i)
                
                'Confirm that the character is valid
                If IsValidChar(.ToCharIndex) Then
                        
                    'Get the angle between the two
                    a = Math_GetAngle(.X + .Width \ 2, .Y + .Height \ 2, CharList(.ToCharIndex).X + CharList(.ToCharIndex).Width \ 2, _
                        CharList(.ToCharIndex).Y) * DegreeToRadian
                        
                    'Move the object
                    .X = .X + Sin(a) * ElapsedTime * 0.3
                    .Y = .Y - Cos(a) * ElapsedTime * 0.3
                    
                    'Check if the point has been reached
                    If Abs(.X + (.Width \ 2) - (CharList(.ToCharIndex).X + CharList(.ToCharIndex).Width \ 2)) _
                        + Abs(.Y + (.Height \ 2) - CharList(.ToCharIndex).Y) < 5 Then
                        .Used = False
                    End If
                    
                End If
                    
            End With
            
        End If
    Next i

End Sub

Public Function IsValidChar(ByVal CharIndex As Integer) As Boolean
'*********************************************************************************
'Checks if a character is of a valid index and is in use
'*********************************************************************************

    'Valid range
    If CharIndex < 1 Then Exit Function
    If CharIndex > CharListUBound Then Exit Function
    
    'Character in use
    If CharList(CharIndex).Used = 0 Then Exit Function
    
    'All is good
    IsValidChar = True

End Function

Public Sub AddFadeItem(ByVal X As Integer, ByVal Y As Integer, ByRef Grh As tGrh)
'*********************************************************************************
'Creates a fading item, used for items on a map whos life has expired
'*********************************************************************************
Dim Index As Integer

    'Find the array index to use
    Index = 0
    Do While FadeItem(Index).Alpha > 0
        Index = Index + 1
        If Index > MAXFADEITEMS Then Exit Sub
    Loop
    
    'Fill in the values
    FadeItem(Index).Alpha = 255
    FadeItem(Index).Grh = Grh
    FadeItem(Index).X = X
    FadeItem(Index).Y = Y

End Sub

Public Sub AddPickedupItem(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, _
    ByVal ToCharIndex As Integer, ByRef Grh As tGrh)
'*********************************************************************************
'Adds an item to the PickedupItem() array
'*********************************************************************************
Dim Index As Integer

    'Find the array index to use
    Index = 0
    Do While PickedupItem(Index).Used
        Index = Index + 1
        If Index > MAXPICKEDUPITEMS Then Exit Sub
    Loop
    
    'Fill in the values
    PickedupItem(Index).Used = True
    PickedupItem(Index).X = X
    PickedupItem(Index).Y = Y
    PickedupItem(Index).Width = Width
    PickedupItem(Index).Height = Height
    PickedupItem(Index).ToCharIndex = ToCharIndex
    PickedupItem(Index).Grh = Grh
    
End Sub

Public Sub UpdateMapItems()
'*********************************************************************************
'Updates the movement of map items
'*********************************************************************************
Dim i As Long

    'Check through all the items
    For i = 0 To LastMapItem
    
        'Confirm the item is moving
        If MapItems(i).Moving Then
        
            'Decrease the Y velocity
            MapItems(i).Yv = MapItems(i).Yv + ElapsedTime * (GRAVITY * 0.02)
            
            'Apply the velocity changes
            If Abs(MapItems(i).DestX - MapItems(i).X) < 32 Then
                MapItems(i).X = MapItems(i).X + MapItems(i).Xv
            End If
            MapItems(i).Y = MapItems(i).Y + MapItems(i).Yv
            
            'Confirm the item is still in the screen
            If MapItems(i).X < 0 Then
                MapItems(i).X = 0
                MapItems(i).Xv = -MapItems(i).Xv    'Wall bounce
            End If
            If MapItems(i).X + GrhData(Items(MapItems(i).ItemIndex).GrhIndex).Width > MapInfo.TileWidth * GRIDSIZE Then
                MapItems(i).X = MapInfo.TileWidth * GRIDSIZE - GrhData(Items(MapItems(i).ItemIndex).GrhIndex).Width
                MapItems(i).Xv = -MapItems(i).Xv    'Wall bounce
            End If

            'Check if the destination has been reached
            If MapItems(i).Y >= MapItems(i).DestY Then
                MapItems(i).Moving = False
            End If
            
        End If

    Next i

End Sub

Public Sub AddDamage(ByVal X As Integer, ByVal Y As Integer, ByVal Value As Integer)
'*********************************************************************************
'Adds damage to the DamageList() array
'*********************************************************************************
Dim i As Long

    'Add new slots to the array
    If DamageListFreeIndexes = 0 Then
        DamageListUBound = DamageListUBound + 20
        ReDim Preserve DamageList(1 To DamageListUBound)
        DamageListFreeIndexes = DamageListFreeIndexes + 20
    End If
    
    'Find the next free index
    For i = 1 To DamageListUBound
        If DamageList(i).Value = 0 Then
            DamageList(i).Value = Value
            DamageList(i).X = X
            DamageList(i).Y = Y
            DamageList(i).Life = DamageLife
            Exit Sub
        End If
    Next i

End Sub

Public Sub EraseMapItem(ByVal ItemSlot As Integer)
'*********************************************************************************
'Erases a map item off the map
'*********************************************************************************

    'Check for a valid item index
    If ItemSlot > MapItemsUBound Then Exit Sub

    'Erase the item
    ZeroMemory MapItems(ItemSlot), LenB(MapItems(ItemSlot))
    
    'Check if the highest used index has lowered
    If ItemSlot >= LastMapItem Then
    
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

End Sub

Public Sub ClearInputQueue()
'*********************************************************************************
'Clears the GetAsyncKeyState queue to prevent key presses from a long time
' ago falling into "have been pressed"
'*********************************************************************************
Dim i As Long

    For i = 1 To 145
        GetKey i
    Next i

End Sub

Public Sub Main()
'*********************************************************************************
'Primary entry point for the program
'*********************************************************************************
Dim PacketKeys() As String

    Randomize

    'Open the log file
    Log_Open App.Path & "\clientlog.txt"

    'Show the form
    Load frmMain
    frmMain.Hide
    frmMain.Visible = False

    'Resize the form to fit the screen
    frmMain.Width = Screen.TwipsPerPixelX * ScreenWidth
    frmMain.Height = Screen.TwipsPerPixelY * ScreenHeight
    
    'Make the common IDs
    InitCommonIDs
    
    'Create the classes
    Set sndBuf = New ByteBuffer
    Set rBuf = New ByteBuffer
    
    'Create the messages
    Set Msg = New Messages
    Msg.Load "English"
    
    'Set the socket encryption
    GenerateEncryptionKeys PacketKeys()
    frmMain.GOREsock.ClearPicture
    frmMain.GOREsock.SetEncryption PacketEncTypeServerIn, PacketEncTypeServerOut, PacketKeys()
    Erase PacketKeys
    
    'Load the NPC templates
    IO_ClientNPCs_Load NPCTemplateUBound, NPCTemplate()
    
    'Load the item information
    IO_ClientItems_Load ItemsUBound, Items()
    
    'Load the Grh data
    IO_GrhData_Load GrhData(), NumGrhs
    
    'Load the paper-dolling information
    IO_PD_Body_Load PDBody, NumPDBody
    IO_PD_Sprite_Load PDSprite, NumPDSprite
    
    'Set the initial swap slot
    InvSwapSlot = -1
    
    'Load up the connect form
    Load frmConnect
    frmConnect.Show

End Sub

Public Sub Data_Send()
'*********************************************************************************
'Send the data in the send buffer to the server
'*********************************************************************************

    'Check that theres data in the buffer
    If sndBuf.HasBuffer Then
        
        'Send the data
        frmMain.GOREsock.SendData LocalSocketID, sndBuf.Get_Buffer()

        'Add the buffer size to the bandwidth count
        'BufferSize (UBound + 1) + TCP (20) + IPv4 (20)
        BytesOut = BytesOut + UBound(sndBuf.Get_Buffer()) + 41

        'Clear the buffer
        sndBuf.Clear
        
    End If

End Sub

Public Sub InitEngine()
'*********************************************************************************
'Inits the engine
'*********************************************************************************
Dim sFullscreen As Boolean
Dim s32Bit As Boolean
Dim sVSync As Boolean

    'Show the main form
    frmMain.Show
    frmMain.Visible = True
    DoEvents
    
    'Get the settings
    sFullscreen = (Val(IO_INI_Read(App.Path & "\Data\Settings.ini", "GRAPHICS", "Fullscreen")) <> 0)
    s32Bit = (Val(IO_INI_Read(App.Path & "\Data\Settings.ini", "GRAPHICS", "32Bit")) <> 0)
    sVSync = (Val(IO_INI_Read(App.Path & "\Data\Settings.ini", "GRAPHICS", "VSync")) <> 0)
    
    'If fullscreen, load the custom cursor
    If sFullscreen Then
        CursorSpeed = Val(IO_INI_Read(App.Path & "\Data\Settings.ini", "GENERAL", "CursorSpeed"))
        Graphics_SetGrh CursorGrh, 5, ANIMTYPE_LOOP
    End If
    
    'Load up the engine
    Engine_Init
    Graphics_Init frmMain.hWnd, sFullscreen, s32Bit, sVSync
    Input_Init frmMain
    Audio_Init frmMain
    
    'Close frmConnect
    Unload frmConnect

End Sub

Public Sub GameLoop()
'*********************************************************************************
'Handles the main game loop
'*********************************************************************************
Dim FailedUnloads As Byte
Dim CLTime_Ping As Long
    
    'Load up the engine
    InitEngine
    
    'Clear the input queue
    ClearInputQueue

    Do While EngineRunning
    
        If CurrMapIndex > 0 Then
    
            'Handle input
            Input_Keys_Polling
                    
            'Update items on the map
            UpdateMapItems
            
            'Update the fading items
            UpdateFadeItems
            
            'Update the picked up items
            UpdatePickedupItems
                
            'Draw the screen
            DrawGame

        End If
        
        'Update the sound effects
        Sfx_Update
        
        'Make sure the music is looping
        Music_Loop
        
        'Check to ping the server
        If CLTime_Ping < timeGetTime Then
            CLTime_Ping = timeGetTime + 3000
            PingSentTime = timeGetTime
            sndBuf.Put_Byte PId.CS_Ping
        End If
        
        'Send the packet buffer
        Data_Send
            
        'Let windows do its events
        DoEvents
        
        'Update the FPS
        Engine_UpdateFPS
        
    Loop
    
    'Close the connection to the server
    Do While frmMain.GOREsock.Shut(LocalSocketID) = soxERROR
        DoEvents
        FailedUnloads = FailedUnloads + 1
        If FailedUnloads > 5 Then Exit Do
    Loop
    FailedUnloads = 0
    
    'Shut down the socket
    Do While frmMain.GOREsock.ShutDown <> soxERROR
        DoEvents
        FailedUnloads = FailedUnloads + 1
        If FailedUnloads > 5 Then Exit Do
    Loop
    FailedUnloads = 0

    'Unhook the socket
    frmMain.GOREsock.UnHook
    
    'Destroy the engine
    Engine_Destroy
    
    'Destroy the game
    DestroyGame
    
    'Unload the form
    Unload frmMain
    
    'Close
    End

End Sub

Private Sub DestroyGame()
'*********************************************************************************
'Destroy the aspects of the game not part of the engine
'*********************************************************************************

    'Destroy the buffers
    Set rBuf = Nothing
    Set sndBuf = Nothing

End Sub

Public Sub DrawGame()
'*********************************************************************************
'Draw the game screen
'*********************************************************************************
Dim tmpGrh As tGrh
Dim i As Long
Dim l As Long

    Graphics_BeginScene
        
        'Background
        Engine_DrawBackground
        
        'Map tiles that are behind the characters
        Engine_DrawMap 1
        
        'The characters
        Engine_DrawChars
        
        'Fading items
        For i = 0 To MAXFADEITEMS
            If FadeItem(i).Alpha > 0 Then
                l = D3DColorARGB(FadeItem(i).Alpha, 255, 255, 255)
                Graphics_DrawGrh FadeItem(i).Grh, FadeItem(i).X - ScreenX, FadeItem(i).Y - ScreenY, l, l, l, l
            End If
        Next i
        
        'Items that are being picked up
        For i = 0 To MAXPICKEDUPITEMS
            If PickedupItem(i).Used Then
                Graphics_DrawGrh PickedupItem(i).Grh, PickedupItem(i).X - ScreenX, PickedupItem(i).Y - ScreenY
            End If
        Next i
        
        'Items on the map
        Engine_DrawMapItems
        
        'Map tiles in front of the characters
        Engine_DrawMap 0
        
        'Draw damage
        For i = 1 To DamageListUBound
            If DamageList(i).Value > 0 Then
                FontDefault.Draw DamageList(i).Value, DamageList(i).X - ScreenX, _
                    DamageList(i).Y - ScreenY - ((DamageLife - DamageList(i).Life) \ 50), D3DColorARGB(200, 255, 0, 0)
                DamageList(i).Life = DamageList(i).Life - ElapsedTime
                If DamageList(i).Life < 0 Then DamageList(i).Value = 0
            End If
        Next i
        
        'Debug information
        Engine_DrawDebugInfo
        
        'GUI
        GUI_Draw
        
        'Draw the tooltip
        If TooltipItemIndex > 0 Then Engine_DrawItemToolTip TooltipItemIndex
        
        'Draw swap item
        If InvSwapSlot <> -1 Then
            tmpGrh = UserInv(InvSwapSlot).Grh
            Graphics_DrawGrh tmpGrh, CursorPos.X - 10, CursorPos.Y - 10
        End If

        'Cursor
        If Fullscreen Then
            Graphics_DrawGrh CursorGrh, CursorPos.X, CursorPos.Y
        End If
    
    Graphics_EndScene

End Sub

Public Sub MakePC(ByVal CharIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Name As String, ByVal Heading As Byte, _
    ByVal Body As Byte, Optional ByVal MoveDir As Byte = 0)
'*********************************************************************************
'Makes a new player character
'*********************************************************************************

    'Resize the charlist array if needed
    If CharIndex > CharListUBound Then
        CharListUBound = CharIndex
        ReDim Preserve CharList(1 To CharListUBound)
    End If
    
    'Empty out the old variables
    ZeroMemory CharList(CharIndex), LenB(CharList(CharIndex))
    
    'Create the character
    With CharList(CharIndex)
        .Name = Name
        .X = X
        .Y = Y
        .DrawX = X
        .DrawY = Y
        .Used = 1
        .Jump = 0
        .LastTileX = -1
        .LastTileY = -1
        .BodyIndex = Body
        .IsNPC = False
        
        'If the NPC is moving, start their movement
        Select Case MoveDir
        Case 0
            .Heading = Heading
        Case EAST
            .Heading = EAST
            Engine_ModifyXForLag X, EAST, MOVESPEED
            .X = X
        Case WEST
            .Heading = WEST
            Engine_ModifyXForLag X, WEST, MOVESPEED
            .X = X
        End Select
        .MoveDir = MoveDir
        
        'Size information
        .Width = PDBody(.BodyIndex).Width
        .Height = PDBody(.BodyIndex).Height
        
        'Start the animation
        Graphics_SetGrh .Body, PDBody(.BodyIndex).Stand, ANIMTYPE_LOOP
        
    End With

End Sub

Public Sub MakeNPC(ByVal CharIndex As Integer, ByVal X As Integer, ByVal Y As Integer, _
    ByVal TemplateID As Integer, Optional ByVal MoveDir As Byte = 0)
'*********************************************************************************
'Make a new non-player character
'*********************************************************************************

    'Resize the charlist array if needed
    If CharIndex > CharListUBound Then
        CharListUBound = CharIndex
        ReDim Preserve CharList(1 To CharListUBound)
    End If
    
    'Empty out the old variables
    ZeroMemory CharList(CharIndex), LenB(CharList(CharIndex))
  
    'Create the NPC character
    With CharList(CharIndex)
    
        'Set the starting position and movement variables
        .X = X
        .Y = Y
        .DrawX = X
        .DrawY = Y
        .Used = 1
        .Jump = 0
        .LastTileX = -1
        .LastTileY = -1
        
        'If the NPC is moving, start their movement
        Select Case MoveDir
        Case 0
            .Heading = EAST
        Case EAST
            .Heading = EAST
            Engine_ModifyXForLag X, EAST, MOVESPEED
            .X = X
        Case WEST
            .Heading = WEST
            Engine_ModifyXForLag X, WEST, MOVESPEED
            .X = X
        End Select
        .MoveDir = MoveDir
        
        'Template information
        .Name = NPCTemplate(TemplateID).Name
        .BodyIndex = NPCTemplate(TemplateID).Sprite
        .Width = NPCTemplate(TemplateID).Width
        .Height = NPCTemplate(TemplateID).Height
        .Heading = NPCTemplate(TemplateID).Heading
        
        'Set the character as a NPC
        .IsNPC = True

        'Start the animation
        Graphics_SetGrh .Body, PDSprite(.BodyIndex).Stand, ANIMTYPE_LOOP
        
    End With

End Sub
