Attribute VB_Name = "Engine"
'*********************************************************************************
'Handles all common aspects of the different components, namely the public
'functions and variables used between the different engine parts
'*********************************************************************************

Option Explicit

'DirectX
Public DX As DirectX8

'States if the engine is running
Public EngineRunning As Boolean

'Grh data
Public GrhData() As tGrhData
Public NumGrhs As Long

'Map graphics
Public MapGrh() As tMapGrh
Public NumMapGrhs As Long

'Paper-dolling information
Public PDBody() As tPDBody
Public NumPDBody As Byte
Public PDSprite() As tPDSprite
Public NumPDSprite As Byte

'Index of the currently loaded map
Public CurrMapIndex As Long

'Screen location (top-left corner's pixel offset from the top-left of the screen)
Public ScreenX As Long
Public ScreenY As Long

'The value of ScreenX/Y at which the map render buffer was last calculated
Private Const ScreenCalcSize As Integer = 64    'How often (in pixel distance) to calculate the MapGrhPtr
Private LastScreenXCalc As Integer
Private LastScreenYCalc As Integer
Private LastScreenCalcMapIndex As Integer       'Used to update if the map changes

'Map graphics to loop through to check to render
Private MapGrhPtr() As Integer
Public NumMapGrhPtrs As Integer
Public UpdateMapGrhPtr As Boolean

'Background information
Public BG(1 To NumBGLayers) As tBackground
Public BGSizeX As Byte
Public BGSizeY As Byte

'Debug mode - displays debug information
Private DebugMode As Boolean

'Bandwidth information
Public BytesIn As Long
Public BytesOut As Long

'Timing
Public EngineStartTime As Long  'When the engine started
Public FPS As Long              'The current FPS value
Private FPSCounter As Long      'Counts up the FPS
Private FPSUpdateTime As Long   'When to update the FPS
Public ElapsedTime As Long      'How much time elapsed on the last frame
Private LastFrame As Long       'The tick time of the last frame

'Ping information
Public Const TOTALPINGS As Byte = 3 'Total number of pings to average over
Public Ping As Single               'Average of the last few pings
Public NumPings As Byte             'Number of pings we have stored
Public PingSentTime As Long         'When the ping was sent to the server
Public PingTime(0 To TOTALPINGS - 1) As Long    'The time the past few pings took

'Map items
Public MapItems() As tClientMapItem     'Information on each item on the map
Public MapItemsUBound As Integer        'UBound of the MapItems() array
Public LastMapItem As Integer           'Highest MapItems() index in use

'Item list
Public Items() As tClientItem
Public ItemsUBound As Integer

'Input functions
Public Declare Function GetKey Lib "user32.dll" Alias "GetAsyncKeyState" (ByVal vKey As Long) As Integer

Private Sub Engine_MakeMapGrhPtr()
'*********************************************************************************
'Make the optimized array for the map grh list
'*********************************************************************************
Dim GrhIndex As Long
Dim i As Integer
    
    'Check that there are even map grhs
    If NumMapGrhs < 1 Then Exit Sub
    
    'Check for force over-ride
    If Not UpdateMapGrhPtr Then
        
        'Check if the position has changed enough
        If CurrMapIndex = LastScreenCalcMapIndex Then
            If Abs(ScreenX - LastScreenXCalc) < ScreenCalcSize Then
                If Abs(ScreenY - LastScreenYCalc) < ScreenCalcSize Then
                    Exit Sub
                End If
            End If
        End If
        
    Else
        UpdateMapGrhPtr = False
    End If
    
    'Update the last calculation position
    LastScreenXCalc = ScreenX
    LastScreenYCalc = ScreenY
    LastScreenCalcMapIndex = CurrMapIndex
    
    'Resize the array to fit all components if needed
    ReDim MapGrhPtr(1 To NumMapGrhs)
    NumMapGrhPtrs = 0
    
    'Loop through all the graphics
    For i = 1 To NumMapGrhs
        With MapGrh(i)
            GrhIndex = Graphics_GetGrhIndex(.Grh)
            If GrhIndex > 0 Then
                
                'Check if the graphic is in bounds
                If .X + GrhData(GrhIndex).Width - ScreenX > -ScreenCalcSize Then
                    If .Y + GrhData(GrhIndex).Height - ScreenY > -ScreenCalcSize Then
                        If .X - ScreenX < ScreenWidth + ScreenCalcSize Then
                            If .Y - ScreenY < ScreenHeight + ScreenCalcSize Then
                                
                                'Add to the array
                                NumMapGrhPtrs = NumMapGrhPtrs + 1
                                MapGrhPtr(NumMapGrhPtrs) = i
                                
                            End If
                        End If
                    End If
                End If
            
            End If
        End With
    Next i
    
    'Scale the array back down
    If NumMapGrhPtrs > 0 Then ReDim Preserve MapGrhPtr(1 To NumMapGrhPtrs) Else Erase MapGrhPtr
    
End Sub

Public Sub Engine_DrawTileInfo()
'*********************************************************************************
'Draws the tile information
'*********************************************************************************
Dim Lines(0 To 4) As TLVERTEX
Dim InfoGrh As tGrh
Dim TileX As Long
Dim TileY As Long
Dim X As Long
Dim Y As Long
Dim l As Long

    'Create the temp grh
    Graphics_SetGrh InfoGrh, 2
    
    'Loop through the screen
    For X = 0 To (ScreenWidth \ GRIDSIZE) + 1
        For Y = 0 To (ScreenHeight \ GRIDSIZE) + 1
        
            'Get the tile locations
            TileX = (ScreenX \ GRIDSIZE) + X
            TileY = (ScreenY \ GRIDSIZE) + Y
                    
            'Set the light value
            Select Case Engine_GetTileInfo(TileX, TileY)
                Case TILETYPE_PLATFORM: l = D3DColorARGB(255, 0, 255, 0)
                Case TILETYPE_BLOCKED: l = D3DColorARGB(255, 255, 0, 0)
                Case TILETYPE_LADDER: l = D3DColorARGB(255, 0, 0, 255)
                Case TILETYPE_SPAWN: l = D3DColorARGB(255, 255, 255, 0)
                Case TILETYPE_NOTHING: l = 0
            End Select
            
            'Draw the tile
            If l <> 0 Then
                Graphics_DrawGrh InfoGrh, (TileX * GRIDSIZE) - ScreenX, (TileY * GRIDSIZE) - ScreenY, l, l, l, l
            End If

        Next Y
    Next X
    
End Sub

Public Sub Engine_DrawBackground()
'*********************************************************************************
'Draw all of the background layers
'*********************************************************************************
Dim GrhIndex As Long
Dim PixelX As Long
Dim PixelY As Long
Dim X As Long
Dim Y As Long
Dim i As Long

    'Loop through all of the layers in reverse
    For i = NumBGLayers To 1 Step -1
        
        'Loop through each segment of the layer
        For X = 0 To BGSizeX
            For Y = 0 To BGSizeY
                
                'Check that the graphic exists
                If BG(i).Segment(X, Y).GrhIndex > 0 Then
                    
                    'Find the pixel X/Y co-ordinate
                    PixelX = X * BGGridSize - (ScreenX \ i)
                    PixelY = Y * BGGridSize - (ScreenY \ (i * 2))
                    
                    'Get the GrhIndex
                    GrhIndex = Graphics_GetGrhIndex(BG(i).Segment(X, Y))
                    
                    'In-bounds check
                    If PixelX + GrhData(GrhIndex).Width >= 0 Then
                        If PixelY + GrhData(GrhIndex).Height >= 0 Then
                            If PixelX <= ScreenWidth Then
                                If PixelY <= ScreenHeight Then
                    
                                    'Draw
                                    Graphics_DrawGrh BG(i).Segment(X, Y), PixelX, PixelY
                                    
                                End If
                            End If
                        End If
                    End If

                End If
                
            Next Y
        Next X
        
    Next i

End Sub

Public Sub Engine_DrawGrid()
'*********************************************************************************
'Draws the grid over the screen
'*********************************************************************************
Const GridLight As Long = -1761607681   'ARGB 150/255/255/255
Dim GridGrh As tGrh
Dim X As Long
Dim Y As Long

    'Create the temp grh
    Graphics_SetGrh GridGrh, 1

    'Loop through the screen
    For X = 0 To ScreenWidth + GRIDSIZE Step GRIDSIZE
        For Y = 0 To ScreenHeight + GRIDSIZE Step GRIDSIZE
            
            'Draw the grid
            Graphics_DrawGrh GridGrh, X + (GRIDSIZE - (ScreenX Mod GRIDSIZE)) - GRIDSIZE, _
                Y + (GRIDSIZE - (ScreenY Mod GRIDSIZE)) - GRIDSIZE, GridLight, GridLight, GridLight, GridLight
            
        Next Y
    Next X
    
End Sub

Public Sub Engine_ModifyXForLag(ByRef X As Integer, ByVal Heading As Byte, ByVal Speed As Single)
'*********************************************************************************
'Modifies the X position to take into consideration network lag
'*********************************************************************************

    Select Case Heading
        Case EAST
            'X = X + (Speed * Ping * 0.5)
        Case WEST
            'X = X - (Speed * Ping * 0.5)
    End Select

End Sub

Public Function Engine_CharCollision(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, _
    Optional ByVal NPCOnly As Boolean = False) As Integer
'*********************************************************************************
'Checks if the given rectangular area collides with any NPCs
'*********************************************************************************
Dim ClosestDistance As Single
Dim TempDistance As Single
Dim CenterX As Integer
Dim CenterY As Integer
Dim i As Integer

    'Starting closest distance
    ClosestDistance = -1

    'Find the center location
    CenterX = X + (Width \ 2)
    CenterY = Y + (Height \ 2)

    'Loop through all the characters
    For i = 1 To CharListUBound
        
        'Check for NPC-only collision
        If NPCOnly Then
            If Not CharList(i).IsNPC Then GoTo NextI
        End If
        
        With CharList(i)
            If .Used Then
            
                'Check for rectangular collision
                If Math_Collision_Rect(.X, .Y, .Width, .Height, X, Y, Width, Height) Then
                
                    'Collision made, find the distance
                    TempDistance = Abs(Math_Distance(CenterX, CenterY, .X + (.Width \ 2), .Y + (.Width \ 2)))
                    
                    'Check if the distance found is closer than what we had before
                    If ClosestDistance = -1 Or TempDistance < ClosestDistance Then
                        
                        'This is our new star player!
                        Engine_CharCollision = i
                        ClosestDistance = TempDistance
                    
                    End If
                
                End If
            
            End If
        End With
        
NextI:
    Next i

End Function

Public Sub Engine_DrawMapItems()
'*********************************************************************************
'Draws the items on the map
'*********************************************************************************
Dim i As Long

    'Draw the items
    For i = 0 To LastMapItem
        If MapItems(i).ItemIndex > 0 Then
            Graphics_DrawGrh MapItems(i).Grh, MapItems(i).X - ScreenX, MapItems(i).Y - ScreenY
        End If
    Next i

End Sub

Public Sub Engine_DrawItemToolTip(ByVal ItemIndex As Integer)
'*********************************************************************************
'Draws the tooltip for an item at the cursor
'*********************************************************************************
Const BackdropColor As Long = -1778384896
Const DrawXOffset As Long = 15
Const DrawYOffset As Long = 15
Dim v(0 To 3) As TLVERTEX
Dim BackdropHeight As Integer
Dim BackdropWidth As Integer
Dim LineText(35) As String
Dim LineColor(35) As Long
Dim Lines As Integer
Dim i As Long
Dim j As Long
Dim YOffset As Long
Dim G As tGrh

    'Check for a valid item
    If ItemIndex < 1 Then Exit Sub
    If ItemIndex > ItemsUBound Then Exit Sub
    
    'Set the initial lines value
    Lines = -1
    
    'Name
    If Items(ItemIndex).Name <> vbNullString Then
        Lines = Lines + 1
        LineText(Lines) = Items(ItemIndex).Name
        LineColor(Lines) = D3DColorARGB(255, 0, 255, 0)
    End If
    
    'Item type
    Lines = Lines + 1
    Select Case Items(ItemIndex).ItemType
    Case ITEMTYPE_USEONCE
        LineText(Lines) = "Consumeable"
    Case ITEMTYPE_CLOTHES
        LineText(Lines) = "Clothes"
    Case ITEMTYPE_WEAPON
        LineText(Lines) = "Weapon"
    Case ITEMTYPE_CAP
        LineText(Lines) = "Cap"
    Case ITEMTYPE_CLOTHES
        LineText(Lines) = "Clothes"
    Case ITEMTYPE_EARACC
        LineText(Lines) = "Ear Accessory"
    Case ITEMTYPE_EYEACC
        LineText(Lines) = "Eye Accessory"
    Case ITEMTYPE_FOREHEAD
        LineText(Lines) = "Forehead"
    Case ITEMTYPE_GLOVES
        LineText(Lines) = "Gloves"
    Case ITEMTYPE_MANTLE
        LineText(Lines) = "Mantle"
    Case ITEMTYPE_PANTS
        LineText(Lines) = "Pants"
    Case ITEMTYPE_PENDANT
        LineText(Lines) = "Pendant"
    Case ITEMTYPE_RING
        LineText(Lines) = "Ring"
    Case ITEMTYPE_SHIELD
        LineText(Lines) = "Shield"
    Case ITEMTYPE_SHOES
        LineText(Lines) = "Shoes"
    End Select
    LineColor(Lines) = D3DColorARGB(255, 255, 255, 0)
    
    'Description
    If LenB(Items(ItemIndex).Desc) > 0 Then
        Lines = Lines + 1
        LineText(Lines) = Items(ItemIndex).Desc
        LineColor(Lines) = D3DColorARGB(255, 255, 255, 255)
    End If
    
    'Damage
    If Items(ItemIndex).MinHit <> 0 Or Items(ItemIndex).MaxHit <> 0 Then
        Lines = Lines + 1
        LineText(Lines) = "Damage: " & Items(ItemIndex).MinHit & "/" & Items(ItemIndex).MaxHit
        LineColor(Lines) = D3DColorARGB(255, 255, 255, 255)
    End If
    
    'Strength
    If Items(ItemIndex).Str <> 0 Then
        Lines = Lines + 1
        LineText(Lines) = "Str: " & Items(ItemIndex).Str
        LineColor(Lines) = D3DColorARGB(255, 0, 255, 0)
    End If
    
    'Dexterity
    If Items(ItemIndex).Dex <> 0 Then
        Lines = Lines + 1
        LineText(Lines) = "Dex: " & Items(ItemIndex).Dex
        LineColor(Lines) = D3DColorARGB(255, 0, 255, 0)
    End If
    
    'Intelligence
    If Items(ItemIndex).Intl <> 0 Then
        Lines = Lines + 1
        LineText(Lines) = "Int: " & Items(ItemIndex).Intl
        LineColor(Lines) = D3DColorARGB(255, 0, 255, 0)
    End If
    
    'Luck
    If Items(ItemIndex).Luk <> 0 Then
        Lines = Lines + 1
        LineText(Lines) = "Luk: " & Items(ItemIndex).Luk
        LineColor(Lines) = D3DColorARGB(255, 0, 255, 0)
    End If
    
    'Find the width
    For i = 0 To Lines
        j = FontDefault.Width(LineText(i))
        If j > BackdropWidth Then
            BackdropWidth = j
        End If
    Next i
    BackdropWidth = BackdropWidth + 10
    
    'Find the height
    BackdropHeight = (Lines + 1) * FontDefault.Height + 10
    
    'Draw the backdrop
    Graphics_SetTexture 0
    Graphics_DrawRect CursorPos.X - 5 + DrawXOffset, CursorPos.Y - 5 + DrawYOffset, BackdropWidth, BackdropHeight, _
        BackdropColor, BackdropColor, BackdropColor, BackdropColor

    'Draw
    For i = 0 To Lines
        FontDefault.Draw LineText(i), CursorPos.X + DrawXOffset, CursorPos.Y + YOffset + DrawYOffset, LineColor(i)
        YOffset = YOffset + FontDefault.Height
    Next i

End Sub

Private Sub Engine_DrawChar(ByVal CharIndex As Integer)
'*********************************************************************************
'Update and draw a character
'*********************************************************************************
Dim OldOnGround As Byte
Dim CheckTileX As Integer
Dim CheckTileY As Integer
Dim NewCheckTile As Integer
Dim UpdateX As Boolean
Dim UpdateY As Boolean
Dim TileChangeX As Integer
Dim MoveDistanceX As Single
Dim MoveDistanceY As Single
Dim s As Single
Dim RemainderX As Long
Dim i As Long

    With CharList(CharIndex)
    
        'Check to play the death animation
        If Not .Used Then
            If .Action = eDeath Then
                If .Body.AnimType = ANIMTYPE_STATIONARY Then .Action = eNone
                GoTo SkipToDraw
            Else
                Exit Sub
            End If
        End If

        'Get the X co-ordinate remainder
        RemainderX = .X Mod GRIDSIZE

        'Store the old OnGround
        OldOnGround = .OnGround
        
        'Update the user's position
        MoveDistanceX = ElapsedTime * MOVESPEED
        MoveDistanceY = ElapsedTime * MOVESPEED
        Select Case .MoveDir
            Case EAST
                .X = .X + MoveDistanceX
            Case WEST
                .X = .X - MoveDistanceX
        End Select
        
        'User is jumping
        If .Jump > 0 Then
            .Y = .Y - (ElapsedTime * .Jump * MOVESPEED)
            MoveDistanceY = ElapsedTime * .Jump * MOVESPEED
            .Jump = .Jump - (ElapsedTime * JUMPDECAY)
            If .Jump < 0 Then .Jump = 0
            .OnGround = 0
        End If
        
        'Check if the tile X has changed
        If .LastTileX <> .X \ GRIDSIZE Then
            UpdateX = True
            TileChangeX = ((.X \ GRIDSIZE) - .LastTileX)
            .LastTileX = .X \ GRIDSIZE
        End If

        If UpdateX Then
    
            'Check for blocking to the right
            If TileChangeX > 0 Then
                CheckTileX = ((.X + .Width - 5) \ GRIDSIZE)
                If CheckTileX <= MapInfo.TileWidth Then
                    For i = 0 To .Height \ GRIDSIZE
                        CheckTileY = i + (.Y \ GRIDSIZE)
                        If Engine_GetTileInfo(CheckTileX, CheckTileY) = TILETYPE_BLOCKED Then
                            .X = (CheckTileX * GRIDSIZE) - .Width - 1
                            .LastTileX = .X \ GRIDSIZE
                            Exit For
                        End If
                    Next i
                End If
                
            'Check for blocking to the left
            Else
                CheckTileX = (.X + 5) \ GRIDSIZE
                If CheckTileX <= MapInfo.TileWidth Then
                    For i = 0 To .Height \ GRIDSIZE
                        CheckTileY = i + (.Y \ GRIDSIZE)
                        If Engine_GetTileInfo(CheckTileX, CheckTileY) = TILETYPE_BLOCKED Then
                            .X = ((CheckTileX + 1) * GRIDSIZE) + 1
                            .LastTileX = .X \ GRIDSIZE
                            Exit For
                        End If
                    Next i
                End If
            End If

            'Check if the user will be dropping
            CheckTileY = ((.Y + .Height) \ GRIDSIZE) + 1    'Get the tile below the user
            If CheckTileY <= MapInfo.TileHeight Then
                For i = 0 To (.Width + RemainderX) \ GRIDSIZE
                    CheckTileX = i + (.X \ GRIDSIZE)
                    If Engine_GetTileInfo(CheckTileX, CheckTileY) = TILETYPE_BLOCKED Or _
                        Engine_GetTileInfo(CheckTileX, CheckTileY) = TILETYPE_PLATFORM Then
                        i = 0
                        Exit For
                    End If
                Next i
                If i > 0 Then
                    .OnGround = 0
                End If
            End If
            
        End If
        
        'Check for dropping
        If .Jump < 1 And .OnGround = 0 Then
            .Y = .Y + (ElapsedTime * MOVESPEED)
        End If
        
        'Check if the tile Y has changed
        If .LastTileY <> .Y \ GRIDSIZE Then
            UpdateY = True
            .LastTileY = .Y \ GRIDSIZE
        End If

        If .Jump < 1 Then
        
            'Dropping handling
            If .Jump = 0 Then
                NewCheckTile = ((.Y + .Height + 1) \ GRIDSIZE)
                If NewCheckTile <> CheckTileY Then
                    For i = 0 To (.Width + RemainderX) \ GRIDSIZE
                        CheckTileX = i + (.X \ GRIDSIZE)
                        Select Case Engine_GetTileInfo(CheckTileX, NewCheckTile)
                            Case TILETYPE_BLOCKED, TILETYPE_PLATFORM
                                .Y = NewCheckTile * GRIDSIZE - .Height - 1
                                .OnGround = 1
                                .Jump = 0
                                Exit For
                        End Select
                    Next i
                End If
            End If
            
        End If
        
        If .Jump >= 1 Then

            'Head-hitting handling
            If UpdateY Then
                NewCheckTile = (.Y \ GRIDSIZE)
                For i = 0 To (.Width + RemainderX) \ GRIDSIZE
                    CheckTileX = i + (.X \ GRIDSIZE)
                    If Engine_GetTileInfo(CheckTileX, NewCheckTile) = TILETYPE_BLOCKED Then
                        .Y = (NewCheckTile * GRIDSIZE) + GRIDSIZE - 1
                        .OnGround = 0
                        .Jump = 0
                        Exit For
                    End If
                Next i
            End If

        End If
        
        '*** Animations ***
        
        'Check if to end an active action
        If .Action <> eNone Then
            If .Body.AnimType = ANIMTYPE_STATIONARY Then
                .Action = eNone
            End If
        End If
        
        'If theres an action going, only use that animation
        If .Action = eNone Then
            
            'Jumping / falling
            If .OnGround = 0 Then
                
                'Going up
                If .Jump > 0 Then
                    If .IsNPC Then
                        If .Body.GrhIndex <> PDSprite(.BodyIndex).JumpUp Then Graphics_SetGrh .Body, PDSprite(.BodyIndex).JumpUp, ANIMTYPE_LOOP
                    Else
                        If .Body.GrhIndex <> PDBody(.BodyIndex).JumpUp Then Graphics_SetGrh .Body, PDBody(.BodyIndex).JumpUp, ANIMTYPE_LOOP
                    End If
                    
                'Going down
                Else
                    If .IsNPC Then
                        If .Body.GrhIndex <> PDSprite(.BodyIndex).JumpDown Then Graphics_SetGrh .Body, PDSprite(.BodyIndex).JumpDown, ANIMTYPE_LOOP
                    Else
                        If .Body.GrhIndex <> PDBody(.BodyIndex).JumpDown Then Graphics_SetGrh .Body, PDBody(.BodyIndex).JumpDown, ANIMTYPE_LOOP
                    End If
                End If
            
            Else
                
                'Walking / idling
                If .LastX <> .X Then
                
                    'Is moving
                    .IdleFrames = 0
                    If .IsNPC Then
                        If .Body.GrhIndex <> PDSprite(.BodyIndex).Walk Then Graphics_SetGrh .Body, PDSprite(.BodyIndex).Walk, ANIMTYPE_LOOP
                    Else
                        If .Body.GrhIndex <> PDBody(.BodyIndex).Walk Then Graphics_SetGrh .Body, PDBody(.BodyIndex).Walk, ANIMTYPE_LOOP
                    End If
                    
                Else
                    
                    'Just landed from a jump
                    If OldOnGround = 0 Then GoTo ForceStand
    
                    'Is not moving
                    .IdleFrames = .IdleFrames + 1
                    If .IdleFrames > 3 Then
ForceStand:
                        .IdleFrames = 0
                        If .IsNPC Then
                            If .Body.GrhIndex <> PDSprite(.BodyIndex).Stand Then Graphics_SetGrh .Body, PDSprite(.BodyIndex).Stand, ANIMTYPE_LOOP
                        Else
                            If .Body.GrhIndex <> PDBody(.BodyIndex).Stand Then Graphics_SetGrh .Body, PDBody(.BodyIndex).Stand, ANIMTYPE_LOOP
                        End If

                    End If
                    
                End If
                
            End If
            
        End If
        
        'Set the last X/Y values
        .LastX = .X
        .LastY = .Y
        
        'In-bounds check
        If .X < 0 Then .X = 0
        If .X > MapInfo.TileWidth * GRIDSIZE Then .X = MapInfo.TileWidth * GRIDSIZE
        If .Y + .Height > MapInfo.TileHeight * GRIDSIZE Then
            .Y = MapInfo.TileHeight * GRIDSIZE - .Height
            .Jump = 0
            .OnGround = 1
        End If

        'Update the draw location
        If .DrawX <> .X Then
            s = (1 + (Abs(Abs(.DrawX - .X) - MoveDistanceX) / 96))
            If s < 1 Then s = 1
            MoveDistanceX = MoveDistanceX * s
            If .DrawX > .X Then
                .DrawX = .DrawX - MoveDistanceX
            Else
                .DrawX = .DrawX + MoveDistanceX
            End If
            If Abs(.DrawX - .X) < MoveDistanceX Then .DrawX = .X
        End If
        If .DrawY <> .Y Then
            s = (1 + (Abs(Abs(.DrawY - .Y) - MoveDistanceY) / 96))
            If s < 1 Then s = 1
            MoveDistanceY = MoveDistanceY * s
            If .DrawY > .Y Then
                .DrawY = .DrawY - MoveDistanceY
            Else
                .DrawY = .DrawY + MoveDistanceY
            End If
            If Abs(.DrawY - .Y) < MoveDistanceY Then .DrawY = .Y
        End If
        
        'Check to update the screen position
        If CharIndex = UserCharIndex Then
            ScreenX = .DrawX - (ScreenWidth \ 2)
            ScreenY = .DrawY - (ScreenHeight \ 2)
            If ScreenX < 0 Then ScreenX = 0
            If ScreenY < 0 Then ScreenY = 0
            If ScreenX + ScreenWidth > MapInfo.TileWidth * GRIDSIZE Then ScreenX = MapInfo.TileWidth * GRIDSIZE - ScreenWidth
            If ScreenY + ScreenHeight > MapInfo.TileHeight * GRIDSIZE Then ScreenY = MapInfo.TileHeight * GRIDSIZE - ScreenHeight
        End If
        
SkipToDraw:

        'Draw the character
        Graphics_DrawGrh .Body, .DrawX - ScreenX, .DrawY - ScreenY, , , , , (.Heading = WEST) * -.Width
        
        'Write the character's name
        If .Name <> vbNullString Then
            If Not .IsNPC Then
                i = FontDefault.Width(.Name)
                FontDefault.Draw .Name, .DrawX + (.Width \ 2) - (i \ 2) - ScreenX, .DrawY + .Height - ScreenY, -1
            End If
        End If
        
    End With

End Sub

Public Sub Engine_DrawChars()
'*********************************************************************************
'Updates and draws all of the characters (wrapper for Engine_DrawChar)
'*********************************************************************************
Dim i As Long

    'Loop through all of the characters
    For i = 1 To CharListUBound
        If CharList(i).Used = True Or CharList(i).Action = eDeath Then
            Engine_DrawChar i
        End If
    Next i

End Sub

Public Sub Engine_DrawMap(ByVal Behind As Byte)
'*********************************************************************************
'Draws all of the map graphics (Behind parameter must = 0 for in front, 1 for behind)
'*********************************************************************************
Dim GrhIndex As Long
Dim i As Long

    'Update the MapGrhPtr
    Engine_MakeMapGrhPtr

    'Loop through all of the graphics on the map
    For i = 1 To NumMapGrhPtrs
        With MapGrh(MapGrhPtr(i))
        
            'Check if we have the correct "behind" value
            If .Behind = Behind Then
                
                'Get the GrhIndex from the Grh
                GrhIndex = Graphics_GetGrhIndex(.Grh)
                If GrhIndex > 0 Then
                    
                    'Check if the graphic will be in the view of the map
                    If .X - ScreenX < ScreenWidth Then
                        If .Y - ScreenY < ScreenHeight Then
                            If .X + GrhData(GrhIndex).Width - ScreenX > 0 Then
                                If .Y + GrhData(GrhIndex).Height - ScreenY > 0 Then
                                
                                    'Draw the graphic
                                    Graphics_DrawGrh .Grh, .X - ScreenX, .Y - ScreenY
                                    
                                End If
                            End If
                        End If
                    End If
                
                End If
            End If
        
        End With
    Next i
    
    'Render the particle effects
    If Behind Then
        Effect_Render EFFECTRENDER_MAPBACK
    Else
        Effect_Render EFFECTRENDER_MAPFRONT
    End If

End Sub

Public Sub Engine_UpdateFPS(Optional ByVal FrameLimit As Long = 66)
'*********************************************************************************
'Update the FPS and the elapsed time
'*********************************************************************************

    'Raise the FPS count
    FPSCounter = FPSCounter + 1
    If FPSUpdateTime < timeGetTime Then
        FPSUpdateTime = timeGetTime + 1000
        FPS = FPSCounter
        FPSCounter = 0
    End If
    
    'Raise the elapsed time
    ElapsedTime = timeGetTime - LastFrame
    LastFrame = timeGetTime
    If ElapsedTime > 33 Then ElapsedTime = 33   'Don't run slower than 30 FPS
    
    'Check if to sleep
    If FrameLimit > 0 Then
        If ElapsedTime < (2000 / FrameLimit) Then
            Sleep (2000 / FrameLimit) - ElapsedTime
        End If
    End If

End Sub

Public Sub Engine_Init(Optional ByVal UseGraphics As Boolean = True)
'*********************************************************************************
'Load up all aspects of the engine
'*********************************************************************************

    Randomize

    'Set the engine start time
    EngineStartTime = timeGetTime

    'Set the timer frequency to as high as possible
    timeBeginPeriod 1
    
    'Set the initial values for the map items
    LastMapItem = -1
    MapItemsUBound = -1
    
    'Get the settings
    DebugMode = (Val(IO_INI_Read(App.Path & "\Data\Settings.ini", "GENERAL", "DebugMode")) <> 0)

    'Create the root DirectX object
    Set DX = New DirectX8

    'Engine is running
    EngineRunning = True
    
    'Set the update time
    FPSUpdateTime = timeGetTime + 1000

End Sub

Public Sub Engine_DrawDebugInfo()
'*********************************************************************************
'Draws the information for the Debug Mode
'*********************************************************************************
Dim Lines(0 To 4) As TLVERTEX
Dim v As Long
Dim t As Long
Dim s As String
Dim i As Long

    'Check if we're in debug mode
    If Not DebugMode Then Exit Sub
    
    'Tile information
    Engine_DrawTileInfo

    'Character location and collision area
    If False Then
        For i = 0 To 3
            Lines(i).Rhw = 1
            Lines(i).Color = D3DColorARGB(255, 255, 255, 255)
        Next i
        For i = 1 To CharListUBound
            If CharList(i).Used Then
                Lines(0).X = CharList(i).X - ScreenX
                Lines(0).Y = CharList(i).Y - ScreenY
                Lines(1).X = Lines(0).X + CharList(i).Width
                Lines(1).Y = Lines(0).Y
                Lines(2).X = Lines(1).X
                Lines(2).Y = Lines(1).Y + CharList(i).Height
                Lines(3).X = Lines(0).X
                Lines(3).Y = Lines(2).Y
                Lines(4) = Lines(0)
                Graphics_DrawLineStrip Lines(), 4
                s = i
                s = "Pixel:(" & Int(CharList(i).X) & "," & Int(CharList(i).Y) & ")" & _
                    " Tile:(" & CharList(i).X \ GRIDSIZE & "," & CharList(i).Y \ GRIDSIZE & ")"
                FontDefault.Draw s, CharList(i).X - ScreenX - (FontDefault.Width(s) \ 2) + (CharList(i).Width \ 2), _
                    CharList(i).Y - ScreenY - FontDefault.Height, D3DColorARGB(255, 255, 255, 255)
            End If
        Next i
    End If

    'Top-right stat display
    With FontDefault
        
        'FPS and ping
        .Draw "FPS: " & FPS, ScreenWidth - .Width("FPS: " & FPS) - 5, 25, D3DColorARGB(255, 255, 255, 255)
        .Draw "Ping: " & Int(Ping), ScreenWidth - .Width("Ping: " & Int(Ping)) - 5, 25 + .Height, D3DColorARGB(255, 255, 255, 255)
        
        'Store the time, and make sure we don't have 0 so we don't divide by 0
        t = ((timeGetTime - EngineStartTime) \ 1000)
        If t < 1 Then t = 1
        
        'Bandwidth information
        v = BytesIn \ t
        .Draw "Bytes in/sec: " & v, ScreenWidth - .Width("Bytes In/sec: " & v) - 5, 25 + (.Height * 2), D3DColorARGB(255, 255, 255, 255)

        v = BytesOut \ t
        .Draw "Bytes out/sec: " & v, ScreenWidth - .Width("Bytes Out/sec: " & v) - 5, 25 + (.Height * 3), D3DColorARGB(255, 255, 255, 255)

    End With
    
End Sub

Public Sub Engine_Destroy()
'*********************************************************************************
'Unload all aspects of the engine
'*********************************************************************************
    
    Log "Destroy engine"

    'Destroy the DirectX device
    If Not DX Is Nothing Then Set DX = Nothing
    
    'Close the log file
    Log_Close
    
    'Engine closed
    EngineRunning = False

End Sub

Public Function Engine_GetTileInfo(ByVal TileX As Integer, ByVal TileY As Integer)
'*********************************************************************************
'Get the information from a tile - recommended you use this instead of direct access
'since this has error checking
'*********************************************************************************

    'Check for out-of-bounds
    If TileX < 0 Then
        Engine_GetTileInfo = TILETYPE_BLOCKED
        Exit Function
    End If
    If TileY < 0 Then
        Engine_GetTileInfo = TILETYPE_BLOCKED
        Exit Function
    End If
    If TileX > MapInfo.TileWidth Then
        Engine_GetTileInfo = TILETYPE_BLOCKED
        Exit Function
    End If
    If TileY > MapInfo.TileHeight Then
        Engine_GetTileInfo = TILETYPE_BLOCKED
        Exit Function
    End If
    
    'Return the tile info
    Engine_GetTileInfo = MapInfo.TileInfo(TileX, TileY)

End Function
