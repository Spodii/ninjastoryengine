Attribute VB_Name = "General"
Option Explicit

'Holds all of the map information
Public MapInfo As tMapInfo

'Whether or not to draw the grid
Public DrawGrid As Boolean

'Whether or not to draw the tile info
Public DrawTileInfo As Boolean

'Character array
Public CharList() As tClientChar
Public CharListUBound As Integer    'Size of the CharList() array
Public UserCharIndex As Integer     'Not used by map editor

Public Sub GUI_Init()
'Just ignore this
End Sub

Public Function HasFloatingBlocks() As Boolean
'*********************************************************************************
'Checks if the map has any floating blocks (blocks that do not have another block
'or the edge of the map to the left, right or bottom of the map)
'*********************************************************************************
Dim X As Long
Dim Y As Long

    'Loop through every tile
    For X = 0 To MapInfo.TileWidth
        For Y = 0 To MapInfo.TileHeight
        
            'Check for a blocked tile
            If MapInfo.TileInfo(X, Y) = TILETYPE_BLOCKED Then
                
                'Check to the left
                If X - 1 >= 0 Then
                    If MapInfo.TileInfo(X - 1, Y) <> TILETYPE_BLOCKED Then
                        HasFloatingBlocks = True
                        frmScreen.Caption = "Screen (Warning: Floating block found at [" & X - 1 & "," & Y & "])"
                        Exit Function
                    End If
                End If
                
                'Check to the right
                If X + 1 <= MapInfo.TileWidth Then
                    If MapInfo.TileInfo(X + 1, Y) <> TILETYPE_BLOCKED Then
                        HasFloatingBlocks = True
                        frmScreen.Caption = "Screen (Warning: Floating block found at [" & X + 1 & "," & Y & "])"
                        Exit Function
                    End If
                End If
                
                'Check to the bottom
                If Y + 1 <= MapInfo.TileHeight Then
                    If MapInfo.TileInfo(X, Y + 1) <> TILETYPE_BLOCKED Then
                        HasFloatingBlocks = True
                        frmScreen.Caption = "Screen (Warning: Floating block found at [" & X & "," & Y + 1 & "])"
                        Exit Function
                    End If
                End If
                
            End If
            
        Next Y
    Next X

    frmScreen.Caption = "Screen"

End Function

Public Sub CalcBGSizes()
'*********************************************************************************
'Calculates the background array sizes
'*********************************************************************************
Dim i As Long
    
    BGSizeX = (MapInfo.TileWidth * GRIDSIZE) \ BGGridSize
    BGSizeY = (MapInfo.TileHeight * GRIDSIZE) \ BGGridSize
    For i = 1 To NumBGLayers
        ReDim BG(i).Segment(0 To BGSizeX, 0 To BGSizeY)
    Next i
        
End Sub

Public Function IsValidValue(ByVal KeyAscii As Integer) As Boolean
'*********************************************************************************
'Returns if a value is valid for a text box that accepts only numbers
'*********************************************************************************

    If Not IsNumeric(Chr$(KeyAscii)) Then
        If KeyAscii <> 8 Then Exit Function
    End If
    IsValidValue = True

End Function

Public Function GetValueFromCD(Optional ByVal InitDir As String = vbNullString, Optional ByVal Filter As String = vbNullString) As String
'*********************************************************************************
'Returns the value from a Common Dialog box
'*********************************************************************************

    'Check the InitDir
    If InitDir = vbNullString Then InitDir = App.Path

    'Set up the CD
    With frmMain.CD
        .Filter = Filter
        .DialogTitle = "Load"
        .FileName = vbNullString
        .InitDir = InitDir
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
        .ShowOpen
    End With
    
    'Return the value
    GetValueFromCD = frmMain.CD.FileName

End Function

Private Sub MakeNPCSpawnArray(ByRef pNPCSpawn() As tMapSpawn, ByRef pNPCSpawnUBound As Integer)
'*********************************************************************************
'Creates the array of NPC spawns
'*********************************************************************************
Dim i As Long
Dim s() As String
Dim tID As Integer
Dim tAmount As Integer

    'Set the initial size of the array
    pNPCSpawnUBound = -1

    'Loop through the NPC list
    For i = 0 To frmSpawn.SpawnLst.ListCount - 1
        If frmSpawn.SpawnLst.List(i) <> vbNullString Then
        
            'Get the values
            s() = Split(frmSpawn.SpawnLst.List(i), " ")
            tID = Val(s(0))
            tAmount = Val(Mid$(s(1), 2, Len(s(1)) - 2))
            
            'Check for valid values
            If tID > 0 Then
                If tAmount > 0 Then
                
                    'Increase the size of the array
                    pNPCSpawnUBound = pNPCSpawnUBound + 1
                    ReDim Preserve pNPCSpawn(0 To pNPCSpawnUBound)
                    
                    'Set the values
                    pNPCSpawn(pNPCSpawnUBound).NPCID = tID
                    pNPCSpawn(pNPCSpawnUBound).Amount = tAmount
                    
                End If
            End If
            
        End If
    Next i

End Sub

Private Sub MakeSpawnTileArray(ByRef pSpawnTile() As tTilePos, ByRef pSpawnTileUBound As Integer)
'*********************************************************************************
'Creates the spawn tile array
'*********************************************************************************
Dim tempSize As Integer
Dim X As Long
Dim Y As Long

    'Set the initial values
    pSpawnTileUBound = -1
    tempSize = -1

    'Loop through the tiles
    For X = 0 To MapInfo.TileWidth
        For Y = 0 To MapInfo.TileHeight
        
            'Check the tile attribute
            If MapInfo.TileInfo(X, Y) = TILETYPE_SPAWN Then
                
                'Increase the spawn tile count
                pSpawnTileUBound = pSpawnTileUBound + 1
            
                'Check if we need to resize the array
                If tempSize < pSpawnTileUBound Then
                    tempSize = tempSize + 50
                    ReDim Preserve pSpawnTile(0 To tempSize)
                End If
                
                'Write the information into the array
                pSpawnTile(pSpawnTileUBound).TileX = X
                pSpawnTile(pSpawnTileUBound).TileY = Y
                
            End If
        
        Next Y
    Next X
    
    'Check if we have any tiles
    If pSpawnTileUBound > -1 Then
        
        'Scale the array down to pSpawnTileUBound
        ReDim Preserve pSpawnTile(0 To pSpawnTileUBound)
                
    End If
                
End Sub

Public Sub ToolBar_Click(ByVal Button As Long)
'*********************************************************************************
'Handles clicking of toolbar buttons
'*********************************************************************************
Dim i As Long
Dim s As String
Dim sp() As String
Dim SpawnTile() As tTilePos
Dim SpawnTileUBound As Integer
Dim NPCSpawn() As tMapSpawn
Dim NPCSpawnUBound As Integer

    Select Case Button
    
        'New
        Case 1
            If MsgBox("Are you sure you wish to create a new map?" & vbNewLine & _
                "All changes to the current map will be lost.", vbYesNo) = vbYes Then
                NewMap
            End If
        
        'Load
        Case 2
            If MsgBox("Are you sure you wish to load a map?" & vbNewLine & _
                "All changes to the current map will be lost.", vbYesNo) = vbYes Then
                s = GetValueFromCD(App.Path & "\Maps\", ".nsmi")
                If s = vbNullString Then Exit Sub
                sp = Split(s, "\")
                i = Val(Left$(sp(UBound(sp)), Len(sp(UBound(sp))) - 5))
                If i <= 0 Then Exit Sub
                IO_Map_LoadInfo i, MapInfo
                IO_Map_LoadGrhs i, MapGrh, NumMapGrhs, BG(), BGSizeX, BGSizeY
                IO_Map_LoadServerInfo i, SpawnTile(), SpawnTileUBound, NPCSpawn(), NPCSpawnUBound
                frmSpawn.SpawnLst.Clear
                For i = 0 To NPCSpawnUBound
                    frmSpawn.SpawnLst.AddItem NPCSpawn(i).NPCID & " [" & NPCSpawn(i).Amount & "]"
                Next i
                HasFloatingBlocks
                CurrMapIndex = i
                frmMapSettings.MusicTxt.Text = MapInfo.MusicID
                frmMain.MapNameLbl.Caption = MapInfo.Name
                ScreenX = 0
                ScreenY = 0
            End If
        
        'Save
        Case 3
            If CurrMapIndex <= 0 Then
                CurrMapIndex = Val(InputBox("Please enter the index for this new map."))
                If CurrMapIndex <= 0 Then Exit Sub
            End If
            OptimizeMapGrhs
            MapInfo.HasFloatingBlocks = HasFloatingBlocks
            IO_Map_SaveInfo CurrMapIndex, MapInfo
            IO_Map_SaveGrhs CurrMapIndex, MapGrh(), NumMapGrhs, BG(), BGSizeX, BGSizeY
            MakeSpawnTileArray SpawnTile(), SpawnTileUBound
            MakeNPCSpawnArray NPCSpawn(), NPCSpawnUBound
            IO_Map_SaveServerInfo CurrMapIndex, SpawnTile(), SpawnTileUBound, NPCSpawn(), NPCSpawnUBound
            UpdateHighestMapIndex
            MsgBox "Map successfully saved.", vbOKOnly
            
        'Save as
        Case 4
            CurrMapIndex = 0
            ToolBar_Click 3
        
        Case 5: 'Sep
        
        'Set graphics
        Case 6
            frmGraphics.Visible = (Not frmGraphics.Visible)
            
        Case 7: 'Blocks
        Case 8: 'Floods
        
        'Tile info
        Case 9
            frmTileInfo.Visible = (Not frmTileInfo.Visible)
        
        Case 10: 'Exits
        
        'NPCs
        Case 11
            frmSpawn.Visible = (Not frmSpawn.Visible)
        
        Case 12: 'Particles
        Case 13: 'Sfx
        
        'Map info
        Case 14
            frmMapSettings.Visible = (Not frmMapSettings.Visible)
        
        Case 15: 'Sep
        Case 16: 'Toggle weather display
        Case 17: 'Toggle chars display
        
        'Toggle the grid display
        Case 18
            DrawGrid = (Not DrawGrid)
            
        'Toggle tile info display
        Case 19
            DrawTileInfo = (Not DrawTileInfo)
        
        Case 20: 'Toggle mini-map display
    End Select

End Sub

Private Sub UpdateHighestMapIndex()
'*********************************************************************************
'Finds the highest map index and sets it in maps.ini
'*********************************************************************************
Dim HighIndex As Integer
Dim Files() As String
Dim s() As String
Dim sr As String
Dim i As Long

    Files() = AllFilesInFolders(App.Path & "\Maps\", False)
    For i = 0 To UBound(Files)
        s = Split(Files(i), "\")
        sr = s(UBound(s))
        If Len(sr) > 5 Then
            sr = Left$(sr, Len(sr) - 5)
            If Val(sr) > HighIndex Then HighIndex = Val(sr)
        End If
    Next i
    IO_INI_Write App.Path & "\Maps\maps.ini", "GENERAL", "NumMaps", HighIndex

End Sub

Sub Main()
'*********************************************************************************
'Entry point for the map editor
'*********************************************************************************

    'Load the main form
    Load frmMain
    frmMain.Show
    DoEvents
    
    'Display settings
    DrawGrid = False
    DrawTileInfo = True
    
    'Load the other forms
    Load frmScreen
    Load frmGraphics
    Load frmTileInfo
    Load frmMapSettings
    Load frmBG
    Load frmSpawn
    
    'Load the engine
    Engine_Init
    IO_GrhData_Load GrhData(), NumGrhs
    Graphics_Init frmScreen.hWnd, False

    'Start the game loop
    GameLoop

End Sub

Sub NewMap()
'*********************************************************************************
'Create a new map
'*********************************************************************************

    'Clear the map index
    CurrMapIndex = 0
    
    'Make a new map
    NumMapGrhs = 0
    NumMapGrhPtrs = 0
    Erase MapGrh
    MapInfo.TileWidth = 400
    MapInfo.TileHeight = 70
    MapInfo.Name = "New map"
    ReDim MapInfo.TileInfo(0 To MapInfo.TileWidth, 0 To MapInfo.TileHeight)
    CalcBGSizes
    
    'Clear the GUI information
    frmMapSettings.NameTxt.Text = MapInfo.Name
    frmMapSettings.WidthTxt.Text = MapInfo.TileWidth
    frmMapSettings.HeightTxt.Text = MapInfo.TileHeight
    
    'Reset the screen location
    ScreenX = 0
    ScreenY = 0

End Sub

Sub OptimizeMapGrhs()
'*********************************************************************************
'Runs through the map grh list and deletes any empty entries, making sure every index is used
'*********************************************************************************
Dim OldNumMapGrhs As Long
Dim i As Long
Dim j As Long

    'Store the old NumMapGrhs value
    OldNumMapGrhs = NumMapGrhs

    'Loop through all the MapGrhs and look for an unused index
    For i = 1 To NumMapGrhs
        If i <= NumMapGrhs Then
            If MapGrh(i).Grh.GrhIndex = 0 Then
                
                'The index is not used, so move every entry down one to remove it
                For j = i To NumMapGrhs - 1
                    MapGrh(j) = MapGrh(j + 1)
                Next j
                
                'Lower NumMapGrhs, since our array will be 1 smaller
                NumMapGrhs = NumMapGrhs - 1
                
            End If
        End If
    Next i
    
    'Resize the array with the new size
    If OldNumMapGrhs <> NumMapGrhs Then ReDim Preserve MapGrh(1 To NumMapGrhs)
    
End Sub

Sub SetInfo(ByVal s As String, Optional ByVal Critical As Boolean = False)
'*********************************************************************************
'Displays information in the info bar on the bottom of the screen
'*********************************************************************************

    If Critical Then
        frmMain.CritTimer.Enabled = False
        frmMain.CritTimer.Enabled = True
        frmMain.InfoLbl.Caption = s
    Else
        If Not frmMain.CritTimer.Enabled Then frmMain.InfoLbl.Caption = s
    End If

End Sub

Sub HandleInput()
'*********************************************************************************
'Handle input from GetAsyncKeyState()
'*********************************************************************************
Const SpeedMult As Single = 0.5 'Screen movement speed multiplier

    'Screen movement
    If GetKey(vbKeyUp) Then
        ScreenY = ScreenY - ElapsedTime * SpeedMult
        If ScreenY < 0 Then ScreenY = 0
    End If
    If GetKey(vbKeyRight) Then
        ScreenX = ScreenX + ElapsedTime * SpeedMult
        If ScreenX + ScreenWidth > MapInfo.TileWidth * GRIDSIZE Then ScreenX = MapInfo.TileWidth * GRIDSIZE - ScreenWidth
    End If
    If GetKey(vbKeyDown) Then
        ScreenY = ScreenY + ElapsedTime * SpeedMult
        If ScreenY + ScreenHeight > MapInfo.TileHeight * GRIDSIZE Then ScreenY = MapInfo.TileHeight * GRIDSIZE - ScreenHeight
    End If
    If GetKey(vbKeyLeft) Then
        ScreenX = ScreenX - ElapsedTime * SpeedMult
        If ScreenX < 0 Then ScreenX = 0
    End If
    
End Sub

Sub GameLoop()
'*********************************************************************************
'Update all the rendering and such
'*********************************************************************************
Dim UpdateFPSTime As Long
Dim f As Form

    EngineRunning = True
    Do While EngineRunning
    
        'Handle input
        HandleInput
    
        Graphics_BeginScene
            
            'Draw the background
            Engine_DrawBackground
            
            'Draw the map
            Engine_DrawMap 1
            Engine_DrawMap 0
            
            'Draw the grid
            If DrawGrid Then Engine_DrawGrid
            
            'Draw the tile information
            If DrawTileInfo Then Engine_DrawTileInfo
            
        Graphics_EndScene
        
        'Update the frame rate
        Engine_UpdateFPS
        
        'Check to update the FPS display
        If UpdateFPSTime < timeGetTime Then
            UpdateFPSTime = timeGetTime + 1000
            frmMain.FPSLbl.Caption = "FPS: " & FPS
        End If
        
        'Let windows do its events
        DoEvents
        
    Loop

    'Unload the engine
    Engine_Destroy
    
    'Unload the forms
    For Each f In VB.Forms
        Set f = Nothing
    Next f
    
    'Finish off
    End

End Sub
