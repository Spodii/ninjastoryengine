VERSION 5.00
Begin VB.Form frmScreen 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Screen"
   ClientHeight    =   11520
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    'Restore old settings
    Me.Left = Val(IO_INI_Read(App.Path & "\Dev\Map Editor Settings.ini", Me.Name, "Left"))
    Me.Top = Val(IO_INI_Read(App.Path & "\Dev\Map Editor Settings.ini", Me.Name, "Top"))

End Sub

Private Function NextFreeMapGrh() As Long
Dim i As Long

    'Check if theres a free slot already available
    For i = 1 To NumMapGrhs
        If MapGrh(i).Grh.GrhIndex = 0 Then
            NextFreeMapGrh = i
            Exit Function
        End If
    Next i
    
    'Resize the array to fit the new map grh and return the new high index
    NumMapGrhs = NumMapGrhs + 1
    ReDim Preserve MapGrh(1 To NumMapGrhs)
    NextFreeMapGrh = NumMapGrhs

End Function

Private Function GraphicAtPixel(ByVal X As Long, ByVal Y As Long) As Long
Dim i As Long

    'Check if a graphic already exists at that pixel location
    For i = 1 To NumMapGrhs
        If MapGrh(i).X = X Then
            If MapGrh(i).Y = Y Then
                GraphicAtPixel = i
                Exit Function
            End If
        End If
    Next i

End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim BGIndex As Long         'Selected background index from frmBG (if placing backgrounds)
Dim MapGrhIndex As Long     'Index in the MapGrh() array to use (if placing map grhs)
Dim GrhIndex As Long        'GrhIndex from frmGraphics or from frmBG
Dim TileX As Long           'Tile X co-ordinate of the pixel
Dim TileY As Long           'Tile Y co-ordinate of the pixel
Dim DestX As Long           'Graphic destination tile X (if placing map grhs)
Dim DestY As Long           'Graphic destination tile Y (if placing map grhs)
Dim GraphicAtTile As Long   'GrhIndex already at the MapGrh(DestX, DestY) (if placing map grhs)
Dim BGX As Long             'Background X co-ordinate (if placing backgrounds)
Dim BGY As Long             'Background Y co-ordinate (if placing backgrounds)
Dim i As Long

    'Get the tile location
    TileX = (X + ScreenX) \ GRIDSIZE
    TileY = (Y + ScreenY) \ GRIDSIZE
    If TileX > MapInfo.TileWidth Then Exit Sub
    If TileY > MapInfo.TileHeight Then Exit Sub
    If TileX < 0 Then Exit Sub
    If TileY < 0 Then Exit Sub
    
    'Update the tile and pixel location display
    frmMain.TileLbl.Caption = "(" & TileX & "," & TileY & ")"
    frmMain.PixelLbl.Caption = "(" & Int(X + ScreenX) & "," & Int(Y + ScreenY) & ")"
    
    'Check for a left-click
    If Button = vbLeftButton Then
    
        If frmGraphics.Visible Then
        
            '*** Set graphics ***
            If frmGraphics.SetOpt.Value Then
            
                'Find the destination location
                If frmGraphics.SnapChk.Value = 0 Then
                    DestX = X + ScreenX
                    DestY = Y + ScreenY
                Else
                    DestX = TileX * GRIDSIZE
                    DestY = TileY * GRIDSIZE
                End If
            
                GrhIndex = Val(frmGraphics.GrhIndexTxt.Text)
                If GrhIndex > 0 Then
                    If GrhIndex <= NumGrhs Then
                        If GrhData(GrhIndex).NumFrames > 0 Or GrhData(GrhIndex).TextureID > 0 Then
                        
                            'Check if the graphic already exists at that pixel - if so, delete it
                            GraphicAtTile = GraphicAtPixel(DestX, DestY)
                            If GraphicAtTile > 0 Then MapGrh(GraphicAtTile).Grh.GrhIndex = 0
                            
                            'Resize the array to fit the new map graphic
                            MapGrhIndex = NextFreeMapGrh
                            
                            'Check whether to animate the tile or not
                            If GrhData(GrhIndex).NumFrames > 1 Then
                                Graphics_SetGrh MapGrh(MapGrhIndex).Grh, GrhIndex, ANIMTYPE_LOOP
                            Else
                                Graphics_SetGrh MapGrh(MapGrhIndex).Grh, GrhIndex, ANIMTYPE_STATIONARY
                            End If
                            
                            'Check whether to use absolute positioning or snapped to the grid
                            If frmGraphics.SnapChk.Value = 0 Then
                                MapGrh(MapGrhIndex).X = X + ScreenX
                                MapGrh(MapGrhIndex).Y = Y + ScreenY
                            Else
                                MapGrh(MapGrhIndex).X = TileX * GRIDSIZE
                                MapGrh(MapGrhIndex).Y = TileY * GRIDSIZE
                            End If
                            
                            'Update the map
                            UpdateMapGrhPtr = True
                            
                            'Set the "behind" value
                            MapGrh(MapGrhIndex).Behind = frmGraphics.BehindChk.Value
                            Me.Caption = " Screen - " & NumMapGrhs & " tiles"
                            
                        End If
                    End If
                End If
            
            '*** Erase graphics ***
            ElseIf frmGraphics.EraseOpt.Value Then
                
                'Check for any graphics at the tile
                For i = 1 To NumMapGrhs
                    If MapGrh(i).X \ GRIDSIZE = TileX Then
                        If MapGrh(i).Y \ GRIDSIZE = TileY Then
                            
                            'Erase the map graphic
                            If MapGrh(i).Grh.GrhIndex <> 0 Then
                                MapGrh(i).Grh.GrhIndex = 0
                                UpdateMapGrhPtr = True
                            End If
                            
                        End If
                    End If
                Next i
                
            End If
        End If
        
        If frmBG.Visible Then
            '*** Set background ***
            If frmBG.SetOpt.Value Then
                BGIndex = frmBG.LayerCmb.ListIndex + 1
                BGX = (X + (ScreenX \ BGIndex)) \ BGGridSize
                BGY = (Y + (ScreenY \ BGIndex)) \ BGGridSize
                GrhIndex = Val(frmBG.GrhTxt.Text)
                If GrhIndex >= 0 Then
                    Graphics_SetGrh BG(BGIndex).Segment(BGX, BGY), GrhIndex
                End If
            End If
        End If

        If frmTileInfo.Visible Then
            '*** Tile info values ***
            If frmTileInfo.SetOpt.Value Then
                If frmTileInfo.AttLst.ListIndex > -1 Then
                    MapInfo.TileInfo(TileX, TileY) = frmTileInfo.AttLst.ListIndex
                    If frmTileInfo.AttLst.ListIndex = TILETYPE_BLOCKED Or frmTileInfo.AttLst.ListIndex = TILETYPE_NOTHING Then
                        HasFloatingBlocks
                    End If
                End If
            End If
        End If
        
    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Forward to MouseDown
    Form_MouseDown Button, Shift, X, Y

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If EngineRunning Then
        
        'Don't close if the application is still running, just hide
        Me.Visible = False
        Cancel = 1
        
    Else
        
        'Save settings
        IO_INI_Write App.Path & "\Dev\Map Editor Settings.ini", Me.Name, "Left", Me.Left
        IO_INI_Write App.Path & "\Dev\Map Editor Settings.ini", Me.Name, "Top", Me.Top

    End If

End Sub
