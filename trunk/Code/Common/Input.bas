Attribute VB_Name = "Input"
'*********************************************************************************
'Handles all the input related to the game (ie events from input)
'*********************************************************************************

Option Explicit

Private Declare Function GetActiveWindow Lib "user32" () As Long

Public Sub Input_Mouse_WheelUp()
'*********************************************************************************
'Mouse wheel scroll up events
'*********************************************************************************

End Sub

Public Sub Input_Mouse_WheelDown()
'*********************************************************************************
'Mouse wheel scroll down events
'*********************************************************************************

End Sub

Public Sub Input_Mouse_MiddleDown()
'*********************************************************************************
'Mouse Middle-down events
'*********************************************************************************

End Sub

Public Sub Input_Mouse_MiddleUp()
'*********************************************************************************
'Mouse Middle-up events
'*********************************************************************************

End Sub

Public Sub Input_Mouse_Move()
'*********************************************************************************
'Mouse move events
'*********************************************************************************
Dim TargetWindow As Byte
    
    'Find the target window
    TargetWindow = GUI_FindTargetWindow
    
    'Clear the item tooltip
    TooltipItemIndex = 0
    
    'Move a window
    Select Case SelectedWID
        
    Case WID_StatsWindow
        StatsWindow.X = StatsWindow.X + CursorDiff.X
        StatsWindow.Y = StatsWindow.Y + CursorDiff.Y
    
    Case WID_InvWindow
        InvWindow.X = InvWindow.X + CursorDiff.X
        InvWindow.Y = InvWindow.Y + CursorDiff.Y
    
    Case WID_EquipWindow
        EquipWindow.X = EquipWindow.X + CursorDiff.X
        EquipWindow.Y = EquipWindow.Y + CursorDiff.Y

    End Select
    
    'Forward to the corresponding class
    Select Case TargetWindow

    Case WID_InvWindow
        InvWindow.MoveCursor CursorPos.X, CursorPos.Y
        
    Case WID_EquipWindow
        EquipWindow.MoveCursor CursorPos.X, CursorPos.Y

    End Select

End Sub

Public Sub Input_Mouse_RightUp()
'*********************************************************************************
'Mouse Right-up events
'*********************************************************************************

End Sub

Public Sub Input_Mouse_RightDown()
'*********************************************************************************
'Mouse Right-down events
'*********************************************************************************
Dim TargetWindow As Byte
    
    'Find the target window
    TargetWindow = GUI_FindTargetWindow
    
    'Clear the selected window
    SelectedWID = 0
    
    'Forward to the corresponding class
    Select Case TargetWindow

    Case WID_InvWindow
        InvWindow.RightClick CursorPos.X, CursorPos.Y
        
    Case WID_EquipWindow
        EquipWindow.RightClick CursorPos.X, CursorPos.Y
        
    End Select

End Sub

Public Sub Input_Mouse_LeftUp()
'*********************************************************************************
'Mouse left-up events
'*********************************************************************************

    'Clear the selected window
    SelectedWID = 0

End Sub

Public Sub Input_Mouse_LeftDown()
'*********************************************************************************
'Mouse left-down events
'*********************************************************************************
Dim TargetWindow As Byte
    
    'Find the target window
    TargetWindow = GUI_FindTargetWindow
    
    'Clear the selected window
    SelectedWID = 0
    
    'Forward to the corresponding class
    Select Case TargetWindow
    
    Case WID_ChatBox
        ChatBox.LeftClick CursorPos.X, CursorPos.Y
        
    Case WID_StatsWindow
        StatsWindow.LeftClick CursorPos.X, CursorPos.Y
        
    Case WID_InvWindow
        InvWindow.LeftClick CursorPos.X, CursorPos.Y
        
    Case WID_EquipWindow
        EquipWindow.LeftClick CursorPos.X, CursorPos.Y
        
    Case WID_HUD
        HUD.LeftClick CursorPos.X, CursorPos.Y
        
    End Select

End Sub

Public Sub Input_Mouse_WheelScroll()
'*********************************************************************************
'Mouse wheel scroll events
'*********************************************************************************

End Sub

Public Sub Input_Keys_Down(ByVal KeyCode As Integer)
'*********************************************************************************
'Handles KeyDown events
'*********************************************************************************

    Select Case KeyCode

    Case vbKeyS
        StatsWindow.Visible = Not StatsWindow.Visible
        
    Case vbKeyI
        InvWindow.Visible = Not InvWindow.Visible
        
    Case vbKeyE
        EquipWindow.Visible = Not EquipWindow.Visible
        
    End Select

End Sub

Public Sub Input_Keys_Press(ByVal KeyAscii As Integer)
'*********************************************************************************
'Handles KeyPress events
'*********************************************************************************

    Select Case KeyAscii
        
    'Delete
    Case 8  'Backspace
        If LenB(ChatBox.InputText) > 0 Then
            ChatBox.InputText = Left$(ChatBox.InputText, Len(ChatBox.InputText) - 1)
        End If
        
    'Return
    Case 13 'Enter
        If IsEnteringChat Then
            If LenB(ChatBox.InputText) > 0 Then
                sndBuf.Put_Byte PId.CS_Chat_Say
                sndBuf.Put_String ChatBox.InputText
                ChatBox.InputText = vbNullString
                IsEnteringChat = False
            End If
        Else
            IsEnteringChat = True
        End If
        
    'Text
    Case Else
        If IsEnteringChat Then
            If IsLegalString(Chr$(KeyAscii)) Then
                ChatBox.InputText = ChatBox.InputText & Chr$(KeyAscii)
            End If
        End If
        
    End Select

End Sub

Public Sub Input_Keys_Polling()
'*********************************************************************************
'Handles key events through key polling (checks every frame, not at the rate which
'Windows returns key presses)
'*********************************************************************************
Dim bMoved As Boolean
Dim CheckTileX As Integer
Dim CheckTileY As Integer
Dim i As Long

    'Check that our window has the focus
    If GetActiveWindow = 0 Then Exit Sub
    
    'Pick up items from the ground
    If GetKey(vbKeyAlt) Then
        If CharList(UserCharIndex).Action = eNone Then
            If LastInputPickupTime < timeGetTime - PICKUPITEMTIME Then
                For i = 0 To MapItemsUBound
                    If MapItems(i).ItemIndex > 0 Then
                        If MapItems(i).LastPickupAttempt + PICKUPSAMEITEMTIME < timeGetTime Then
                            If Math_Collision_PointRect(MapItems(i).X, MapItems(i).Y, CharList(UserCharIndex).X + _
                                CharList(UserCharIndex).Width \ 2 - (PICKUPITEMDISTCLIENT \ 2), CharList(UserCharIndex).Y + _
                                CharList(UserCharIndex).Height - (PICKUPITEMDISTCLIENT \ 2), _
                                PICKUPITEMDISTCLIENT, PICKUPITEMDISTCLIENT) Then
                                sndBuf.Put_Byte PId.CS_Pickup
                                sndBuf.Put_Integer i
                                LastInputPickupTime = timeGetTime
                                MapItems(i).LastPickupAttempt = timeGetTime
                                Exit For
                            End If
                        End If
                    End If
                Next i
            End If
        End If
    End If
    
    'Punch
    If GetKey(vbKeyControl) Then
        If CharList(UserCharIndex).Action = eNone Then
            If LastInputAttackTime < timeGetTime - 100 Then
                With CharList(UserCharIndex)
                    i = Engine_CharCollision(.X - (PDBody(.BodyIndex).PunchWidth * -(.Heading = WEST)), .Y, _
                        PDBody(.BodyIndex).PunchWidth, .Height, True)
                End With
                If i > 0 Then
                    sndBuf.Put_Byte PId.CS_Punch_Hit
                    sndBuf.Put_Integer i
                Else
                    sndBuf.Put_Byte PId.CS_Punch
                End If
                LastInputAttackTime = timeGetTime
            End If
        End If
    End If

    'Move right
    If GetKey(vbKeyRight) Then
        If CharList(UserCharIndex).Action = eNone Then
            If CharList(UserCharIndex).MoveDir <> EAST Then
                If LastInputMoveTime < timeGetTime - 100 Then
                    CheckTileX = ((CharList(UserCharIndex).X + (ElapsedTime * MOVESPEED) + CharList(UserCharIndex).Width) \ GRIDSIZE)
                    For i = 0 To CharList(UserCharIndex).Height \ GRIDSIZE
                        CheckTileY = (CharList(UserCharIndex).Y \ GRIDSIZE) + i
                        If Engine_GetTileInfo(CheckTileX, CheckTileY) = TILETYPE_BLOCKED Then
                            GoTo DontMoveRight
                        End If
                    Next i
                    sndBuf.Put_Byte PId.CS_MoveEast
                    LastInputMoveTime = timeGetTime
                    LastSentMoveDir = EAST
                End If
            End If
            bMoved = True
        End If
    End If
DontMoveRight:
    
    'Move left
    If GetKey(vbKeyLeft) Then
        If CharList(UserCharIndex).Action = eNone Then
            If CharList(UserCharIndex).MoveDir <> WEST Then
                If LastInputMoveTime < timeGetTime - 100 Then
                    CheckTileX = ((CharList(UserCharIndex).X - (ElapsedTime * MOVESPEED)) \ GRIDSIZE)
                    For i = 0 To (CharList(UserCharIndex).Height \ GRIDSIZE)
                        CheckTileY = (CharList(UserCharIndex).Y \ GRIDSIZE) + i
                        If Engine_GetTileInfo(CheckTileX, CheckTileY) = TILETYPE_BLOCKED Then
                            GoTo DontMoveLeft
                        End If
                    Next i
                    sndBuf.Put_Byte PId.CS_MoveWest
                    LastInputMoveTime = timeGetTime
                    LastSentMoveDir = WEST
                End If
            End If
            bMoved = True
        End If
    End If
DontMoveLeft:

    'Jump
    If GetKey(vbKeyUp) Then
        If CharList(UserCharIndex).Action <> eHit Then
            If CharList(UserCharIndex).OnGround = 1 Then
                If CharList(UserCharIndex).Jump = 0 Then
                    If LastInputJumpTime < timeGetTime - 100 Then
                        CheckTileY = (CharList(UserCharIndex).Y \ GRIDSIZE)
                        For i = 0 To (CharList(UserCharIndex).Width \ GRIDSIZE)
                            CheckTileX = (CharList(UserCharIndex).X \ GRIDSIZE) + i
                            If Engine_GetTileInfo(CheckTileX, CheckTileY) = TILETYPE_BLOCKED Then
                                GoTo DontJump
                            End If
                        Next i
                        sndBuf.Put_Byte PId.CS_Jump
                        LastInputJumpTime = timeGetTime
                    End If
                End If
            End If
        End If
    End If
DontJump:

    'Check if the user stopped moving
    If Not bMoved Then
        If LastSentMoveDir <> 0 Then
            sndBuf.Put_Byte PId.CS_MoveStop
            LastSentMoveDir = 0
        End If
    End If
    
    'Quit
    If GetKey(vbKeyEscape) Then EngineRunning = False
    
End Sub

