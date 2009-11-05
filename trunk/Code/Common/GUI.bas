Attribute VB_Name = "GUI"
Option Explicit

'GUI classes
Public HUD As GUIHUD
Public InfoBox As GUIInfoBox
Public ChatBox As GUIChatBox
Public StatsWindow As GUIStatsWindow
Public InvWindow As GUIInvWindow
Public EquipWindow As GUIEquipWindow

'Window IDs for the GUIs
Public Const WID_HUD As Byte = 0
Public Const WID_ChatBox As Byte = 1
Public Const WID_StatsWindow As Byte = 2
Public Const WID_InvWindow As Byte = 3
Public Const WID_EquipWindow As Byte = 4
Public Const NumWIDs As Byte = 4

'The WID of the window with the focus
Public FocusWID As Byte

'Selected window (0 = none since you can't drag the GUI)
Public SelectedWID As Byte

'If the user is in the mode for entering text into the chat box
Public IsEnteringChat As Boolean

Public Sub GUI_Init()
'*********************************************************************************
'Load up the GUI
'*********************************************************************************

    'Create the classes
    Set HUD = New GUIHUD
    Set InfoBox = New GUIInfoBox
    Set ChatBox = New GUIChatBox
    Set StatsWindow = New GUIStatsWindow
    Set InvWindow = New GUIInvWindow
    Set EquipWindow = New GUIEquipWindow
    
    'Load
    HUD.Load
    InfoBox.Load
    ChatBox.Load
    StatsWindow.Load
    InvWindow.Load
    EquipWindow.Load

End Sub

Public Sub GUI_Draw()
'*********************************************************************************
'Draw the different GUI components
'*********************************************************************************
Dim i As Long

    'HUD and related components always come first
    HUD.Draw
    InfoBox.Draw
    ChatBox.Draw

    'Windows without focus
    For i = NumWIDs To WID_StatsWindow Step -1
        If i <> FocusWID Then GUI_DrawWID i
    Next i
    
    'Window with focus
    If FocusWID >= WID_StatsWindow Then
        GUI_DrawWID FocusWID
    End If

End Sub

Private Sub GUI_DrawWID(ByVal WID As Byte)
'*********************************************************************************
'Draws a window by WID
'*********************************************************************************

    Select Case WID
    
    Case WID_HUD
        HUD.Draw
        
    Case WID_InvWindow
        InvWindow.Draw
        
    Case WID_StatsWindow
        StatsWindow.Draw
        
    Case WID_ChatBox
        ChatBox.Draw
        
    Case WID_EquipWindow
        EquipWindow.Draw
        
    End Select

End Sub

Public Function GUI_FindTargetWindow() As Byte
'*********************************************************************************
'Finds the window that the mouse is over
'*********************************************************************************
Dim i As Byte

    'Start with the focus window
    If FocusWID <> 0 Then
        If GUI_CursorOverWindow(FocusWID) Then
            GUI_FindTargetWindow = FocusWID
            Exit Function
        End If
    End If
    
    'Check everything else but the FocusWID and the HUD
    For i = 1 To NumWIDs
        If GUI_CursorOverWindow(i) Then
            GUI_FindTargetWindow = i
            Exit Function
        End If
    Next i
    
    'Wasn't over any of the windows, so just return the HUD
    GUI_FindTargetWindow = WID_HUD

End Function

Private Function GUI_CursorOverWindow(ByVal WID As Byte) As Boolean
'*********************************************************************************
'Checks if the cursor is over a window
'*********************************************************************************
Dim X As Integer
Dim Y As Integer
Dim Width As Integer
Dim Height As Integer
    
    'Get the window location based off of what window it is
    Select Case WID
        
    Case WID_HUD
        X = 0
        Y = 0
        Width = ScreenWidth
        Height = ScreenHeight
        
    Case WID_ChatBox
        X = ChatBox.X
        Y = ChatBox.Y
        Width = ChatBox.Width
        Height = ChatBox.Height
        
    Case WID_StatsWindow
        If Not StatsWindow.Visible Then Exit Function
        X = StatsWindow.X
        Y = StatsWindow.Y
        Width = StatsWindow.Width
        Height = StatsWindow.Height
        
    Case WID_InvWindow
        If Not InvWindow.Visible Then Exit Function
        X = InvWindow.X
        Y = InvWindow.Y
        Width = InvWindow.Width
        Height = InvWindow.Height
        
    Case WID_EquipWindow
        If Not EquipWindow.Visible Then Exit Function
        X = EquipWindow.X
        Y = EquipWindow.Y
        Width = EquipWindow.Width
        Height = EquipWindow.Height
        
    End Select
    
    'Check if the cursor is over the window
    GUI_CursorOverWindow = Math_Collision_PointRect(CursorPos.X, CursorPos.Y, X, Y, Width, Height)

End Function
