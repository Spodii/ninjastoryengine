VERSION 5.00
Object = "{D21493D3-ED81-46BA-B7BC-6C76771C54C2}#1.0#0"; "GOREsockClient.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Ninja Story"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Begin GOREsock.GOREsockClient GOREsock 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements DirectXEvent8

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Sub DirectXEvent8_DXCallback(ByVal EventID As Long)
'*********************************************************************************
'Handles mouse device events (movement, clicking, mouse wheel scrolling, etc)
'*********************************************************************************
Dim DevData(1 To 50) As DIDEVICEOBJECTDATA
Dim NumEvents As Long
Dim i As Long

Dim bMoved As Boolean
Dim bLeftDown As Boolean
Dim bLeftUp As Boolean
Dim bRightDown As Boolean
Dim bRightUp As Boolean
Dim bMiddleDown As Boolean
Dim bMiddleUp As Boolean

    On Error GoTo ErrOut

    'Check if message is for us
    If EventID <> MouseEvent Then Exit Sub
    If GetActiveWindow = 0 Then Exit Sub
    
    'Retrieve data
    NumEvents = Input_GetDeviceData(DevData())
    
    'Loop through the data
    For i = 1 To NumEvents
        Select Case DevData(i).lOfs
        
        'Moved on the X axis
        Case DIMOFS_X
            If Not Fullscreen Then
                LastCursorPos = CursorPos
                GetCursorPos CursorPos
                CursorPos.X = CursorPos.X - (Me.Left \ Screen.TwipsPerPixelX)
                CursorPos.Y = CursorPos.Y - (Me.Top \ Screen.TwipsPerPixelY)
                CursorDiff.X = CursorDiff.X + (CursorPos.X - LastCursorPos.X)
                CursorDiff.Y = CursorDiff.Y + (CursorPos.Y - LastCursorPos.Y)
            Else
                CursorDiff.X = CursorDiff.X + (DevData(i).lData * CursorSpeed)
            End If
            bMoved = True
        
        'Moved on the Y axis
        Case DIMOFS_Y
            If Not Fullscreen Then
                LastCursorPos = CursorPos
                GetCursorPos CursorPos
                CursorPos.X = CursorPos.X - (Me.Left \ Screen.TwipsPerPixelX)
                CursorPos.Y = CursorPos.Y - (Me.Top \ Screen.TwipsPerPixelY)
                CursorDiff.X = CursorDiff.X + (CursorPos.X - LastCursorPos.X)
                CursorDiff.Y = CursorDiff.Y + (CursorPos.Y - LastCursorPos.Y)
            Else
                CursorDiff.Y = CursorDiff.Y + (DevData(i).lData * CursorSpeed)
            End If
            bMoved = True
        
        'Mouse wheel moved
        Case DIMOFS_Z
            If DevData(i).lData > 0 Then
                Input_Mouse_WheelUp
            ElseIf DevData(i).lData < 0 Then
                Input_Mouse_WheelDown
            End If
        
        'Left button event
        Case DIMOFS_BUTTON0
            If DevData(i).lData = 0 Then
                If LeftButtonDown Then
                    LeftButtonDown = False
                    bLeftUp = True
                End If
            Else
                If Not LeftButtonDown Then
                    LeftButtonDown = True
                    bLeftDown = True
                End If
            End If
        
        'Right button event
        Case DIMOFS_BUTTON1
            If DevData(i).lData = 0 Then
                If RightButtonDown Then
                    RightButtonDown = False
                    bRightUp = True
                End If
            Else
                If Not RightButtonDown Then
                    RightButtonDown = True
                    bRightDown = True
                End If
            End If
            
        'Middle button event
        Case DIMOFS_BUTTON2
            If DevData(i).lData = 0 Then
                If MiddleButtonDown Then
                    MiddleButtonDown = False
                    bMiddleUp = True
                End If
            Else
                If Not MiddleButtonDown Then
                    MiddleButtonDown = True
                    bMiddleDown = True
                End If
            End If
        
        End Select
    Next i
    
    'When in fullscreen mode, perform extra in-bound checks and position updates
    If Fullscreen Then
    
        'Set the new cursor position
        CursorPos.X = CursorPos.X + CursorDiff.X
        CursorPos.Y = CursorPos.Y + CursorDiff.Y
        
        'Make sure the cursor is located in the screen
        If CursorPos.X < 0 Then CursorPos.X = 0
        If CursorPos.X > Me.ScaleWidth Then CursorPos.X = Me.ScaleWidth
        If CursorPos.Y < 0 Then CursorPos.Y = 0
        If CursorPos.Y > Me.ScaleHeight Then CursorPos.Y = Me.ScaleHeight
        
    End If
    
    'Mouse movement handle forwarding
    If bMoved Then Input_Mouse_Move
    If bLeftDown Then Input_Mouse_LeftDown
    If bLeftUp Then Input_Mouse_LeftUp
    If bRightDown Then Input_Mouse_RightDown
    If bRightUp Then Input_Mouse_RightUp
    If bMiddleDown Then Input_Mouse_MiddleDown
    If bMiddleUp Then Input_Mouse_MiddleUp
    
    'Clear the last difference
    CursorDiff.X = 0
    CursorDiff.Y = 0
    
    Exit Sub
    
ErrOut:

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    'Forward to the KeyDown events
    Input_Keys_Down KeyCode

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    'Forward to the KeyPress events
    Input_Keys_Press KeyAscii

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Make sure the input device is acquired
    Input_Acquire

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Close down the engine
    EngineRunning = False
    Cancel = 1

End Sub

Private Sub GOREsock_OnConnecting(inSox As Long)

    If GettingAccount Then
        
        'Give the socket some time to be fully loaded
        DoEvents
        Sleep 50
        DoEvents
        
        'Build the packet
        sndBuf.Clear
        sndBuf.Put_Byte AccountPId.CS_GetChars
        sndBuf.Put_String Trim$(frmConnect.NameTxt.Text)
        sndBuf.Put_String Trim$(frmConnect.PassTxt.Text)
        
        'Send the packet to the account server
        frmMain.GOREsock.SendData LocalSocketID, sndBuf.Get_Buffer()
        
        'Clear the buffer
        sndBuf.Clear

    Else
    
        'Check if the socket is already open
        If Not SocketOpen Then
            
            'Wait for a little bit so the socket is ready
            DoEvents
            Sleep 50
            DoEvents
            
            'Send the initial connection packet
            sndBuf.Clear
            sndBuf.Put_Byte PId.CS_Connect
            sndBuf.Put_String Trim$(frmConnect.NameTxt.Text)
            sndBuf.Put_String Trim$(frmConnect.CharsLst.List(frmConnect.CharsLst.ListIndex))
            sndBuf.Put_String Trim$(frmConnect.PassTxt.Text)
            
            'Send the data
            Data_Send
                
            'Make sure the engine is loaded
            If Not EngineRunning Then GameLoop
                
        End If

    End If

End Sub

Private Sub GOREsock_OnDataArrival(inSox As Long, inData() As Byte)
Dim PacketID As Byte
Dim i As Byte
Dim s As String

    If GettingAccount Then
    
        'Set the buffer
        rBuf.Set_Buffer inData()
    
        'Get the packet ID
        PacketID = rBuf.Get_Byte
        
        'Forward to the corresponding packet handling method
        With AccountPId
            Select Case PacketID
            
            Case .SC_BadPass:
                MsgBox "Invalid password.", vbOKOnly
            
            Case .SC_NoChars:
                MsgBox "Account does not exist.", vbOKOnly
            
            Case .SC_SendChars
                frmConnect.CharsLst.Clear
                For i = 1 To 5
                    s = rBuf.Get_String
                    If s <> vbNullString Then frmConnect.CharsLst.AddItem s
                Next i
                frmConnect.CharsLst.Enabled = True
                frmConnect.LoginCmd.Enabled = True
                frmConnect.NameTxt.Enabled = False
                frmConnect.PassTxt.Enabled = False
                frmConnect.ConnectCmd.Enabled = False
                frmConnect.CharsLst.ListIndex = Val(IO_INI_Read(App.Path & "\Data\Settings.ini", "CONNECT", "CharID"))
            
            Case 0: rBuf.Overflow       'An error occured
            Case Else: rBuf.Overflow    'An error occured
            End Select
        End With
        
        'Close down the connection to the account server
        frmMain.GOREsock.Shut LocalSocketID
        frmMain.GOREsock.ShutDown
    
    Else
        
        'Forward to the packet handler
        Data_Handle inSox, inData()
        
    End If

End Sub
