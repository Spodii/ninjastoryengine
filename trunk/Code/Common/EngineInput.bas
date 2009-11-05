Attribute VB_Name = "EngineInput"
Option Explicit

'Input objects (DirectInput)
Private DI As DirectInput8
Private DIDevice As DirectInputDevice8

'Mouse event id
Public MouseEvent As Long

'Cursor information
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public CursorPos As POINTAPI        'Cursor position at the current frame
Public LastCursorPos As POINTAPI    'Cursor position at the last frame
Public CursorDiff As POINTAPI       'Difference between the last position and current
Public CursorSpeed As Single        'Cursor speed multiplier
Public CursorGrh As tGrh            'Cursor graphic

'Mouse state
Public LeftButtonDown As Boolean    'If the left mouse button is down
Public RightButtonDown As Boolean   'If the right mouse button is down
Public MiddleButtonDown As Boolean  'If the middle (wheel) button is down

Public Function Input_GetDeviceData(ByRef DevData() As DIDEVICEOBJECTDATA) As Long
'*********************************************************************************
'Returns the array of events that have happened
'*********************************************************************************

    Input_GetDeviceData = DIDevice.GetDeviceData(DevData(), DIGDD_DEFAULT)

End Function

Public Sub Input_Acquire()
'*********************************************************************************
'Acquire the input device
'*********************************************************************************
    
    On Error Resume Next
    DIDevice.Acquire
    On Error GoTo 0
    
End Sub

Public Sub Input_Init(ByRef TargetForm As Form)
'*********************************************************************************
'Load up the input
'*********************************************************************************
Dim diProp As DIPROPLONG

    'Create the input objects
    Set DI = DX.DirectInputCreate
    Set DIDevice = DI.CreateDevice("guid_SysMouse")
    Call DIDevice.SetCommonDataFormat(DIFORMAT_MOUSE)
    
    'If in windowed mode, free the mouse from the screen
    If Fullscreen Then
        
        'Take complete control of the cursor
        Call DIDevice.SetCooperativeLevel(TargetForm.hWnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE)
        
    Else
    
        'Use the system cursor
        Call DIDevice.SetCooperativeLevel(TargetForm.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE)
        
    End If
    
    'Set the input information (buffer size, event handler object, etc)
    diProp.lHow = DIPH_DEVICE
    diProp.lObj = 0
    diProp.lData = 50
    Call DIDevice.SetProperty("DIPROP_BUFFERSIZE", diProp)
    MouseEvent = DX.CreateEvent(TargetForm)
    DIDevice.SetEventNotification MouseEvent
    
    'Acquire the input
    Input_Acquire

End Sub
