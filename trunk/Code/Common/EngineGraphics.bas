Attribute VB_Name = "EngineGraphics"
'*********************************************************************************
'Handles all of the graphical parts of the engine
'*********************************************************************************

Option Explicit

'Constants
Public Fullscreen As Boolean                    'Use fullscreen mode
Private Depth32Bit As Boolean                   'Use 32-bit pixel depth (fullscreen only, windowed defaults to desktop settings)
Private VSync As Boolean                        'Use VSync refreshing (fullscreen only)
Private Const TexturePath As String = "\Graphics\"
Private Const ColorKey As Long = -65281         'Color key for textures
Private Const TextureLife As Long = 600000      'How long (in milliseconds) a texture can remain in memory unused

'Different fonts
Public FontDefault As EngineFont

'Describes a flexible vertex format
Private Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    Rhw As Single
    Color As Long
    tU As Single
    tV As Single
End Type
Public Const FVF_Size As Long = 28

'DirectX 8 Objects
Private D3D As Direct3D8
Private D3DX As D3DX8
Private D3DDevice As Direct3DDevice8

'Device variables
Private D3DWindow As D3DPRESENT_PARAMETERS      'Describes the viewport and used to restore when in fullscreen
Private UsedCreateFlags As CONST_D3DCREATEFLAGS 'The flags we used to create the device when it first succeeded
Private DispMode As D3DDISPLAYMODE              'Describes the display mode

'Texture information
Private Type Textures
    Tex As Direct3DTexture8
    UnloadTime As Long
    Width As Long
    Height As Long
    Loaded As Boolean
End Type
Private Textures() As Textures
Private NumTextures As Long 'Total number of textures
Private LastTexture As Long 'The last texture ID to be used

'Describes the return from a texture init
Private Type D3DXIMAGE_INFO_A
    Width As Long
    Height As Long
    Depth As Long
    MipLevels As Long
    Format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
End Type

'Misc
Private TargethWnd As Long  'The handle to render to

Public Sub Graphics_CreateTexture(ByRef Tex As Direct3DTexture8, ByVal FilePath As String, Optional ByRef SrcWidth As Long, Optional ByRef SrcHeight As Long, Optional ByVal ColorKey As Long = &HFF000000)
'*********************************************************************************
'Sets a texture with the specified parameters
'*********************************************************************************
Dim TexInfo As D3DXIMAGE_INFO_A
Dim Data() As Byte

    'Check if the texture is encrypted
    If LCase$(Right$(FilePath, 1)) = "e" Then
        
        'Create encrypted texture (load the file into memory, decrypt the memory, then create the texture
        'using the array of bytes in memory)
        IO_LoadEncryptedTexture FilePath, Data()
        Set Tex = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, Data(0), UBound(Data) + 1, D3DX_DEFAULT, _
            D3DX_DEFAULT, 1, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, ColorKey, TexInfo, ByVal 0)
    
    Else
    
        'Create unencrypted texture (direct copy from the file)
        Set Tex = D3DX.CreateTextureFromFileEx(D3DDevice, FilePath, D3DX_DEFAULT, D3DX_DEFAULT, 1, _
            0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, ColorKey, TexInfo, ByVal 0)
        
    End If

    'Return the source texture size
    SrcWidth = TexInfo.Width
    SrcHeight = TexInfo.Height

End Sub

Public Sub Graphics_DrawTriangleList(ByRef v() As TLVERTEX, ByVal Primitives As Long)
'*********************************************************************************
'Renders a triangle list
'*********************************************************************************

    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, Primitives, v(0), FVF_Size

End Sub

Public Sub Graphics_DrawRect(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, _
    ByVal Color0 As Long, ByVal Color1 As Long, ByVal Color2 As Long, ByVal Color3 As Long)
'*********************************************************************************
'Renders a TLVertex with 4 points in a triangle strip to draw a rectangle
'*********************************************************************************
Dim v(0 To 3) As TLVERTEX

    v(0).X = X
    v(0).Y = Y
    v(1).X = X + Width
    v(1).Y = Y
    v(2).X = X
    v(2).Y = Y + Height
    v(3).X = X + Width
    v(3).Y = Y + Height
    v(0).Color = Color0
    v(1).Color = Color1
    v(2).Color = Color2
    v(3).Color = Color3
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, v(0), FVF_Size

End Sub

Public Sub Graphics_SetTextureEX(ByRef Tex As Direct3DTexture8)
'*********************************************************************************
'Sets a texture that is not part of the internal Textures() array
'*********************************************************************************

    'Check for a valid texture
    If Tex Is Nothing Then Exit Sub
    
    'Set the texture
    D3DDevice.SetTexture 0, Tex
    
    'Clear the last texture value
    LastTexture = 0

End Sub

Public Sub Graphics_SetTexture(ByVal TextureID As Long)
'*********************************************************************************
'Set the texture to be used
'*********************************************************************************

    If TextureID = 0 Then
    
        'Clear the texture
        D3DDevice.SetTexture 0, Nothing
        
    Else
    
        'Check if the texture is loaded into memory
        If Not Textures(TextureID).Loaded Then Graphics_LoadTexture TextureID
    
        'Set the texture if it is not already set to TextureID
        If LastTexture <> TextureID Then D3DDevice.SetTexture 0, Textures(TextureID).Tex
        
        'Update the texture's life timer
        Textures(TextureID).UnloadTime = timeGetTime + TextureLife

    End If

End Sub

Public Sub Graphics_FlushTextures()
'*********************************************************************************
'Check to flush any textures from memory that haven't been used for TextureLife time
'*********************************************************************************
Dim i As Long

    For i = 1 To NumTextures
        If Textures(i).UnloadTime < timeGetTime Then
            Graphics_UnloadTexture i
        End If
    Next i

End Sub

Public Sub Graphics_SetGrh(ByRef Grh As tGrh, ByVal GrhIndex As Long, Optional ByVal AnimType As Byte = ANIMTYPE_STATIONARY)
'*********************************************************************************
'Set the values for a Grh so it can be used
'*********************************************************************************
    
    Grh.GrhIndex = GrhIndex
    Grh.LastUpdated = timeGetTime
    Grh.AnimType = AnimType
    Grh.Frame = 1

End Sub

Public Sub Graphics_BeginScene()
'*********************************************************************************
'Begin the rendering scene
'*********************************************************************************

    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    D3DDevice.BeginScene

End Sub

Public Sub Graphics_EndScene()
'*********************************************************************************
'End the rendering scene and present the screen
'*********************************************************************************
    
    D3DDevice.EndScene
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

End Sub

Public Function Graphics_GetGrhIndex(ByRef Grh As tGrh) As Long
'*********************************************************************************
'Get the GrhIndex
'*********************************************************************************
    
    If Grh.AnimType = ANIMTYPE_STATIONARY Then
        Graphics_GetGrhIndex = Grh.GrhIndex
    Else
        If GrhData(Grh.GrhIndex).NumFrames > 1 Then
            Graphics_GetGrhIndex = GrhData(Grh.GrhIndex).Frames(Int(Grh.Frame))
        Else
            Graphics_GetGrhIndex = Grh.GrhIndex
        End If
    End If

End Function

Public Function Graphics_FToDW(f As Single) As Long
'*********************************************************************************
'Converts a float to a D-Word, or in Visual Basic terms, a Single to a Long
'*********************************************************************************
Dim Buf As D3DXBuffer

    'Converts a single into a long (Float to DWORD)
    Set Buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData Buf, 0, 4, 1, f
    D3DX.BufferGetData Buf, 0, 4, 1, Graphics_FToDW
    Set Buf = Nothing

End Function

Public Sub Graphics_SetRenderState(ByVal StateType As CONST_D3DRENDERSTATETYPE, ByVal Value As Long)
'*********************************************************************************
'Sets a render state for the D3D Device
'*********************************************************************************
    
    D3DDevice.SetRenderState StateType, Value

End Sub

Public Sub Graphics_DrawLineStrip(ByRef v() As TLVERTEX, ByVal Primitives As Long)
'*********************************************************************************
'Draws a line strip
'*********************************************************************************

    D3DDevice.DrawPrimitiveUP D3DPT_LINESTRIP, Primitives, v(0), FVF_Size

End Sub

Public Sub Graphics_DrawParticles(ByRef v() As TLVERTEX, ByVal ParticleCount As Long)
'*********************************************************************************
'Draws a point list for particle effects
'*********************************************************************************

    D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, ParticleCount, v(0), FVF_Size

End Sub

Public Sub Graphics_DrawGrh(ByRef Grh As tGrh, ByVal X As Long, ByVal Y As Long, Optional ByVal Light0 As Long = -1, _
    Optional ByVal Light1 As Long = -1, Optional ByVal Light2 As Long = -1, Optional ByVal Light3 As Long = -1, Optional ByVal FlipWidth As Long = 0)
'*********************************************************************************
'Draws based off of Grh information
'*********************************************************************************
Dim v(0 To 3) As TLVERTEX
Dim GrhIndex As Long
Dim TextureID As Long
Dim w As Single   'Shortcut to the width
Dim H As Single   'Shortcut to the height

    'Check for a valid GrhIndex
    If Grh.GrhIndex = 0 Then Exit Sub

    'Update
    Graphics_UpdateGrh Grh
    
    'Check that the Grh info is valid
    If GrhData(Grh.GrhIndex).NumFrames > 1 Then
        If Int(Grh.Frame) > GrhData(Grh.GrhIndex).NumFrames Then Exit Sub
        If Int(Grh.Frame) < 1 Then Exit Sub
    End If
    
    'Get the GrhIndex
    GrhIndex = Graphics_GetGrhIndex(Grh)
    
    'Store the output width and height
    w = GrhData(GrhIndex).Width
    H = GrhData(GrhIndex).Height
    
    'Perform in-bounds check if needed
    If X + w > 0 Then
        If Y + H > 0 Then
            If X < ScreenWidth Then
                If Y < ScreenHeight Then
                            
                    'Set the texture
                    TextureID = GrhData(GrhIndex).TextureID
                    Graphics_SetTexture TextureID
                    
                    'Check to flip
                    If FlipWidth > 0 Then
                    
                        'Set up the TLVertex flipped horizontally
                        v(0) = Graphics_CreateTLVertex(X + FlipWidth, Y, Light0, GrhData(GrhIndex).X / Textures(TextureID).Width, GrhData(GrhIndex).Y / Textures(TextureID).Height)
                        v(1) = Graphics_CreateTLVertex(X - w + FlipWidth, Y, Light1, (GrhData(GrhIndex).X + w + 1) / Textures(TextureID).Width, v(0).tV)
                        v(2) = Graphics_CreateTLVertex(X + FlipWidth, Y + H, Light2, v(0).tU, (GrhData(GrhIndex).Y + H + 1) / Textures(TextureID).Height)
                        v(3) = Graphics_CreateTLVertex(X - w + FlipWidth, Y + H, Light3, v(1).tU, v(2).tV)
                    
                    Else
                        
                        'Set up the TLVertex normally
                        v(0) = Graphics_CreateTLVertex(X, Y, Light0, GrhData(GrhIndex).X / Textures(TextureID).Width, GrhData(GrhIndex).Y / Textures(TextureID).Height)
                        v(1) = Graphics_CreateTLVertex(X + w, Y, Light1, (GrhData(GrhIndex).X + w + 1) / Textures(TextureID).Width, v(0).tV)
                        v(2) = Graphics_CreateTLVertex(X, Y + H, Light2, v(0).tU, (GrhData(GrhIndex).Y + H + 1) / Textures(TextureID).Height)
                        v(3) = Graphics_CreateTLVertex(X + w, Y + H, Light3, v(1).tU, v(2).tV)
                        
                    End If
                    
                    'Draw
                    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, v(0), FVF_Size
                    
                End If
            End If
        End If
    End If

End Sub

Private Function Graphics_CreateTLVertex(ByVal X As Long, ByVal Y As Long, ByVal Color As Long, ByVal tU As Single, ByVal tV As Single) As TLVERTEX
'*********************************************************************************
'Wrapper to create a TLVertex in less lines
'*********************************************************************************
    
    With Graphics_CreateTLVertex
        .X = X
        .Y = Y
        .Color = Color
        .Rhw = 1
        .tU = tU
        .tV = tV
    End With

End Function

Private Sub Graphics_UpdateGrh(ByRef Grh As tGrh)
'*********************************************************************************
'Updates an animated Grh
'*********************************************************************************

    'Check for a valid GrhIndex
    If Grh.GrhIndex = 0 Then Exit Sub

    'Check for an animated Grh
    If Grh.AnimType = ANIMTYPE_LOOP Or Grh.AnimType = ANIMTYPE_LOOPONCE Then
        
        'Check that the Grh is animated
        If GrhData(Grh.GrhIndex).NumFrames < 2 Then Exit Sub
        
        'Update the frames
        Grh.Frame = Grh.Frame + ((timeGetTime - Grh.LastUpdated) * GrhData(Grh.GrhIndex).Speed)
        Grh.LastUpdated = timeGetTime
        
        'Check if to roll over
        If Int(Grh.Frame) > GrhData(Grh.GrhIndex).NumFrames Then

            If Grh.AnimType = ANIMTYPE_LOOP Then
            
                'Roll over
                Do While Grh.Frame > GrhData(Grh.GrhIndex).NumFrames
                    Grh.Frame = Grh.Frame - GrhData(Grh.GrhIndex).NumFrames
                Loop
                
            Else
            
                'Done animating
                Grh.AnimType = ANIMTYPE_STATIONARY
                Grh.Frame = 1
            
            End If
            
        End If
        
    End If

End Sub

Private Sub Graphics_UnloadTexture(ByVal TextureID As Long)
'*********************************************************************************
'Unload a texture from memory
'*********************************************************************************

    'Unload the texture
    Set Textures(TextureID).Tex = Nothing
    
    'Set the loaded value to false
    Textures(TextureID).Loaded = False

End Sub

Private Sub Graphics_SetRenderStates()
'*********************************************************************************
'Set up the render states
'*********************************************************************************

    With D3DDevice
        
        'Set the shader to be used
        D3DDevice.SetVertexShader FVF
    
        'Set the render states
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE

        'Particle engine settings
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
        'Set the texture stage stats (filters)
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
        
    End With
    
End Sub

Private Sub Graphics_InitTextures()
'*********************************************************************************
'Load the information concerning textures
'*********************************************************************************
    
    Log "Loading texture settings"

    'Get the number of textures
    NumTextures = Val(IO_INI_Read(App.Path & "\Data\Grh.ini", "GRAPHICS", "NumTextures"))
    Log "Found " & NumTextures & " textures"
    
    'Check for a valid number
    If NumTextures <= 0 Then
        MsgBox "Error retrieving the number of textures from " & vbNewLine & _
            App.Path & "\Data\Grh.ini", vbOKOnly
        EngineRunning = False
        Exit Sub
    End If
    
    'Resize the array to fit the values
    ReDim Textures(1 To NumTextures)

End Sub

Private Sub Graphics_LoadTexture(ByVal TextureID As Long)
'*********************************************************************************
'Load a texture
'*********************************************************************************
Dim TexInfo As D3DXIMAGE_INFO_A
Dim FilePath As String
Dim Data() As Byte

    Log "Loading texture " & TextureID

    'Confirm that the file exists and the texture number is valid
    If TextureID > NumTextures Then
        Log "Texture " & TextureID & " was greater than NumTextures (" & NumTextures & ")"
        Exit Sub
    End If
    
    'Store the path to the texture
    FilePath = App.Path & TexturePath & TextureID & ".png"
    If Not IO_FileExist(FilePath) Then
        
        'If the encrypted version was not found, look for the decrypted version
        FilePath = FilePath & "e"
        If Not IO_FileExist(FilePath) Then
            
            'Neither encrypted nor decrypted file could be found
            Log "File for texture " & TextureID & " could not be found"
            Exit Sub
            
        End If
    End If
    
    'Load up the texture
    Graphics_CreateTexture Textures(TextureID).Tex, FilePath, Textures(TextureID).Width, Textures(TextureID).Height, ColorKey

    'Set the texture as loaded in memory and reset its timer
    Textures(TextureID).Loaded = True
    Textures(TextureID).UnloadTime = timeGetTime + TextureLife
    
    Log "Texture " & TextureID & " created"
        
End Sub

Public Sub Graphics_Init(ByVal hWnd As Long, Optional ByVal bFullscreen As Boolean = False, Optional ByVal bUse32Bit As Boolean = False, _
    Optional ByVal bVSync As Boolean = False)
'*********************************************************************************
'Create the graphics engine
'*********************************************************************************

    'Store the engine settings
    Fullscreen = bFullscreen
    Depth32Bit = bUse32Bit
    VSync = bVSync

    'Create the root Direct3D objects
    Log "Creating D3D object"
    Set D3D = DX.Direct3DCreate()
    Log "Creating D3DX object"
    Set D3DX = New D3DX8
    
    'Load the texture information
    Graphics_InitTextures

    'Store the render target handle
    TargethWnd = hWnd
    Log "TargethWnd set to " & TargethWnd

    'Create the device
    Graphics_CreateDevice
    
    'Set up the render states
    Graphics_SetRenderStates
    
    'Create the particle engine
    Effect_Init
    
    'Create the fonts
    Set FontDefault = New EngineFont
    FontDefault.Load App.Path & "\Data\texdefault.dat", App.Path & "\Data\texdefault.png"
    
    'Load the GUI
    GUI_Init
    
    'Set the starting FPS
    FPS = 60
    ElapsedTime = 1000 / 60

End Sub

Private Sub Graphics_CreateDevice()
'*********************************************************************************
'Create the Direcct3D8 device
'*********************************************************************************

    'Retrieve current display mode
    Log "Retrieving adapter display mode"
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode

    'Set up the device for windowed or fullscreen
    Log "Creating D3DWindow"
    With D3DWindow
        If Fullscreen Then
            If Depth32Bit Then DispMode.Format = D3DFMT_X8R8G8B8 Else DispMode.Format = D3DFMT_R5G6B5
            If VSync Then .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC Else .SwapEffect = D3DSWAPEFFECT_COPY
            .BackBufferCount = 1
            .BackBufferFormat = DispMode.Format
            .BackBufferWidth = ScreenWidth
            .BackBufferHeight = ScreenHeight
            .hDeviceWindow = TargethWnd
        Else
            .Windowed = 1
            .SwapEffect = D3DSWAPEFFECT_COPY
            .BackBufferFormat = DispMode.Format
        End If
    End With
    
    Log "Creating D3DDevice"
    
    'Set the primary format to a dummy format, forcing it to be changed
    UsedCreateFlags = D3DCREATE_MULTITHREADED
    
    'Loop through all possible formats
    On Error GoTo NextFormat
NextFormat:
    
    'Move down to the next format
    Select Case UsedCreateFlags
        Case D3DCREATE_MULTITHREADED: UsedCreateFlags = D3DCREATE_PUREDEVICE
        Case D3DCREATE_PUREDEVICE: UsedCreateFlags = D3DCREATE_HARDWARE_VERTEXPROCESSING
        Case D3DCREATE_HARDWARE_VERTEXPROCESSING: UsedCreateFlags = D3DCREATE_MIXED_VERTEXPROCESSING
        Case D3DCREATE_MIXED_VERTEXPROCESSING: UsedCreateFlags = D3DCREATE_SOFTWARE_VERTEXPROCESSING
        Case D3DCREATE_SOFTWARE_VERTEXPROCESSING
            'If we hit an error trying this format, we ran out of options - the device can't be made
            MsgBox "Error creating the Direct3D device.", vbOKOnly
            Exit Sub
    End Select

    'Set the D3DDevices
    If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
    Log "Attempting format " & UsedCreateFlags
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, TargethWnd, UsedCreateFlags, D3DWindow)
    
    'If we made it this far with no error, we have a working format
    Log "D3DDevice created with format " & UsedCreateFlags
    On Error GoTo 0

End Sub

Public Sub Graphics_Destroy()
'*********************************************************************************
'Destroy the objects used by the class
'*********************************************************************************
Dim i As Long

    Log "Destroying graphics engine"
    
    'Unload the particle engine
    Effect_Destroy
    
    'Destroy the textures
    For i = 1 To NumTextures
        If Not Textures(i).Tex Is Nothing Then Set Textures(i).Tex = Nothing
    Next i
    Erase Textures()
    
    'Destroy the DirectX objects
    If Not D3D Is Nothing Then Set D3D = Nothing
    If Not D3DX Is Nothing Then Set D3DX = Nothing
    If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
    
    'Destroy the fonts
    Set FontDefault = Nothing

End Sub

