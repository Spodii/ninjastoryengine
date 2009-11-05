Attribute VB_Name = "EngineParticles"
'*********************************************************************************
'Holds all the information for the different particle effects
'*********************************************************************************

Option Explicit

Private Type Effect
    X As Single                 'Location of effect
    Y As Single
    Gfx As Long                 'Particle texture used
    RenderTime As Byte          'When the effect is rendered
    AdjustToScreen As Boolean   'If ScreenX and ScreenY is taken into the position calculations
    Used As Boolean             'If the effect is in use
    Life As Long                'The life of the effect
    BlendOne As Boolean         'If to use BlendOne rendering method
    EffectType As Long          'What number of effect that is used
    FloatSize As Long           'The size of the particles
    Direction As Single         'Direction the effect is animating in
    PreviousFrame As Long       'Tick time of the last frame
    ParticleCount As Integer    'Number of particles total
    ParticlesLeft As Integer    'Number of particles left - only for non-repetitive effects
    PartVertex() As TLVERTEX    'Used to point render particles
    Particles() As EngineParticle   'Information on each particle
End Type
Private MaxParticles As Long    'Maximum number of particles allowed at once
Private NumParticles As Long    'Current number of active particles
Private NumEffects As Long      'Index of the highest used effect
Private NumFree As Long         'Number of free slots in the Effects() array
Private EffectsUBound As Long   'UBound of the Effects() array
Private Effects() As Effect     'List of all the active effects

'Particle effect textures
Private NumPartTextures As Long             'Number of particle effect textures
Private PartTexture() As Direct3DTexture8   'All of the individual textures

'Effect rendering locations
Public Const EFFECTRENDER_MAPBACK As Byte = 0   'Render after the map tiles in the back
Public Const EFFECTRENDER_MAPFRONT As Byte = 1  'Render after the map tiles in the front

'Effect numbers
Public Const EffectNum_Fire As Long = 1

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)

Public Sub Effect_Init()
'*********************************************************************************
'Creates the particle engine
'*********************************************************************************

    'Get the settings
    MaxParticles = Val(IO_INI_Read(App.Path & "\Data\Settings.ini", "GRAPHICS", "MaxParticles"))
    NumPartTextures = Val(IO_INI_Read(App.Path & "\Data\Grh.ini", "GRAPHICS", "NumParticles"))
    EffectsUBound = -1

    'Set the initial size of the array
    Effect_IncreaseArray
    
    'Load the texture information
    Effect_LoadTextures
    
    '//!!
    Dim i As Long
    i = Effect_Create(500, 1535, 1, 200, EffectNum_Fire)
    Effect_SetAdjustToScreen i, True

End Sub

Public Sub Effect_Destroy()
'*********************************************************************************
'Unloads the particle effects engine
'*********************************************************************************
Dim i As Long

    'Unload the textures
    For i = 0 To NumPartTextures
        Set PartTexture(i) = Nothing
    Next i
    
    'Unload the effects array
    Erase Effects

End Sub

Private Sub Effect_LoadTextures()
'*********************************************************************************
'Loads all of the textures for the particle effects
'*********************************************************************************
Dim FilePath As String
Dim i As Long
    
    'Resize the texture array
    ReDim PartTexture(0 To NumPartTextures)
    
    'Loop through all the textures
    For i = 0 To NumPartTextures
        
        'Set the file path to the encrypted texture
        FilePath = App.Path & "\Graphics\p" & i & ".png"
        If Not IO_FileExist(FilePath) Then
    
            'If the encrypted version was not found, look for the decrypted version
            FilePath = FilePath & "e"
            If Not IO_FileExist(FilePath) Then
                
                'Neither encrypted nor decrypted file could be found
                Log "File for particle texture " & i & " could not be found"
                GoTo NextI
            
            End If
        End If
    
        'Load the texture
        Graphics_CreateTexture PartTexture(i), FilePath
    
NextI:
    
    Next i

End Sub

Private Sub Effect_IncreaseArray()
'*********************************************************************************
'Increases the size of the effects array
'*********************************************************************************

    'Allocate room for 5 more particle effects
    EffectsUBound = EffectsUBound + 5
    ReDim Preserve Effects(0 To EffectsUBound)
    
    'We just made 5 slots, so we have 5 free slots
    NumFree = NumFree + 5

End Sub

Private Function Effect_NextOpenSlot() As Integer
'*********************************************************************************
'Returns the next open effect slot
'*********************************************************************************
Dim i As Long

    'Check if theres any slots free already
    If NumFree = 0 Then
        
        'Allocate memory since we must have run out
        Effect_IncreaseArray
        
    End If
    
    'Find the free slot
    For i = 0 To EffectsUBound
        If Not Effects(i).Used Then
            Effect_NextOpenSlot = i
            Exit For
        End If
    Next i
    
    'Check if the index returned is highest than that which is currently the highest index
    If Effect_NextOpenSlot > NumEffects Then NumEffects = Effect_NextOpenSlot

End Function

Public Sub Effect_Render(ByVal RenderTime As Byte)
'*********************************************************************************
'Draws the effects based off of their rendering time
'*********************************************************************************
Dim i As Long

    'Loop through the effects
    For i = 0 To NumEffects
        
        'Check if the effect is in use
        If Effects(i).Used Then
            
            'Check if the render time matches
            If Effects(i).RenderTime = RenderTime Then
                
                'Update the effect
                Effect_Update i
                
                'Render the effect
                Effect_RenderEffect i
                
            End If
            
        End If
        
    Next i
    
End Sub

Private Sub Effect_RenderEffect(ByVal EffectIndex As Integer)
'*********************************************************************************
'Draws the particle effect
'*********************************************************************************
    
    'Confirm the effect is in use
    If Not Effects(EffectIndex).Used Then Exit Sub
    
    'Set the render states
    Graphics_SetRenderState D3DRS_POINTSIZE, Effects(EffectIndex).FloatSize
    If Effects(EffectIndex).BlendOne Then Graphics_SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    
    'Set the texture
    Graphics_SetTextureEX PartTexture(Effects(EffectIndex).Gfx)
    
    'Draw the effect
    Graphics_DrawParticles Effects(EffectIndex).PartVertex(), Effects(EffectIndex).ParticleCount
    
    'Return the render state back to normal
    If Effects(EffectIndex).BlendOne Then Graphics_SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

End Sub

Public Sub Effect_SetAdjustToScreen(ByVal EffectIndex As Long, ByVal Value As Boolean)
'*********************************************************************************
'Set if an effect adjusts itself to the screen
'*********************************************************************************

    Effects(EffectIndex).AdjustToScreen = Value

End Sub

Public Sub Effect_SetRenderTime(ByVal EffectIndex As Long, ByVal Value As Byte)
'*********************************************************************************
'Set the time an effect is rendered
'*********************************************************************************

    Effects(EffectIndex).RenderTime = Value

End Sub

Public Sub Effect_SetLife(ByVal EffectIndex As Long, ByVal Life As Long)
'*********************************************************************************
'Set the life of an effect
'*********************************************************************************

    If Life < 1 Then
        Effects(EffectIndex).Life = -1
    Else
        Effects(EffectIndex).Life = timeGetTime + Life
    End If

End Sub

Public Function Effect_GetLife(ByVal EffectIndex As Long) As Long
'*********************************************************************************
'Get the life of an effect
'*********************************************************************************

    Effect_GetLife = Effects(EffectIndex).Life

End Function

Public Sub Effect_SetSize(ByVal EffectIndex As Long, ByVal Size As Single)
'*********************************************************************************
'Set the size of an effect
'*********************************************************************************

    Effects(EffectIndex).FloatSize = Graphics_FToDW(Size)

End Sub

Public Function Effect_Create(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, ByVal EffectType As Long, Optional ByVal Direction As Integer = 180) As Integer
'*********************************************************************************
'Creates an effect
'*********************************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot

    'Auto-adjust to keep within the particles limit
    If NumParticles + Effects(EffectIndex).ParticleCount > MaxParticles Then
        Particles = MaxParticles - NumParticles
        If Particles <= 0 Then
            Effect_Create = -1
            Exit Function
        End If
    End If
    
    With Effects(EffectIndex)
        
        'Set the number of particles left to the total avaliable
        .ParticlesLeft = Particles
        NumParticles = NumParticles + Particles
    
        'Return the index of the used slot
        Effect_Create = EffectIndex
    
        'Set the effect's variables
        .EffectType = EffectType        'Set the effect number
        .ParticleCount = Particles      'Set the number of particles
        .Used = True                    'Enabled the effect
        .X = X                          'Set the effect's X coordinate
        .Y = Y                          'Set the effect's Y coordinate
        .Gfx = Gfx                      'Set the graphic
        .Life = -1                      'Default to infinite life
        .BlendOne = True                'Default to BLENDONE rendering
        .RenderTime = EFFECTRENDER_MAPBACK  'Default to rendering in front of the map background
                      
        
        'Size of the particles
        .FloatSize = Graphics_FToDW(15)
    
        'Redim the number of particles
        ReDim .Particles(0 To Particles)
        ReDim .PartVertex(0 To Particles)
    
        'Create the particles
        For LoopC = 0 To .ParticleCount
            Set .Particles(LoopC) = New EngineParticle
            .Particles(LoopC).Used = True
            .PartVertex(LoopC).Rhw = 1
            Effect_Reset EffectIndex, LoopC
        Next LoopC
    
        'Set The Initial Time
        .PreviousFrame = timeGetTime

    End With

End Function

Private Sub Effect_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*********************************************************************************
'Sets the particle's position
'*********************************************************************************

    With Effects(EffectIndex)
        Select Case .EffectType
            
        'Fire
        Case EffectNum_Fire
    
            .Particles(Index).ResetColor 1, 0.2, 0.2, 0.4 + (Rnd * 0.2), 0.03 + (Rnd * 0.07)
            .Particles(Index).ResetIt .X - 10 + Rnd * 20, .Y - 10 + Rnd * 20, _
                -Sin((.Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, _
                Cos((.Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, _
                0, 0
        
        End Select
    End With
    
End Sub

Private Sub Effect_Update(ByVal EffectIndex As Integer)
'*********************************************************************************
'Update the effect
'*********************************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long
Dim EffectDieing As Boolean

    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effects(EffectIndex).PreviousFrame) * 0.01
    Effects(EffectIndex).PreviousFrame = timeGetTime

    'Check if the effect is dieing
    If Effects(EffectIndex).Life <> -1 Then
        If Effects(EffectIndex).Life < timeGetTime Then EffectDieing = True
    End If
    
    'Go through the particle loop
    For LoopC = 0 To Effects(EffectIndex).ParticleCount
    
        With Effects(EffectIndex).Particles(LoopC)
    
            'Check if particle is in use
            If .Used Then
    
                'Update the particle
                .UpdateParticle ElapsedTime
    
                'Check if the particle is ready to die
                If .sngA <= 0 Then
    
                    'Check if the effect is ending
                    If Not EffectDieing Then
    
                        'Reset the particle
                        Effect_Reset EffectIndex, LoopC
    
                    Else
    
                        'Disable the particle
                        .Used = False
    
                        'Subtract from the total particle count
                        Effects(EffectIndex).ParticlesLeft = Effects(EffectIndex).ParticlesLeft - 1
                        NumParticles = NumParticles - 1

                        'Check if the effect is out of particles
                        If Effects(EffectIndex).ParticlesLeft = 0 Then Exit For
    
                        'Clear the color (dont leave behind any artifacts)
                        Effects(EffectIndex).PartVertex(LoopC).Color = 0
    
                    End If
    
                Else
    
                    'Set the particle information on the particle vertex
                    Effects(EffectIndex).PartVertex(LoopC).Color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
                    If Effects(EffectIndex).AdjustToScreen Then
                        Effects(EffectIndex).PartVertex(LoopC).X = .sngX - ScreenX
                        Effects(EffectIndex).PartVertex(LoopC).Y = .sngY - ScreenY
                    Else
                        Effects(EffectIndex).PartVertex(LoopC).X = .sngX
                        Effects(EffectIndex).PartVertex(LoopC).Y = .sngY
                    End If
                    
                End If
    
            End If
            
        End With

    Next LoopC
    
    'If the effect is no longer used, then erase it
    If Effects(EffectIndex).ParticlesLeft = 0 Then Effect_Kill EffectIndex

End Sub

Private Sub Effect_Kill(ByVal EffectIndex As Long)
'*********************************************************************************
'Set the effect as no longer used
'*********************************************************************************

    'Make sure the effect is even in use
    If Not Effects(EffectIndex).Used Then Exit Sub
    
    'Erase the effect's information from memory
    Erase Effects(EffectIndex).PartVertex
    Erase Effects(EffectIndex).Particles
    
    'Set the effect as unused
    Effects(EffectIndex).Used = False
    
    'Increase the count of number of free effects
    NumFree = NumFree + 1

    'If this effect was the highest one used, then find the next highest index in use
    Do While Effects(NumEffects).Used = True And NumEffects > 0
        NumEffects = NumEffects - 1
    Loop
        
End Sub
