Attribute VB_Name = "EngineAudio"
Option Explicit

Private DS As DirectSound8
Private DSBDesc As DSBUFFERDESC

'Buffers used to hold the sounds being played
Private Const MAXSFX As Integer = 50    'Maximum number of sounds active at once
Private SfxBuffer(0 To MAXSFX) As DirectSoundSecondaryBuffer8

'Buffers used to hold the sound templates (what SfxBuffer indicies loads from)
Private Type tSoundTemplate
    Buffer As DirectSoundSecondaryBuffer8
    LastUsed As Long
    Loaded As Boolean
End Type
Private SfxBufferTemplate() As tSoundTemplate
Private NumSfxTemplates As Integer

'Music variables
Private Const MUSICMAXVOLUME As Long = 100
Private Const MUSICMAXBALANCE As Long = 100
Private Const MUSICMAXSPEED As Long = 226
Private MusicEvent As IMediaEvent
Private MusicControl As IMediaControl
Private MusicPosition As IMediaPosition
Private MusicAudio As IBasicAudio
Private MusicPlaying As Boolean

'The update rate for the sound effects
Private Const SFXUPDATERATE As Long = 500
Private LastSfxUpdateTime As Long

'How frequently we check to see if a template is to be unloaded
Private Const SFXTEMPLATELIFE As Long = 600000  'How long a sfx template stays in memory unused
Private LastTemplateUpdateTime As Long

Public Sub Music_Volume(ByVal Volume As Long)
'*********************************************************************************
'Sets the music's volume
'*********************************************************************************

    'Check if the music is playing
    If Not MusicPlaying Then Exit Sub

    'Set the volume
    If Volume >= MUSICMAXVOLUME Then Volume = MUSICMAXVOLUME
    If Volume <= 0 Then Volume = 0
    MusicAudio.Volume = (Volume * MUSICMAXVOLUME) - 10000

End Sub

Public Sub Music_Loop()
'*********************************************************************************
'Makes sure the music is looping
'*********************************************************************************
    
    'Check if the music is playing
    If Not MusicPlaying Then Exit Sub
    
    'Check if the end has been reached
    If MusicPosition.CurrentPosition >= MusicPosition.StopTime Then
        
        'Set the position back to the start
        MusicPosition.CurrentPosition = 0
        
    End If

End Sub

Public Sub Music_Stop()
'*********************************************************************************
'Stops the music
'*********************************************************************************
    
    If Not MusicPlaying Then Exit Sub
    MusicControl.Stop
    Set MusicControl = Nothing
    Set MusicAudio = Nothing
    Set MusicEvent = Nothing
    Set MusicPosition = Nothing
    MusicPlaying = False
    
End Sub

Public Sub Music_Play(ByVal MusicID As Integer)
'*********************************************************************************
'Plays music
'*********************************************************************************

    'Make sure the file exists
    If Not IO_FileExist(App.Path & "\Music\" & MusicID & ".mp3") Then
        Music_Stop
        Exit Sub
    End If
    
    'Create the music objects
    Set MusicControl = New FilgraphManager
    MusicControl.RenderFile App.Path & "\Music\" & MusicID & ".mp3"
    
    Set MusicAudio = MusicControl
    MusicAudio.Volume = 0
    MusicAudio.Balance = 0
    
    Set MusicEvent = MusicControl
    
    Set MusicPosition = MusicControl
    MusicPosition.Rate = 1
    MusicPosition.CurrentPosition = 0
    
    'Play the music
    MusicControl.Run
    
    MusicPlaying = True
 
End Sub

Public Sub Sfx_Update()
'*********************************************************************************
'Updates the sound effects
'*********************************************************************************
Dim f As CONST_DSBSTATUSFLAGS
Dim i As Long

    If LastSfxUpdateTime + SFXUPDATERATE < timeGetTime Then
        LastSfxUpdateTime = timeGetTime
        
        'Loop through all the active sounds
        For i = 0 To MAXSFX
        
            'Check that the buffer is even set
            If Not SfxBuffer(i) Is Nothing Then
                
                'Check if the sound has ended (must be done last)
                If SfxBuffer(i).GetStatus = DSBSTATUS_TERMINATED Or SfxBuffer(i).GetStatus = 0 Then
                    Set SfxBuffer(i) = Nothing
                End If
                
            End If

        Next i
        
    End If
    
    If LastTemplateUpdateTime + SFXTEMPLATELIFE < timeGetTime Then
        LastTemplateUpdateTime = timeGetTime
    
        'Loop through all the sound templates
        For i = 0 To NumSfxTemplates
            If SfxBufferTemplate(i).Loaded Then
                
                'Check if it is time to unload the template
                If SfxBufferTemplate(i).LastUsed + SFXTEMPLATELIFE < timeGetTime Then
                    Sfx_UnloadTemplate i
                End If
                
            End If
        Next i
    End If
    
End Sub

Public Sub Audio_Init(ByRef TargetForm As Form)
'*********************************************************************************
'Loads up the audio-related information
'*********************************************************************************
Dim i As Long

    'Get the number of templates
    NumSfxTemplates = Val(IO_INI_Read(App.Path & "\Sfx\Sfx.ini", "SFX", "NumSfx"))
    
    'Set the buffer descriptions that we will use
    DSBDesc.lFlags = DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
    
    'Create the array for the templates
    ReDim SfxBufferTemplate(0 To NumSfxTemplates)
    
    'Create the DirectSound object
    Set DS = DX.DirectSoundCreate("")
    DS.SetCooperativeLevel TargetForm.hWnd, DSSCL_PRIORITY

End Sub

Private Sub Sfx_UnloadTemplate(ByVal SfxID As Integer)
'*********************************************************************************
'Unloads a sound effects template
'*********************************************************************************

    'Check if the template is already unloaded
    If Not SfxBufferTemplate(SfxID).Loaded Then Exit Sub
    
    'Unload the template
    Set SfxBufferTemplate(SfxID).Buffer = Nothing
    SfxBufferTemplate(SfxID).Loaded = False

End Sub

Private Sub Sfx_LoadTemplate(ByVal SfxID As Integer)
'*********************************************************************************
'Loads a sound template / confirms the template is loaded
'*********************************************************************************

    'Check if the template is loaded alrady
    If SfxBufferTemplate(SfxID).Loaded Then Exit Sub

    'Confirm the file exists
    If IO_FileExist(App.Path & "\Sfx\" & SfxID & ".wav") Then

        'Set the template buffer
        Set SfxBufferTemplate(SfxID).Buffer = DS.CreateSoundBufferFromFile(App.Path & "\Sfx\" & SfxID & ".wav", DSBDesc)
    
    End If
    
End Sub

Public Sub Sfx_Stop()
'*********************************************************************************
'Stops all of the active sound effects
'*********************************************************************************
Dim i As Long

    'Loop through all the sounds
    For i = 0 To MAXSFX
    
        'If the sfx is playing, stop it
        If Not SfxBuffer(i) Is Nothing Then
            SfxBuffer(i).Stop
            Set SfxBuffer(i) = Nothing
        End If
    
    Next i
    
End Sub

Public Sub Sfx_Play(ByVal SfxID As Integer, Optional ByVal PlayFlags As CONST_DSBPLAYFLAGS = DSBPLAY_DEFAULT)
'*********************************************************************************
'Adds a sound to the sound buffer and plays it
'*********************************************************************************
Dim BufferIndex As Integer

    'Confirm the Sfx ID
    If SfxID > NumSfxTemplates Then Exit Sub
    Sfx_LoadTemplate SfxID
    If SfxBufferTemplate(SfxID).Buffer Is Nothing Then Exit Sub

    'Find a free buffer index
    BufferIndex = 0
    Do While Not SfxBuffer(BufferIndex) Is Nothing
        BufferIndex = BufferIndex + 1
        If BufferIndex > MAXSFX Then
            Exit Sub    'No free buffers, can't create the sound
        End If
    Loop
    
    'Play the sound
    Set SfxBuffer(BufferIndex) = DS.DuplicateSoundBuffer(SfxBufferTemplate(SfxID).Buffer)
    SfxBuffer(BufferIndex).SetCurrentPosition 0
    SfxBuffer(BufferIndex).SetPan 0
    SfxBuffer(BufferIndex).SetVolume 0
    SfxBuffer(BufferIndex).Play PlayFlags
    
End Sub
