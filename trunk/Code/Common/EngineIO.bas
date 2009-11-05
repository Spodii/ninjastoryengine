Attribute VB_Name = "EngineIO"
'*********************************************************************************
'Handles all the input and output of files, databases, etc
'*********************************************************************************

Option Explicit

'File index for the log file
Private LogFileNum As Byte

'INI I/O functions
Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long
Public Declare Function MakeFilePath Lib "imagehlp.dll" Alias "MakeSureDirectoryPathExists" (ByVal lpPath As String) As Long

Public Sub Log_Open(ByVal LogFile As String)
'*********************************************************************************
'Open the log file
'*********************************************************************************

    'Check if the log file is already open
    If LogFileNum > 0 Then Exit Sub

    'Delete the old file
    If IO_FileExist(LogFile) Then Kill LogFile

    'Get the file number
    LogFileNum = FreeFile
    
    'Open the file
    Open LogFile For Binary Access Write As #LogFileNum

End Sub

Public Sub Log_Close()
'*********************************************************************************
'Close the log file
'*********************************************************************************
    
    'Close the log file
    Close #LogFileNum
    
    'Set the file number down to 0, telling us it isn't loaded
    LogFileNum = 0
    
End Sub

Public Sub Log(ByVal LogInfo As String)
'*********************************************************************************
'Write to the log file
'*********************************************************************************
    
    'Check if the log file is open
    If LogFileNum = 0 Then Exit Sub
    
    'Write to the file
    Put #LogFileNum, , vbNewLine & LogInfo

End Sub

Public Sub IO_LoadEncryptedTexture(ByVal FilePath As String, ByRef Data() As Byte)
'*********************************************************************************
'Loads an encrypted texture file into memory and decrypts it
'*********************************************************************************
Dim FileNum As Byte

    'Get the file number
    FileNum = FreeFile
    
    'Make sure the file exists
    If Not IO_FileExist(FilePath) Then
        Log "Error! File " & FilePath & " not found!"
        Exit Sub
    End If
    
    'Open the file
    Open FilePath For Binary Access Read As #FileNum
        
        'Resize the array
        ReDim Data(0 To LOF(FileNum) - 1)
        
        'Get the data
        Get #FileNum, , Data()
        
    Close #FileNum
    
    'Decrypt the data
    Encryption_RC4_DecryptByte Data(), "as!JKLmxvc2341zsdfasfd!#@)(*"

End Sub

Public Sub IO_ClientItems_Save(ByVal HighestIndex As Integer, ByRef pItems() As tClientItem)
'*********************************************************************************
'Saves all of the item templates to a single file that the client can read
'*********************************************************************************
Dim FileNum As Byte

    'Get the file number
    FileNum = FreeFile
    
    'Make sure the file does not already exist
    If IO_FileExist(App.Path & "\Data\Items.dat") Then Kill App.Path & "\Data\Items.dat"
    
    'Open the file
    Open App.Path & "\Data\Items.dat" For Binary Access Write As #FileNum
    
        'Save the highest index first
        Put #FileNum, , HighestIndex
        
        'Save the items array
        Put #FileNum, , pItems()
    
    Close #FileNum

End Sub

Public Sub IO_ClientNPCs_Save(ByVal HighestIndex As Integer, ByRef pNPCs() As tClientNPC)
'*********************************************************************************
'Saves all of the NPC templates to a single file that the client can read
'*********************************************************************************
Dim FileNum As Byte

    'Get the file number
    FileNum = FreeFile
    
    'Make sure the file does not already exist
    If IO_FileExist(App.Path & "\Data\NPCs.dat") Then Kill App.Path & "\Data\NPCs.dat"
    
    'Open the file
    Open App.Path & "\Data\NPCs.dat" For Binary Access Write As #FileNum
    
        'Save the highest index first
        Put #FileNum, , HighestIndex
        
        'Save the NPCs array
        Put #FileNum, , pNPCs()
    
    Close #FileNum

End Sub

Public Sub IO_ClientItems_Load(ByRef HighestIndex As Integer, ByRef pItems() As tClientItem)
'*********************************************************************************
'Loads all of the NPC template information from the NPCs.dat file
'*********************************************************************************
Dim bDecrypting As Boolean
Dim FileNum As Byte

    'Get the file number
    FileNum = FreeFile
    
    'Make sure the file exists
    If Not IO_FileExist(App.Path & "\Data\Items.dat") Then
        
        'If the .dat file does not exist, check if the .edat file, an encrypted version
        'of the .dat file, exists
        If Not IO_FileExist(App.Path & "\Data\Items.edat") Then
            Log "Error! File " & App.Path & "\Data\Items.dat not found!"
            Exit Sub
        Else
            
            'Decrypt the .edat file
            bDecrypting = True
            Encryption_Twofish_DecryptFile App.Path & "\Data\Items.edat", App.Path & "\Data\Items.dat", "$#34098FDS:JLka12asfZASDASDafswd"
        
        End If
            
    End If
    
    'Open the file
    Open App.Path & "\Data\Items.dat" For Binary Access Read As #FileNum
    
        'Get the highest index
        Get #FileNum, , HighestIndex
        
        'Resize the array to fit all of the items
        ReDim pItems(0 To HighestIndex)
        
        'Get all of the item information
        Get #FileNum, , pItems()
    
    Close #FileNum
    
    'Kill the decrypted file if we made it from the encrypted version
    If bDecrypting Then
        Kill App.Path & "\Data\Items.dat"
    End If

End Sub

Public Sub IO_ClientNPCs_Load(ByRef HighestIndex As Integer, ByRef pNPCs() As tClientNPC)
'*********************************************************************************
'Loads all of the NPC template information from the NPCs.dat file
'*********************************************************************************
Dim bDecrypting As Boolean
Dim FileNum As Byte

    'Get the file number
    FileNum = FreeFile
    
    'Make sure the file exists
    If Not IO_FileExist(App.Path & "\Data\NPCs.dat") Then
    
        'If the .dat file does not exist, check if the .edat file, an encrypted version
        'of the .dat file, exists
        If Not IO_FileExist(App.Path & "\Data\NPCs.edat") Then
            Log "Error! File " & App.Path & "\Data\NPCs.dat not found!"
            Exit Sub
        Else
        
            'Decrypt the .edat file
            bDecrypting = True
            Encryption_Twofish_DecryptFile App.Path & "\Data\NPCs.edat", App.Path & "\Data\NPCs.dat", "ZXC23asdfASDKJL123SGDkl;asdf1234"
        
        End If
        
    End If
    
    'Open the file
    Open App.Path & "\Data\NPCs.dat" For Binary Access Read As #FileNum
    
        'Get the highest index
        Get #FileNum, , HighestIndex
        
        'Resize the array to fit all of the NPCs
        ReDim pNPCs(1 To HighestIndex)
        
        'Get all of the NPC information
        Get #FileNum, , pNPCs()
    
    Close #FileNum

    'Kill the decrypted file if we made it from the encrypted version
    If bDecrypting Then
        Kill App.Path & "\Data\NPCs.dat"
    End If

End Sub

Public Function IO_ServerNPC_Load(ByVal NPCID As Integer, ByRef NPC As tServerNPC) As Boolean
'*********************************************************************************
'Loads a NPC from a binary file (used by the server)
'*********************************************************************************
Dim FileNum As Byte

    'Get the file number
    FileNum = FreeFile
    
    'Make sure the file exists
    If Not IO_FileExist(App.Path & "\Server Data\temp\" & NPCID & ".npc") Then
        Log "Error! File " & App.Path & "\Server Data\temp\" & NPCID & ".npc not found!"
        Exit Function
    End If
    
    'Open the file
    Open App.Path & "\Server Data\temp\" & NPCID & ".npc" For Binary Access Read As #FileNum
    
        'Load the NPC
        Get #FileNum, , NPC
        
    Close #FileNum
    
    'Load successful
    IO_ServerNPC_Load = True
    
End Function

Public Sub IO_ServerNPC_Save(ByVal NPCID As Integer, ByRef NPC As tServerNPC)
'*********************************************************************************
'Saves a NPC to a binary file (used by the server)
'*********************************************************************************
Dim FileNum As Byte

    'Get the file number
    FileNum = FreeFile
    
    'Make sure the file does not already exist
    If IO_FileExist(App.Path & "\Server Data\temp\" & NPCID & ".npc") Then Kill App.Path & "\Server Data\temp\" & NPCID & ".npc"
    
    'Open the file
    Open App.Path & "\Server Data\temp\" & NPCID & ".npc" For Binary Access Write As #FileNum
    
        'Load the NPC
        Put #FileNum, , NPC
        
    Close #FileNum

End Sub

Public Sub IO_PD_Sprite_Load(ByRef PDSprite() As tPDSprite, ByRef NumPDSprite As Byte)
'*********************************************************************************
'Load the sprite paper-dolling information
'*********************************************************************************
Dim i As Byte
    
    'Get the number of paper-dolls for this layer
    NumPDSprite = Val(IO_INI_Read(App.Path & "\Data\Sprite.dat", "GENERAL", "NumSprites"))
    
    'Resize the array
    ReDim PDSprite(0 To NumPDSprite)
    
    'Loop through the layers, grabbing the GrhIndexes
    For i = 1 To NumPDSprite
        With PDSprite(i)
            .Width = Val(IO_INI_Read(App.Path & "\Data\Sprite.dat", i, "Width"))
            .Height = Val(IO_INI_Read(App.Path & "\Data\Sprite.dat", i, "Height"))
            .Stand = Val(IO_INI_Read(App.Path & "\Data\Sprite.dat", i, "Stand"))
            .Walk = Val(IO_INI_Read(App.Path & "\Data\Sprite.dat", i, "Walk"))
            .JumpUp = Val(IO_INI_Read(App.Path & "\Data\Sprite.dat", i, "JumpUp"))
            .JumpDown = Val(IO_INI_Read(App.Path & "\Data\Sprite.dat", i, "JumpDown"))
            .Punch = Val(IO_INI_Read(App.Path & "\Data\Sprite.dat", i, "Punch"))
            .PunchTime = Val(IO_INI_Read(App.Path & "\Data\Sprite.dat", i, "PunchTime"))
            .PunchWidth = Val(IO_INI_Read(App.Path & "\Data\Sprite.dat", i, "PunchWidth"))
            .Hit = Val(IO_INI_Read(App.Path & "\Data\Sprite.dat", i, "Hit"))
            .Death = Val(IO_INI_Read(App.Path & "\Data\Sprite.dat", i, "Death"))
        End With
    Next i
    
End Sub

Public Sub IO_PD_Body_Load(ByRef PDBody() As tPDBody, ByRef NumPDBody As Byte)
'*********************************************************************************
'Load the body paper-dolling information
'*********************************************************************************
Dim i As Byte
    
    'Get the number of paper-dolls for this layer
    NumPDBody = Val(IO_INI_Read(App.Path & "\Data\Body.dat", "GENERAL", "NumBodies"))
    
    'Resize the array
    ReDim PDBody(0 To NumPDBody)
    
    'Loop through the layers, grabbing the GrhIndexes
    For i = 1 To NumPDBody
        With PDBody(i)
            .Width = Val(IO_INI_Read(App.Path & "\Data\Body.dat", i, "Width"))
            .Height = Val(IO_INI_Read(App.Path & "\Data\Body.dat", i, "Height"))
            .Stand = Val(IO_INI_Read(App.Path & "\Data\Body.dat", i, "Stand"))
            .Walk = Val(IO_INI_Read(App.Path & "\Data\Body.dat", i, "Walk"))
            .JumpUp = Val(IO_INI_Read(App.Path & "\Data\Body.dat", i, "JumpUp"))
            .JumpDown = Val(IO_INI_Read(App.Path & "\Data\Body.dat", i, "JumpDown"))
            .Punch = Val(IO_INI_Read(App.Path & "\Data\Body.dat", i, "Punch"))
            .PunchTime = Val(IO_INI_Read(App.Path & "\Data\Body.dat", i, "PunchTime"))
            .PunchWidth = Val(IO_INI_Read(App.Path & "\Data\Body.dat", i, "PunchWidth"))
            .Hit = Val(IO_INI_Read(App.Path & "\Data\Body.dat", i, "Hit"))
            .Death = Val(IO_INI_Read(App.Path & "\Data\Body.dat", i, "Death"))
        End With
    Next i
    
End Sub

Public Sub IO_Map_LoadInfo(ByVal MapIndex As Long, ByRef pMapInfo As tMapInfo)
'*********************************************************************************
'Load a map's info from a file (.nsmi)
'*********************************************************************************
Dim FileNum As Byte
Dim MapNameSize As Byte

    Log "Loading map info for map " & MapIndex

    'Make sure the file exists
    If Not IO_FileExist(App.Path & "\Maps\" & MapIndex & ".nsmi") Then
        Log "Error! File " & App.Path & "\Maps\" & MapIndex & ".nsmi not found!"
        Exit Sub
    End If
    
    'Open the file
    FileNum = FreeFile
    Open App.Path & "\Maps\" & MapIndex & ".nsmi" For Binary Access Read As #FileNum
        
        'Load the map info
        Get #FileNum, , MapNameSize
        pMapInfo.Name = Space$(MapNameSize)
        Get #FileNum, , pMapInfo.Name
        Get #FileNum, , pMapInfo.MusicID
        Get #FileNum, , pMapInfo.TileWidth
        Get #FileNum, , pMapInfo.TileHeight
        Get #FileNum, , pMapInfo.HasFloatingBlocks
        ReDim pMapInfo.TileInfo(0 To pMapInfo.TileWidth, 0 To pMapInfo.TileHeight)
        Get #FileNum, , pMapInfo.TileInfo
        
    Close #FileNum

End Sub

Public Sub IO_Map_SaveInfo(ByVal MapIndex As Long, ByRef pMapInfo As tMapInfo)
'*********************************************************************************
'Save a map's info to a file (.nsmi)
'*********************************************************************************
Dim FileNum As Byte
Dim MapNameSize As Byte

    Log "Saving map info for map " & MapIndex

    'Make sure the file does not already exist
    If IO_FileExist(App.Path & "\Maps\" & MapIndex & ".nsmi") Then Kill App.Path & "\Maps\" & MapIndex & ".nsmi"

    'Open the file
    FileNum = FreeFile
    Open App.Path & "\Maps\" & MapIndex & ".nsmi" For Binary Access Write As #FileNum
    
        'Save the map info
        MapNameSize = Len(pMapInfo.Name)
        Put #FileNum, , MapNameSize
        Put #FileNum, , pMapInfo.Name
        Put #FileNum, , pMapInfo.MusicID
        Put #FileNum, , pMapInfo.TileWidth
        Put #FileNum, , pMapInfo.TileHeight
        Put #FileNum, , pMapInfo.HasFloatingBlocks
        Put #FileNum, , pMapInfo.TileInfo
        
    Close #FileNum

End Sub

Public Sub IO_Map_SaveServerInfo(ByVal MapIndex As Integer, ByRef pSpawnTile() As tTilePos, ByVal pSpawnTileUBound As Integer, ByRef pNPCSpawn() As tMapSpawn, ByVal pNPCSpawnUBound As Integer)
'*********************************************************************************
'Saves a map's server information to a file (.nsms)
'*********************************************************************************
Dim FileNum As Byte

    Log "Saving map server information for map " & MapIndex
    
    'Open the file
    FileNum = FreeFile
    Open App.Path & "\Maps\" & MapIndex & ".nsms" For Binary Access Write As #FileNum
        
        'Put the UBounds
        Put #FileNum, , pSpawnTileUBound
        Put #FileNum, , pNPCSpawnUBound
        
        'Check if there is an array to put
        If pSpawnTileUBound > -1 Then
            
            'Put the array
            Put #FileNum, , pSpawnTile()
            
        End If

        'Same as above but with the spawn list
        If pNPCSpawnUBound > -1 Then
            Put #FileNum, , pNPCSpawn()
        End If
        
    Close #FileNum
    
End Sub

Public Sub IO_Map_LoadServerInfo(ByVal MapIndex As Integer, ByRef pSpawnTile() As tTilePos, ByRef pSpawnTileUBound As Integer, ByRef pNPCSpawn() As tMapSpawn, ByRef pNPCSpawnUBound As Integer)
'*********************************************************************************
'Loads a map's server information from a file (.nsms)
'*********************************************************************************
Dim FileNum As Byte

    Log "Loading map server information for map " & MapIndex
    
    'Make sure the file exists
    If Not IO_FileExist(App.Path & "\Maps\" & MapIndex & ".nsms") Then
        Log "Error! File " & App.Path & "\Maps\" & MapIndex & ".nsms not found!"
        Exit Sub
    End If
    
    'Open the file
    FileNum = FreeFile
    Open App.Path & "\Maps\" & MapIndex & ".nsms" For Binary Access Read As #FileNum
    
        'Get the UBounds
        Get #FileNum, , pSpawnTileUBound
        Get #FileNum, , pNPCSpawnUBound
        
        'Check if there is an array
        If pSpawnTileUBound = -1 Then
        
            'No spawn tiles, so no need for an array
            Erase pSpawnTile
            
        Else
            
            'Resize the array to fit all the tiles
            ReDim pSpawnTile(0 To pSpawnTileUBound)
            
            'Get the tile locations
            Get #FileNum, , pSpawnTile()
            
        End If
        
        'Same thing as above, but with the NPC spawn
        If pNPCSpawnUBound = -1 Then
            Erase pNPCSpawn
        Else
            ReDim pNPCSpawn(0 To pNPCSpawnUBound)
            Get #FileNum, , pNPCSpawn()
        End If
        
    Close #FileNum

End Sub

Public Sub IO_Map_LoadGrhs(ByVal MapIndex As Integer, ByRef pMapGrh() As tMapGrh, ByRef pNumMapGrhs As Long, ByRef pBG() As tBackground, ByRef pBGSizeX As Byte, ByRef pBGSizeY As Byte)
'*********************************************************************************
'Load a map's grhs from a file (.nsmg)
'*********************************************************************************
Dim TempGrhIndex As Long
Dim FileNum As Byte
Dim i As Long
Dim X As Long
Dim Y As Long

    Log "Loading map grhs for map " & MapIndex
    
    'Make sure the file exists
    If Not IO_FileExist(App.Path & "\Maps\" & MapIndex & ".nsmg") Then
        Log "Error! File " & App.Path & "\Maps\" & MapIndex & ".nsmg not found!"
        Exit Sub
    End If
    
    'Open the file
    FileNum = FreeFile
    Open App.Path & "\Maps\" & MapIndex & ".nsmg" For Binary Access Read As #FileNum
    
        'Load the map graphics
        Get #FileNum, , pNumMapGrhs
        If pNumMapGrhs > 0 Then
            ReDim pMapGrh(1 To pNumMapGrhs)
            Get #FileNum, , pMapGrh
        End If
        
        'Load the map backgrounds
        Get #FileNum, , pBGSizeX
        Get #FileNum, , pBGSizeY
        For i = 1 To NumBGLayers
            ReDim pBG(i).Segment(0 To pBGSizeX, 0 To pBGSizeY)
            For X = 0 To pBGSizeX
                For Y = 0 To pBGSizeY
                    Get #FileNum, , TempGrhIndex
                    Graphics_SetGrh pBG(i).Segment(X, Y), TempGrhIndex
                Next Y
            Next X
        Next i

    Close #FileNum
    
End Sub

Public Sub IO_Map_SaveGrhs(ByVal MapIndex As Integer, ByRef pMapGrh() As tMapGrh, ByRef pNumMapGrhs As Long, ByRef pBG() As tBackground, _
    ByRef pBGSizeX As Byte, ByRef pBGSizeY As Byte)
'*********************************************************************************
'Saves a map's grhs to a file (.nsmg)
'*********************************************************************************
Dim FileNum As Byte
Dim i As Long
Dim X As Long
Dim Y As Long

    Log "Saving map grhs for map " & MapIndex

    'Make sure the file does not already exist
    If IO_FileExist(App.Path & "\Maps\" & MapIndex & ".nsmg") Then Kill App.Path & "\Maps\" & MapIndex & ".nsmg"

    'Open the file
    FileNum = FreeFile
    Open App.Path & "\Maps\" & MapIndex & ".nsmg" For Binary Access Write As #FileNum
    
        'Save the map graphics
        Put #FileNum, , pNumMapGrhs
        Put #FileNum, , pMapGrh
        
        'Save the map backgrounds
        Put #FileNum, , pBGSizeX
        Put #FileNum, , pBGSizeY
        For i = 1 To NumBGLayers
            For X = 0 To pBGSizeX
                For Y = 0 To pBGSizeY
                    Put #FileNum, , pBG(i).Segment(X, Y).GrhIndex
                Next Y
            Next X
        Next i
        
    Close #FileNum

End Sub

Public Function IO_INI_Read(ByVal FilePath As String, ByVal Main As String, ByVal Var As String) As String
'*********************************************************************************
'Reads a variable from a INI file and returns it as a string
'*********************************************************************************

    'Create the buffer string and grab the value
    IO_INI_Read = Space$(500)   'The two 500s define the maximum size we will be able to get
    getprivateprofilestring Main, Var, vbNullString, IO_INI_Read, 500, FilePath
    
    'Trim off the empty spaces on the right side of the string
    IO_INI_Read = RTrim$(IO_INI_Read)
    
    'If there is a string left, trim off the right-most character (terminating character)
    If LenB(IO_INI_Read) <> 0 Then IO_INI_Read = Left$(IO_INI_Read, Len(IO_INI_Read) - 1)

End Function

Public Sub IO_INI_Write(ByVal FilePath As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
'*********************************************************************************
'Writes a variable to an INI file under a specified variable and category
'*********************************************************************************

    writeprivateprofilestring Main, Var, Value, FilePath

End Sub

Public Function IO_FileExist(ByVal FilePath As String) As Boolean
'*********************************************************************************
'Returns if a file exists or not
'*********************************************************************************
    
    On Error GoTo ErrOut

    If LenB(Dir$(FilePath, vbNormal)) <> 0 Then IO_FileExist = True

    On Error GoTo 0

Exit Function

'An error will most likely be caused by invalid filenames (those that do not follow the file name rules)
ErrOut:

    IO_FileExist = False

End Function

Public Sub IO_GrhData_Load(ByRef pGrhData() As tGrhData, ByRef pNumGrhs As Long)
'*********************************************************************************
'Load the Grh data from Grh.dat
'*********************************************************************************
Dim GrhIndex As Long
Dim FileNum As Byte

    'Get the number of Grhs
    pNumGrhs = Val(IO_INI_Read(App.Path & "\Data\Grh.ini", "GRAPHICS", "NumGrhs"))
    If pNumGrhs <= 0 Then
        MsgBox "Error retrieving NumGrhs from \Data\Grh.ini!", vbOKOnly
        Exit Sub
    End If
    
    'Resize the GrhData array to fit all the GrhData
    ReDim pGrhData(1 To pNumGrhs)
    
    'Open the file
    Log "Loading Grh.dat"
    If Not IO_FileExist(App.Path & "\Data\Grh.dat") Then
        MsgBox "Unable to find \Data\Grh.dat!", vbOKOnly
        Exit Sub
    End If
    FileNum = FreeFile
    Open App.Path & "\Data\Grh.dat" For Binary Access Read As #FileNum
    
    Do While Not EOF(FileNum)
    
        'Get the Grh index
        Get #FileNum, , GrhIndex
        If GrhIndex = 0 Then Exit Do
        Log "Loading Grh " & GrhIndex
        
        'Load the Grh
        Get #FileNum, , pGrhData(GrhIndex).X
        Get #FileNum, , pGrhData(GrhIndex).Y
        Get #FileNum, , pGrhData(GrhIndex).Width
        Get #FileNum, , pGrhData(GrhIndex).Height
        Get #FileNum, , pGrhData(GrhIndex).TextureID
        Get #FileNum, , pGrhData(GrhIndex).Speed
        Get #FileNum, , pGrhData(GrhIndex).NumFrames
        If pGrhData(GrhIndex).NumFrames > 1 Then
            ReDim pGrhData(GrhIndex).Frames(1 To pGrhData(GrhIndex).NumFrames)
            Get #FileNum, , pGrhData(GrhIndex).Frames
        End If
    
    Loop
    
    'Close the file
    Close #FileNum

End Sub
