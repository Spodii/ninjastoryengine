Attribute VB_Name = "General"
Option Explicit

'INI I/O functions
Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Private Function Engine_INI_Read(ByVal FilePath As String, ByVal Main As String, ByVal Var As String) As String
'*********************************************************************************
'Reads a variable from a INI file and returns it as a string
'*********************************************************************************

    'Create the buffer string and grab the value
    Engine_INI_Read = Space$(500)    'The two 500s define the maximum size we will be able to get
    getprivateprofilestring Main, Var, vbNullString, Engine_INI_Read, 500, FilePath
    
    'Trim off the empty spaces on the right side of the string
    Engine_INI_Read = RTrim$(Engine_INI_Read)
    
    'If there is a string left, trim off the right-most character (terminating character)
    If LenB(Engine_INI_Read) <> 0 Then Engine_INI_Read = Left$(Engine_INI_Read, Len(Engine_INI_Read) - 1)

End Function

Private Sub Engine_INI_Write(ByVal FilePath As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
'*********************************************************************************
'Writes a variable to an INI file under a specified variable and category
'*********************************************************************************

    writeprivateprofilestring Main, Var, Value, FilePath

End Sub

Private Function Engine_FileExist(ByVal FilePath As String) As Boolean
'*********************************************************************************
'Returns if a file exists or not
'*********************************************************************************
On Error GoTo ErrOut

    If LenB(Dir$(FilePath, vbNormal)) <> 0 Then Engine_FileExist = True

Exit Function

'An error will most likely be caused by invalid filenames (those that do not follow the file name rules)
ErrOut:

    Engine_FileExist = False

End Function

Sub Main()
'*********************************************************************************
'Create the Grh.dat file with all the Grh information
'*********************************************************************************
Dim HighestGrhIndex As Long
Dim GrhRawFileNum As Byte
Dim GrhDatFileNum As Byte
Dim RetStr As String
Dim s() As String
Dim CurrentGrh As tGrhData
Dim GrhIndex As Long
Dim i As Long

Dim GrhBuf() As Long
Dim GrhBufUBound As Long
Dim i1 As Long
Dim i2 As Long

    'Check for duplicate grh indexes
    i = -1
    GrhBufUBound = -1
    GrhRawFileNum = FreeFile
    Open App.Path & "\Dev\GrhRaw.txt" For Input As #GrhRawFileNum
    Do While Not EOF(GrhRawFileNum)
        Line Input #GrhRawFileNum, RetStr
        GrhIndex = GetGrhIndexFromString(RetStr)
        If GrhIndex > 0 Then
            i = i + 1
            If i > GrhBufUBound Then
                GrhBufUBound = GrhBufUBound + 500
                ReDim Preserve GrhBuf(0 To GrhBufUBound)
            End If
            GrhBuf(i) = GrhIndex
        End If
    Loop
    Close #GrhRawFileNum
    For i1 = 1 To i
        For i2 = i1 + 1 To i
            If GrhBuf(i1) = GrhBuf(i2) Then
                If MsgBox("Duplicate entries of GrhIndex " & GrhBuf(i1) & " has been found. Do you wish to continue compiling?" & vbNewLine & _
                    "Duplicate Grh numbers can lead to graphical display failures and artifacts.", vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
        Next i2
    Next i1
    i = 0
    Erase GrhBuf
    GrhIndex = 0
    
    'Delete any old Grh.dat file
    If Engine_FileExist(App.Path & "\Data\Grh.dat") = True Then Kill App.Path & "\Data\Grh.dat"

    'Open the GrhDat file
    If Engine_FileExist(App.Path & "\Data\Grh.dat") Then Kill App.Path & "\Data\Grh.dat"
    GrhDatFileNum = FreeFile
    Open App.Path & "\Data\Grh.dat" For Binary Access Write As #GrhDatFileNum
    
    'Open the GrhRaw file
    GrhRawFileNum = FreeFile
    Open App.Path & "\Dev\GrhRaw.txt" For Input As #GrhRawFileNum
    
    On Error GoTo ErrHandler
    
    'Loop through the file
    Do While Not EOF(GrhRawFileNum)
        
        'Get the line
        Line Input #GrhRawFileNum, RetStr
        
        'Get the Grh Index
        GrhIndex = GetGrhIndexFromString(RetStr)
        If GrhIndex > 0 Then

            'Check for a new highest grh index
            If GrhIndex > HighestGrhIndex Then HighestGrhIndex = GrhIndex
        
            'Split the rest of the values
            s() = Split(RetStr, "=")
            s() = Split(s(1), "-")
            
            'Check whether it is animated or not
            If Val(s(0)) = 1 Then
                
                'Stationary
                CurrentGrh.TextureID = Val(s(1))
                CurrentGrh.X = Val(s(2))
                CurrentGrh.Y = Val(s(3))
                CurrentGrh.Width = Val(s(4))
                CurrentGrh.Height = Val(s(5))
                If CurrentGrh.X < 0 Then GoTo ErrHandler
                If CurrentGrh.Y < 0 Then GoTo ErrHandler
                If CurrentGrh.Width <= 0 Then GoTo ErrHandler
                If CurrentGrh.Height <= 0 Then GoTo ErrHandler
                
            ElseIf Val(s(0)) > 1 Then
            
                'Animated
                CurrentGrh.NumFrames = Val(s(0))
                ReDim CurrentGrh.Frames(1 To CurrentGrh.NumFrames)
                For i = 1 To CurrentGrh.NumFrames
                    CurrentGrh.Frames(i) = Val(s(i))
                Next i
                CurrentGrh.Speed = Val(s(i)) / 10000
                
            Else
            
                'Error
                GoTo ErrHandler
                
            End If
            
            'Write the Grh data to the Grh.dat file
            Put #GrhDatFileNum, , GrhIndex
            Put #GrhDatFileNum, , CurrentGrh.X
            Put #GrhDatFileNum, , CurrentGrh.Y
            Put #GrhDatFileNum, , CurrentGrh.Width
            Put #GrhDatFileNum, , CurrentGrh.Height
            Put #GrhDatFileNum, , CurrentGrh.TextureID
            Put #GrhDatFileNum, , CurrentGrh.Speed
            Put #GrhDatFileNum, , CurrentGrh.NumFrames
            If CurrentGrh.NumFrames > 1 Then Put #GrhDatFileNum, , CurrentGrh.Frames

        End If
        
    Loop
    
    'Close the files
    Close #GrhRawFileNum
    Close #GrhDatFileNum
    
    'Update the Grh.ini file
    Engine_INI_Write App.Path & "\Data\Grh.ini", "GRAPHICS", "NumGrhs", HighestGrhIndex
    Engine_INI_Write App.Path & "\Data\Grh.ini", "GRAPHICS", "NumTextures", GetNumTextures
    
    'Done
    MsgBox "Grh.dat successfully written!", vbOKOnly
                    
    Exit Sub
                    
ErrHandler:

    MsgBox "Error loading GrhIndex " & GrhIndex, vbOKOnly

End Sub

Private Function GetGrhIndexFromString(ByVal s As String) As Long
'*********************************************************************************
'Looks to see if a line is valid in the Grh entry format, and returns the Grh Index
'from the line if valid (0 is returned for invalid)
'*********************************************************************************
Dim sp() As String

    'Ignore comment lines
    If Left$(s, 2) <> "//" Then
        
        'Check for an equal sign
        If InStr(1, s, "=") > 0 Then
        
            'Check for a delimiter
            If InStr(1, s, "-") > 0 Then
                
                'Split the GrhIndex from the rest of the string
                sp = Split(s, "=")
                If UBound(sp) = 1 Then
                    
                    'Get the value
                    GetGrhIndexFromString = Val(sp(0))
                    If GetGrhIndexFromString < 0 Then GetGrhIndexFromString = 0
                
                End If
                
            End If
        
        End If
        
    End If

End Function

Private Function GetNumTextures() As Long
'*********************************************************************************
'Returns the highest texture file index
'*********************************************************************************
Dim s() As String
Dim s2() As String
Dim i As Long

    'Get a list of all the files
    s = AllFilesInFolders(App.Path & "\Graphics\")
    
    For i = 0 To UBound(s)
        
        'Get the file number
        s2 = Split(s(i), "\")
        s2 = Split(s2(UBound(s2)), ".")
        
        'Check if it is the highest file number
        If Val(s2(0)) > GetNumTextures Then GetNumTextures = Val(s2(0))

    Next i

End Function
