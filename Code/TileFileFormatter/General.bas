Attribute VB_Name = "General"
Option Explicit

Sub Main()
Dim bDeleteSource As Boolean
Dim f() As String
Dim i As Long

    'Confirm
    If MsgBox("Are you sure you wish to create the encrypted file formats?" & vbNewLine & vbNewLine & _
        "This process may take a while, depending on the amount of content." & vbNewLine & _
        "You will be notified by a message box when the process is complete.", vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    'Check to delete the source
    'If MsgBox("Do you wish to delete the source files?" & vbNewLine & vbNewLine & _
    '    "This is recommended only for a client-only copy. The server still" & vbNewLine & _
    '    "uses the unecrypted files.", vbYesNo) = vbNo Then
        bDeleteSource = False
    'Else
    '    bDeleteSource = True
    'End If
    
    'Begin encryption
    
    'Textures
    f() = AllFilesInFolders(App.Path & "\Graphics\")
    For i = 0 To UBound(f)
        If isSuffix(f(i), "png") Then
            Encryption_RC4_EncryptFile f(i), f(i) & "e", "as!JKLmxvc2341zsdfasfd!#@)(*"
            If bDeleteSource Then Kill f(i)
        End If
    Next i
    
    'NPCs.dat
    If IO_FileExist(App.Path & "\Data\NPCs.dat") Then
        If IO_FileExist(App.Path & "\Data\NPCs.edat") Then Kill App.Path & "\Data\NPCs.edat"
        Encryption_Twofish_EncryptFile App.Path & "\Data\NPCs.dat", App.Path & "\Data\NPCs.edat", "ZXC23asdfASDKJL123SGDkl;asdf1234"
        If bDeleteSource Then Kill App.Path & "\Data\NPCs.dat"
    End If
    
    'Items.dat
    If IO_FileExist(App.Path & "\Data\Items.dat") Then
        If IO_FileExist(App.Path & "\Data\Items.edat") Then Kill App.Path & "\Data\Items.edat"
        Encryption_Twofish_EncryptFile App.Path & "\Data\Items.dat", App.Path & "\Data\Items.edat", "$#34098FDS:JLka12asfZASDASDafswd"
        If bDeleteSource Then Kill App.Path & "\Data\Items.dat"
    End If
        
    'All done
    MsgBox "File formatting completed.", vbOKOnly

End Sub

Private Function isSuffix(ByVal File As String, ByVal Suffix As String) As Boolean
Dim s() As String

    'Check for a specified suffix
    Suffix = UCase$(Suffix)
    s() = Split(File, ".")
    isSuffix = (UCase$(s(UBound(s))) = Suffix)

End Function

Private Function IO_FileExist(ByVal FilePath As String) As Boolean
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
