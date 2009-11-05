Attribute VB_Name = "General"
Option Explicit

'ID of the local socket
Public LocalSocketID As Long

'If the server is in high priority mode
Public IsHighPriority As Boolean

'Priority handling functions
Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

'INI I/O functions
Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Sub Main()
'*********************************************************************************
'Entry point for the account server
'*********************************************************************************

    'Check for already loaded
    If App.PrevInstance Then Exit Sub

    'Load the form
    Load frmMain
    frmMain.Hide
    frmMain.Visible = False
    
    'Put the form in the taskbar
    TrayAdd frmMain, frmMain.Caption, MouseMove

    'Create the command IDs
    InitCommonIDs

    'Init the MySQL connection
    MySQL_Init
    
    'Create the conversion and receiving buffer
    Set ConBuf = New ByteBuffer
    Set rBuf = New ByteBuffer
    
    'Create the socket
    CreateSocket

End Sub

Private Sub CreateSocket()
'*********************************************************************************
'Create the socket and open it for connections
'*********************************************************************************
Dim PacketKeys() As String
Dim IsPublic As Boolean
Dim PublicIP As String
Dim Port As Integer

    'Get the settings
    Port = Val(IO_INI_Read(App.Path & "\Server Data\Settings.ini", "ACCOUNT", "Port"))
    IsPublic = (Val(IO_INI_Read(App.Path & "\Server Data\Settings.ini", "GENERAL", "Public")) <> 0)

    'Clear the socket image (cleans up a little RAM)
    frmMain.GOREsock.ClearPicture
    
    'Set the encryption
    GenerateEncryptionKeys PacketKeys
    frmMain.GOREsock.ClearPicture
    frmMain.GOREsock.SetEncryption PacketEncTypeServerIn, PacketEncTypeServerOut, PacketKeys()
    Erase PacketKeys
    
    'Load the socket
    If IsPublic Then
        PublicIP = Trim$(IO_INI_Read(App.Path & "\Server Data\Settings.ini", "GENERAL", "PublicIP"))
        LocalSocketID = frmMain.GOREsock.Listen(PublicIP, Port)
    Else
        LocalSocketID = frmMain.GOREsock.Listen("127.0.0.1", Port)
    End If
    
    'Check for a valid connection
    If frmMain.GOREsock.Address(LocalSocketID) = "-1" Then
        MsgBox "Error binding the socket to the IP address. Please make sure port " & Port & " is not in use." & vbNewLine & _
        "IsPublic is set to " & IsPublic & ". If this value is " & True & " and you are behind a router, make sure you " & vbNewLine & _
        "properly forward the port " & Port & " in your router's configuration.", vbOKOnly
        Unload frmMain
        End
    End If

End Sub

Public Sub Server_Unload()
'*********************************************************************************
'Unloads the server - this may end up being called a few times before it finally
'goes through since GOREsock can be picky on the shutdown
'*********************************************************************************

    'Terminate the socket
    If frmMain.GOREsock.ShutDown <> soxERROR Then
        frmMain.GOREsock.UnHook
        
        'Remove the icon from the system tray
        TrayDelete
        
        'Unload the forms
        Unload frmSettings
        Unload frmMain

        'Finish up
        End
        
    End If

End Sub

Private Sub SetPriority(ByVal High As Boolean)
'*********************************************************************************
'Sets the priority of the server between high and normal
'*********************************************************************************
    
    'Set the new priority
    If Not (IsHighPriority = High) Then
        If High Then
            SetThreadPriority GetCurrentThread, 2
            SetPriorityClass GetCurrentProcess, &H80
        Else
            SetThreadPriority GetCurrentThread, 0
            SetPriorityClass GetCurrentProcess, &H20
        End If
    End If
    
    'Store the priority
    IsHighPriority = High
    
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
