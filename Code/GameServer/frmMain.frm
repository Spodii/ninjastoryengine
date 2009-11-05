VERSION 5.00
Object = "{BA320567-4AFB-40AE-8844-350337894A61}#1.0#0"; "GOREsockServer.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Ninja Story Server"
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1965
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1455
   ScaleWidth      =   1965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer UnloadTmr 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   720
      Top             =   120
   End
   Begin GOREsock.GOREsockServer GOREsock 
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Show the pop-up menu
    If X = 7695 Then
        Load frmSettings
        frmSettings.Show
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Remove the tray icon
    TrayDelete

End Sub

Private Sub GOREsock_OnClose(inSox As Long)

    'Clear off the user
    If inSox > 0 Then
        If inSox <= UserListUBound Then
            If Not UserList(inSox) Is Nothing Then
                
                'Close down the user
                Log "User " & UserList(inSox).Name & " disconnected"
                UserList(inSox).Unload
                Set UserList(inSox) = Nothing
                
                'Check if to size down the LastUser value
                'We only scale down the array if the user that was just closed
                'was the highest index in use
                If inSox = LastUser Then
                
                    'Find the highest index that is in use
                    Do While UserList(LastUser) Is Nothing
                        LastUser = LastUser - 1
                        If LastUser = 0 Then Exit Do
                    Loop
                    
                End If
                            
            End If
        End If
    End If

End Sub

Private Sub GOREsock_OnConnection(inSox As Long)

    'Make sure Nagling is off
    frmMain.GOREsock.SetOption inSox, soxSO_TCP_NODELAY, True
    
    'Check for enough room in the user array
    If inSox > UserListUBound Then
        UserListUBound = inSox
        ReDim Preserve UserList(1 To UserListUBound)
    End If

End Sub

Private Sub GOREsock_OnDataArrival(inSox As Long, inData() As Byte)

    'Check for a valid inSox index
    If inSox > UserListUBound Then Exit Sub

    'Forward the data to the packet handler sub
    Data_Handle inSox, inData()
    
End Sub

Private Sub UnloadTmr_Timer()

    'Constantly tries to close down the server
    Server_Unload

End Sub
