VERSION 5.00
Object = "{BA320567-4AFB-40AE-8844-350337894A61}#1.0#0"; "GOREsockServer.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Ninja Story Account Server"
   ClientHeight    =   1410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2445
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1410
   ScaleWidth      =   2445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer KeepAlive 
      Interval        =   60000
      Left            =   720
      Top             =   120
   End
   Begin VB.Timer UnloadTmr 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   1200
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

Private Type ConnectionList
    LastPacket As Long
End Type
Private LastConnection As Long
Private ConnectionList() As ConnectionList

Private KeepAliveTicks As Long
Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Sub GOREsock_OnConnection(inSox As Long)

    'Resize the array to fit the new connection
    If inSox > LastConnection Then
        LastConnection = LastConnection + 10
        ReDim Preserve ConnectionList(1 To LastConnection)
    End If
    
    'Clear the last packet time
    ConnectionList(inSox).LastPacket = timeGetTime

End Sub

Private Sub GOREsock_OnDataArrival(inSox As Long, inData() As Byte)

    'Check the last packet time
    If ConnectionList(inSox).LastPacket + 5000 < timeGetTime Then Exit Sub

    'Forward the data to the packet handler sub
    Data_Handle inSox, inData()

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    'Show the pop-up menu
    If X = 7695 Then
        Load frmSettings
        frmSettings.Show
    End If

End Sub

Private Sub KeepAlive_Timer()
Dim i As Long

    'Database keep-alive
    KeepAliveTicks = KeepAliveTicks + 1
    If KeepAliveTicks = 10 Then
        MySQL_KeepAlive
        KeepAliveTicks = 0
    End If

    'Remove the sockets that have been open too long (>5 seconds)
    For i = 1 To LastConnection
        If ConnectionList(i).LastPacket <> 0 Then
            If timeGetTime - ConnectionList(i).LastPacket > 5000 Then
                GOREsock.Shut i
            End If
        End If
    Next i
    
End Sub

Private Sub ShutTmr_Timer()

    
End Sub

Private Sub UnloadTmr_Timer()

    'Constantly tries to close down the server
    Server_Unload

End Sub
