VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connect to Ninja Story"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2895
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   2895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox RemChk 
      Caption         =   "Remember"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton QuitCmd 
      Caption         =   "Quit"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton LoginCmd 
      Caption         =   "Log In"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton ConnectCmd 
      Caption         =   "Connect"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox NameTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox PassTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.ListBox CharsLst 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1200
      ItemData        =   "frmConnect.frx":17D2A
      Left            =   120
      List            =   "frmConnect.frx":17D2C
      TabIndex        =   4
      Top             =   1560
      Width           =   2655
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Characters:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Account:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   645
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CharsLst_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then LoginCmd_Click

End Sub

Private Sub ConnectCmd_Click()
Dim IP As String
Dim Port As Integer

    'Check for valid name
    If Not IsLegalName(NameTxt.Text) Then
        MsgBox "Please enter a valid account name.", vbOKOnly
        NameTxt.SetFocus
        Exit Sub
    End If
    
    'Check for a valid password
    If Not IsLegalPassword(PassTxt.Text) Then
        MsgBox "Please enter a valid password.", vbOKOnly
        PassTxt.SetFocus
        Exit Sub
    End If

    'Log in to the account server
    GettingAccount = True
    LocalSocketID = 0
    If frmMain.GOREsock.ShutDown <> soxERROR Then
        IP = Trim$(IO_INI_Read(App.Path & "\Data\Settings.ini", "GENERAL", "AccountIP"))
        Port = Val(IO_INI_Read(App.Path & "\Data\Settings.ini", "GENERAL", "AccountPort"))
        LocalSocketID = frmMain.GOREsock.Connect(IP, Port)
        If LocalSocketID = -1 Then
            MsgBox "Unable to connect to the account server!" & vbCrLf & "Either the server is down or you are not connected to the internet.", vbOKOnly
        End If
    End If

End Sub

Private Sub Form_Load()

    'Get the frmConnect settings
    NameTxt.Text = IO_INI_Read(App.Path & "\Data\Settings.ini", "CONNECT", "Name")
    PassTxt.Text = IO_INI_Read(App.Path & "\Data\Settings.ini", "CONNECT", "Password")
    RemChk.Value = Val(IO_INI_Read(App.Path & "\Data\Settings.ini", "CONNECT", "Remember"))

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'Save the settings
    If RemChk.Value Then
        IO_INI_Write App.Path & "\Data\Settings.ini", "CONNECT", "Name", NameTxt.Text
        IO_INI_Write App.Path & "\Data\Settings.ini", "CONNECT", "Password", PassTxt.Text
    Else
        IO_INI_Write App.Path & "\Data\Settings.ini", "CONNECT", "Name", vbNullString
        IO_INI_Write App.Path & "\Data\Settings.ini", "CONNECT", "Password", vbNullString
    End If
    IO_INI_Write App.Path & "\Data\Settings.ini", "CONNECT", "Remember", RemChk.Value

    'Check if we're unloading everything
    If Not EngineRunning Then
        Unload frmMain
        Unload Me
        End
    End If

End Sub

Private Sub GOREsockAcct_OnConnection(inSox As Long)

End Sub

Private Sub LoginCmd_Click()
Dim IP As String
Dim Port As Integer

    'Check for valid name
    If Not IsLegalName(NameTxt.Text) Then
        MsgBox "Please enter a valid account name.", vbOKOnly
        NameTxt.SetFocus
        Exit Sub
    End If
    
    'Check for a valid password
    If Not IsLegalPassword(PassTxt.Text) Then
        MsgBox "Please enter a valid password.", vbOKOnly
        PassTxt.SetFocus
        Exit Sub
    End If
    
    'Check for a valid character
    If Not IsLegalName(CharsLst.List(CharsLst.ListIndex)) Then
        MsgBox "Please select a valid character.", vbOKOnly
        CharsLst.SetFocus
        Exit Sub
    End If
    
    'Connect to the game server
    GettingAccount = False
    If frmMain.GOREsock.ShutDown <> soxERROR Then
        LocalSocketID = 0
        IP = Trim$(IO_INI_Read(App.Path & "\Data\Settings.ini", "GENERAL", "GameIP"))
        Port = Val(IO_INI_Read(App.Path & "\Data\Settings.ini", "GENERAL", "GamePort"))
        LocalSocketID = frmMain.GOREsock.Connect(IP, Port)
        If LocalSocketID = -1 Then
            MsgBox "Unable to connect to the game server!" & vbCrLf & "Either the server is down or you are not connected to the internet.", vbOKOnly
        Else
            frmMain.GOREsock.SetOption LocalSocketID, soxSO_TCP_NODELAY, True
        End If
    End If
    
End Sub

Private Sub NameTxt_KeyDown(KeyCode As Integer, Shift As Integer)

    'Press Connect
    If KeyCode = vbKeyReturn Then ConnectCmd_Click

End Sub

Private Sub PassTxt_KeyDown(KeyCode As Integer, Shift As Integer)

    'Press Connect
    If KeyCode = vbKeyReturn Then ConnectCmd_Click

End Sub

Private Sub QuitCmd_Click()

    'Close down
    Unload frmMain
    Unload Me
    End

End Sub
