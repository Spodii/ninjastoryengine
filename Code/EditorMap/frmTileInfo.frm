VERSION 5.00
Begin VB.Form frmTileInfo 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Tile Info"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   1665
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox FloorTxt 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Text            =   "0"
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton SetCmd 
      Caption         =   "Set"
      Height          =   255
      Left            =   960
      TabIndex        =   4
      Top             =   2160
      Width           =   615
   End
   Begin VB.ListBox AttLst 
      Height          =   1035
      ItemData        =   "frmTileInfo.frx":0000
      Left            =   120
      List            =   "frmTileInfo.frx":0002
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.OptionButton SetOpt 
      BackColor       =   &H80000005&
      Caption         =   "Set"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.OptionButton IgnoreOpt 
      BackColor       =   &H80000005&
      Caption         =   "Ignore"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Floor size:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Attributes:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   705
   End
End
Attribute VB_Name = "frmTileInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AttLst_Click()

    SetOpt_Click

End Sub

Private Sub Form_Load()

    'Restore old settings
    Me.Left = Val(IO_INI_Read(App.Path & "\Dev\Map Editor Settings.ini", Me.Name, "Left"))
    Me.Top = Val(IO_INI_Read(App.Path & "\Dev\Map Editor Settings.ini", Me.Name, "Top"))
    Me.Visible = Val(IO_INI_Read(App.Path & "\Dev\Map Editor Settings.ini", Me.Name, "Visible"))
    Me.Show
    
    'Set the attributes
    AttLst.Clear
    AttLst.AddItem "None"
    AttLst.AddItem "Blocked"
    AttLst.AddItem "Platform"
    AttLst.AddItem "Ladder"
    AttLst.AddItem "Spawn"
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If EngineRunning Then
        
        'Don't close if the application is still running, just hide
        Me.Visible = False
        Cancel = 1
        
    Else
        
        'Save settings
        IO_INI_Write App.Path & "\Dev\Map Editor Settings.ini", Me.Name, "Left", Me.Left
        IO_INI_Write App.Path & "\Dev\Map Editor Settings.ini", Me.Name, "Top", Me.Top
        IO_INI_Write App.Path & "\Dev\Map Editor Settings.ini", Me.Name, "Visible", Me.Visible

    End If

End Sub

Private Sub IgnoreOpt_Click()

    SetOpt.Value = False
    IgnoreOpt.Value = True

End Sub

Private Sub SetCmd_Click()
Dim t As Long
Dim X As Long
Dim Y As Long

    'Get the size
    t = Val(FloorTxt.Text)
    If t < 1 Then
        MsgBox "Invalid floor size.", vbOKOnly
        Exit Sub
    End If

    'Confirm
    If MsgBox("Are you sure you wish to set the floor size to " & t & " tiles?" & vbNewLine & _
        "The size can not be shrunken after being set. Setting this will cover the bottom " & t & " tiles as blocked." _
        , vbYesNo) = vbNo Then
        Exit Sub
    End If
    
    'Set the floor size
    For X = 0 To MapInfo.TileWidth
        For Y = MapInfo.TileHeight - t To MapInfo.TileHeight
            MapInfo.TileInfo(X, Y) = TILETYPE_BLOCKED
        Next Y
    Next X
    HasFloatingBlocks

End Sub

Private Sub SetOpt_Click()

    SetOpt.Value = True
    IgnoreOpt.Value = False

End Sub
