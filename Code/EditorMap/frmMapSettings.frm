VERSION 5.00
Begin VB.Form frmMapSettings 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Map Settings"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox MusicTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      TabIndex        =   8
      Text            =   "1"
      ToolTipText     =   "ID of the map's music"
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox NameTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      TabIndex        =   6
      ToolTipText     =   "The name of the map"
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton SetCmd 
      Caption         =   "Set"
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      ToolTipText     =   "Apply map size changes"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox HeightTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Text            =   "500"
      ToolTipText     =   "Height of the map in tiles"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox WidthTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Text            =   "500"
      ToolTipText     =   "Width of the map in tiles"
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Music:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Height:"
      Height          =   195
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Width:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   465
   End
End
Attribute VB_Name = "frmMapSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    'Restore old settings
    Me.Left = Val(IO_INI_Read(App.Path & "\Dev\Map Editor Settings.ini", Me.Name, "Left"))
    Me.Top = Val(IO_INI_Read(App.Path & "\Dev\Map Editor Settings.ini", Me.Name, "Top"))
    Me.Visible = Val(IO_INI_Read(App.Path & "\Dev\Map Editor Settings.ini", Me.Name, "Visible"))
    Me.Show
    
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

Private Sub HeightTxt_KeyPress(KeyAscii As Integer)

    If Not IsValidValue(KeyAscii) Then KeyAscii = 0

End Sub

Private Sub MusicTxt_Change()

    MapInfo.MusicID = Val(MusicTxt.Text)

End Sub

Private Sub NameTxt_Change()

    If Len(NameTxt.Text) > 30 Then NameTxt.Text = Left$(NameTxt.Text, 30)
    MapInfo.Name = NameTxt.Text
    frmMain.MapNameLbl.Caption = MapInfo.Name

End Sub

Private Sub SetCmd_Click()
Dim i As Long
Dim w As Long
Dim H As Long

    'Check for valid values
    w = Val(WidthTxt.Text)
    H = Val(HeightTxt.Text)
    If w < 10 Then Exit Sub
    If H < 10 Then Exit Sub
    If w > MAP_MAXSIZE Then Exit Sub
    If H > MAP_MAXSIZE Then Exit Sub

    'Resize the array
    ReDim MapInfo.TileInfo(0 To w, 0 To H)
    
    'Check if the width or height was decreased
    If w < MapInfo.TileWidth Or H < MapInfo.TileHeight Then
        
        'Check if any graphics went out of the map's range
        For i = 1 To NumMapGrhs
            If MapGrh(i).X > w * 32 Or MapGrh(i).Y > H * 32 Then
                MapGrh(i).Grh.GrhIndex = 0
            End If
        Next i
        
        'Clean up the array
        OptimizeMapGrhs
    
    End If
    
    'Set the new sizes
    MapInfo.TileWidth = w
    MapInfo.TileHeight = H
    CalcBGSizes

End Sub

Private Sub WidthTxt_KeyPress(KeyAscii As Integer)

    If Not IsValidValue(KeyAscii) Then KeyAscii = 0

End Sub
