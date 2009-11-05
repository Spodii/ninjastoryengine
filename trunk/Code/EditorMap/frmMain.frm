VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Ninja Story Map Editor"
   ClientHeight    =   10500
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13380
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CD 
      Left            =   120
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer CritTimer 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   720
      Top             =   480
   End
   Begin VB.PictureBox picInfo 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      FillColor       =   &H80000009&
      ForeColor       =   &H80000009&
      Height          =   225
      Left            =   0
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   890
      TabIndex        =   0
      Top             =   10275
      Width           =   13380
      Begin VB.Label FPSLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "FPS: 0"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   12960
         TabIndex        =   5
         ToolTipText     =   "Frames per second"
         Top             =   0
         Width           =   780
      End
      Begin VB.Label PixelLbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(0,0)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   11640
         TabIndex        =   4
         ToolTipText     =   "Pixel the cursor is hovering over"
         Top             =   0
         Width           =   1035
      End
      Begin VB.Label TileLbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(0,0)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10680
         TabIndex        =   3
         ToolTipText     =   "Tile the cursor is hovering over"
         Top             =   0
         Width           =   675
      End
      Begin VB.Line LineFPS 
         X1              =   856
         X2              =   856
         Y1              =   0
         Y2              =   16
      End
      Begin VB.Line LineMouse 
         X1              =   768
         X2              =   768
         Y1              =   0
         Y2              =   16
      End
      Begin VB.Line LineTile 
         X1              =   704
         X2              =   704
         Y1              =   0
         Y2              =   16
      End
      Begin VB.Label MapNameLbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "No Map Loaded"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8520
         TabIndex        =   2
         ToolTipText     =   "Name of your currently loaded map"
         Top             =   0
         Width           =   2010
      End
      Begin VB.Line LineName 
         X1              =   560
         X2              =   560
         Y1              =   0
         Y2              =   16
      End
      Begin VB.Label InfoLbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Information"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Random information of goodie-ness!"
         Top             =   0
         Width           =   930
      End
   End
   Begin EditorMap.ucToolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   13380
      _ExtentX        =   23601
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CritFlashCount As Byte

Private Sub CreateToolbar()
    
    'Load the toolbar settings
    DoEvents
    With Toolbar
        .Initialize 16, True, False, True
        .AddBitmap LoadResPicture("TOOLBARICONS", vbResBitmap), vbMagenta
 
        .AddButton , 7, "New (Ctrl + N)"
        .AddButton , 10, "Load (Ctrl + L)"
        .AddButton , 15, "Save (Ctrl + S)"
        .AddButton , 16, "Save As (Ctrl + A)"
        
        .AddButton , , , eSeparator
        
        .AddButton , 12, "Set Tiles (Ctrl + 1 or F1)"
        .AddButton , 0, "Blocks (Ctrl + 2 or F2)"
        .AddButton , 13, "Floods (Ctrl + 3 or F3)"
        .AddButton , 6, "Tile Info (Ctrl + 4 or F4)"
        .AddButton , 1, "Exits (Ctrl + 5 or F5)"
        .AddButton , 8, "NPCs (Ctrl + 6 or F6)"
        .AddButton , 14, "Particles (Ctrl + 7 or F7)"
        .AddButton , 18, "Sfx (Ctrl + 8 or F8)"
        .AddButton , 4, "Map Info (Ctrl + 9 or F9)"
        
        .AddButton , , , eSeparator
        
        .AddButton , 20, "Toggle Weather (Ctrl + W)"
        .AddButton , 9, "Toggle Characters (Ctrl + C)"
        .AddButton , 3, "Toggle Grid (Ctrl + G)"
        .AddButton , 4, "Toggle Tile Info (Ctrl + I)"
        .AddButton , 21, "Toggle Mini-Map (Ctrl + M)"
        
    End With

End Sub

Private Sub CritTimer_Timer()
    
    'Update the critical information rate
    If InfoLbl.ForeColor = vbRed Then InfoLbl.ForeColor = &H80000008 Else InfoLbl.ForeColor = vbRed
    CritFlashCount = CritFlashCount + 1
    If CritFlashCount > 7 Then
        CritFlashCount = 0
        CritTimer.Enabled = False
    End If

End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Clear the information bar
    SetInfo vbNullString

End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As Long)

    'Forward the call to a public function
    ToolBar_Click Button

End Sub

Private Sub MDIForm_Resize()

    If picInfo.ScaleWidth < 380 Then Exit Sub
    
    FPSLbl.Left = picInfo.ScaleWidth - 56
    LineFPS.x1 = picInfo.ScaleWidth - 64
    LineFPS.x2 = picInfo.ScaleWidth - 64
    PixelLbl.Left = picInfo.ScaleWidth - 144
    LineMouse.x1 = picInfo.ScaleWidth - 152
    LineMouse.x2 = picInfo.ScaleWidth - 152
    TileLbl.Left = picInfo.ScaleWidth - 208
    LineTile.x1 = picInfo.ScaleWidth - 216
    LineTile.x2 = picInfo.ScaleWidth - 216
    MapNameLbl.Left = picInfo.ScaleWidth - 350
    LineName.x1 = picInfo.ScaleWidth - 358
    LineName.x2 = picInfo.ScaleWidth - 358
    InfoLbl.Width = picInfo.ScaleWidth - 374

End Sub

Private Sub MDIForm_Load()
    
    'Load the other forms
    Load frmScreen
    
    'Make a new map
    NewMap
    
    'Ready the toolbar
    CreateToolbar

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    EngineRunning = False

End Sub
