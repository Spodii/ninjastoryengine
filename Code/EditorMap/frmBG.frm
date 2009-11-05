VERSION 5.00
Begin VB.Form frmBG 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Background"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   2310
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton SetOpt 
      BackColor       =   &H80000005&
      Caption         =   "Set"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin VB.OptionButton IgnoreOpt 
      BackColor       =   &H80000005&
      Caption         =   "Ignore"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.TextBox GrhTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Text            =   "0"
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox LayerCmb 
      Height          =   315
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grh Index:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Layer:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   435
   End
End
Attribute VB_Name = "frmBG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Long
    
    'Restore old settings
    Me.Left = Val(IO_INI_Read(App.Path & "\Dev\Map Editor Settings.ini", Me.Name, "Left"))
    Me.Top = Val(IO_INI_Read(App.Path & "\Dev\Map Editor Settings.ini", Me.Name, "Top"))
    Me.Visible = Val(IO_INI_Read(App.Path & "\Dev\Map Editor Settings.ini", Me.Name, "Visible"))
    Me.Show
    
    'Set the combo box
    LayerCmb.Clear
    LayerCmb.AddItem "1 (closest)"
    For i = 2 To NumBGLayers - 1
        LayerCmb.AddItem i
    Next i
    LayerCmb.AddItem NumBGLayers & " (farthest)"
    LayerCmb.ListIndex = 0
    
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
