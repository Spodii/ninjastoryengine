VERSION 5.00
Begin VB.Form frmGraphics 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Set Graphics"
   ClientHeight    =   1200
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   80
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   171
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox BehindChk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Behind"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Value           =   1  'Checked
      Width           =   855
   End
   Begin VB.CheckBox SnapChk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Snap to grid"
      Height          =   255
      Left            =   1275
      TabIndex        =   5
      Top             =   840
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.TextBox GrhIndexTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Text            =   "0"
      Top             =   480
      Width           =   1575
   End
   Begin VB.OptionButton IgnoreOpt 
      BackColor       =   &H80000005&
      Caption         =   "Ignore"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   855
   End
   Begin VB.OptionButton EraseOpt 
      BackColor       =   &H80000005&
      Caption         =   "Erase"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.OptionButton SetOpt 
      BackColor       =   &H80000005&
      Caption         =   "Set"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GrhIndex:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   690
   End
End
Attribute VB_Name = "frmGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub EraseOpt_Click()

    SetOpt.Value = False
    EraseOpt.Value = True
    IgnoreOpt.Value = False

End Sub

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

Private Sub GrhIndexTxt_Change()

    SetOpt_Click

End Sub

Private Sub GrhIndexTxt_Click()

    SetOpt_Click

End Sub

Private Sub GrhIndexTxt_KeyPress(KeyAscii As Integer)

    If Not IsValidValue(KeyAscii) Then KeyAscii = 0

End Sub

Private Sub IgnoreOpt_Click()

    SetOpt.Value = False
    EraseOpt.Value = False
    IgnoreOpt.Value = True

End Sub

Private Sub SetOpt_Click()

    SetOpt.Value = True
    EraseOpt.Value = False
    IgnoreOpt.Value = False

End Sub
