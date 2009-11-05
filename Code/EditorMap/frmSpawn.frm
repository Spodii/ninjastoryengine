VERSION 5.00
Begin VB.Form frmSpawn 
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " NPC Spawning"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   3360
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton DelCmd 
      Caption         =   "Delete"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton AddCmd 
      Caption         =   "Add"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox AmountTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox IDTxt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.ListBox SpawnLst 
      Height          =   1035
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NPC ID:"
      Height          =   195
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   585
   End
End
Attribute VB_Name = "frmSpawn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AddCmd_Click()

    'Add an entry
    SpawnLst.AddItem "Free [0]"

End Sub

Private Sub DelCmd_Click()

    'Delete an entry
    If SpawnLst.ListIndex > -1 Then
        SpawnLst.RemoveItem (SpawnLst.ListIndex)
    End If

End Sub

Private Sub AmountTxt_Change()

    UpdateEntry

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

Private Sub IDTxt_Change()

    UpdateEntry

End Sub

Private Sub UpdateEntry()

    'Update the entry
    If SpawnLst.ListIndex > -1 Then
        SpawnLst.List(SpawnLst.ListIndex) = IDTxt.Text & " [" & AmountTxt.Text & "]"
    End If

End Sub

Private Sub SpawnLst_Click()
Dim s() As String

    'Select a new item
    If SpawnLst.ListIndex > -1 Then
        
        'Break apart the NPC ID and the amount
        s() = Split(SpawnLst.List(SpawnLst.ListIndex), " ")
        IDTxt.Text = s(0)
        AmountTxt.Text = Mid$(s(1), 2, Len(s(1)) - 2)
    
    End If

End Sub
