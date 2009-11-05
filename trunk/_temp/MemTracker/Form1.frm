VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   4875
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3480
      TabIndex        =   4
      Text            =   "10000"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Text            =   "10000"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Get"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m As MemTracker
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long

Private Sub Command1_Click()
Dim i As Long
Dim j As Long
    
    j = timeGetTime
    For i = Val(Text2.Text) To 1 Step -1
        m.Add 1, i
    Next i
    MsgBox timeGetTime - j

End Sub

Private Sub Command2_Click()
Dim aType() As Byte
Dim aIndex() As Long
Dim i As Long

Dim j As Long
    
    j = timeGetTime
    m.GetOldest Val(Text3.Text), aType(), aIndex()
    MsgBox timeGetTime - j
    
    Text1.Text = vbNullString
    'For i = 0 To UBound(aType)
    '    Text1.Text = Text1.Text & "Type: " & aType(i) & "    Index: " & aIndex(i) & vbNewLine
    'Next i
    
End Sub

Private Sub Command3_Click()

    m.Update 3

End Sub

Private Sub Form_Load()

    Set m = New MemTracker
    timeBeginPeriod 1
    Command1_Click
    Command2_Click
    
End Sub
