VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GUIChatBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'General information
Private SkinPath As String

'Chat box
Public cX As Integer
Public cY As Integer
Public cWidth As Integer
Public cHeight As Integer
Private cGrh As tGrh

'Input box
Private cInputX As Integer
Private cInputY As Integer
Private cInputWidth As Integer
Private cInputText As String

'History box
Private Const HISTORYSIZE As Integer = 100  'Number of lines held in the history
Private Const HISTORYLINES As Integer = 3   'Number of history lines to draw
Private cHistoryX As Integer
Private cHistoryY As Integer
Private cHistoryWidth As Integer
Private cHistoryText(0 To HISTORYSIZE) As String
Private cHistoryOffset As Integer   'History line offset

Public Property Get X() As Integer
    X = cX
End Property

Public Property Let X(ByVal Value As Integer)
    cX = Value
    If STAYINSCREEN Then
        If cX < 0 Then cX = 0
        If cX + Width > ScreenWidth Then cX = ScreenWidth - Width
    End If
End Property

Public Property Get Y() As Integer
    Y = cY
End Property

Public Property Let Y(ByVal Value As Integer)
    cY = Value
    If STAYINSCREEN Then
        If cY < 0 Then cY = 0
        If cY + Height > ScreenHeight Then cY = ScreenHeight - Height
    End If
End Property

Public Property Get Width() As Integer
    Width = cWidth
End Property

Public Property Get Height() As Integer
    Height = cHeight
End Property

Public Sub LeftClick(ByVal X As Integer, ByVal Y As Integer)
'*********************************************************************************
'Handles left-clicking the window
'*********************************************************************************
Dim HistoryHeight As Integer

    'Check to set the typing input mode
    HistoryHeight = HISTORYLINES * FontDefault.Height
    If Math_Collision_PointRect(X, Y, cX + cHistoryX, cY + cHistoryY - HistoryHeight, cHistoryWidth, HistoryHeight) Or _
        Math_Collision_PointRect(X, Y, cX + cInputX, cY + cInputY, cInputWidth, FontDefault.Height) Then
        IsEnteringChat = True
        Exit Sub
    End If

End Sub

Public Property Get HistoryOffset() As Integer
'*********************************************************************************
'Return the history offset
'*********************************************************************************

    HistoryOffset = cHistoryOffset

End Property

Public Property Let HistoryOffset(ByVal Value As Integer)
'*********************************************************************************
'Set the history offset
'*********************************************************************************

    cHistoryOffset = Value

End Property

Public Property Get InputText() As String
'*********************************************************************************
'Return the input text
'*********************************************************************************

    InputText = cInputText

End Property

Public Property Let InputText(ByVal Value As String)
'*********************************************************************************
'Set the input text
'*********************************************************************************

    cInputText = Value
    If Len(cInputText) > MAXCHATLENGTH Then cInputText = Left$(cInputText, MAXCHATLENGTH)

End Property

Public Sub Load(Optional ByVal GUIName As String = "Default")
'*********************************************************************************
'Load the chatbox
'*********************************************************************************

    'Set the skin path
    SkinPath = App.Path & "\Data\GUI\" & GUIName & ".ini"

    'Set the HUD information
    cX = Val(IO_INI_Read(SkinPath, "CHATBOX", "X"))
    cY = Val(IO_INI_Read(SkinPath, "CHATBOX", "Y"))
    cWidth = Val(IO_INI_Read(SkinPath, "CHATBOX", "Width"))
    cHeight = Val(IO_INI_Read(SkinPath, "CHATBOX", "Height"))
    
    'Set the input
    cInputX = Val(IO_INI_Read(SkinPath, "CHATBOX", "InputX"))
    cInputY = Val(IO_INI_Read(SkinPath, "CHATBOX", "InputY"))
    cInputWidth = Val(IO_INI_Read(SkinPath, "CHATBOX", "InputWidth"))
    
    'Set the history
    cHistoryX = Val(IO_INI_Read(SkinPath, "CHATBOX", "HistoryX"))
    cHistoryY = Val(IO_INI_Read(SkinPath, "CHATBOX", "HistoryY"))
    cHistoryWidth = Val(IO_INI_Read(SkinPath, "CHATBOX", "HistoryWidth"))
    
    'Set the HUD Grh
    Graphics_SetGrh cGrh, Val(IO_INI_Read(SkinPath, "CHATBOX", "BackGrh")), ANIMTYPE_LOOP
    
End Sub

Public Sub AddHistory(ByVal Text As String)
'*********************************************************************************
'Adds a line to the chat history
'*********************************************************************************
Dim i As Long

    'Shift everything up in the array
    For i = HISTORYSIZE To 1 Step -1
        cHistoryText(i) = cHistoryText(i - 1)
    Next i
    
    'Add the new line
    cHistoryText(0) = Text

End Sub

Public Sub Draw()
'*********************************************************************************
'Draw the chatbox
'*********************************************************************************
Dim i As Long

    'Backdrop
    Graphics_DrawGrh cGrh, cX, cY
    
    'History text
    For i = cHistoryOffset To cHistoryOffset + HISTORYLINES - 1
        If i > HISTORYLINES Then Exit For
        FontDefault.Draw cHistoryText(i), cX + cHistoryX, cY + cHistoryY - ((i + 1) * FontDefault.Height), -1
    Next i
    
    'Input text
    FontDefault.Draw cInputText, cX + cInputX, cY + cInputY, -1

End Sub
