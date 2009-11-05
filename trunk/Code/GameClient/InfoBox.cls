VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GUIInfoBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MAXMESSAGES As Integer = 25
Private Const MESSAGELIFE As Long = 8000

'Messages
Private Type tMessage
    Text As String
    Life As Long
    Width As Integer
End Type
Private cMessage(0 To MAXMESSAGES) As tMessage

'General information
Private SkinPath As String

'Location
Private cY As Integer

'Color
Private cR As Byte
Private cG As Byte
Private cB As Byte

Public Sub Load(Optional ByVal GUIName As String = "Default")
'*********************************************************************************
'Load the info box
'*********************************************************************************

    'Set the skin path
    SkinPath = App.Path & "\Data\GUI\" & GUIName & ".ini"

    'Set the HUD information
    cY = Val(IO_INI_Read(SkinPath, "INFOBOX", "Y"))
    cR = Val(IO_INI_Read(SkinPath, "INFOBOX", "Red"))
    cG = Val(IO_INI_Read(SkinPath, "INFOBOX", "Green"))
    cB = Val(IO_INI_Read(SkinPath, "INFOBOX", "Blue"))

End Sub

Public Sub Add(ByVal Message As String)
'*********************************************************************************
'Add a message to the info box
'*********************************************************************************
Dim i As Long

    'Move all the elements in the array up (3 = 2, 2 = 1, 1 = 0, etc)
    For i = MAXMESSAGES To 1 Step -1
        cMessage(i) = cMessage(i - 1)
    Next i
    
    'Add the message to the lowest index
    cMessage(0).Text = Message
    cMessage(0).Life = timeGetTime + MESSAGELIFE
    cMessage(0).Width = FontDefault.Width(Message)
    
End Sub

Public Sub Draw()
'*********************************************************************************
'Updates and draws all of the messages
'*********************************************************************************
Dim i As Long
Dim a As Long

    For i = 0 To MAXMESSAGES
        
        'Check that the message exists
        If cMessage(i).Life <> 0 Then
            
            'Check for removal
            If cMessage(i).Life < timeGetTime Then
                
                'Remove
                cMessage(i).Life = 0
                cMessage(i).Text = vbNullString
                Exit For
                
            Else
                
                'Calculate the alpha
                a = 10 + ((cMessage(i).Life - timeGetTime) * 0.5)
                If a > 255 Then a = 255
                
                'Draw
                FontDefault.Draw cMessage(i).Text, ScreenWidth - cMessage(i).Width, _
                    cY - (FontDefault.Height * (i + 1)), D3DColorARGB(a, cR, cG, cB)
        
            End If
        
        Else
        
            'Since all items have the same life, if we hit one no longer set, the rest won't be set, either
            Exit For
            
        End If
        
    Next i

End Sub
