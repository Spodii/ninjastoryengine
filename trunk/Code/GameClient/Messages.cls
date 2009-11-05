VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Messages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Message() As String
Private MessageUBound As Integer

Public Sub Load(Optional ByVal Language As String = "English")
'*********************************************************************************
'Load the messages
'*********************************************************************************
Dim f As String
Dim i As Long

    'File path
    f = App.Path & "\Data\Messages.ini"

    'Get the number of messages
    MessageUBound = Val(IO_INI_Read(f, "GENERAL", "NumMessages"))
    ReDim Message(0 To MessageUBound)
    
    'load the individual messages
    For i = 0 To MessageUBound
        Message(i) = Trim$(IO_INI_Read(f, Language, i))
    Next i
    
End Sub

Public Function Grab(ByVal MessageID As Byte)
'*********************************************************************************
'Grab a message
'*********************************************************************************

    Grab = Message(MessageID)

End Function

Public Function GrabReplace(ByVal MessageID As Byte, ByRef Params() As String) As String
'*********************************************************************************
'Grab a message and replace its parameters
'*********************************************************************************
Dim i As Long

    'Grab the message
    GrabReplace = Grab(MessageID)
    
    'Replace the parameters
    For i = 0 To UBound(Params)
        GrabReplace = Replace$(GrabReplace, "<" & i & ">", Params(i))
    Next i

End Function
