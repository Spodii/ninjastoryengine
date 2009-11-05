VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ByteBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private PutBytePos As Long
Private GetBytePos As Long
Private ByteBuffer() As Byte
Private ByteBufferUbound As Long
Private Const DefaultSize As Long = 256 'Default size of the buffer

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Sub Overflow()

    'Force the buffer to overflow to break the packet reading loop
    GetBytePos = ByteBufferUbound + 10

End Sub

Public Function HasBuffer() As Boolean

    'If there is a buffer or not
    If PutBytePos > 0 Then HasBuffer = True

End Function

Public Sub Clear()

    'Clear all the values so we can use the buffer from the start again
    PutBytePos = 0
    GetBytePos = 0
    ByteBufferUbound = DefaultSize
    ReDim ByteBuffer(ByteBufferUbound)

End Sub

Public Function Get_Buffer_Remainder() As Byte()
Dim b() As Byte

    'Return the remainder of the byte array buffer

    'Check if we already hit the end of the buffer
    If UBound(ByteBuffer) - GetBytePos + 1 = 0 Then
        Get_Buffer_Remainder = b
    Else
        ReDim b(0 To (UBound(ByteBuffer) - GetBytePos + 1))
        CopyMemory b(0), ByteBuffer(GetBytePos), UBound(ByteBuffer) - GetBytePos + 1
        Get_Buffer_Remainder = ByteBuffer
    End If
    
End Function

Public Function Get_Buffer() As Byte()

    'Scale the buffer down
    If ByteBufferUbound <> PutBytePos - 1 Then
        ByteBufferUbound = PutBytePos - 1
        ReDim Preserve ByteBuffer(0 To ByteBufferUbound)
    End If
    
    'Return the byte array buffer
    Get_Buffer = ByteBuffer

End Function

Public Function Get_Byte() As Byte

    'Retrieve a byte from the buffer (1 byte)
    If GetBytePos > ByteBufferUbound Then Exit Function
    CopyMemory Get_Byte, ByteBuffer(GetBytePos), 1
    GetBytePos = GetBytePos + 1

End Function

Public Function Get_Integer() As Integer

    'Retrieve an integer from the buffer (2 bytes)
    If GetBytePos + 1 > ByteBufferUbound Then Exit Function
    CopyMemory Get_Integer, ByteBuffer(GetBytePos), 2
    GetBytePos = GetBytePos + 2

End Function

Public Function Get_Long() As Long

    'Retrieve a long from the buffer (4 bytes)
    If GetBytePos + 3 > ByteBufferUbound Then Exit Function
    CopyMemory Get_Long, ByteBuffer(GetBytePos), 4
    GetBytePos = GetBytePos + 4

End Function

Private Function Get_PutPos() As Long

    'Return the put byte position
    Get_PutPos = PutBytePos

End Function

Public Function Get_ReadPos() As Long

    'Return the read byte position
    Get_ReadPos = GetBytePos

End Function

Public Function Get_String() As String
Dim tempB() As Byte
Dim ArraySize As Byte

    'Retrieve a string from the buffer
    ArraySize = Get_Byte    'Get the size of the string

    'Check for a valid size before sizing the array
    If ArraySize = 0 Then Exit Function

    'Resize the temp byte array to fit the size of the string
    ReDim tempB(ArraySize - 1)

    'Copy the bytes for the string in the buffer to the temp byte array
    CopyMemory tempB(0), ByteBuffer(GetBytePos), ArraySize

    'Convert the byte array to Unicode
    Get_String = StrConv(tempB, vbUnicode)
    GetBytePos = GetBytePos + ArraySize

End Function

Public Function Get_StringEX() As String
Dim tempB() As Byte
Dim ArraySize As Integer

    'Retrieve a very long string from the buffer
    ArraySize = Get_Integer 'Get the size of the string

    'Check for a valid size before sizing the array
    If ArraySize = 0 Then Exit Function

    'Resize the temp byte array to fit the size of the string
    ReDim tempB(ArraySize - 1)

    'Copy the bytes for the string in the buffer to the temp byte array
    CopyMemory tempB(0), ByteBuffer(GetBytePos), ArraySize

    'Convert the byte array to Unicode
    Get_StringEX = StrConv(tempB, vbUnicode)
    GetBytePos = GetBytePos + ArraySize

End Function

Public Sub Put_Byte(ByVal Value As Byte)

    'Store a byte (1 byte)
    If ByteBufferUbound < PutBytePos Then
        ReDim Preserve ByteBuffer(0 To PutBytePos)
        ByteBufferUbound = PutBytePos
    End If
    CopyMemory ByteBuffer(PutBytePos), Value, 1
    PutBytePos = PutBytePos + 1

End Sub

Public Sub Put_Integer(ByVal Value As Integer)

    'Store an integer (2 bytes)
    If ByteBufferUbound < PutBytePos + 1 Then
        ReDim Preserve ByteBuffer(0 To PutBytePos + 1)
        ByteBufferUbound = PutBytePos + 1
    End If
    CopyMemory ByteBuffer(PutBytePos), Value, 2
    PutBytePos = PutBytePos + 2

End Sub

Public Sub Put_Long(ByVal Value As Long)

    'Store a long (4 bytes)
    If ByteBufferUbound < PutBytePos + 3 Then
        ReDim Preserve ByteBuffer(0 To PutBytePos + 3)
        ByteBufferUbound = PutBytePos + 3
    End If
    CopyMemory ByteBuffer(PutBytePos), Value, 4
    PutBytePos = PutBytePos + 4

End Sub

Public Sub Put_String(ByRef Value As String)
Dim tempB() As Byte
Dim i As Long

    'Store a string

    'Check for an empty string
    If Value = vbNullString Then
        Put_Byte 0
        Exit Sub
    End If

    'Cache the UBound
    i = Len(Value) - 1
    
    'Store a byte-long value that represents the size of the string
    If i > 254 Then Exit Sub
    Put_Byte i + 1
    
    'Convert the string to a byte array
    tempB = StrConv(Value, vbFromUnicode)

    'Resize the array to fit the string
    If ByteBufferUbound < PutBytePos + i Then
        ReDim Preserve ByteBuffer(0 To PutBytePos + i)
        ByteBufferUbound = PutBytePos + i
    End If
    
    'Store the byte array of the string into the buffer byte array
    CopyMemory ByteBuffer(PutBytePos), tempB(0), i + 1
    PutBytePos = PutBytePos + i + 1

End Sub

Public Sub Put_StringEX(ByRef Value As String)
Dim tempB() As Byte
Dim i As Long

    'Store a very long string
    
    'Check for an empty string
    If Value = vbNullString Then
        Put_Integer 0
        Exit Sub
    End If
    
    'Cache the UBound
    i = Len(Value) - 1

    'Store a byte-long value that represents the size of the string
    If i > 32760 Then Exit Sub
    Put_Integer i + 1
    
    'Convert the string to a byte array
    tempB = StrConv(Value, vbFromUnicode)
 
    'Resize the array to fit the string
    If ByteBufferUbound < PutBytePos + i Then
        ReDim Preserve ByteBuffer(0 To PutBytePos + i)
        ByteBufferUbound = PutBytePos + i
    End If
    
    'Store the byte array of the string into the buffer byte array
    CopyMemory ByteBuffer(PutBytePos), tempB(0), i + 1
    PutBytePos = PutBytePos + i + 1

End Sub

Public Sub Set_Buffer(ByRef Value() As Byte)

    'Clear the values
    PutBytePos = 0
    GetBytePos = 0

    'Cache the UBound
    ByteBufferUbound = UBound(Value)

    'Set the byte buffer to the size of the array being sent in
    ReDim ByteBuffer(0 To ByteBufferUbound)

    'Copy to the byte buffer
    CopyMemory ByteBuffer(0), Value(0), ByteBufferUbound + 1

End Sub

Private Sub Class_Initialize()
    
    'Set the buffer UBound
    ByteBufferUbound = -1

End Sub
