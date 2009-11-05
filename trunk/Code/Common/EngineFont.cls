VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EngineFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'The vertex array for each character in the font
Private Type CharVA
    Vertex(0 To 5) As TLVERTEX
End Type

'Header information for the font
Private Type VFH
    BitmapWidth As Long         'Size of the bitmap itself
    BitmapHeight As Long
    CellWidth As Long           'Size of the cells (area for each character)
    CellHeight As Long
    BaseCharOffset As Byte      'The character we start from
    CharWidth(0 To 255) As Byte 'The actual factual width of each character
    CharVA(0 To 255) As CharVA
End Type

'General information of the font
Private HeaderInfo As VFH           'Holds the header information
Private RowPitch As Integer         'Number of characters per row
Private RowFactor As Single         'Percentage of the texture width each character takes
Private ColFactor As Single         'Percentage of the texture height each character takes
Private CharHeight As Long          'Height to use for the text

'Font texture information
Private Texture As Direct3DTexture8 'Holds the texture of the text
Private TextureWidth As Long        'Width of the texture
Private TextureHeight As Long       'Height of the texture

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Property Get Height() As Long
'*********************************************************************************
'Returns the height of the text (this is a global value)
'*********************************************************************************

    Height = CharHeight

End Property

Public Function Width(ByVal Text As String) As Long
'*********************************************************************************
'Returns the width of specified text
'*********************************************************************************
Dim b() As Byte
Dim i As Long

    'Check for valid text
    If Text = vbNullString Then Exit Function
    
    'Convert the string into a byte array
    b() = StrConv(Text, vbFromUnicode)
    
    'Loop through the text
    For i = 0 To Len(Text) - 1
    
        'Add up the widths
        Width = Width + HeaderInfo.CharWidth(b(i))
        
    Next i

End Function

Public Sub Draw(ByVal Text As String, ByVal X As Long, ByVal Y As Long, ByVal Color As Long)
'*********************************************************************************
'Draws the text as specified in the parameter
'*********************************************************************************
Dim TempVA() As TLVERTEX
Dim XOffset As Long
Dim YOffset As Long
Dim Lines() As String
Dim Ascii() As Byte
Dim i As Long
Dim j As Long
Dim o As Long   'Array offset

    'Check for a valid text
    If Text = vbNullString Then Exit Sub
    
    'Set the texture
    Graphics_SetTextureEX Texture

    'Break up the text into individual lines separated by vbNewLine
    Lines() = Split(Text, vbNewLine)
    
    'Resize the TempVA to fit every verticy
    ReDim TempVA(0 To Len(Text) * 6)
    
    Dim asdf As Long
    Dim f As Long
    asdf = timeGetTime

    'Loop through the lines
    For i = 0 To UBound(Lines)
        If LenB(Lines(i)) Then
        
            'Set the Y co-ordinate offset
            YOffset = i * CharHeight
            
            'Reset the X co-ordinate offset
            XOffset = 0
        
            'Convert the string into a byte array
            Ascii() = StrConv(Text, vbFromUnicode)
 
            'Loop through each character
            For j = 0 To Len(Lines(i)) - 1

                'Copy from the cached vertex array for this character to the temp vertex array
                CopyMemory TempVA(o + 0), HeaderInfo.CharVA(Ascii(j)).Vertex(0), FVF_Size * 6
                
                'Set up the screen locations
                TempVA(o + 0).X = X + XOffset
                TempVA(o + 0).Y = Y + YOffset
                
                TempVA(o + 1).X = TempVA(o + 1).X + X + XOffset
                TempVA(o + 1).Y = TempVA(o + 0).Y

                TempVA(o + 2).X = TempVA(o + 0).X
                TempVA(o + 2).Y = TempVA(o + 2).Y + TempVA(o + 0).Y

                TempVA(o + 4).X = TempVA(o + 1).X
                TempVA(o + 4).Y = TempVA(o + 2).Y

                'Set the colors
                TempVA(o + 0).Color = Color
                TempVA(o + 1).Color = Color
                TempVA(o + 2).Color = Color
                TempVA(o + 4).Color = Color
                
                'Set the two duplicate verticies
                TempVA(o + 3) = TempVA(o + 1)
                TempVA(o + 5) = TempVA(o + 2)

                'Shift over the the position to render the next character
                XOffset = XOffset + HeaderInfo.CharWidth(Ascii(j))
                
                'Increase the offset
                o = o + 6
                
            Next j
            
        End If
    Next i

    'Draw the rectangle
    Graphics_DrawTriangleList TempVA(), (o \ 3)

End Sub

Public Sub Load(ByVal DatFile As String, ByVal TextureFile As String, Optional ByVal HeightMod As Long = 0)
'*********************************************************************************
'Load a font and all of its information
'HeightMod can be used to modify the CharHeight
'*********************************************************************************
Dim FileNum As Byte
Dim Row As Single
Dim u As Single
Dim v As Single
Dim i As Long

    'Load the texture
    Graphics_CreateTexture Texture, TextureFile, TextureWidth, TextureHeight, D3DColorARGB(255, 255, 0, 255)

    'Load the header information
    FileNum = FreeFile
    Open DatFile For Binary As #FileNum
        Get #FileNum, , HeaderInfo
    Close #FileNum
    
    'Calculate the common values
    CharHeight = HeaderInfo.CellHeight + HeightMod
    RowPitch = HeaderInfo.BitmapWidth \ HeaderInfo.CellWidth
    ColFactor = HeaderInfo.CellWidth / HeaderInfo.BitmapWidth
    RowFactor = HeaderInfo.CellHeight / HeaderInfo.BitmapHeight
    
    'Cache the verticies used to draw the character so we only
    'have to set the color and screen X/Y values
    For i = 0 To 255

        'Set the tU and tV values
        Row = (i - HeaderInfo.BaseCharOffset) \ RowPitch
        u = ((i - HeaderInfo.BaseCharOffset) - (Row * RowPitch)) * ColFactor
        v = Row * RowFactor
        
        'Set up the verticies
        With HeaderInfo.CharVA(i)
            .Vertex(0).Rhw = 1
            .Vertex(0).tU = u
            .Vertex(0).tV = v
            .Vertex(0).X = 0
            .Vertex(0).Y = 0
            .Vertex(0).Z = 0
            
            .Vertex(1).Rhw = 1
            .Vertex(1).tU = u + ColFactor
            .Vertex(1).tV = v
            .Vertex(1).X = HeaderInfo.CellWidth
            .Vertex(1).Y = 0
            .Vertex(1).Z = 0

            .Vertex(2).Rhw = 1
            .Vertex(2).tU = u
            .Vertex(2).tV = v + RowFactor
            .Vertex(2).X = 0
            .Vertex(2).Y = HeaderInfo.CellHeight
            .Vertex(2).Z = 0
            
            .Vertex(4).Rhw = 1
            .Vertex(4).tU = u + ColFactor
            .Vertex(4).tV = v + RowFactor
            .Vertex(4).X = HeaderInfo.CellWidth
            .Vertex(4).Y = HeaderInfo.CellHeight
            .Vertex(4).Z = 0
            
            .Vertex(3) = .Vertex(1)
            .Vertex(5) = .Vertex(2)
        End With
    
    Next i

End Sub

Private Sub Class_Terminate()
'*********************************************************************************
'Unload the font texture
'*********************************************************************************
    
    Set Texture = Nothing

End Sub

