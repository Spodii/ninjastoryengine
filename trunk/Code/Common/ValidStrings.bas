Attribute VB_Name = "ValidStrings"
Option Explicit

'Maximum length of a single chat message
Public Const MAXCHATLENGTH As Long = 75

Public Function IsLegalName(ByVal s As String) As Boolean
'*********************************************************************************
'Check if a string is legal for a name
'*********************************************************************************
    
    'Check the size
    If s = vbNullString Then Exit Function
    If Len(s) > 10 Then Exit Function
    If Len(s) < 3 Then Exit Function
    
    'Check the contents
    If Not IsLegalString(s, True, True, True) Then Exit Function
    
    'All checks out
    IsLegalName = True

End Function

Public Function IsLegalPassword(ByVal s As String) As Boolean
'*********************************************************************************
'Check if a string is legal for a password
'*********************************************************************************
    
    'Check the size
    If s = vbNullString Then Exit Function
    If Len(s) > 10 Then Exit Function
    If Len(s) < 3 Then Exit Function
    
    'Check the contents
    If Not IsLegalString(s, True, True, True) Then Exit Function
    
    'All checks out
    IsLegalPassword = True

End Function

Public Function IsLegalString(ByVal s As String, Optional ByVal AllowNumeric As Boolean = True, _
    Optional ByVal AllowAlphaUpper As Boolean = True, Optional ByVal AllowAlphaLower As Boolean = True) As Boolean
'*********************************************************************************
'Check if a string contains any unwanted characters
'*********************************************************************************
Dim b() As Byte
Dim i As Long
Dim IsNumeric As Boolean
Dim IsLowerCase As Boolean
Dim IsUpperCase As Boolean

    On Error GoTo ErrOut
    
    'Check for invalid string
    If s = vbNullString Then Exit Function

    'Convert the string into a byte array
    b() = StrConv(s, vbFromUnicode)
    
    'Loop through the string and check the values
    For i = 0 To UBound(b)
        
        'Check for numeric
        IsNumeric = Char_IsNumeric(b(i))
        If IsNumeric Then
            If Not AllowNumeric Then
                Exit Function
            End If
        End If
        
        'Check for lowercase
        IsLowerCase = Char_IsAlpha_LowerCase(b(i))
        If IsLowerCase Then
            If Not AllowAlphaLower Then
                Exit Function
            End If
        End If
        
        'Check for uppercase
        IsUpperCase = Char_IsAlpha_UpperCase(b(i))
        If IsUpperCase Then
            If Not AllowAlphaUpper Then
                Exit Function
            End If
        End If
        
        'Check if not numeric, lowercase or uppercase
        If Not IsNumeric Then
            If Not IsLowerCase Then
                If Not IsUpperCase Then
                    Exit Function
                End If
            End If
        End If
        
    Next i
        
    'Valid string
    IsLegalString = True
    
    Exit Function
    
ErrOut:

    'There was some kind of error :(
    IsLegalString = False

End Function

Private Function Char_IsAlpha_LowerCase(ByVal Ascii As Byte) As Boolean
'*********************************************************************************
'Checks if a character is part of the alphabet and lowercase
'*********************************************************************************

    'Characters a to z
    If Ascii >= 97 Then
        If Ascii <= 122 Then
            Char_IsAlpha_LowerCase = True
            Exit Function
        End If
    End If

End Function

Private Function Char_IsAlpha_UpperCase(ByVal Ascii As Byte) As Boolean
'*********************************************************************************
'Checks if a character is part of the alphabet and uppercase
'*********************************************************************************

    'Characters A to Z
    If Ascii >= 65 Then
        If Ascii <= 90 Then
            Char_IsAlpha_UpperCase = True
            Exit Function
        End If
    End If

End Function

Private Function Char_IsNumeric(ByVal Ascii As Byte) As Boolean
'*********************************************************************************
'Checks if a character is numeric
'*********************************************************************************
    
    'Numbers 0 to 9
    If Ascii >= 48 Then
        If Ascii <= 57 Then
            Char_IsNumeric = True
            Exit Function
        End If
    End If

End Function
