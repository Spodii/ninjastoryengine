Attribute VB_Name = "EngineMath"
'*********************************************************************************
'Handles a lot of the common math routines
'*********************************************************************************

Option Explicit

Public Function Math_Collision_PointRect(ByVal PointX As Long, ByVal PointY As Long, ByVal RectX As Long, ByVal RectY As Long, _
    ByVal RectWidth As Long, ByVal RectHeight As Long) As Boolean
'*********************************************************************************
'Checks if a point is located inside of a rectangle
'*********************************************************************************

    If PointX >= RectX Then
        If PointX <= RectX + RectWidth Then
            If PointY >= RectY Then
                If PointY <= RectY + RectHeight Then
                    Math_Collision_PointRect = True
                End If
            End If
        End If
    End If
    
End Function

Public Function Math_Distance(ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Single
'*********************************************************************************
'Distance formula, what else!? :)
'*********************************************************************************

    Math_Distance = Sqr(((Y1 - Y2) ^ 2 + (x1 - x2) ^ 2))

End Function

Public Function Math_Random(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
'*********************************************************************************
'Find a random number between a fixed range
'*********************************************************************************

    Math_Random = Fix((UpperBound - LowerBound + 1) * Rnd) + LowerBound
    
End Function

Public Function Math_GetAngle(ByVal CenterX As Long, ByVal CenterY As Long, ByVal TargetX As Long, ByVal TargetY As Long) As Single
'*********************************************************************************
'Finds the angle (in degrees) between two points
'*********************************************************************************
Dim SideA As Single
Dim SideC As Single

    On Error GoTo ErrOut

    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = TargetY Then

        'Check for going right (90 degrees)
        If CenterX < TargetX Then
            Math_GetAngle = 90

            'Check for going left (270 degrees)
        Else
            Math_GetAngle = 270
        End If

        'Exit the function
        Exit Function

    End If

    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = TargetX Then

        'Check for going up (360 degrees)
        If CenterY > TargetY Then
            Math_GetAngle = 360

            'Check for going down (180 degrees)
        Else
            Math_GetAngle = 180
        End If

        'Exit the function
        Exit Function

    End If

    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)

    'Side B = CenterY

    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)

    'Calculate the angle
    Math_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Math_GetAngle = (Atn(-Math_GetAngle / Sqr(-Math_GetAngle * Math_GetAngle + 1)) + 1.5708) * 57.29583

    'If the angle is >180, subtract from 360
    If TargetX < CenterX Then Math_GetAngle = 360 - Math_GetAngle

    'Exit function

Exit Function

    'Check for error
ErrOut:

    'Return a 0 saying there was an error
    Math_GetAngle = 0

Exit Function

End Function

Private Function Math_Between(ByVal Value As Single, ByVal Bound1 As Single, ByVal Bound2 As Single) As Byte
'*********************************************************************************
'Find if a value is between two other values (used for line collision)
'*********************************************************************************

    'Checks if a value lies between two bounds
    If Bound1 > Bound2 Then
        If Value >= Bound2 Then
            If Value <= Bound1 Then Math_Between = 1
        End If
    Else
        If Value >= Bound1 Then
            If Value <= Bound2 Then Math_Between = 1
        End If
    End If
    
End Function

Public Function Math_Collision_Line(ByVal L1X1 As Long, ByVal L1Y1 As Long, ByVal L1X2 As Long, ByVal L1Y2 As Long, ByVal L2X1 As Long, ByVal L2Y1 As Long, ByVal L2X2 As Long, ByVal L2Y2 As Long) As Byte
'*********************************************************************************
'Check if two lines intersect (return 1 if true)
'*********************************************************************************
Dim m1 As Single
Dim M2 As Single
Dim B1 As Single
Dim B2 As Single
Dim IX As Single

    'This will fix problems with vertical lines
    If L1X1 = L1X2 Then L1X1 = L1X1 + 1
    If L2X1 = L2X2 Then L2X1 = L2X1 + 1

    'Find the first slope
    m1 = (L1Y2 - L1Y1) / (L1X2 - L1X1)
    B1 = L1Y2 - m1 * L1X2

    'Find the second slope
    M2 = (L2Y2 - L2Y1) / (L2X2 - L2X1)
    B2 = L2Y2 - M2 * L2X2
    
    'Check if the slopes are the same
    If M2 - m1 = 0 Then
    
        If B2 = B1 Then
            'The lines are the same
            Math_Collision_Line = 1
        Else
            'The lines are parallel (can never intersect)
            Math_Collision_Line = 0
        End If
        
    Else
        
        'An intersection is a point that lies on both lines. To find this, we set the Y equations equal and solve for X.
        'M1X+B1 = M2X+B2 -> M1X-M2X = -B1+B2 -> X = B1+B2/(M1-M2)
        IX = ((B2 - B1) / (m1 - M2))
        
        'Check for the collision
        If Math_Between(IX, L1X1, L1X2) Then
            If Math_Between(IX, L2X1, L2X2) Then Math_Collision_Line = 1
        End If
        
    End If
    
End Function

Public Function Math_Collision_LineRect(ByVal SX As Long, ByVal SY As Long, ByVal SW As Long, ByVal SH As Long, ByVal x1 As Long, ByVal Y1 As Long, ByVal x2 As Long, ByVal Y2 As Long) As Byte
'*********************************************************************************
'Check if a line intersects with a rectangle (returns 1 if true)
'*********************************************************************************

    'Top line
    If Math_Collision_Line(SX, SY, SX + SW, SY, x1, Y1, x2, Y2) Then
        Math_Collision_LineRect = 1
        Exit Function
    End If
    
    'Right line
    If Math_Collision_Line(SX + SW, SY, SX + SW, SY + SH, x1, Y1, x2, Y2) Then
        Math_Collision_LineRect = 1
        Exit Function
    End If

    'Bottom line
    If Math_Collision_Line(SX, SY + SH, SX + SW, SY + SH, x1, Y1, x2, Y2) Then
        Math_Collision_LineRect = 1
        Exit Function
    End If

    'Left line
    If Math_Collision_Line(SX, SY, SX, SY + SW, x1, Y1, x2, Y2) Then
        Math_Collision_LineRect = 1
        Exit Function
    End If

End Function

Function Math_Collision_Rect(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal Width1 As Integer, ByVal Height1 As Integer, ByVal x2 As Integer, ByVal Y2 As Integer, ByVal Width2 As Integer, ByVal Height2 As Integer) As Boolean
'*********************************************************************************
'Check for collision between two rectangles
'*********************************************************************************
 
    If x1 + Width1 >= x2 Then
        If x1 <= x2 + Width2 Then
            If Y1 + Height1 >= Y2 Then
                If Y1 <= Y2 + Height2 Then
                    Math_Collision_Rect = True
                End If
            End If
        End If
    End If
 
End Function
