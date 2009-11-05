Attribute VB_Name = "Packets"
Option Explicit

'Conversion buffer
Public ConBuf As ByteBuffer

'Received packet buffer
Public rBuf As ByteBuffer

Public Sub Data_Handle(ByVal inSox As Long, ByRef Data() As Byte)
'*********************************************************************************
'Global routine for handling incoming packets. Packets are forwarded to another
'method corresponding to their PacketID.
'*********************************************************************************
Dim PacketID As Byte

    'Set the buffer
    rBuf.Set_Buffer Data()

    'Get the packet ID
    PacketID = rBuf.Get_Byte
    
    'Forward to the corresponding packet handling method
    With AccountPId
        Select Case PacketID
        
        Case .CS_GetChars: Data_GetChars inSox
        
        Case 0: rBuf.Overflow       'An error occured
        Case Else: rBuf.Overflow    'An error occured
        End Select
    End With

End Sub

Public Sub Data_GetChars(ByVal inSox As Long)
'*********************************************************************************
'User requests the list of characters from an account
'<Account(S)><Password(S)>
'*********************************************************************************
Dim Account As String
Dim Password As String

    'Get the values
    Account = rBuf.Get_String
    Password = rBuf.Get_String
    
    'Check for valid information
    If Not IsLegalName(Account) Then Exit Sub
    If Not IsLegalPassword(Password) Then Exit Sub
    
    'Perform the query
    DB_RS.Open "SELECT pass,user1,user2,user3,user4,user5 FROM accounts WHERE `name`='" & Account & "'", _
        DB_Conn, adOpenStatic, adLockOptimistic
    
    'Check if the name exists
    If DB_RS.EOF Then
        ConBuf.Clear
        ConBuf.Put_Byte AccountPId.SC_NoChars
        frmMain.GOREsock.SendData inSox, ConBuf.Get_Buffer()
        DB_RS.Close
        Exit Sub
    End If
    
    'MD5 the password
    Password = MD5_String(Password)
    
    'Check if the password is correct
    If DB_RS!Pass <> Password Then
        ConBuf.Clear
        ConBuf.Put_Byte AccountPId.SC_BadPass
        frmMain.GOREsock.SendData inSox, ConBuf.Get_Buffer()
        DB_RS.Close
        Exit Sub
    End If
    
    'Create the list of characters
    ConBuf.Clear
    ConBuf.Put_Byte AccountPId.SC_SendChars
    ConBuf.Put_String DB_RS!user1
    ConBuf.Put_String DB_RS!user2
    ConBuf.Put_String DB_RS!user3
    ConBuf.Put_String DB_RS!user4
    ConBuf.Put_String DB_RS!user5
    frmMain.GOREsock.SendData inSox, ConBuf.Get_Buffer()
    
    'Close the recordset
    DB_RS.Close

End Sub
