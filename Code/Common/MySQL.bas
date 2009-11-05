Attribute VB_Name = "MySQL"
Option Explicit

'Connection objects
Public DB_Conn As ADODB.Connection
Public DB_RS As ADODB.Recordset

Public Function MySQL_Init() As Boolean
Dim DB_User As String   'The database username - (default "root")
Dim DB_Pass As String   'Password to your database for the corresponding username
Dim DB_Name As String   'Name of the table in the database (default "vbgore")
Dim DB_Host As String   'IP of the database server - use localhost if hosted locally! Only host remotely for multiple servers!
Dim DB_Port As Integer  'Port of the database (default "3306")
Dim ErrorString As String
Dim i As Long
 
    On Error GoTo ErrOut
 
    'Create the database connection object
    Set DB_Conn = New ADODB.Connection
    Set DB_RS = New ADODB.Recordset
 
    'Get the variables
    DB_User = Trim$(IO_INI_Read(App.Path & "\Server Data\Settings.ini", "MYSQL", "User"))
    DB_Pass = Trim$(IO_INI_Read(App.Path & "\Server Data\Settings.ini", "MYSQL", "Password"))
    DB_Name = Trim$(IO_INI_Read(App.Path & "\Server Data\Settings.ini", "MYSQL", "Database"))
    DB_Host = Trim$(IO_INI_Read(App.Path & "\Server Data\Settings.ini", "MYSQL", "Host"))
    DB_Port = Val(IO_INI_Read(App.Path & "\Server Data\Settings.ini", "MYSQL", "Port"))
 
    'Create the connection
    DB_Conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & DB_Host & _
        ";DATABASE=" & DB_Name & ";PORT=" & DB_Port & ";UID=" & DB_User & ";PWD=" & DB_Pass & ";OPTION=3"
    DB_Conn.CursorLocation = adUseClient
    DB_Conn.Open
 
    On Error GoTo 0
    
    'Init was successful
    MySQL_Init = True
 
    Exit Function
 
ErrOut:
 
    'Refresh the errors
    DB_Conn.Errors.Refresh
 
    'Get the error string if there is one
    If DB_Conn.Errors.Count > 0 Then ErrorString = DB_Conn.Errors.Item(0).Description
 
End Function

Public Sub MySQL_KeepAlive()

    'Send a query to the database to keep the connection alive
    DB_RS.Open "SELECT name FROM accounts WHERE 0=1", DB_Conn, adOpenStatic, adLockOptimistic
    DB_RS.Close

End Sub

Public Sub MySQL_Optimize()

    'Optimize the database tables
    DB_Conn.Execute "OPTIMIZE TABLE accounts,users"

End Sub
