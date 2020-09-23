VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MYSQL_CONNECTION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'###############################################################################
'###############################################################################
'
' MyVBQL - Visual Basic library to interface with a MySQL database
' Copyright (C) 2000,2001 icarz, Inc.
'
' VBMySQLDirect - Extension of the original MyVBQL library
' Copyright (C) 2004 Robert Rowe
'
' This library is free software; you can redistribute it and/or
' modify it under the terms of the GNU Library General Public
' License as published by the Free Software Foundation; either
' version 2 of the License, or (at your option) any later version.
'
' This library is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
' Library General Public License for more details.
'
' You should have received a copy of the GNU Library General Public
' License along with this library; if not, write to the Free
' Software Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.
'
'###############################################################################
'###############################################################################
'
' Written by Eric Grau (with additions and changes by Robert Rowe)
'
' Please send questions, comments, and changes to robert_rowe@yahoo.com
'
'###############################################################################
'###############################################################################
'

Option Explicit

'===============================================================================================
'========= private variables ===================================================================
'===============================================================================================

Private Type MyOption                                       'option variable to send to the server during OpenConnection
    eOption As MYSQL_OPTION
    sArg As String
End Type

Private MyOptArr() As MyOption                             'options to send to the server during OpenConnection
Private mMyError As MYSQL_ERR                              'error object containing error # and description
Private msHostname As String                               'for OpenConnection function
Private msUsername As String                               'for OpenConnection function
Private msPassword As String                               'for OpenConnection function
Private msDbName As String                                 'for OpenConnection function
Private mlPortNum As Long                                  'for OpenConnection function
Private mlFlags As Long                                    'for OpenConnection function
Private myState As MYSQL_CONNECTION_STATE                  'state of this connection object
Private mlMYSQL As Long                                    'pointer to mysql connection
Private msQueryString As String                            'last sql statement queried on this connection
Private mbAutoUpdate As Boolean                            'determines if the auto update feature is enabled
Private mlLOCAL_INFILE As Long                             'used to provide a pointer to turn Local Infile on/off

'===============================================================================================
'========= connection state enum ===============================================================
'===============================================================================================

Public Enum MYSQL_CONNECTION_STATE
    MY_CONN_OPEN = 0
    MY_CONN_CLOSED = 1
End Enum

'===============================================================================================
'========= server options enum =================================================================
'===============================================================================================

Public Enum MYSQL_OPTION
    MYSQL_OPT_CONNECT_TIMEOUT = 0
    MYSQL_OPT_COMPRESS = 1
    MYSQL_OPT_NAMED_PIPE = 2
    MYSQL_INIT_COMMAND = 3
    MYSQL_READ_DEFAULT_FILE = 4
    MYSQL_READ_DEFAULT_GROUP = 5
    'MYSQL_SET_CHARSET_DIR = 6
    'MYSQL_SET_CHARSET_NAME = 7
    'MYSQL_OPT_LOCAL_INFILE = 8
End Enum

'===============================================================================================
'========= connection flags enum ===============================================================
'===============================================================================================

Public Enum MYSQL_FLAG
    CLIENT_LONG_PASSWORD = 1                               'new more secure passwords
    CLIENT_FOUND_ROWS = 2                                  'Found instead of affected rows
    CLIENT_LONG_FLAG = 4                                   'Get all column flags
    CLIENT_CONNECT_WITH_DB = 8                             'One can specify db on connect
    CLIENT_NO_SCHEMA = 16                                  'Don't allow database.table.column
    CLIENT_COMPRESS = 32                                   'Can use compression protocol
    CLIENT_ODBC = 64                                       'Odbc client
    CLIENT_LOCAL_FILES = 128                               'Can use LOAD DATA LOCAL
    CLIENT_IGNORE_SPACE = 256                              'Ignore spaces before '('
    CLIENT_CHANGE_USER = 512                               'Support the mysql_change_user()
    CLIENT_INTERACTIVE = 1024                              'This is an interactive client
    CLIENT_SSL = 2048                                      'Switch to SSL after handshake
    CLIENT_IGNORE_SIGPIPE = 4096                           'IGNORE sigpipes
    CLIENT_TRANSACTIONS = 8192                             'Client knows about transactions
    CLIENT_RESERVED = 16384                                'Old flag for 4.1 protocol
    CLIENT_SECURE_CONNECTION = 32768                       'New 4.1 authentication
    CLIENT_MULTI_STATEMENTS = 65536                        'Enable/disable multi-stmt support
    CLIENT_MULTI_RESULTS = 131072                          'Enable/disable multi-results
End Enum

'===============================================================================================
'========= flush types enum ====================================================================
'===============================================================================================

Public Enum MYSQL_FLUSH
    MY_FLUSH_ALL = 0
    MY_FLUSH_HOSTS = 1
    MY_FLUSH_LOGS = 2
    MY_FLUSH_PRIVILEGES = 3
    MY_FLUSH_STATUS = 4
    MY_FLUSH_TABLES = 5
    MY_FLUSH_TABLES_WITH_READ_LOCK = 6
End Enum

Public Enum MYSQL_LOCKTYPE
    MY_LOCK_READ = 1
    MY_LOCK_READ_LOCAL = 2
    MY_LOCK_WRITE = 3
    MY_LOCK_LOW_PRIORITY_WRITE = 4
End Enum

Public Enum MYSQL_SHOWTYPE
    MY_SHOW_DATABASES = 1
    MY_SHOW_TABLES = 2
    MY_SHOW_COLUMNS = 3
    MY_SHOW_INDEX = 4
    MY_SHOW_TABLE_STATUS = 5
    MY_SHOW_STATUS = 6
    MY_SHOW_VARIABLES = 7
    MY_SHOW_LOGS = 8
    MY_SHOW_PROCESSLIST = 9
    MY_SHOW_GRANTS_FOR = 10
    MY_SHOW_CREATE_TABLE = 11
    MY_SHOW_MASTER_STATUS = 12
    MY_SHOW_MASTER_LOGS = 13
    MY_SHOW_SLAVE_STATUS = 14
    'MY_SHOW_WARNINGS = 15
    'MY_SHOW_ERRORS = 16
    MY_SHOW_TABLE_TYPES = 17
    MY_SHOW_PRIVILEGES = 18
End Enum

'===============================================================================================
'========= ADO compatibility enums =============================================================
'===============================================================================================

Public Enum CursorLocationEnum
    adUseNone = 1
    adUseServer = 2
    adUseClient = 3
End Enum

'Property Locals
Private mvarConnectionTimeout As Long
Private mvarCommandTimeout As Long
Private mvarDefaultDatabase As String
Private mvarConnectionString As String
Private mvarAllowLocalInfile As Boolean
Private mvarCursorLocation As CursorLocationEnum

Public Property Let CursorLocation(ByVal vData As CursorLocationEnum)
'---------------------------------------------------------------------------------------
' Procedure : CursorLocation
' DateTime  : 02/13/04 18:32
' Author    : Robert Rowe
' Purpose   : Does nothing. ADO compatibility property.
'---------------------------------------------------------------------------------------

    mvarCursorLocation = vData
End Property

Public Property Get CursorLocation() As CursorLocationEnum
'---------------------------------------------------------------------------------------
' Procedure : CursorLocation
' DateTime  : 02/13/04 18:32
' Author    : Robert Rowe
' Purpose   : Does nothing. ADO compatibility property.
'---------------------------------------------------------------------------------------

    CursorLocation = mvarCursorLocation
End Property

Public Property Let AllowLocalInfile(ByVal vData As Boolean)
'---------------------------------------------------------------------------------------
' Procedure : AllowLocalInfile
' DateTime  : 02/12/04 21:29
' Author    : Robert Rowe
' Purpose   : Switch to turn on/off support for Local Infile (defaults off)
'             Not allowed on an open connection
'---------------------------------------------------------------------------------------

    If myState = MY_CONN_OPEN Then
        Err.Raise 40017, "AllowLocalInfile", E40017 '"Operation not allowed when Connection is opened."
        Exit Property
    End If
    mvarAllowLocalInfile = vData
    mlLOCAL_INFILE = IIf(mvarAllowLocalInfile, 1, 0)
End Property

Public Property Get AllowLocalInfile() As Boolean
'---------------------------------------------------------------------------------------
' Procedure : AllowLocalInfile
' DateTime  : 02/12/04 21:28
' Author    : Robert Rowe
' Purpose   : Get current setting of the Local Infile option
'---------------------------------------------------------------------------------------

    AllowLocalInfile = mvarAllowLocalInfile
End Property

Public Property Let ConnectionString(ByVal vData As String)
'---------------------------------------------------------------------------------------
' Procedure : ConnectionString
' DateTime  : 02/12/04 21:59
' Author    : Robert Rowe
' Purpose   : Sets the Connection String for the Connection. ADO compatibility property.
'             Not allowed on an open connection
'---------------------------------------------------------------------------------------

    If myState = MY_CONN_OPEN Then
        Err.Raise 40017, "ConnectionTimeout", E40017 '"Operation not allowed when Connection is opened."
        Exit Property
    End If

    mvarConnectionString = vData
End Property

Public Property Get ConnectionString() As String
'---------------------------------------------------------------------------------------
' Procedure : ConnectionString
' DateTime  : 02/12/04 21:59
' Author    : Robert Rowe
' Purpose   : Gives the value of the current ConnectionString. ADO compatibility property.
'---------------------------------------------------------------------------------------

    ConnectionString = mvarConnectionString
End Property

Public Property Let DefaultDatabase(ByVal vData As String)
'---------------------------------------------------------------------------------------
' Procedure : DefaultDatabase
' DateTime  : 02/12/04 21:32
' Author    : Robert Rowe
' Purpose   : Set the current database. Same as SelectSb. ADO compatibility property.
'             Not allowed on a closed connection
'---------------------------------------------------------------------------------------
    
    If myState = MY_CONN_CLOSED Then
        Err.Raise 40002, "DefaultDatabase", E40002 '"Connection Closed."
        Exit Property
    End If
    mvarDefaultDatabase = vData
    SelectDb mvarDefaultDatabase
    mvarDefaultDatabase = msDbName
End Property

Public Property Get DefaultDatabase() As String
'---------------------------------------------------------------------------------------
' Procedure : DefaultDatabase
' DateTime  : 02/12/04 21:32
' Author    : Robert Rowe
' Purpose   : Gives the name of the Current Database. Same as DbName. ADO compatibility property.
'---------------------------------------------------------------------------------------

    mvarDefaultDatabase = msDbName
    DefaultDatabase = mvarDefaultDatabase
End Property

Public Property Let CommandTimeout(ByVal vData As Long)
'---------------------------------------------------------------------------------------
' Procedure : CommandTimeout
' DateTime  : 02/12/04 21:34
' Author    : Robert Rowe
' Purpose   : Does nothing. ADO compatibility property.
'---------------------------------------------------------------------------------------

    mvarCommandTimeout = vData
End Property

Public Property Get CommandTimeout() As Long
'---------------------------------------------------------------------------------------
' Procedure : CommandTimeout
' DateTime  : 02/12/04 21:34
' Author    : Robert Rowe
' Purpose   : Does nothing. ADO compatibility property.
'---------------------------------------------------------------------------------------

    CommandTimeout = mvarCommandTimeout
End Property

Public Property Let ConnectionTimeout(ByVal vData As Long)
'---------------------------------------------------------------------------------------
' Procedure : ConnectionTimeout
' DateTime  : 02/12/04 21:34
' Author    : Robert Rowe
' Purpose   : Sets the Connection timeout for the Connection. ADO compatibility property.
'             Not allowed on an open connection
'---------------------------------------------------------------------------------------

    If myState = MY_CONN_OPEN Then
        Err.Raise 40017, "ConnectionTimeout", E40017 '"Operation not allowed when Connection is opened."
        Exit Property
    End If
    SetOption MYSQL_OPT_CONNECT_TIMEOUT, vData
    mvarConnectionTimeout = vData
End Property

Public Property Get ConnectionTimeout() As Long
'---------------------------------------------------------------------------------------
' Procedure : ConnectionTimeout
' DateTime  : 02/12/04 21:34
' Author    : Robert Rowe
' Purpose   : Get the Connection timeout for the Connection. ADO compatibility property.
'---------------------------------------------------------------------------------------

    ConnectionTimeout = mvarConnectionTimeout
End Property

Public Sub DisableKeys(TableName As String)
'---------------------------------------------------------------------------------------
' Procedure : DisableKeys
' DateTime  : 02/12/04 19:16
' Author    : Robert Rowe
' Purpose   : Turns the updating of non-unique index off (works with MyISAM tables only)
'---------------------------------------------------------------------------------------

    If Len(TableName) = 0 Then
        Err.Raise 40016, "DisableKeys", E40016 '"No Table specified."
        Exit Sub
    End If
    Execute "ALTER TABLE " & TableName & " DISABLE KEYS"
End Sub
Public Sub EnbleKeys(TableName As String)
'---------------------------------------------------------------------------------------
' Procedure : EnbleKeys
' DateTime  : 02/12/04 19:15
' Author    : Robert Rowe
' Purpose   : Turns the updating of non-unique index back on after a call to
'             Disable Keys has been made (works with MyISAM tables only)
'---------------------------------------------------------------------------------------

    If Len(TableName) = 0 Then
        Err.Raise 40016, "DisableKeys", E40016 '"No Table specified."
        Exit Sub
    End If
    Execute "ALTER TABLE " & TableName & " ENABLE KEYS"
End Sub

Public Function Esc(dirtystring As Variant, Optional AddDelimiter As Boolean) As Variant
'---------------------------------------------------------------------------------------
' Procedure : Esc (renamed from mysql_escape_string by Robert Rowe on 02/07/04)
' DateTime  : 1/2/2004 08:39
' Author    : MIKE HILLYER
' Purpose   : THIS MODULE HOLDS THE CODE NEEDED TO IMPLEMENT MYSQL_ESCAPE_STRING
'             IN A VISUAL BASIC 6 PROJECT. IT SHOULD BE NOTED THAT MYSQL HAS DEPRECIATED
'             MYSQL_ESCAPE_STRING IN FAVOR OF MYSQL_REAL_ESCAPE_STRING BUT SINCE
'             THE LATTER REQUIRES AN ESTABLISHED CONNECTION IT ADDS A LOT OF
'             COMPLEXITY FOR A DEVELOPER WHO IS CURRENTLY USING MYODBC. WE COULD USE
'             MYSQL_REAL_ESCAPE_STRING WHEN USING AN ALL-API APPROACH TO USING MYSQL
'             AND OMITTING MYODBC ENTIRELY.
'                  SEE http://www.mysql.com/doc/en/mysql_escape_string.html
'
'             USER BEWARE: THIS FUNCTION IF FOR USE WITH LATIN(DEFAULT)
'                          CHARACTER SETS ONLY, IT WILL NOT NECESSCARILY
'                          WORK WITH NON-LATIN CHARACTER SETS!
'
' License   : Copyright 2003 Mike Hillyer (www.vbmysql.com)
'
'             This program is free software; you can redistribute it and/or modify
'             it under the terms of the GNU General Public License as published by
'             the Free Software Foundation; either version 2 of the License, or
'             (at your option) any later version.
'
'             This program is distributed in the hope that it will be useful,
'             but WITHOUT ANY WARRANTY; without even the implied warranty of
'             MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'             GNU General Public License for more details.
'
'             You should have received a copy of the GNU General Public License
'             along with this program; if not, write to the Free Software
'             Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'             Please forward all questions/suggestions to mike@vbmysql.com
'
'---------------------------------------------------------------------------------------
    
Dim strFrom As String           'SOURCE STRING PASSED TO FUNCTION
Dim lngFromLength As String     'LENGTH OF SOURCE STRING
Dim strTo As String             'DESTINATION STRING COMING FROM FUNCTION
Dim lngToLength As Long         'LENGTH OF DESTINATION STRING
    
    strFrom = dirtystring           'STORE FUNCTION INPUT
    lngFromLength = Len(strFrom)    'GET LENGTH OF INPUT
    
    strTo = Space(lngFromLength * 2 + 1) 'ALLOCATE A BUFFER FOR OUTPUT OF FUNCTION
                                         '2 BYTES PER CHARACTER PLUS A BYTE FOR NULL
                                         'TERMINATOR USED BY FUNCTION
    
    If mlMYSQL = 0 Then
        lngToLength = mysql_escape_string(strTo, strFrom, lngFromLength) 'CALL API
    Else
        lngToLength = mysql_real_escape_string(mlMYSQL, strTo, strFrom, lngFromLength) 'CALL API
    End If
    
    If AddDelimiter Then
        Esc = "'" & Left(strTo, lngToLength) & "'" 'TRIM NULL TERMINATOR
    Else
        Esc = Left(strTo, lngToLength) 'TRIM NULL TERMINATOR
    End If
End Function

Public Property Get AutoUpdate() As Boolean
Attribute AutoUpdate.VB_Description = "Returns/sets the auto update option."
    AutoUpdate = mbAutoUpdate
End Property

Public Property Let AutoUpdate(bAuto As Boolean)
    mbAutoUpdate = bAuto
End Property

Public Property Get QueryString() As String
Attribute QueryString.VB_Description = "Returns the last sql statement executed on this connection."
    QueryString = msQueryString
End Property

Public Property Get Error() As MYSQL_ERR
Attribute Error.VB_Description = "Returns the error object associated with this connection."
    Set Error = mMyError
End Property

Public Property Get State() As MYSQL_CONNECTION_STATE
Attribute State.VB_Description = "Returns the state of the current connection."
    State = myState
End Property

Public Property Get DbName() As String
Attribute DbName.VB_Description = "Returns the database name used to establish the current connection."
    DbName = msDbName
End Property

Public Property Get Password() As String
Attribute Password.VB_Description = "Returns the password used to establish the current connection."
    Password = msPassword
End Property

Public Property Get Username() As String
Attribute Username.VB_Description = "Returns the user name used to establish the current connection."
    Username = msUsername
End Property

Public Property Get Hostname() As String
Attribute Hostname.VB_Description = "Returns the host name used to establish the current connection."
    Hostname = msHostname
End Property

Public Sub SetOption(ByVal eOption As MYSQL_OPTION, Optional ByVal sArg As String = "")
Attribute SetOption.VB_Description = "Set an option to send to the server during the OpenConnection method."
'---------------------------------------------------------------------------------------
' Procedure : SetOption
' DateTime  : ???
' Author    : Eric Grau with modifications by Robert Rowe
' Purpose   : Set various options that define the behaviour of a MySQL connection
'---------------------------------------------------------------------------------------

Dim FoundIt As Boolean
Dim i As Long
On Error Resume Next
    
    'Look for passed option on option list
    FoundIt = False
    For i = 1 To UBound(MyOptArr)
        If MyOptArr(i).eOption = eOption Then
            FoundIt = True
            Exit For
        End If
    Next i
    
    If FoundIt Then
        MyOptArr(i).eOption = eOption
        MyOptArr(i).sArg = sArg
    Else
        ReDim Preserve MyOptArr(UBound(MyOptArr) + 1)
        MyOptArr(UBound(MyOptArr)).eOption = eOption
        MyOptArr(UBound(MyOptArr)).sArg = sArg
    End If
End Sub

Public Sub ClearOptions()
Attribute ClearOptions.VB_Description = "Clear any options set by the SetOption method."
'---------------------------------------------------------------------------------------
' Procedure : ClearOptions
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Clears any connection options previously set by the SetOption method.
'---------------------------------------------------------------------------------------

    ReDim MyOptArr(0)
End Sub

Public Function OpenConnection(Optional ByVal sHostName As String = "", Optional ByVal sUserName As String = "", Optional ByVal sPassword As String = "", Optional ByVal sDbName As String = "", Optional ByVal lPortNum As Long = 3306) As MYSQL_CONNECTION_STATE ', Optional ByVal lFlags As MYSQL_FLAG
Attribute OpenConnection.VB_Description = "Open connection to a MySQL server."
'---------------------------------------------------------------------------------------
' Procedure : OpenConnection
' DateTime  : ???
' Author    : Eric Grau with modifications by Robert Rowe
' Purpose   : Opens a new MySQL connections.
'             ConnectionString parsing added by Robert Rowe
'---------------------------------------------------------------------------------------

Dim TempArray() As String
Dim i As Long
Dim POS As Long
Dim theTemp As String

    'clear any existing errors
    Me.Error.Clear
    
    'close any pending connections
    Me.CloseConnection
    
    'start with a return value of conn state closed
    OpenConnection = myState
    
    'Parse ConnectionString if provided
    If Len(mvarConnectionString) > 0 Then
        TempArray = Split(mvarConnectionString, ";")
        For i = 0 To UBound(TempArray)
            Select Case True
            Case Left$(UCase$(Trim$(TempArray(i))), 6) = "SERVER"
                POS = InStr(Trim$(TempArray(i)), "=")
                If POS > 0 Then
                    theTemp = Trim$(Mid$(Trim$(TempArray(i)), POS + 1))
                    If Len(theTemp) > 0 Then msHostname = theTemp
                End If
            Case Left$(UCase$(Trim$(TempArray(i))), 8) = "DATABASE"
                POS = InStr(Trim$(TempArray(i)), "=")
                If POS > 0 Then
                    theTemp = Trim$(Mid$(Trim$(TempArray(i)), POS + 1))
                    If Len(theTemp) > 0 Then msDbName = theTemp
                End If
            Case Left$(UCase$(Trim$(TempArray(i))), 3) = "UID"
                POS = InStr(Trim$(TempArray(i)), "=")
                If POS > 0 Then
                    theTemp = Trim$(Mid$(Trim$(TempArray(i)), POS + 1))
                    If Len(theTemp) > 0 Then msUsername = theTemp
                End If
            Case Left$(UCase$(Trim$(TempArray(i))), 4) = "PWD"
                POS = InStr(Trim$(TempArray(i)), "=")
                If POS > 0 Then
                    theTemp = Trim$(Mid$(Trim$(TempArray(i)), POS + 1))
                    If Len(theTemp) > 0 Then msPassword = theTemp
                End If
            End Select
        Next i
    End If

    'Override with passed variables if provided
    If Len(sHostName) > 0 Then msHostname = sHostName
    If Len(sUserName) > 0 Then msUsername = sUserName
    If Len(sPassword) > 0 Then msPassword = sPassword
    If Len(sDbName) > 0 Then msDbName = sDbName
    If Len(lPortNum) > 0 Then mlPortNum = lPortNum

    'initialize a connection handler
    mlMYSQL = mysql_init(mlMYSQL)
    If mlMYSQL = 0 Then
        Me.Error.Add -1, "Couldn't obtain a connection handler."
        Exit Function
    Else
        If UBound(MyOptArr) > 0 Then
            'setup options for mysql_real_connect()
            'set options
            For i = 1 To UBound(MyOptArr)
                Select Case MyOptArr(i).eOption
                Case MYSQL_OPT_CONNECT_TIMEOUT
                    mysql_options mlMYSQL, MYSQL_OPT_CONNECT_TIMEOUT, MyOptArr(i).sArg
                Case MYSQL_OPT_COMPRESS
                    mysql_options mlMYSQL, MYSQL_OPT_COMPRESS, 0
                Case MYSQL_OPT_NAMED_PIPE
                    mysql_options mlMYSQL, MYSQL_OPT_NAMED_PIPE, 0
                Case MYSQL_INIT_COMMAND
                    mysql_options mlMYSQL, MYSQL_INIT_COMMAND, MyOptArr(i).sArg
                Case MYSQL_READ_DEFAULT_FILE
                    mysql_options mlMYSQL, MYSQL_READ_DEFAULT_FILE, MyOptArr(i).sArg
                Case MYSQL_READ_DEFAULT_GROUP
                    mysql_options mlMYSQL, MYSQL_READ_DEFAULT_GROUP, MyOptArr(i).sArg
                End Select
            Next i
        End If
        
        'Set Local Infile option
        mysql_options mlMYSQL, 8, mlLOCAL_INFILE

        'connect to server
        If mysql_real_connect(mlMYSQL, msHostname, msUsername, msPassword, msDbName, mlPortNum, "", mlFlags) = 0 Then
            'connection attempt NG ...
            CheckForError
            RealClose
        Else
            'connection established ... state is now open
            myState = MY_CONN_OPEN
            'Debug.Print "MySQL connection openned at address " & mlMYSQL
        End If
    End If

    'setup the return value
    OpenConnection = myState
End Function

Public Sub CloseConnection()
Attribute CloseConnection.VB_Description = "Close connection to MySQL server."
'---------------------------------------------------------------------------------------
' Procedure : CloseConnection
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Closes an Open connection
'---------------------------------------------------------------------------------------

    'close connection if it's currently open
    If myState = MY_CONN_OPEN Then RealClose
End Sub

Public Sub Ping()
Attribute Ping.VB_Description = "Checks if connection is still active.  If not, reconnect using the same parameters used to initially make the connection."
'---------------------------------------------------------------------------------------
' Procedure : Ping
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Make sure the connection is still active ... if not it will reconnect
'---------------------------------------------------------------------------------------

    'make sure the connection is still active ... if not it will reconnect
    If myState = MY_CONN_OPEN Then mysql_ping mlMYSQL
End Sub

Public Function SelectDb(ByVal sDbName As String) As Boolean
Attribute SelectDb.VB_Description = "Selects the current database."
'---------------------------------------------------------------------------------------
' Procedure : SelectDb
' DateTime  : ???
' Author    : Eric Grau with modifications by Robert Rowe
' Purpose   : Changes the database that the connection uses.
'             Error trapping added by Robert Rowe
'---------------------------------------------------------------------------------------

    If myState = MY_CONN_CLOSED Then
        Err.Raise 40002, "SelectDb", E40002 '"Connection Closed."
        Exit Function
    End If
    
    'clear any existing errors
    Me.Error.Clear
    
    'attempt to change the db
    If mysql_select_db(mlMYSQL, sDbName) = 0 Then
        'success
        SelectDb = True
        'setup module wide variable
        msDbName = sDbName
    Else
        'failure
        SelectDb = False
        CheckForError
    End If
End Function

Public Function Execute(ByVal sSQL As String, Optional ByRef lRecordsAffected As Long = 0) As MYSQL_RS
Attribute Execute.VB_Description = "Execute a query on the current connection."
'---------------------------------------------------------------------------------------
' Procedure : Execute
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Public interface to the InternalExecute method
'---------------------------------------------------------------------------------------

    'create a new result set object
    Set Execute = GetNewRs
    
    'execute the statement
    InternalExecute Execute, sSQL, lRecordsAffected
End Function

Public Function QueryInfo() As String
Attribute QueryInfo.VB_Description = "Returns info about the effect of certain types of queries."
'---------------------------------------------------------------------------------------
' Procedure : QueryInfo
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Returns information about the effect of the most recently executed query.
'---------------------------------------------------------------------------------------

Dim lMYSQL_INFO As Long                                'pointer to mysql info

    'can't get query info if connection is closed
    If Me.State = MY_CONN_CLOSED Then
        Err.Raise 40002, "QueryInfo", E40002 '"Connection Closed."
        Exit Function
    End If

    'get query info
    lMYSQL_INFO = mysql_info(mlMYSQL)
    'Debug.Print "Pointer to character array for mysql_info = " & lMYSQL_INFO

    'convertion pointer to string
    QueryInfo = MySqlStringPointerToVbString(lMYSQL_INFO)
End Function

Public Sub Flush(FlushType As MYSQL_FLUSH)
Attribute Flush.VB_Description = "Execute a FLUSH statement on the server."
'---------------------------------------------------------------------------------------
' Procedure : Flush
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Runs the MySQL FLUSH statement on the database.
'---------------------------------------------------------------------------------------

Dim sFlush As String

    'check the flush type
    Select Case FlushType
    Case MY_FLUSH_ALL
        sFlush = "HOSTS, LOGS, PRIVILEGES, STATUS, TABLES"
    Case MY_FLUSH_HOSTS
        sFlush = "HOSTS"
    Case MY_FLUSH_LOGS
        sFlush = "LOGS"
    Case MY_FLUSH_PRIVILEGES
        sFlush = "PRIVILEGES"
    Case MY_FLUSH_STATUS
        sFlush = "STATUS"
    Case MY_FLUSH_TABLES
        sFlush = "TABLES"
    Case MY_FLUSH_TABLES_WITH_READ_LOCK
        sFlush = "TABLES WITH READ LOCK"
    Case Else
        Err.Raise 40004, "Flush", E40004 '"Invalid Flush Option."
        Exit Sub
    End Select

    'execute the flush
    Execute "FLUSH " & sFlush
End Sub

Public Function LastInsertID() As Long
Attribute LastInsertID.VB_Description = "Return the AUTO_INCREMENT value generated by the most recently executed query."
'---------------------------------------------------------------------------------------
' Procedure : LastInsertID
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Returns the AUTO_INCREMENT value generated by the most recently executed query.
'---------------------------------------------------------------------------------------

    'can't get last insert id if connection is closed
    If Me.State = MY_CONN_CLOSED Then
        Err.Raise 40002, "LastInsertID", E40002 '"Connection Closed."
        Exit Function
    End If
    
    'get the id
    LastInsertID = mysql_insert_id(mlMYSQL)
End Function

Friend Sub InternalExecute(ByRef rs As MYSQL_RS, ByRef sSQL As String, Optional ByRef lRecordsAffected As Long = 0)
'---------------------------------------------------------------------------------------
' Procedure : InternalExecute
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Executes a query returning a recordset if the query returned results.
'---------------------------------------------------------------------------------------

Dim bGotError As Boolean

    'clear any existing error
    Me.Error.Clear
     
    'save the query string
    msQueryString = sSQL
    
    'check to make sure query isn't a zero length string
    sSQL = Trim(sSQL)
    If Len(sSQL) = 0 Then
        Err.Raise 40000, "InternalExecute", E40000 '"No Query Specified."
        Exit Sub
    End If
    
    'can't execute query is connection is closed
    If myState = MY_CONN_CLOSED Then
        Err.Raise 40002, "InternalExecute", E40002 '"Connection Closed."
        Exit Sub
    End If

    'execute the query on the rs object
    rs.OpenRecordset sSQL, mlMYSQL, bGotError
    
    'return # records affected
    lRecordsAffected = rs.AffectedRecords

    CheckForError
End Sub

Friend Function MySqlStringPointerToVbString(ByVal lpString As Long) As String
'---------------------------------------------------------------------------------------
' Procedure : MySqlStringPointerToVbString
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Returns a VB string from a Null terminated string
'---------------------------------------------------------------------------------------

Dim b As Byte
Dim sVal As String

    'check for null pointer
    If lpString = 0 Then Exit Function

    sVal = ""
    Do
        'copy string 1 byte at a time until we get null terminator
        CopyMemory b, ByVal lpString, BYTE_SIZE
        If b = 0 Then Exit Do
        sVal = sVal & Chr(b)
        lpString = lpString + BYTE_SIZE
    Loop

    'return the string
    MySqlStringPointerToVbString = sVal
End Function

Private Sub Class_Initialize()
    'Defaults for property values
    mvarConnectionTimeout = 15
    mvarAllowLocalInfile = False
    
    'Local Infile default off
    mlLOCAL_INFILE = 0
    
    'start with no options
    ClearOptions
    
    'Set default connection Timeout
    SetOption MYSQL_OPT_CONNECT_TIMEOUT, mvarConnectionTimeout
    
    'start with connection state closed
    myState = MY_CONN_CLOSED
    
    'creat new error object
    Set mMyError = New MYSQL_ERR
End Sub

Private Sub Class_Terminate()
    'gotta close open connections
    Me.CloseConnection
    'release my error object
    Set mMyError = Nothing
End Sub

Private Function GetNewRs() As MYSQL_RS
'---------------------------------------------------------------------------------------
' Procedure : GetNewRs
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Create and return a new rs object
'---------------------------------------------------------------------------------------
    
Dim MyRS As MYSQL_RS

    Set MyRS = New MYSQL_RS
    Set MyRS.Connection = Me
    Set GetNewRs = MyRS
End Function

Private Sub RealClose()
'---------------------------------------------------------------------------------------
' Procedure : RealClose
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Close the connection and release the MySQL object
'---------------------------------------------------------------------------------------

    'close the current connection
    mysql_close mlMYSQL
    
    'set connection pointer to null ... need to do this inclase OpenConnection is called again
    mlMYSQL = 0
    
    'clear the last query
    msQueryString = ""
    
    'set the connection state to closed
    myState = MY_CONN_CLOSED
End Sub

Private Sub CheckForError()
'---------------------------------------------------------------------------------------
' Procedure : CheckForError
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Queries MySQL to see if the last action caused an error
'---------------------------------------------------------------------------------------

Dim lNumber As Long
Dim sDesc As String
Dim LPSTR As Long
'Dim b As Byte

    'get error code
    lNumber = mysql_errno(mlMYSQL)

    'is there an error
    If lNumber <> 0 Then
        'yes ... get pointer to error string
        LPSTR = mysql_error(mlMYSQL)
        'Debug.Print "Pointer to character array for error message = " & LPSTR

        'add the error to the error object
        sDesc = MySqlStringPointerToVbString(LPSTR)
        Me.Error.Add lNumber, sDesc
    End If
End Sub
Public Sub BeginTrans()
'---------------------------------------------------------------------------------------
' Procedure : BeginTrans
' DateTime  : 02/13/04 18:37
' Author    : Robert Rowe
' Purpose   : Starts a Transaction. Will only affect InnoDB tables.
'---------------------------------------------------------------------------------------

    Execute "BEGIN"
End Sub
Public Sub CommitTrans()
'---------------------------------------------------------------------------------------
' Procedure : CommitTrans
' DateTime  : 02/13/04 18:37
' Author    : Robert Rowe
' Purpose   : Commits a Transaction. Will only affect InnoDB tables.
'---------------------------------------------------------------------------------------
    
    Execute "COMMIT"
End Sub
Public Sub RollbackTrans()
'---------------------------------------------------------------------------------------
' Procedure : RollbackTrans
' DateTime  : 02/13/04 18:37
' Author    : Robert Rowe
' Purpose   : Rollsback a Transaction. Will only affect InnoDB tables.
'---------------------------------------------------------------------------------------

    Execute "ROLLBACK"
End Sub
Public Sub LockTable(TableName As String, LockType As MYSQL_LOCKTYPE)
'---------------------------------------------------------------------------------------
' Procedure : LockTable
' DateTime  : 02/13/04 19:29
' Author    : Robert Rowe
' Purpose   : Set a lock on a given table using the lock type provided. TableName can
'             optionally contain [tablename as alias]
'---------------------------------------------------------------------------------------

Dim sLock As String

    TableName = Trim$(TableName)
    If Len(TableName) = 0 Then
        Err.Raise 40016, "LockTable", E40016 'No Table specified.
        Exit Sub
    End If

    Select Case LockType
    Case MY_LOCK_READ
        sLock = " READ"
    Case MY_LOCK_READ_LOCAL
        sLock = " READ LOCAL"
    Case MY_LOCK_WRITE
        sLock = " WRITE"
    Case MY_LOCK_LOW_PRIORITY_WRITE
        sLock = " LOW_PRIORITY WRITE"
    Case Else
        Err.Raise 40018, "LockTable", E40018 'Invalid Lock Option.
        Exit Sub
    End Select

    Execute "LOCK TABLES " & TableName & sLock
End Sub
Public Sub UnlockTables()
'---------------------------------------------------------------------------------------
' Procedure : UnlockTables
' DateTime  : 02/13/04 19:28
' Author    : Robert Rowe
' Purpose   : Unlocks any Locked tables
'---------------------------------------------------------------------------------------

    Execute "UNLOCK TABLES"
End Sub
Public Function Show(ShowType As MYSQL_SHOWTYPE, Optional TableName As String = "", Optional LikeClause As String = "", Optional User As String = "") As MYSQL_RS 'Optional LimitCount As Long = 0,
'---------------------------------------------------------------------------------------
' Procedure : Show
' DateTime  : 02/13/04 23:58
' Author    : Robert Rowe
' Purpose   : Returns various list of information depending on which Showtype is specified
'             Some ShowTypes require a TableName while others may have other optional clauses
'---------------------------------------------------------------------------------------

Dim ShowCommand As String

    'Determine Show Type
    Select Case ShowType
    Case MY_SHOW_DATABASES
        ShowCommand = "DATABASES "
    Case MY_SHOW_TABLES
        ShowCommand = "TABLES "
    Case MY_SHOW_COLUMNS
        ShowCommand = "COLUMNS FROM "
    Case MY_SHOW_INDEX
        ShowCommand = "INDEX FROM "
    Case MY_SHOW_TABLE_STATUS
        ShowCommand = "TABLE STATUS "
    Case MY_SHOW_STATUS
        ShowCommand = "STATUS "
    Case MY_SHOW_VARIABLES
        ShowCommand = "VARIABLES "
    Case MY_SHOW_LOGS
        ShowCommand = "LOGS "
    Case MY_SHOW_PROCESSLIST
        ShowCommand = "PROCESSLIST"
    Case MY_SHOW_GRANTS_FOR
        ShowCommand = "GRANTS FOR "
    Case MY_SHOW_CREATE_TABLE
        ShowCommand = "CREATE TABLE "
    Case MY_SHOW_MASTER_STATUS
        ShowCommand = "MASTER STATUS "
    Case MY_SHOW_MASTER_LOGS
        ShowCommand = "MASTER LOGS "
    Case MY_SHOW_SLAVE_STATUS
        ShowCommand = "SLAVE STATUS "
'    Case MY_SHOW_WARNINGS
'        ShowCommand = "WARNINGS "
'    Case MY_SHOW_ERRORS
'        ShowCommand = "ERRORS "
    Case MY_SHOW_TABLE_TYPES
        ShowCommand = "TABLE TYPES "
    Case MY_SHOW_PRIVILEGES
        ShowCommand = "PRIVILEGES "
    Case Else
        Err.Raise 40021, "Show", E40021 'Invalid Show Option.
        Exit Function
    End Select
        
    'Check for a table name
    Select Case ShowType
    Case MY_SHOW_COLUMNS, MY_SHOW_INDEX, MY_SHOW_CREATE_TABLE
        TableName = Trim$(TableName)
        If Len(TableName) > 0 Then
            ShowCommand = ShowCommand & TableName
        Else
            Err.Raise 40019, "Show", E40019  'Missing Parameter. The specified Show Type requires a Table Name."
            Exit Function
        End If
    End Select
    
    'Check for a user
    Select Case ShowType
    Case MY_SHOW_GRANTS_FOR
        User = Trim$(User)
        If Len(User) > 0 Then
            ShowCommand = ShowCommand & User
        Else
            Err.Raise 40020, "Show", E40020  'Missing Parameter. The specified Show Type requires a User."
            Exit Function
        End If
    End Select
    
    'Check for a like clause
    Select Case ShowType
    Case MY_SHOW_DATABASES, MY_SHOW_TABLES, MY_SHOW_COLUMNS, MY_SHOW_TABLE_STATUS, MY_SHOW_STATUS, MY_SHOW_VARIABLES
        LikeClause = Trim$(LikeClause)
        If Len(LikeClause) > 0 Then
            If Left$(UCase$(LikeClause), 4) <> "LIKE" Then LikeClause = "LIKE " & LikeClause
            ShowCommand = ShowCommand & LikeClause
        End If
    End Select
    
'    'Check for a limit clause
'    Select Case ShowType
'    Case MY_SHOW_WARNINGS, MY_SHOW_ERRORS
'        If LimitCount > 0 Then
'            ShowCommand = ShowCommand & "LIMIT " & LimitCount
'        End If
'    End Select

    Set Show = Execute("SHOW " & ShowCommand)
End Function

Public Sub CreateDatabase(DatabaseName As String)
'---------------------------------------------------------------------------------------
' Procedure : CreateDatabase
' DateTime  : 02/14/04 19:41
' Author    : Robert Rowe
' Purpose   : Creates a database using the name provided
'---------------------------------------------------------------------------------------

    DatabaseName = Trim$(DatabaseName)
    If Len(DatabaseName) = 0 Then
        Err.Raise 40022, "CreateDatabase", E40022 'No Database specified.
        Exit Sub
    End If
    
    Execute "CREATE DATABASE " & DatabaseName
End Sub
