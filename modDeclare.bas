Attribute VB_Name = "modDeclare"
'###############################################################################
'###############################################################################
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

Public Const LONG_SIZE = 4
Public Const INT_SIZE = 2
Public Const BYTE_SIZE = 1

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDestination As Any, lpSource As Any, ByVal lLength As Long)

'connection management routines
Public Declare Sub mysql_close Lib "libmySQL" (ByVal lMYSQL As Long)
Public Declare Function mysql_init Lib "libmySQL" (ByVal lMYSQL As Long) As Long
Public Declare Function mysql_options Lib "libmySQL" (ByVal lMYSQL As Long, ByVal lOption As Long, ByVal sArg As String) As Long
Public Declare Function mysql_ping Lib "libmySQL" (ByVal lMYSQL As Long) As Long
Public Declare Function mysql_real_connect Lib "libmySQL" (ByVal lMYSQL As Long, ByVal sHostName As String, ByVal sUserName As String, ByVal sPassword As String, ByVal sDbName As String, ByVal lPortNum As Long, ByVal sSocketName As String, ByVal lFlags As Long) As Long

'status and error-reporting routines
Public Declare Function mysql_errno Lib "libmySQL" (ByVal lMYSQL As Long) As Long
Public Declare Function mysql_error Lib "libmySQL" (ByVal lMYSQL As Long) As Long

'query contruction and execution routines
Public Declare Function mysql_query Lib "libmySQL" (ByVal lMYSQL As Long, ByVal sQueryString As String) As Long
Public Declare Function mysql_select_db Lib "libmySQL" (ByVal lMYSQL As Long, ByVal sDbName As String) As Long

'string escaping - Added by Robert Rowe - 02/07/04
Public Declare Function mysql_escape_string Lib "libmySQL.dll" (ByVal strTo As String, ByVal strFrom As String, ByVal lngLength As Long) As Long
Public Declare Function mysql_real_escape_string Lib "libmySQL.dll" (ByVal lMYSQL As Long, ByVal strTo As String, ByVal strFrom As String, ByVal lngLength As Long) As Long
Public Declare Function mysql_real_query Lib "libmySQL" (ByVal lMYSQL As Long, ByVal sQueryString As String, ByVal lLength As Long) As Long
        
'result set processing routines
Public Declare Sub mysql_data_seek Lib "libmySQL" (ByVal lMYSQL_RES As Long, ByVal lOffset As Currency)
Public Declare Sub mysql_free_result Lib "libmySQL" (ByVal lMYSQL As Long)
Public Declare Function mysql_affected_rows Lib "libmySQL" (ByVal lMYSQL_RES As Long) As Long
Public Declare Function mysql_fetch_field_direct Lib "libmySQL" (ByVal lMYSQL_RES As Long, ByVal lFieldNum As Long) As Long
Public Declare Function mysql_fetch_lengths Lib "libmySQL" (ByVal lMYSQL_RES As Long) As Long
Public Declare Function mysql_fetch_row Lib "libmySQL" (ByVal lMYSQL_RES As Long) As Long
Public Declare Function mysql_field_count Lib "libmySQL" (ByVal lMYSQL As Long) As Long
Public Declare Function mysql_info Lib "libmySQL" (ByVal lMYSQL As Long) As Long
Public Declare Function mysql_insert_id Lib "libmySQL" (ByVal lMYSQL As Long) As Long
Public Declare Function mysql_num_fields Lib "libmySQL" (ByVal lMYSQL_RES As Long) As Long
Public Declare Function mysql_num_rows Lib "libmySQL" (ByVal lMYSQL_RES As Long) As Long
Public Declare Function mysql_store_result Lib "libmySQL" (ByVal lMYSQL As Long) As Long
Public Declare Function mysql_use_result Lib "libmySQL" (ByVal lMYSQL As Long) As Long

'Error Description Constants
Public Const E40000 As String = "No Query Specified."
Public Const E40001 As String = "A valid Connection object is required to Open a recordset."
Public Const E40002 As String = "Connection Closed."
Public Const E40003 As String = "Server Error Detected." '(Server error appended)
Public Const E40004 As String = "Invalid Flush Option."
Public Const E40005 As String = "Recordset Closed."
Public Const E40006 As String = "Invalid Field Specified."
Public Const E40007 As String = "Add/Edit in progress. Call CancelUpdate first."
Public Const E40008 As String = "Cannot Add or Delete if source query is based on multiple tables."
Public Const E40009 As String = "Could not identify the table to delete from."
Public Const E40010 As String = "Could not identify the record to delete. Include the Primary Key in your query."
Public Const E40011 As String = "Could not identify the table to update."
Public Const E40012 As String = "Could not identify the record to update. Include the Primary Key in your query."
Public Const E40013 As String = "Could not identify the table to Insert Into."
Public Const E40014 As String = "Cannot Insert new record. You must set the value of at least one field."
Public Const E40015 As String = "No Current Record. The requested operation requires a current record and either BOF or EOF are true."
Public Const E40016 As String = "No Table specified."
Public Const E40017 As String = "Operation not allowed when Connection is opened."
Public Const E40018 As String = "Invalid Lock Option."
Public Const E40019 As String = "Missing Parameter. The specified Show Type requires a Table Name."
Public Const E40020 As String = "Missing Parameter. The specified Show Type requires a User."
Public Const E40021 As String = "Invalid Show Option."
Public Const E40022 As String = "No Database specified."
Public Const E40023 As String = "Operation not allowed when Recordset is opened."
Public Const E40024 As String = "Invalid Record Position. The requested position refers to a Deleted record or is greater than RecordCount."
Public Const E40025 As String = "Mismatched Number of Elements. FieldList and Values must contain the same number of elements."
Public Const E40026 As String = "Invalid Save Option."
Public Const E40027 As String = "Invalid Destination File Name."
