VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MYSQL_UPDATE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private mCol As Collection

'===============================================================================================
'========= public methods ======================================================================
'===============================================================================================

Public Function Add(Optional lRow As Long) As MYSQL_UPDATE_FIELDS
    Dim sKey As String
    Dim objNewMember As MYSQL_UPDATE_FIELDS
    Set objNewMember = New MYSQL_UPDATE_FIELDS

    On Local Error GoTo Error_Handle
    If lRow <> 0 Then
        'adding based on row # ... key has to be non-numeric
        sKey = "R" & Trim(lRow)
        mCol.Add objNewMember, sKey
    Else
        mCol.Add objNewMember
    End If

    GoTo Exit_Sub

Error_Handle:
    On Local Error Resume Next
    If lRow = 0 Then GoTo Exit_Sub                         'this should not happen ...

    'already exists ...
    Set objNewMember = Me.Item(lRow, True)

Exit_Sub:
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function

Public Sub Remove(lIndexKey As Long, Optional bIsRowNum As Boolean)
    mCol.Remove IIf(bIsRowNum, "R" & Trim(lIndexKey), lIndexKey)
End Sub

'===============================================================================================
'========= public properties ===================================================================
'===============================================================================================

Public Property Get Item(lIndexKey As Long, Optional bIsRowNum As Boolean) As MYSQL_UPDATE_FIELDS
    On Local Error GoTo Err_Handle

    Set Item = mCol.Item(IIf(bIsRowNum, "R" & Trim(lIndexKey), lIndexKey))
    Exit Property

Err_Handle:
    'item doesn't exist ...
    Set Item = Nothing
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

'===============================================================================================
'========= private methods =====================================================================
'===============================================================================================

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

'===============================================================================================
'===============================================================================================
'===============================================================================================

