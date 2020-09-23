VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MYSQL_ERR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
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

Private msDescription As String
Private mlNumber As Long

Friend Property Let Number(ByVal vData As Long)
    mlNumber = vData
End Property

Friend Property Let Description(ByVal vData As String)
    msDescription = vData
End Property

Public Property Get Number() As Long
    Number = mlNumber
End Property

Public Property Get Description() As String
    Description = msDescription
End Property

Friend Sub Add(ByVal lNumber As Long, ByVal sDescription As String)
    mlNumber = lNumber
    msDescription = sDescription
    If mlNumber <> 0 Then
        Err.Raise 40003, "Server Error", E40003 & vbCrLf & "Number: " & mlNumber & vbCrLf & "Description: " & msDescription '"Server Error Detected."
    End If
End Sub

Public Sub Clear()
    msDescription = ""
    mlNumber = 0
End Sub

