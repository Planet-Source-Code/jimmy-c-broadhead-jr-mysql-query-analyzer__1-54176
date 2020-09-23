VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MYSQL_FIELD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

'===============================================================================================
'========= private variables ===================================================================
'===============================================================================================

Private MyRS As MYSQL_RS                                   'rs object that "owns" this field object
Private mlMYSQL_RES As Long                                'pointer to mysql result set
Private mlMYSQL_ROW As Long                                'pointer to mysql row
Private mlMYSQL_FIELD_LENGTHS As Long                      'pointer to array of column lengths
Private mlFieldCount As Long                               'number of fields in result set
Private mlRequestedField As Long                           'field being requested
Private mNames As Collection                               'field name in result set

'===============================================================================================
'========= field type enum =====================================================================
'===============================================================================================

Public Enum MYSQL_FIELD_TYPES
    MYSQL_DECIMAL = 0
    MYSQL_TINYINT = 1
    MYSQL_SMALLINT = 2
    MYSQL_INT = 3
    MYSQL_FLOAT = 4
    MYSQL_DOUBLE = 5
    MYSQL_NULL = 6
    MYSQL_TIMESTAMP = 7
    MYSQL_BIGINT = 8
    MYSQL_MEDIUMINT = 9
    MYSQL_DATE = 10
    MYSQL_TIME = 11
    MYSQL_DATETIME = 12
    MYSQL_YEAR = 13
    'MYSQL_NEWDATE = 14                                     'HUH... WHAT'S A NEWDATE ???
    MYSQL_ENUM = 247
    MYSQL_SET = 248
    MYSQL_TINYBLOB = 249
    MYSQL_MEDIUMBLOB = 250
    MYSQL_LONGBLOB = 251
    MYSQL_BLOB = 252
    MYSQL_VARCHAR = 253
    MYSQL_CHAR = 254
End Enum

'===============================================================================================
'========= field flags enum ====================================================================
'===============================================================================================

Public Enum MYSQL_FIELD_FLAGS
    MY_NOT_NULL_FLAG = 1
    MY_PRI_KEY_FLAG = 2
    MY_UNIQUE_KEY_FLAG = 4
    MY_MULTIPLE_KEY_FLAG = 8
    MY_BLOB_FLAG = 16
    MY_UNSIGNED_FLAG = 32
    MY_ZEROFILL_FLAG = 64
    MY_BINARY_FLAG = 128
    MY_ENUM_FLAG = 256
    MY_AUTO_INCREMENT_FLAG = 512
    MY_TIMESTAMP_FLAG = 1024
    MY_SET_FLAG = 2048
    'NUM_FLAG = 32768
End Enum

Public Property Get OriginalValue() As Variant
'---------------------------------------------------------------------------------------
' Procedure : OriginalValue
' DateTime  : 02/17/04 23:18
' Author    : Robert Rowe
' Purpose   : Indicates the value of a Field that existed in the record before any changes were made.
'             ADO compatibility property
'---------------------------------------------------------------------------------------

Dim lTmpChar As Long, lCurLen As Long
Dim lRowData As Long
Dim b() As Byte
Dim sVal As String
Dim uFlds As MYSQL_UPDATE_FIELDS
Dim sName As String
Dim uFld As MYSQL_UPDATE_FIELD

    'check if we are in Add Mode
    If MyRS.EditMode = MY_EDIT_ADD Then
        OriginalValue = sVal
        Exit Property
    End If

    'make sure we can get a value ...
    If mlMYSQL_ROW = 0 Or mlMYSQL_FIELD_LENGTHS = 0 Or mlRequestedField < 0 Or mlRequestedField >= mlFieldCount Then Exit Property

    'get pointer to requested field
    CopyMemory lRowData, ByVal (mlMYSQL_ROW + (LONG_SIZE * mlRequestedField)), LONG_SIZE
    'Debug.Print "Pointer to field " & mlRequestedField & " = " & lRowData

    'get length of requested field
    CopyMemory lCurLen, ByVal (mlMYSQL_FIELD_LENGTHS + (LONG_SIZE * mlRequestedField)), LONG_SIZE
    'Debug.Print "Length of field " & mlRequestedField & " = " & lCurLen

    If lRowData = 0 Then
        'null pointer
        Value = Null
        'Debug.Print "Field " & lRequestedField & " = NULL"
    Else
        sVal = ""
        If lCurLen > 0 Then
            ReDim b(0 To (lCurLen - 1))
            'copy string to byte array
            CopyMemory b(0), ByVal lRowData, BYTE_SIZE * lCurLen
            'convert to unicode
            sVal = StrConv(b(), vbUnicode)
        Else
            sVal = ""
        End If

        'should I type the value before returning it ... ie. CLng(sVal), CInt(sVal), etc. ???
        OriginalValue = sVal
        'Debug.Print "Field " & mlRequestedField & " = " & sVal
    End If

End Property

Public Property Get NumericScale() As Byte
'---------------------------------------------------------------------------------------
' Procedure : NumericScale
' DateTime  : 02/17/04 23:14
' Author    : Robert Rowe
' Purpose   : Indicates the scale of numeric values. Same as Decimals.
'             ADO compatibility property
'---------------------------------------------------------------------------------------

    NumericScale = Decimals
End Property

Public Property Get DefinedSize() As Long
'---------------------------------------------------------------------------------------
' Procedure : DefinedSize
' DateTime  : 02/17/04 23:09
' Author    : Robert Rowe
' Purpose   : Returns the largest allowable Length of the requested field.
'             Same as MaxLength. ADO compatibility property
'---------------------------------------------------------------------------------------
    
    DefinedSize = MaxLength
End Property

Public Property Get ActualSize() As Long
'---------------------------------------------------------------------------------------
' Procedure : ActualSize
' DateTime  : 02/17/04 23:04
' Author    : Robert Rowe
' Purpose   : Returns the actual Length of the requested field. Same as Length.
'             ADO compatibility property
'---------------------------------------------------------------------------------------

    ActualSize = Length
End Property

Public Property Get Name() As String
Attribute Name.VB_Description = "Return the name of the field."
    Name = GetFieldName(mlRequestedField)
End Property

Public Property Get Table() As String
Attribute Table.VB_Description = "Returns the name of table that the field belongs to."
    Table = GetFieldTable(mlRequestedField)
End Property

Public Property Get MySqlType() As MYSQL_FIELD_TYPES
Attribute MySqlType.VB_Description = "Return the type of the field."
    MySqlType = GetFieldType(mlRequestedField)
End Property

Public Property Get Length() As Long
Attribute Length.VB_Description = "Returns the length of the field as specified in the CREATE TABLE statement."
On Error Resume Next
    Length = Len(Me.Value)
    'Length = GetFieldLength(mlRequestedField)
End Property

Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns the length of the longest field value present in the result set."
    MaxLength = GetFieldMaxLength(mlRequestedField)
End Property

Public Property Get Flags() As MYSQL_FIELD_FLAGS
Attribute Flags.VB_Description = "Returns the attributes for the field."
    Flags = GetFieldFlags(mlRequestedField)
End Property

Public Property Get Decimals() As Long
Attribute Decimals.VB_Description = "Return the number of decimals for numeric fields, zero for non-numeric fields."
    Decimals = GetFieldDecimals(mlRequestedField)
End Property

Public Property Get Value() As Variant
Attribute Value.VB_Description = "Returns/sets the value of the field.  If the field has been updated, the updated value will be returned."
Attribute Value.VB_UserMemId = 0
Attribute Value.VB_MemberFlags = "200"
Dim lTmpChar As Long, lCurLen As Long
Dim lRowData As Long
Dim b() As Byte
Dim sVal As String
Dim uFlds As MYSQL_UPDATE_FIELDS
Dim sName As String
Dim uFld As MYSQL_UPDATE_FIELD

    'check if we are in Add Mode
    If MyRS.EditMode = MY_EDIT_ADD Then
        'We are in add mode, return the last value that was Let for this field
        If Not MyRS.MyAddObject Is Nothing Then
            Set uFlds = MyRS.MyAddObject.Item(MyRS.MyAddObject.Count)
            BuildNameCollection                            'incase the name collection doesn't exist yet ...
            sName = Name
            If Len(Trim(sName)) = 0 Then Exit Property     'no field name ... that's NG ...

            Set uFld = uFlds.Item(sName)
            If Not uFld Is Nothing Then
                Value = uFld.Value
            End If
        End If
        Exit Property
    End If

    'make sure we can get a value ...
    If mlMYSQL_ROW = 0 Or mlMYSQL_FIELD_LENGTHS = 0 Or mlRequestedField < 0 Or mlRequestedField >= mlFieldCount Then Exit Property

    'check if this field has been updated ...
    If Not MyRS.MyUpdateObject Is Nothing Then
        'we have updates ...
        Set uFlds = MyRS.MyUpdateObject.Item(IIf(MyRS.UpdateRecordNum = 0, MyRS.AbsolutePosition, MyRS.UpdateRecordNum), True)
        If Not uFlds Is Nothing Then
            'we have updates for this row ...
            BuildNameCollection                            'incase the name collection doesn't exist yet ...
            sName = Name
            If Len(Trim(sName)) = 0 Then Exit Property     'no field name ... that's NG ...

            Set uFld = uFlds.Item(sName)
            If Not uFld Is Nothing Then
                'we have a field that we've tried to update ...
                If uFld.Updated Then
                    'this field was successfully updated ...
                    'return the updated value and exit ...
                    Value = uFld.Value
                    Exit Property
                End If
            End If
        End If
    End If

    'get pointer to requested field
    CopyMemory lRowData, ByVal (mlMYSQL_ROW + (LONG_SIZE * mlRequestedField)), LONG_SIZE
    'Debug.Print "Pointer to field " & mlRequestedField & " = " & lRowData

    'get length of requested field
    CopyMemory lCurLen, ByVal (mlMYSQL_FIELD_LENGTHS + (LONG_SIZE * mlRequestedField)), LONG_SIZE
    'Debug.Print "Length of field " & mlRequestedField & " = " & lCurLen

    If lRowData = 0 Then
        'null pointer
        Value = Null
        'Debug.Print "Field " & lRequestedField & " = NULL"
    Else
        sVal = ""
        If lCurLen > 0 Then
            ReDim b(0 To (lCurLen - 1))
            'copy string to byte array
            CopyMemory b(0), ByVal lRowData, BYTE_SIZE * lCurLen
            'convert to unicode
            sVal = StrConv(b(), vbUnicode)
        Else
            sVal = ""
        End If

        'should I type the value before returning it ... ie. CLng(sVal), CInt(sVal), etc. ???
        Value = sVal
        'Debug.Print "Field " & mlRequestedField & " = " & sVal
    End If
End Property

Public Property Let Value(ByVal vData As Variant)
Dim sName As String
Dim uFlds As MYSQL_UPDATE_FIELDS

    If (MyRS.EOF Or MyRS.BOF) And Not MyRS.EditMode = MY_EDIT_ADD Then
        Err.Raise 40015, "Value", E40015 'No Current Record. The requested operation requires a current record and either BOF or EOF are true.
        Exit Property
    End If

    'i can't accept vb objects !!!
    'If IsObject(vData) Then Exit Property

    BuildNameCollection                                    'incase the name collection doesn't exist yet ...
    sName = Name
    If Len(Trim(sName)) = 0 Then Exit Property             'no field name ... that's NG ...

    If Len(Trim(Table)) = 0 Then Exit Property             'must be a calculated field ... can't update it ...

    If MyRS.EditMode = MY_EDIT_NONE Then
        'haven't started an edit yet ... now edit is in progress ...
        MyRS.EditMode = MY_EDIT_IN_PROGRESS
        'if we haven't updated any rows, we need to create update object
        If MyRS.MyUpdateObject Is Nothing Then Set MyRS.MyUpdateObject = New MYSQL_UPDATE
        'set the record number being updated based on current record #
        'this was needed to make AutoUpdate work ...
        MyRS.UpdateRecordNum = MyRS.AbsolutePosition
        'add this current row to the update object
        MyRS.MyUpdateObject.Add MyRS.UpdateRecordNum
    End If

    Select Case MyRS.EditMode
    Case MY_EDIT_IN_PROGRESS
        'get update fields object for the current update in progress
        Set uFlds = MyRS.MyUpdateObject.Item(MyRS.UpdateRecordNum, True)
    Case MY_EDIT_ADD
        'get update fields object for the current add in progress
        Set uFlds = MyRS.MyAddObject.Item(MyRS.MyAddObject.Count)
    Case Else
        Exit Property                                  'shouldn't be here
    End Select

    'add the name and value of the field being updated
    uFlds.Add sName, vData
End Property

Friend Function SetProperties(ByRef vRequestedField As Variant, ByRef lFieldCount As Long, ByRef lMYSQL_RES As Long, ByRef lMYSQL_ROW As Long, ByRef lMYSQL_FIELD_LENGTHS As Long)

    'setup the values required for handling MYSQL_FIELD properties ...

    'set the field count
    mlFieldCount = lFieldCount
    
    'set the pointer to the result set
    mlMYSQL_RES = lMYSQL_RES
    
    'set the pointer to the current row
    mlMYSQL_ROW = lMYSQL_ROW
    
    'set the pointer to the arrary of column lengths
    mlMYSQL_FIELD_LENGTHS = lMYSQL_FIELD_LENGTHS

    If IsNumeric(vRequestedField) Then
        'column referenced by index
        mlRequestedField = vRequestedField
    Else
        'column referenced by name
        BuildNameCollection
        mlRequestedField = GetFieldFromNameCollection(vRequestedField)
    End If

    'check for any invalid values
    If mlRequestedField < 0 Or mlFieldCount = 0 Or mlRequestedField >= mlFieldCount Then
        mlMYSQL_RES = 0
        mlMYSQL_ROW = 0
        mlMYSQL_FIELD_LENGTHS = 0
    End If
End Function

Friend Property Set rs(rs As MYSQL_RS)
    'set the rs object that "owns" this field object
    Set MyRS = rs
End Property

Private Sub Class_Initialize()
    'initialize name collection
    Set mNames = New Collection
End Sub

Private Sub Class_Terminate()
    'release my objects
    Set MyRS = Nothing
    Set mNames = Nothing
End Sub

Private Sub BuildNameCollection()
'attempt to build collection of column names
Dim lCnt As Long
Dim sName As String

    On Local Error Resume Next                             'in case multiple columns in rs have same name

    If mNames.Count = 0 And mlFieldCount > 0 Then
        For lCnt = 1 To mlFieldCount
            sName = GetFieldName(lCnt - 1)
            If Len(sName) > 0 Then mNames.Add lCnt, sName
        Next lCnt
    End If
End Sub

Private Function GetFieldFromNameCollection(ByRef vRequestedField As Variant) As Long
    On Local Error Resume Next                             'in case requested field is not in collection ...

    If mNames.Count > 0 Then GetFieldFromNameCollection = mNames.Item(vRequestedField)
    'collection is 1 based ... fields are 0 based
    GetFieldFromNameCollection = GetFieldFromNameCollection - 1
End Function

Private Function GetFieldName(ByRef lField As Long) As String
Dim lMYSQL_FIELD As Long                               'pointer to mysql field structure
Dim lpName As Long                                     'pointer to field name

    lMYSQL_FIELD = GetFieldStructure(lField)
    If lMYSQL_FIELD = 0 Then Exit Function

    CopyMemory lpName, ByVal lMYSQL_FIELD, LONG_SIZE
    'Debug.Print "Pointer to field name for field " & lfield & " = " & lpName

    GetFieldName = MyRS.Connection.MySqlStringPointerToVbString(lpName)
    'Debug.Print "Name of field " & lfield & " = " & GetFieldName
End Function

Private Function GetFieldTable(ByRef lField As Long) As String
Dim lMYSQL_FIELD As Long                               'pointer to mysql field structure
Dim lpTable As Long                                    'pointer to field table

    lMYSQL_FIELD = GetFieldStructure(lField)
    If lMYSQL_FIELD = 0 Then Exit Function

    '!!!
    'The commented line works with the old DLL
    'CopyMemory lpTable, ByVal (lMYSQL_FIELD + (1 * LONG_SIZE)), LONG_SIZE
    
    'This line works with the new dll
    CopyMemory lpTable, ByVal (lMYSQL_FIELD + (2 * LONG_SIZE)), LONG_SIZE
    'Debug.Print "Pointer to field table for field " & lfield & " = " & lpTable

    GetFieldTable = MyRS.Connection.MySqlStringPointerToVbString(lpTable)
    'Debug.Print "Table of field " & lfield & " = " & GetFieldTable
End Function

Private Function GetFieldType(ByRef lField As Long) As MYSQL_FIELD_TYPES
Dim lMYSQL_FIELD As Long                               'pointer to mysql field structure
Dim eType As MYSQL_FIELD_TYPES                         'enum field type

    lMYSQL_FIELD = GetFieldStructure(lField)
    If lMYSQL_FIELD = 0 Then
        'assume char type ...
        GetFieldType = MYSQL_CHAR
        Exit Function
    End If

    '!!!
    'The commented line works with the old DLL
    'CopyMemory eType, ByVal (lMYSQL_FIELD + (3 * LONG_SIZE)), LONG_SIZE
    
    'This line works with the new dll
    CopyMemory eType, ByVal (lMYSQL_FIELD + (19 * LONG_SIZE)), LONG_SIZE
    'Debug.Print "enum_field_type for field " & lField & " = " & eType

    GetFieldType = eType
End Function

Private Function GetFieldLength(ByRef lField As Long) As Long
Dim lMYSQL_FIELD As Long                               'pointer to mysql field structure

    lMYSQL_FIELD = GetFieldStructure(lField)
    If lMYSQL_FIELD = 0 Then Exit Function

    '!!!
    'The commented line works with the old DLL
    'CopyMemory GetFieldLength, ByVal (lMYSQL_FIELD + (4 * LONG_SIZE)), LONG_SIZE
    
    'This line works with the new dll
    CopyMemory GetFieldLength, ByVal (lMYSQL_FIELD + (8 * LONG_SIZE)), LONG_SIZE
    'Debug.Print "Length for field " & lField & " = " & GetFieldLength
End Function

Private Function GetFieldMaxLength(ByRef lField As Long) As Long
Dim lMYSQL_FIELD As Long                               'pointer to mysql field structure

    lMYSQL_FIELD = GetFieldStructure(lField)
    If lMYSQL_FIELD = 0 Then Exit Function

    '!!!
    'The commented line works with the old DLL
    'CopyMemory GetFieldMaxLength, ByVal (lMYSQL_FIELD + (5 * LONG_SIZE)), LONG_SIZE
    
    'This line works with the new dll
    CopyMemory GetFieldMaxLength, ByVal (lMYSQL_FIELD + (7 * LONG_SIZE)), LONG_SIZE
    'Debug.Print "Max length for field " & lField & " = " & GetFieldMaxLength
End Function

Private Function GetFieldFlags(ByRef lField As Long) As MYSQL_FIELD_FLAGS
Dim lMYSQL_FIELD As Long                               'pointer to mysql field structure
Dim x As Long

    lMYSQL_FIELD = GetFieldStructure(lField)
    If lMYSQL_FIELD = 0 Then Exit Function

    '!!!
    'The commented line works with the old DLL
    'CopyMemory GetFieldFlags, ByVal (lMYSQL_FIELD + (6 * LONG_SIZE)), LONG_SIZE
    
    'This line works with the new dll
    CopyMemory GetFieldFlags, ByVal (lMYSQL_FIELD + (16 * LONG_SIZE)), LONG_SIZE
    '
    'Debug.Print "Flags for field " & lField & " = " & GetFieldFlags
End Function

Private Function GetFieldDecimals(ByRef lField As Long) As Long
Dim lMYSQL_FIELD As Long                               'pointer to mysql field structure

    lMYSQL_FIELD = GetFieldStructure(lField)
    If lMYSQL_FIELD = 0 Then Exit Function

    '!!!
    'The commented line works with the old DLL
    'CopyMemory GetFieldDecimals, ByVal (lMYSQL_FIELD + (7 * LONG_SIZE)), LONG_SIZE

    'This line works with the new dll
    CopyMemory GetFieldDecimals, ByVal (lMYSQL_FIELD + (17 * LONG_SIZE)), LONG_SIZE
    
    'Debug.Print "Decimals for field " & lField & " = " & GetFieldDecimals
End Function

Private Function GetFieldStructure(ByRef lField As Long) As Long
    'get pointer to mysql field structure
    If mlMYSQL_RES = 0 Or lField < 0 Or lField >= mlFieldCount Then
        GetFieldStructure = 0
        Exit Function
    End If

    GetFieldStructure = mysql_fetch_field_direct(mlMYSQL_RES, lField)
    'Debug.Print "Pointer to MYSQL_FIELD structure for field " & lField & " = " & GetFieldStructure
End Function

