VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MYSQL_RS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"MYSQL_FIELD"
Attribute VB_Ext_KEY = "Member0" ,"MYSQL_FIELD"
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
' Please send questions, comments, and changes to robert_rowe@yahoo.com.com
'
'###############################################################################
'###############################################################################
'

Option Explicit

'===============================================================================================
'========= private variables ===================================================================
'===============================================================================================

Private MyCONN As MYSQL_CONNECTION                          'connection object that "owns" this rs object
Private MyFIELDS As MYSQL_FIELD                             'field object
Private mlCurrentRecord As Long                             'current record #
Private mlRecordCount As Long                               'number of records in result set
Private mlFieldCount As Long                                'number of fields in result set
Private mlAffectedRecords As Long                           'number of rows affected by query
Private msQueryString As String                             'last sql statement queried on this rs object
Private myState As MYSQL_RECORDSET_STATE                    'state of this rs object
Private mlMYSQL_RES As Long                                 'pointer to mysql result set
Private mlMYSQL_ROW As Long                                 'pointer to mysql row
Private mlMYSQL_FIELD_LENGTHS As Long                       'pointer to array of column lengths
Private MyAdd As MYSQL_UPDATE                               'used to hold added rows
Private MyReplace As MYSQL_UPDATE                           'used to hold replaced rows
Private MyUpdate As MYSQL_UPDATE                            'used to hold updated rows
Private MyDelete As MYSQL_UPDATE                            'used to hold deleted rows
Private mlUpdateRecord As Long                              'record # currently being editted
Private MyEditMode As MYSQL_EDIT_MODE                       'current edit mode
Private mlNumDeleted As Long                                'Number of deleted rows, used to adjust RecordCount to the correct number of records

'===============================================================================================
'========= rs state enum =======================================================================
'===============================================================================================

Public Enum MYSQL_RECORDSET_STATE
    MY_RS_OPEN = 0
    MY_RS_CLOSED = 1
End Enum

'===============================================================================================
'========= edit mode enum ======================================================================
'===============================================================================================

Public Enum MYSQL_EDIT_MODE
    MY_EDIT_NONE = 0
    MY_EDIT_IN_PROGRESS = 1
    MY_EDIT_ADD = 2
    MY_EDIT_DELETE = 3
    'MY_EDIT_REPLACE = 4
End Enum

'===============================================================================================
'========= absolute position enum ==============================================================
'===============================================================================================

Public Enum MYSQL_ABSOLUTE_POSITION
    MY_POS_BOF = -3
    MY_POS_EOF = -2
    MY_POS_UNKNOWN = -1
End Enum

'===============================================================================================
'========= Find enum ===========================================================================
'===============================================================================================

Public Enum MYSQL_FIND
    MY_FIND_AND_RETURN_POSITION_ONLY = 1
    MY_FIND_GOTO_RECORD_AND_RETURN_POSITION = 2
End Enum

'Added by Robert Rowe 02/07/04

Public Enum MYSQL_SAVE_OPTION
    MY_SAVE_OVERWRITE = 0
    MY_SAVE_APPEND = 1
End Enum

'===============================================================================================
'========= ADO compatibility enums =============================================================
'===============================================================================================

Public Enum CursorTypeEnum
    [adOpenUnspecified] = -1
    adOpenForwardOnly = 0
    adOpenKeyset = 1
    adOpenDynamic = 2
    adOpenStatic = 3
End Enum

Public Enum LockTypeEnum
    [adLockUnspecified] = -1
    adLockReadOnly = 1
    adLockPessimistic = 2
    adLockOptimistic = 3
    adLockBatchOptimistic = 4
End Enum

Public Enum CompareEnum
    adCompareEqual = 1
    adCompareGreaterThan = 2
    adCompareLessThan = 0
    adCompareNotComparable = 4
    adCompareNotEqual = 3
End Enum

'Property Locals
Private mvarCursorType As CursorTypeEnum
Private mvarLockType As LockTypeEnum
Private mvarCursorLocation As CursorLocationEnum
Private mvarBookmark As Long

Public Function CompareBookmarks(Bookmark1 As Variant, Bookmark2 As Variant) As CompareEnum
'---------------------------------------------------------------------------------------
' Procedure : CompareBookmarks
' DateTime  : 02/17/04 12:12
' Author    : Robert Rowe
' Purpose   : Gives the relative positions of 2 Bookmarks. ADO compatibility method.
'---------------------------------------------------------------------------------------

    If IsObject(Bookmark1) Or IsObject(Bookmark2) Then
        CompareBookmarks = adCompareNotComparable
        Exit Function
    End If

    Select Case Bookmark1
    Case Bookmark2
        CompareBookmarks = adCompareEqual
    Case Is > Bookmark2
        CompareBookmarks = adCompareGreaterThan
    Case Is < Bookmark2
        CompareBookmarks = adCompareLessThan
    End Select

End Function

Public Property Let Bookmark(ByVal vData As Long)
'---------------------------------------------------------------------------------------
' Procedure : Bookmark
' DateTime  : 02/17/04 11:43
' Author    : Robert Rowe
' Purpose   : Equivilent to AbsolutePosition. ADO compatibility property.
'---------------------------------------------------------------------------------------

    If myState = MY_RS_CLOSED Then
        Err.Raise 40005, "Bookmark", E40005 '   Recordset Closed.
        Exit Property
    End If

    AbsolutePosition = vData
    mvarBookmark = AbsolutePosition
End Property

Public Property Get Bookmark() As Long
'---------------------------------------------------------------------------------------
' Procedure : Bookmark
' DateTime  : 02/17/04 11:44
' Author    : Robert Rowe
' Purpose   : Equivilent to AbsolutePosition. ADO compatibility property.
'---------------------------------------------------------------------------------------

    mvarBookmark = AbsolutePosition
    Bookmark = mvarBookmark
End Property

Public Property Let CursorLocation(ByVal vData As CursorLocationEnum)
'---------------------------------------------------------------------------------------
' Procedure : CursorLocation
' DateTime  : 02/15/04 02:12
' Author    : Robert Rowe
' Purpose   : Does nothing. ADO compatibility property.
'---------------------------------------------------------------------------------------
    
    mvarCursorLocation = vData
End Property

Public Property Get CursorLocation() As CursorLocationEnum
'---------------------------------------------------------------------------------------
' Procedure : CursorLocation
' DateTime  : 02/15/04 02:13
' Author    : Robert Rowe
' Purpose   : Does nothing. ADO compatibility property.
'---------------------------------------------------------------------------------------

    CursorLocation = mvarCursorLocation
End Property

Public Function GetRows(Optional Rows As Long = 0, Optional Start As Long = 0, Optional Fields As Variant, Optional FieldsInFirstSubscript As Boolean = True, Optional FieldNamesFirst As Boolean = False) As Variant
'---------------------------------------------------------------------------------------
' Procedure : GetRows
' DateTime  : 02/17/04 14:34
' Author    : Robert Rowe
' Purpose   : Returns the data in a recordset as a two two-dimensional array
'---------------------------------------------------------------------------------------

Dim theTemp() As Variant
Dim OrigPosition As Long
Dim NumRecordsWritten As Long
Dim i As Long, j As Long ', v As Variant
Dim NumRows As Long, NumCols As Long
Dim FieldList() As String
On Error GoTo eh

    If myState = MY_RS_CLOSED Then
        Err.Raise 40005, "GetRows", E40005 'Recordset Closed.
        Exit Function
    End If
    
    'If there are no records then bail
    If RecordCount = 0 Then
        GetRows = theTemp
        Exit Function
    End If
    
    ' get the current record ... we are going to return to it
    OrigPosition = mlCurrentRecord
    
    If Start <= 0 Then
        'Reposition to the beginning
        Me.MoveFirst
    Else
        Me.AbsolutePosition = Start
    End If
        
    'Reset row counter
    NumRows = 0
    NumRecordsWritten = 0
        
    'Get the number of columns
    If IsMissing(Fields) Then
        'All fields
        NumCols = mlFieldCount
        ReDim FieldList(1 To NumCols)
        For i = 1 To (mlFieldCount)
            FieldList(i) = Me.Fields(i - 1).Name
        Next i
    Else
        If IsArray(Fields) Then
            NumCols = SafeBoundsCheck(Fields)
            If NumCols = 0 Then
                'We were given a bad array to output all fields
                NumCols = mlFieldCount
                ReDim FieldList(1 To NumCols)
                For i = 1 To (mlFieldCount)
                    FieldList(i) = Me.Fields(i - 1).Name
                Next i
            Else
                'Copy supplied array
                ReDim FieldList(1 To NumCols)
                j = 1
                For i = LBound(Fields) To UBound(Fields)
                    FieldList(j) = Fields(i)
                    j = j + 1
                Next i
            End If
        Else
            If Len(Trim$(Fields)) = 0 Then
                'Empty string so output all fields
                NumCols = mlFieldCount
                ReDim FieldList(1 To NumCols)
                For i = 1 To (mlFieldCount)
                    FieldList(i) = Me.Fields(i - 1).Name
                Next i
            Else
                'Just one field specified
                NumCols = 1
                ReDim FieldList(1 To NumCols)
                FieldList(1) = Trim$(Fields)
            End If
        End If
    End If
    
    'Initialize our array
    ReDim theTemp(0 To NumCols - 1, 0 To 0)
        
    'Put the field names on the first line
    If FieldNamesFirst Then
        For i = 1 To NumCols
            theTemp(i - 1, 0) = Trim$(FieldList(i))
        Next i
        NumRows = 1
    End If
    
    'For each row, loop through the field list and add each value to our array
    'Enlarge the array at the beginning of each row
    Do While Not Me.EOF
        'Increment NumRows
        NumRows = NumRows + 1
        ReDim Preserve theTemp(0 To NumCols - 1, 0 To NumRows - 1)
        
        'Loop through each field appending it to our string
        For i = 1 To NumCols
            theTemp(i - 1, NumRows - 1) = SafeGetValue(FieldList(i))
        Next i
        
        'Increment row counter
        NumRecordsWritten = NumRecordsWritten + 1
                
        'Check for a record limit
        If Rows > 0 Then
            'Check if we just reached the limit
            If NumRecordsWritten = Rows Then Exit Do
        End If
        
        Me.MoveNext
    Loop
    
    'See if we need to swap them
    If Not FieldsInFirstSubscript Then
        theTemp = SwapArrayDim(theTemp)
    End If
    
    'Return result
    GetRows = theTemp

    'return to the record we were at
    mlCurrentRecord = OrigPosition
    
    'reposition the record pointer
    GetRow True
Exit Function
eh:
    GetRows = theTemp
    Err.Raise Err.Number, "GetRows", "Error While Executing GetRows" & vbCrLf & Err.Description
End Function
Public Sub Move(NumRecords As Long, Optional Start As Long = 0)
'---------------------------------------------------------------------------------------
' Procedure : Move
' DateTime  : 02/17/04 22:31
' Author    : Robert Rowe
' Purpose   : Moves the position of the current record.
'---------------------------------------------------------------------------------------

Dim NumRecordsToMove As Long
Dim i As Long

    If myState = MY_RS_CLOSED Then
        Err.Raise 40005, "Move", E40005 'Recordset Closed.
        Exit Sub
    End If
    
    If RecordCount = 0 Then 'No un-deleted records
        Err.Raise 40015, "Move", E40015 'No Current Record. The requested operation requires a current record and either BOF or EOF are true.
        Exit Sub
    End If
    
    NumRecordsToMove = Abs(NumRecords)
    
    Select Case NumRecords
    Case 0
        'No where to go
        Exit Sub
    Case Is < 0
        'Moving backward
        If Start <= 0 Then
            If Me.BOF Then 'Already at BOF
                Err.Raise 40015, "Move", E40015 'No Current Record. The requested operation requires a current record and either BOF or EOF are true.
                Exit Sub
            End If
        Else
            Me.AbsolutePosition = Start
        End If
        
        For i = 1 To NumRecordsToMove
            Me.MovePrevious
            If Me.BOF Then Exit For
        Next i
    Case Is > 0
        'Moving forward
        If Start = 0 Then
            If Me.EOF Then 'Already at EOF
                Err.Raise 40015, "Move", E40015 'No Current Record. The requested operation requires a current record and either BOF or EOF are true.
                Exit Sub
            End If
        Else
            Me.AbsolutePosition = Start
        End If
        For i = 1 To NumRecordsToMove
            Me.MoveNext
            If Me.EOF Then Exit For
        Next i
    End Select

End Sub

Public Sub Save(Destination As String, Optional SaveOption As MYSQL_SAVE_OPTION = MY_SAVE_OVERWRITE, Optional NumRows As Long = 0, Optional ColumnDelimiter As String = vbTab, Optional RowDelimiter As String = vbCrLf, Optional NullExpr As String = "", Optional EscapeValues As Boolean, Optional FieldNamesFirst As Boolean = False)
'---------------------------------------------------------------------------------------
' Procedure : Save
' DateTime  : 02/18/04 20:38
' Author    : Robert Rowe
' Purpose   : Saves the current recordset to a text file.
'---------------------------------------------------------------------------------------

Dim FNum As Integer
Dim theTemp As String
Dim FileIsOpen As Boolean
On Error GoTo eh

    If myState = MY_RS_CLOSED Then
        Err.Raise 40005, "Save", E40005 'Recordset Closed.
        Exit Sub
    End If
    
    Destination = Trim$(Destination)
    If Len(Destination) = 0 Then
        Err.Raise 40027, "Save", E40027  'Invalid Destination File Name.
        Exit Sub
    End If

    'Open file
    FNum = FreeFile
    Select Case SaveOption
    Case MY_SAVE_OVERWRITE
        Open Destination For Output As FNum
    Case MY_SAVE_APPEND
        Open Destination For Append As FNum
    Case Else
        Err.Raise 40026, "Save", E40026 'Invalid Save Option.
        Exit Sub
    End Select
    FileIsOpen = True
    
    'Get the recordset as a string
    theTemp = Me.GetString(NumRows, ColumnDelimiter, RowDelimiter, NullExpr, EscapeValues, FieldNamesFirst)
    
    'Write the recordset to disk
    Print #FNum, theTemp
    
    Close FNum
Exit Sub
eh:
    If FileIsOpen Then Close FNum
    Err.Raise Err.Number, "Save", "Error While Executing Save" & vbCrLf & Err.Description
End Sub

Private Function SwapArrayDim(SourceArray() As Variant) As Variant
'---------------------------------------------------------------------------------------
' Procedure : SwapArrayDim
' DateTime  : 02/17/04 17:08
' Author    : Robert Rowe
' Purpose   : Swaps the dimensions of a two-dimensional array
'---------------------------------------------------------------------------------------

Dim LBRow As Long, LBCol As Long
Dim UBRow As Long, UBCol As Long
Dim CRow As Long, CCol As Long
Dim theTemp() As Variant

    If Not IsArray(SourceArray) Then
        SwapArrayDim = SourceArray
        Exit Function
    End If

    ' get size of original array dimensions
    LBRow = LBound(SourceArray)
    UBRow = UBound(SourceArray)
    LBCol = LBound(SourceArray, 2)
    UBCol = UBound(SourceArray, 2)
    
    ' initialize our new array
    ReDim theTemp(LBCol To UBCol, LBRow To UBRow)
    
    ' swap the dimemsions
    For CRow = LBRow To UBRow
        For CCol = LBCol To UBCol
            theTemp(CCol, CRow) = SourceArray(CRow, CCol)
        Next
    Next
    
    ' return the result
    SwapArrayDim = theTemp

End Function

Public Function GetString(Optional NumRows As Long = 0, Optional ColumnDelimiter As String = vbTab, Optional RowDelimiter As String = vbCrLf, Optional NullExpr As String = "", Optional EscapeValues As Boolean = False, Optional FieldNamesFirst As Boolean = False) As String
'---------------------------------------------------------------------------------------
' Procedure : GetString
' DateTime  : 02/17/04 12:47
' Author    : Robert Rowe
' Purpose   : Returns the data in a recordset as a delimited string.
'---------------------------------------------------------------------------------------

Dim theTemp As String
Dim OrigPosition As Long
Dim NumRecordsWritten As Long
Dim i As Long, v As Variant
On Error GoTo eh

    If myState = MY_RS_CLOSED Then
        Err.Raise 40005, "GetString", E40005 'Recordset Closed.
        Exit Function
    End If
    
    'If there are no records then bail
    If RecordCount = 0 Then
        GetString = ""
        Exit Function
    End If
    
    ' get the current record ... we are going to return to it
    OrigPosition = mlCurrentRecord
    
    'Reposition to the beginning
    Me.MoveFirst
    
    'Initialize local variables
    theTemp = ""
    NumRecordsWritten = 0
    
    'Put the field names on the first line
    If FieldNamesFirst Then
        For i = 0 To (mlFieldCount - 1)
            v = Me.Fields(i).Name
            theTemp = theTemp & v
            
            'Add column delimiter if not on the last field
            If i < mlFieldCount - 1 Then
                theTemp = theTemp & ColumnDelimiter
            End If
        Next i
        
        'Add row delimiter
        theTemp = theTemp & RowDelimiter
    End If
    
    'Loop through each row
    Do While Not Me.EOF
        'Loop through each field appending it to our string
        For i = 0 To (mlFieldCount - 1)
            v = Me.Fields(i).Value
            
            'Mangle value as per provided parameters
            If EscapeValues Then
                theTemp = theTemp & MyCONN.Esc(v)
            Else
                If IsNull(v) Then
                    theTemp = theTemp & NullExpr
                Else
                    theTemp = theTemp & v
                End If
            End If
            
            'Add column delimiter if not on the last field
            If i < mlFieldCount - 1 Then
                theTemp = theTemp & ColumnDelimiter
            End If
        Next i
        
        'Add row delimiter
        theTemp = theTemp & RowDelimiter
        
        'Increment row counter
        NumRecordsWritten = NumRecordsWritten + 1
        
        'Check for a record limit
        If NumRows > 0 Then
            'Check if we just reached the limit
            If NumRecordsWritten = NumRows Then Exit Do
        End If
        
        Me.MoveNext
    Loop
    
    GetString = theTemp

    'return to the record we were at
    mlCurrentRecord = OrigPosition
    
    'reposition the record pointer
    GetRow True

Exit Function
eh:
    GetString = theTemp
    Err.Raise Err.Number, "GetString", "Error While Executing GetString" & vbCrLf & Err.Description
End Function
Public Function GetRowString(Optional NumRows As Long = 0, Optional ColumnDelimiter As String = vbTab, Optional RowDelimiter As String = vbCrLf, Optional NullExpr As String = "", Optional EscapeValues As Boolean = False, Optional FieldNamesFirst As Boolean = False) As String()
'---------------------------------------------------------------------------------------
' Procedure : GetRowString
' DateTime  : 02/17/04 20:51
' Author    : Robert Rowe
' Purpose   : Returns the data in a recordset as an array of delimited strings. Each row
'             is returned as one element in the array with fields delimited.
'---------------------------------------------------------------------------------------

Dim theTemp As String, TempArray() As String
Dim OrigPosition As Long
Dim NumRecordsWritten As Long
Dim i As Long, v As Variant
On Error GoTo eh

    If myState = MY_RS_CLOSED Then
        Err.Raise 40005, "GetString", E40005 'Recordset Closed.
        Exit Function
    End If
    
    'If there are no records then bail
    If RecordCount = 0 Then
        GetRowString = TempArray
        Exit Function
    End If
    
    ' get the current record ... we are going to return to it
    OrigPosition = mlCurrentRecord
    
    'Reposition to the beginning
    Me.MoveFirst
    
    'Initialize local variables
    theTemp = ""
    Erase TempArray
    NumRecordsWritten = 0
    
    'Put the field names on the first line
    If FieldNamesFirst Then
        For i = 0 To (mlFieldCount - 1)
            v = Me.Fields(i).Name
            theTemp = theTemp & v
            
            'Add column delimiter if not on the last field
            If i < mlFieldCount - 1 Then
                theTemp = theTemp & ColumnDelimiter
            End If
        Next i
        
        'Add row delimiter
        theTemp = theTemp & RowDelimiter
    End If
    
    'Loop through each row
    Do While Not Me.EOF
        'Loop through each field appending it to our string
        For i = 0 To (mlFieldCount - 1)
            v = Me.Fields(i).Value
            
            'Mangle value as per provided parameters
            If EscapeValues Then
                theTemp = theTemp & MyCONN.Esc(v)
            Else
                If IsNull(v) Then
                    theTemp = theTemp & NullExpr
                Else
                    theTemp = theTemp & v
                End If
            End If
            
            'Add column delimiter if not on the last field
            If i < mlFieldCount - 1 Then
                theTemp = theTemp & ColumnDelimiter
            End If
        Next i
        
        'Increment row counter
        NumRecordsWritten = NumRecordsWritten + 1
        
        'Add element to array
        ReDim Preserve TempArray(1 To NumRecordsWritten)
        TempArray(NumRecordsWritten) = theTemp
        theTemp = ""
        
        'Check for a record limit
        If NumRows > 0 Then
            'Check if we just reached the limit
            If NumRecordsWritten = NumRows Then Exit Do
        End If
        
        Me.MoveNext
    Loop
    
    GetRowString = TempArray

    'return to the record we were at
    mlCurrentRecord = OrigPosition
    
    'reposition the record pointer
    GetRow True

Exit Function
eh:
    GetRowString = TempArray
    Err.Raise Err.Number, "GetRowString", "Error While Executing GetRowString" & vbCrLf & Err.Description
End Function


Private Function IsOnList(List As Variant, Value As String) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : IsOnList
' DateTime  : 02/17/04 17:06
' Author    : Robert Rowe
' Purpose   : Check if a given value is in a given list (array)
'---------------------------------------------------------------------------------------

Dim i As Long, FoundIt As Boolean
    If IsArray(List) Then
        FoundIt = False
        For i = LBound(List) To UBound(List)
            If Trim$(Value) = Trim$(List(i)) Then
                FoundIt = True
                Exit For
            End If
        Next i
        IsOnList = FoundIt
    Else
        If Trim$(Value) = Trim$(List) Then IsOnList = True Else IsOnList = False
    End If
End Function

Public Property Let LockType(ByVal vData As LockTypeEnum)
'---------------------------------------------------------------------------------------
' Procedure : LockType
' DateTime  : 02/15/04 02:13
' Author    : Robert Rowe
' Purpose   : Does nothing. ADO compatibility property.
'---------------------------------------------------------------------------------------

    mvarLockType = vData
End Property

Public Property Get LockType() As LockTypeEnum
'---------------------------------------------------------------------------------------
' Procedure : LockType
' DateTime  : 02/15/04 02:13
' Author    : Robert Rowe
' Purpose   : Does nothing. ADO compatibility property.
'---------------------------------------------------------------------------------------

    LockType = mvarLockType
End Property

Public Property Let CursorType(ByVal vData As CursorTypeEnum)
'---------------------------------------------------------------------------------------
' Procedure : CursorType
' DateTime  : 02/15/04 02:13
' Author    : Robert Rowe
' Purpose   : Does nothing. ADO compatibility property.
'---------------------------------------------------------------------------------------

    mvarCursorType = vData
End Property

Public Property Get CursorType() As CursorTypeEnum
'---------------------------------------------------------------------------------------
' Procedure : CursorType
' DateTime  : 02/15/04 02:13
' Author    : Robert Rowe
' Purpose   : Does nothing. ADO compatibility property.
'---------------------------------------------------------------------------------------

    CursorType = mvarCursorType
End Property

Private Sub DoCancelUpdate()
'---------------------------------------------------------------------------------------
' Procedure : DoCancelUpdate
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Cancels an update that is in progress. Called by the CancelUpdate method.
'---------------------------------------------------------------------------------------

    Select Case Me.EditMode
    Case MY_EDIT_IN_PROGRESS
        'currently editting ...
        Me.MyUpdateObject.Remove mlUpdateRecord, True
    Case MY_EDIT_ADD
        'currently adding ...
        Me.MyAddObject.Remove Me.MyAddObject.Count
    Case Else
        Exit Sub                                       'shouldn't be here
    End Select

    'done with edit ...
    Me.EditMode = MY_EDIT_NONE
End Sub

Private Function IsArrayByteEmpty(vtData() As Byte) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : IsArrayByteEmpty
' DateTime  : 02/08/04 21:18
' Author    : Robert Rowe
' Purpose   : Checks to see if a Byte Array is empty
'---------------------------------------------------------------------------------------

Dim x As Long
On Error Resume Next
    Err.Clear
    x = UBound(vtData)
    If Err.Number = 0 Then
        IsArrayByteEmpty = False
    Else
        IsArrayByteEmpty = True
    End If
End Function

Private Function IsRowDeleted(RowNumber As Long) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : IsRowDeleted
' DateTime  : 02/15/04 00:48
' Author    : Robert Rowe
' Purpose   : Checks a given row to see if it is Deleted. True = Deleted, False = Valid
'---------------------------------------------------------------------------------------

Dim uDel As MYSQL_UPDATE_FIELDS

    If Not MyDelete Is Nothing Then
        'if this returns an object that's not nothing, this row's been deleted
        Set uDel = MyDelete.Item(RowNumber, True)
    End If
    
    If Not uDel Is Nothing Then
        IsRowDeleted = True
    Else
        IsRowDeleted = False
    End If
End Function

Public Function OpenRs(Optional Source As String, Optional ActiveConnection As MYSQL_CONNECTION, Optional CursorType As CursorTypeEnum = -1, Optional LockType As LockTypeEnum = -1, Optional Options As Long = -1) As MYSQL_RECORDSET_STATE
'---------------------------------------------------------------------------------------
' Procedure : OpenRs
' DateTime  : 2/7/2004 20:54
' Author    : Robert Rowe
' Purpose   : Provides a wrapper for the OpenRecordset method that is more similar to
'             ADO's Open method.
'---------------------------------------------------------------------------------------

    'Validations
    
    'Check for a source
    If Trim$(Source) = "" Then
        'Check for existing QueryString
        If msQueryString = "" Then
            'Trip error
            Err.Raise 40000, "OpenRs", E40000 '"No Query Specified."
        Else
            Source = msQueryString
        End If
    End If
    
    'Check for ActiveConnection or existing connection
    If Not ActiveConnection Is Nothing Then
        If TypeOf ActiveConnection Is MYSQL_CONNECTION Then
            Set MyCONN = ActiveConnection
        End If
    End If
    
    If MyCONN Is Nothing Then
        'Trip error
        Err.Raise 40001, "OpenRs", E40001 '"A valid Connection object is required to Open a recordset."
    Else
        If Not TypeOf MyCONN Is MYSQL_CONNECTION Then
            'Trip error
            Err.Raise 40001, "OpenRs", E40001 '"A valid Connection object is required to Open a recordset."
        End If
    End If
    
    MyCONN.InternalExecute Me, Source
    OpenRs = myState
End Function

Private Function SafeBoundsCheck(v As Variant) As Long
'---------------------------------------------------------------------------------------
' Procedure : SafeBoundsCheck
' DateTime  : 02/17/04 17:10
' Author    : Robert Rowe
' Purpose   : Determines the number of elements in a one-dimensional array
'---------------------------------------------------------------------------------------

On Error GoTo eh
    SafeBoundsCheck = UBound(v) - LBound(v) + 1
Exit Function
eh:
    SafeBoundsCheck = 0
End Function

Private Function SafeGetValue(FieldName As String) As Variant
'---------------------------------------------------------------------------------------
' Procedure : SafeGetValue
' DateTime  : 02/17/04 17:09
' Author    : Robert Rowe
' Purpose   : Gets the value of a field by name or index. Returns empty string if field
'             does not exist
'---------------------------------------------------------------------------------------

On Error GoTo eh
    SafeGetValue = Me.Fields(FieldName).Value
Exit Function
eh:
    SafeGetValue = ""
End Function

Public Property Get State() As MYSQL_RECORDSET_STATE
Attribute State.VB_Description = "Returns the state of the rs object."
'---------------------------------------------------------------------------------------
' Procedure : State
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Returns the state of the MYSQL_RS object.
'---------------------------------------------------------------------------------------

    State = myState
End Property

Public Property Get AffectedRecords() As Long
Attribute AffectedRecords.VB_Description = "Returns the number of records affected by the last query."
'---------------------------------------------------------------------------------------
' Procedure : AffectedRecords
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Returns the number of rows affected by a query (INSERT, UPDATE, DELETE)
'             or the number of rows returned by a query (SELECT).
'---------------------------------------------------------------------------------------

    AffectedRecords = mlAffectedRecords
End Property

Public Property Get QueryString() As String
Attribute QueryString.VB_Description = "Returns the last sql statement executed on this rs."
'---------------------------------------------------------------------------------------
' Procedure : QueryString
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Returns the query used to generate the recordset.
'---------------------------------------------------------------------------------------

    QueryString = msQueryString
End Property
Public Property Get Source() As String
'---------------------------------------------------------------------------------------
' Procedure : Source
' DateTime  : 02/18/04 20:39
' Author    : Robert Rowe
' Purpose   : Same as QueryString. ADO compatibility property.
'---------------------------------------------------------------------------------------

    Source = msQueryString
End Property

Public Property Let Source(sSource As String)
'---------------------------------------------------------------------------------------
' Procedure : Source
' DateTime  : 02/18/04 20:40
' Author    : Robert Rowe
' Purpose   : Same as QueryString. ADO compatibility property.
'---------------------------------------------------------------------------------------

    If myState = MY_RS_OPEN Then
        Err.Raise 40023, "Source", E40023 'Operation not allowed when Recordset is opened.
        Exit Property
    End If

    msQueryString = sSource
End Property

Public Property Get RecordCount() As Long
Attribute RecordCount.VB_Description = "Returns the number of rows in this result set."
'---------------------------------------------------------------------------------------
' Procedure : RecordCount
' DateTime  : ???
' Author    : Eric Grau with modifications by Robert Rowe
' Purpose   : Returns the number of records in a row returning query (SELECT).
'             Change to take deleted records into account added by Robert Rowe.
'---------------------------------------------------------------------------------------

    RecordCount = mlRecordCount - mlNumDeleted
End Property

Public Property Get FieldCount() As Long
Attribute FieldCount.VB_Description = "Returns the number of fields in this result set."
'---------------------------------------------------------------------------------------
' Procedure : FieldCount
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Returns the number of columns in a row returning query (SELECT).
'---------------------------------------------------------------------------------------

    FieldCount = mlFieldCount
End Property

Public Property Get Fields(ByVal vIndexKey As Variant) As MYSQL_FIELD
Attribute Fields.VB_Description = "Returns the requested field object.  Field may be reference by name or by number (0 to FieldCount - 1)."
Attribute Fields.VB_MemberFlags = "200"
'---------------------------------------------------------------------------------------
' Procedure : Fields
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Returns a field object for the specified field.
'---------------------------------------------------------------------------------------

    'The commented code caused an AddNew to fail if the table
    'started out empty. Leave until you make certain that its
    'absence does not cause any other problem
    
    'used when referencing an element in the rs
    'If Me.BOF Or Me.EOF Then
    '    'setup invalid properties ...
    '    MyFIELDS.SetProperties -1, 0, 0, 0, 0
    'Else
        'setup properties for field object
        MyFIELDS.SetProperties vIndexKey, mlFieldCount, mlMYSQL_RES, mlMYSQL_ROW, mlMYSQL_FIELD_LENGTHS
    'End If

    'return the field object
    Set Fields = MyFIELDS
End Property

Public Property Get BOF() As Boolean
Attribute BOF.VB_Description = "Returns True if rs is at the BOF pointer."
'---------------------------------------------------------------------------------------
' Procedure : BOF
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Returns whether or not the MYSQL_RS obejct is at the BOF marker.
'---------------------------------------------------------------------------------------

    BOF = (myState = MY_RS_CLOSED) Or (RecordCount = 0) Or (mlCurrentRecord = 0)
End Property

Public Property Get EOF() As Boolean
Attribute EOF.VB_Description = "Returns True if rs is at the EOF pointer."
'---------------------------------------------------------------------------------------
' Procedure : EOF
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Returns whether or not the MYSQL_RS obejct is at the EOF marker.
'---------------------------------------------------------------------------------------

    EOF = (myState = MY_RS_CLOSED) Or (RecordCount = 0) Or (mlCurrentRecord > mlRecordCount)
End Property

Public Property Get AbsolutePosition() As Long
Attribute AbsolutePosition.VB_Description = "Returns/set the result set position (1 to RecordCount).  Unknown, bof, or eof values may also be returned."
'---------------------------------------------------------------------------------------
' Procedure : AbsolutePosition
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Returns the position of a recordset's current record from 1 to RecordCount,
'             or returns one of the MYSQL_ABSOLUTE_POSITION values.
'---------------------------------------------------------------------------------------

    If myState = MY_RS_CLOSED Or mlRecordCount = 0 Then
        AbsolutePosition = MY_POS_UNKNOWN
    ElseIf Me.BOF Then
        AbsolutePosition = MY_POS_BOF
    ElseIf Me.EOF Then
        AbsolutePosition = MY_POS_EOF
    Else
        AbsolutePosition = mlCurrentRecord
    End If
End Property

Public Property Let AbsolutePosition(ByVal lPosition As Long)
'---------------------------------------------------------------------------------------
' Procedure : AbsolutePosition
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Sets the position of a recordset's current record from 1 to RecordCount, or returns one of the MYSQL_ABSOLUTE_POSITION values.
'---------------------------------------------------------------------------------------

    
    If myState = MY_RS_CLOSED Then
        Err.Raise 40005, "AbsolutePosition", E40005 'Recordset Closed.
        Exit Property
    End If
    
    If mlRecordCount = 0 Then
        'can't set position when rs is closed or rs has no records
        Err.Raise 40024, "AbsolutePosition", E40024 '"Invalid Record Position. The requested position refers to a Deleted record or is greater than RecordCount."
        Exit Property
    ElseIf lPosition > 0 And lPosition <= mlRecordCount Then
        'Check to see if the requested position is deleted
        If IsRowDeleted(lPosition) Then
            Err.Raise 40024, "AbsolutePosition", E40024 '"Invalid Record Position. The requested position refers to a Deleted record or is greater than RecordCount."
            Exit Property
        End If
                
        'set the record #
        mlCurrentRecord = lPosition
        
        'reposition the record pointer
        GetRow
    Else
        Err.Raise 40024, "AbsolutePosition", E40024 '"Invalid Record Position. The requested position refers to a Deleted record or is greater than RecordCount."
        Exit Property
    End If
End Property

Public Property Get TableCount() As Long
Attribute TableCount.VB_Description = "Returns the number of tables in this result set."
'---------------------------------------------------------------------------------------
' Procedure : TableCount
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Returns the number of tables in this result set.
'---------------------------------------------------------------------------------------

Dim vTbl As Variant
    
    vTbl = BuildTableArray

    If IsArray(vTbl) Then
        TableCount = UBound(vTbl) + 1
    Else
        TableCount = 0
    End If
End Property

Public Property Get Table() As String
Attribute Table.VB_Description = "Returns the names of the tables in this result set.  If there are more that one table, the name(s) will be separated by commas."
'---------------------------------------------------------------------------------------
' Procedure : Table
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Returns the name(s) of the tables in this result set.  If there are more
'             that one table, the names will be separated by commas.
'---------------------------------------------------------------------------------------

Dim lCnt As Long
Dim vTbl As Variant
    
    vTbl = BuildTableArray

    If IsArray(vTbl) Then
        lCnt = 0
        Do
            'when there's more than 1 table, send back a comma delimited list
            Table = Table & IIf(lCnt > 0, ",", "") & vTbl(lCnt)
            lCnt = lCnt + 1
        Loop While lCnt <= UBound(vTbl)
    Else
        Table = ""
    End If
End Property

Public Property Get EditMode() As MYSQL_EDIT_MODE
Attribute EditMode.VB_Description = "Returns the edit mode of the result set."
'---------------------------------------------------------------------------------------
' Procedure : EditMode
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Returns the current EditMode of the recordset.
'---------------------------------------------------------------------------------------

Dim uDel As MYSQL_UPDATE_FIELDS

    If Not MyDelete Is Nothing Then
        'if this returns an object that's not nothing, this row's been deleted
        Set uDel = MyDelete.Item(Me.AbsolutePosition, True)
    End If
    
    EditMode = IIf(uDel Is Nothing, MyEditMode, MY_EDIT_DELETE)
End Property

Friend Property Let EditMode(my_ed As MYSQL_EDIT_MODE)
'---------------------------------------------------------------------------------------
' Procedure : EditMode
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Sets the current EditMode of the recordset.
'---------------------------------------------------------------------------------------

    MyEditMode = my_ed
    'when edit mode is set to none ... clear the update record #
    If MyEditMode = MY_EDIT_NONE Then mlUpdateRecord = 0
End Property

Friend Property Get Connection() As MYSQL_CONNECTION
'---------------------------------------------------------------------------------------
' Procedure : Connection
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Returns the connection object used to open the recordset.
'---------------------------------------------------------------------------------------

    Set Connection = MyCONN
End Property

Friend Property Set Connection(cn As MYSQL_CONNECTION)
'---------------------------------------------------------------------------------------
' Procedure : Connection
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Sets the connection object used to open the recordset.
'---------------------------------------------------------------------------------------

    If myState = MY_RS_OPEN Then
        Err.Raise 40023, "Connection", E40023 'Operation not allowed when Recordset is opened.
        Exit Property
    End If

    Set MyCONN = cn
End Property

Public Property Get ActiveConnection() As MYSQL_CONNECTION
'---------------------------------------------------------------------------------------
' Procedure : ActiveConnection
' DateTime  : 02/08/04 22:16
' Author    : Robert Rowe
' Purpose   : ADO compatibility property - equivilent to Connection
'---------------------------------------------------------------------------------------
    
    Set Connection = MyCONN
End Property

Public Property Set ActiveConnection(cn As MYSQL_CONNECTION)
'---------------------------------------------------------------------------------------
' Procedure : ActiveConnection
' DateTime  : 02/08/04 22:16
' Author    : Robert Rowe
' Purpose   : ADO compatibility property - equivilent to the original Connection
'---------------------------------------------------------------------------------------
    
    If myState = MY_RS_OPEN Then
        Err.Raise 40023, "Bookmark", E40023 'Operation not allowed when Recordset is opened.
        Exit Property
    End If

    Set MyCONN = cn
End Property

Friend Property Get MyUpdateObject() As MYSQL_UPDATE
'---------------------------------------------------------------------------------------
' Procedure : MyUpdateObject
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Returns a reference to the Update object.
'---------------------------------------------------------------------------------------

    Set MyUpdateObject = MyUpdate
End Property

Friend Property Set MyUpdateObject(my_upd As MYSQL_UPDATE)
'---------------------------------------------------------------------------------------
' Procedure : MyUpdateObject
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Sets a reference to the Update object.
'---------------------------------------------------------------------------------------

    Set MyUpdate = my_upd
End Property

Friend Property Get MyAddObject() As MYSQL_UPDATE
'---------------------------------------------------------------------------------------
' Procedure : MyAddObject
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Returns a reference to the Add object.
'---------------------------------------------------------------------------------------

    Set MyAddObject = MyAdd
End Property

Friend Property Set MyAddObject(my_add As MYSQL_UPDATE)
'---------------------------------------------------------------------------------------
' Procedure : MyAddObject
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Sets a reference to the Add object.
'---------------------------------------------------------------------------------------

    Set MyAdd = my_add
End Property

Friend Property Get UpdateRecordNum() As Long
'---------------------------------------------------------------------------------------
' Procedure : UpdateRecordNum
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Gets the index of the Update record. Called by Get Value.
'---------------------------------------------------------------------------------------

    'make sure the update record # is 0 when edit mode is set to none
    If MyEditMode = MY_EDIT_NONE Then mlUpdateRecord = 0
    UpdateRecordNum = mlUpdateRecord
End Property

Friend Property Let UpdateRecordNum(lRec As Long)
'---------------------------------------------------------------------------------------
' Procedure : UpdateRecordNum
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Sets the index of the Update record. Called by Let Value.
'---------------------------------------------------------------------------------------

    mlUpdateRecord = lRec
End Property

Friend Function OpenRecordset(ByRef sSQL As String, ByRef lMYSQL As Long, ByRef bGotError As Boolean) As MYSQL_RECORDSET_STATE
'---------------------------------------------------------------------------------------
' Procedure : OpenRecordset
' DateTime  : 02/18/04 20:31
' Author    : Eric Grau with modifications by Robert Rowe
' Purpose   : Opens the recordset.
'             Error traps and the use of mysql_real_query added by Robert Rowe.
'---------------------------------------------------------------------------------------

Dim lRc As Long
Dim lLength As Long

    If myState = MY_RS_OPEN Then
        Err.Raise 40023, "OpenRecordset", E40023 '"Operation not allowed when Recordset is opened."
        Exit Function
    End If

    'Clear out any errors
    Err.Clear
    
    'in case we already have an open rs
    Me.CloseRecordset
    
    'default to zero affected records ...
    mlAffectedRecords = 0
    
    'start with a return value of rs state closed
    OpenRecordset = myState

    'Debug.Print "Sending query: " & sSQL

    'save the query string
    msQueryString = sSQL
    'run the query on the server
    
    'Commented by Robert Rowe 02/07/04
    'lRc = mysql_query(lMYSQL, sSQL)
    
    'Added by Robert Rowe 02/07/04
    lLength = Len(sSQL)
    lRc = mysql_real_query(lMYSQL, sSQL, lLength)
    
    'Debug.Print "Query return value = " & lRc

    If lRc <> 0 Then
        'query failed
        bGotError = True
        'Debug.Print "Query failed... have to check for errors."
        Exit Function
    End If

    'return the result pointer
    mlMYSQL_RES = mysql_store_result(lMYSQL)
    'Debug.Print "Pointer to result = " & mlMYSQL_RES

    If mlMYSQL_RES = 0 Then
        'check if we should have gotten data back
        lRc = mysql_field_count(lMYSQL)

        If lRc = 0 Then
            'no result set, no error
            mlAffectedRecords = mysql_affected_rows(lMYSQL)
            'Debug.Print "Query affected " & mlAffectedRows & " rows."
        Else
            'we have an error
            bGotError = True
            'Debug.Print "Query should have returned result... have to check for errors."
            Exit Function
        End If
    Else
        'we got data back ...
        'setup module wide variables
        mlAffectedRecords = mysql_affected_rows(lMYSQL)
        'Debug.Print "Number of affected rows = " & mlAffectedRecords
        mlRecordCount = mysql_num_rows(mlMYSQL_RES)
        'Debug.Print "Number of rows  = " & mlRecordCount
        mlFieldCount = mysql_num_fields(mlMYSQL_RES)
        'Debug.Print "Number of fields  = " & mlFieldCount

        'start by pointing to row #1
        mlCurrentRecord = 1
        'rs is now open
        myState = MY_RS_OPEN
        
        If mlRecordCount > 0 Then
            'reposition the record pointer
            GetRow
        Else
            
        End If
    End If
End Function

Public Sub MoveFirst()
Attribute MoveFirst.VB_Description = "Move to the first row in the result set."
'---------------------------------------------------------------------------------------
' Procedure : MoveFirst
' DateTime  : 02/18/04 20:26
' Author    : Robert Rowe (originally by Eric Grau)
' Purpose   : Moves the current records position to the first record in the recordset.
'             The code that correctly excludes deleted records and the error traps
'             added by Robert Rowe.
'---------------------------------------------------------------------------------------

Dim i As Long
    If myState = MY_RS_CLOSED Then
        Err.Raise 40005, "MoveFirst", E40005 '"Recordset Closed."
        Exit Sub
    End If
    
    If RecordCount = 0 Then 'No un-deleted records so return BOF
        mlCurrentRecord = 0
        Err.Raise 40015, "MoveFirst", E40015 'No Current Record. The requested operation requires a current record and either BOF or EOF are true.
        Exit Sub
    End If

    'point to first record
    mlCurrentRecord = 1
    'reposition the record pointer
    GetRow
    
    'Check that we are not on a deleted record. Keep moving forward until we are on a good record
    For i = 1 To mlRecordCount
        If IsRowDeleted(mlCurrentRecord) Then
            'Go to the next row because we are on a deleted row
            mlCurrentRecord = mlCurrentRecord + 1
            
            'reposition the record pointer
            GetRow
        Else
            Exit For
        End If
    Next i
End Sub

Public Sub MoveLast()
Attribute MoveLast.VB_Description = "Move to the last row in the result set."
'---------------------------------------------------------------------------------------
' Procedure : MoveLast
' DateTime  : 02/18/04 20:27
' Author    : Robert Rowe (originally by Eric Grau)
' Purpose   : Moves the current records position to the last record in the recordset.
'             The code that correctly excludes deleted records and the error traps
'             added by Robert Rowe.
'---------------------------------------------------------------------------------------

Dim i As Long
    If myState = MY_RS_CLOSED Then
        Err.Raise 40005, "MoveLast", E40005 '"Recordset Closed."
        Exit Sub
    End If
    
    If RecordCount = 0 Then 'No un-deleted records so return BOF
        mlCurrentRecord = mlRecordCount + 1
        Err.Raise 40015, "MoveLast", E40015 'No Current Record. The requested operation requires a current record and either BOF or EOF are true.
        Exit Sub
    End If

    'point to last record
    mlCurrentRecord = mlRecordCount
    'reposition the record pointer
    GetRow
    
    'Check that we are not on a deleted record. Keep moving forward until we are on a good record
    For i = mlRecordCount To 1 Step -1
        If IsRowDeleted(mlCurrentRecord) Then
            'Go to the next row because we are on a deleted row
            mlCurrentRecord = mlCurrentRecord - 1
            
            'reposition the record pointer
            GetRow
        Else
            Exit For
        End If
    Next i
End Sub

Public Sub MovePrevious()
Attribute MovePrevious.VB_Description = "Move to the previous row in the result set."
'---------------------------------------------------------------------------------------
' Procedure : MovePrevious
' DateTime  : 02/18/04 20:28
' Author    : Robert Rowe (originally by Eric Grau)
' Purpose   : Moves the current records position to the previous record in the recordset.
'             The code that correctly excludes deleted records and the error traps
'             added by Robert Rowe.
'---------------------------------------------------------------------------------------

Dim i As Long, pCurrentRecord As Long
    If myState = MY_RS_CLOSED Then
        Err.Raise 40005, "MovePrevious", E40005 '"Recordset Closed."
        Exit Sub
    End If

    If RecordCount = 0 Or Me.BOF Then  'No un-deleted records so return BOF
        mlCurrentRecord = 0
        Err.Raise 40015, "MovePrevious", E40015 'No Current Record. The requested operation requires a current record and either BOF or EOF are true.
        Exit Sub
    End If

    'check if we're already at BOF
    If mlCurrentRecord > 0 Then
        'point to previous record
        mlCurrentRecord = mlCurrentRecord - 1
        'reposition the record pointer
        GetRow
    
        'Check that we are not on a deleted record. Keep moving forward until we are on a good record
        pCurrentRecord = mlCurrentRecord
        For i = pCurrentRecord To 1 Step -1
            If IsRowDeleted(mlCurrentRecord) Then
                'Go to the next row because we are on a deleted row
                mlCurrentRecord = mlCurrentRecord - 1
                
                'reposition the record pointer
                GetRow
            Else
                Exit For
            End If
        Next i
    End If
End Sub

Public Sub MoveNext()
Attribute MoveNext.VB_Description = "Move to the next row in the result set."
'---------------------------------------------------------------------------------------
' Procedure : MoveNext
' DateTime  : 02/18/04 20:27
' Author    : Robert Rowe (originally by Eric Grau)
' Purpose   : Moves the current records position to the next record in the recordset.
'             The code that correctly excludes deleted records and the error traps
'             added by Robert Rowe.
'---------------------------------------------------------------------------------------

Dim i As Long, pCurrentRecord As Long
    If myState = MY_RS_CLOSED Then
        Err.Raise 40005, "MoveNext", E40005 '"Recordset Closed."
        Exit Sub
    End If

    If RecordCount = 0 Or Me.EOF Then  'No un-deleted records so return BOF
        mlCurrentRecord = mlRecordCount + 1
        Err.Raise 40015, "MoveNext", E40015 'No Current Record. The requested operation requires a current record and either BOF or EOF are true.
        Exit Sub
    End If

    'check if we're not at EOF
    If mlCurrentRecord <= mlRecordCount Then
        'point to next record
        mlCurrentRecord = mlCurrentRecord + 1
        'reposition the record pointer
        GetRow
        
        'Check that we are not on a deleted record. Keep moving forward until we are on a good record
        pCurrentRecord = mlCurrentRecord
        For i = pCurrentRecord To mlRecordCount
            If IsRowDeleted(mlCurrentRecord) Then
                'Go to the next row because we are on a deleted row
                mlCurrentRecord = mlCurrentRecord + 1
                
                'reposition the record pointer
                GetRow
            Else
                Exit For
            End If
        Next i
    End If
End Sub

Public Sub CloseRecordset()
Attribute CloseRecordset.VB_Description = "Close the record set."
'---------------------------------------------------------------------------------------
' Procedure : CloseRecordset
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Acts as a public interface to the routine that actually closes the recordset
'---------------------------------------------------------------------------------------

    'close rs if it's currently open
    If myState = MY_RS_OPEN Then
        'check if we need to update first have to update first
        CheckAutoUpdate
        'now close it up
        RealClose
    End If
End Sub

Public Function FindFirst(ByVal vField As Variant, ByVal vData As Variant, Optional ByVal FindMethod As MYSQL_FIND = MY_FIND_GOTO_RECORD_AND_RETURN_POSITION) As Long
Attribute FindFirst.VB_Description = "Find the first row where the given field matches the given value."
'---------------------------------------------------------------------------------------
' Procedure : FindFirst
' DateTime  : 02/18/04 20:16
' Author    : Robert Rowe (originally by Eric Grau)
' Purpose   : Finds the first record containing specified data in the given field.
'             Returns the record number if a match if found otherwise -1 is returned.
'             The code that correctly excludes deleted records and the error traps
'             added by Robert Rowe.
'---------------------------------------------------------------------------------------

Dim lOldPosition As Long
Dim lCount As Long
Dim bToggle As Boolean

    If myState = MY_RS_CLOSED Then
        Err.Raise 40005, "FindFirst", E40005 '"Recordset Closed."
        FindFirst = -1
        Exit Function
    End If

    If RecordCount = 0 Then 'No records or they are all deleted
        ' get out of the function, since we have no records
        FindFirst = -1
        Exit Function
    End If
    
    If IsNumeric(vField) Then
        ' check if it's out of bounds
        If vField >= 0 And vField < Me.FieldCount Then bToggle = True
    Else
        ' check if the field name exists
        For lCount = 0 To Me.FieldCount - 1
            If UCase$(Trim(Me.Fields(lCount).Name)) = UCase$(Trim(vField)) Then
                bToggle = True
                Exit For
            End If
        Next lCount
    End If

    If Not bToggle Then
        ' get out of the function since the field doesnt exist.
        Err.Raise 40006, "FindFirst", E40006 '"Invalid Field Specified."
        FindFirst = -1
        Exit Function
    End If

    If FindMethod = MYSQL_FIND.MY_FIND_GOTO_RECORD_AND_RETURN_POSITION Then
        'auto udpate as needed ...
        CheckAutoUpdate
    End If

    ' get the current record in case we are going to return to it
    lOldPosition = mlCurrentRecord

    ' set this to 0 so our do/loop functions correctly
    mlCurrentRecord = 0

    ' reset this variable
    bToggle = False

    Do Until mlCurrentRecord = mlRecordCount
        ' go to next row.
        mlCurrentRecord = mlCurrentRecord + 1
        'reposition the record pointer
        GetRow True
        If Me.Fields(vField).Value = vData Then
            'Check if found row is deleted
            If Not IsRowDeleted(mlCurrentRecord) Then
                ' set the flag to true, we found it
                bToggle = True
                Exit Do
            End If
        End If
    Loop

    If Not bToggle Then
        ' couldnt find the record. exit with a -1, and return to
        ' the record we were at
        mlCurrentRecord = lOldPosition
        'reposition the record pointer
        GetRow True
        FindFirst = -1
        Exit Function
    End If

    ' if we are here, then the record was found.
    FindFirst = mlCurrentRecord

    Select Case FindMethod
    Case MYSQL_FIND.MY_FIND_AND_RETURN_POSITION_ONLY
        mlCurrentRecord = lOldPosition
        'reposition the record pointer
        GetRow True
    Case MYSQL_FIND.MY_FIND_GOTO_RECORD_AND_RETURN_POSITION
        ' we are already on the record we found ...
    End Select
End Function
Public Function FindLast(ByVal vField As Variant, ByVal vData As Variant, Optional ByVal FindMethod As MYSQL_FIND = MY_FIND_GOTO_RECORD_AND_RETURN_POSITION) As Long
Attribute FindLast.VB_Description = "Find the last row where the given field matches the given value."
'---------------------------------------------------------------------------------------
' Procedure : FindLast
' DateTime  : 02/18/04 20:20
' Author    : Robert Rowe (originally by Eric Grau)
' Purpose   : Finds the last record containing specified data in the given field.
'             Returns the record number if a match if found otherwise -1 is returned.
'             The code that correctly excludes deleted records and the error traps
'             added by Robert Rowe.
'---------------------------------------------------------------------------------------

Dim lOldPosition As Long
Dim lCount As Long
Dim bToggle As Boolean

    If myState = MY_RS_CLOSED Then
        Err.Raise 40005, "FindLast", E40005 '"Recordset Closed."
        FindLast = -1
        Exit Function
    End If
    
    If RecordCount = 0 Then 'No records or they are all deleted
        ' get out of the function, since we have no records
        FindLast = -1
        Exit Function
    End If

    If IsNumeric(vField) Then
        ' check if it's out of bounds
        If vField >= 0 And vField < Me.FieldCount Then bToggle = True
    Else
        ' check if the field name exists
        For lCount = 0 To Me.FieldCount - 1
            If UCase(Trim(Me.Fields(lCount).Name)) = UCase(Trim(vField)) Then
                bToggle = True
                Exit For
            End If
        Next lCount
    End If

    If Not bToggle Then
        ' get out of the function since the field doesnt exist.
        Err.Raise 40006, "FindLast", E40006 '"Invalid Field Specified."
        FindLast = -1
        Exit Function
    End If

    If FindMethod = MYSQL_FIND.MY_FIND_GOTO_RECORD_AND_RETURN_POSITION Then
        'auto udpate as needed ...
        CheckAutoUpdate
    End If

    ' get the current record in case we are going to return to it
    lOldPosition = mlCurrentRecord

    ' set this to mlRecordCount + 1 so our do/loop functions correctly
    mlCurrentRecord = mlRecordCount + 1

    ' reset this variable
    bToggle = False

    Do Until mlCurrentRecord = 1
        ' go to next row.
        mlCurrentRecord = mlCurrentRecord - 1
        'reposition the record pointer
        GetRow True
        If Me.Fields(vField).Value = vData Then
            'Check if found row is deleted
            If Not IsRowDeleted(mlCurrentRecord) Then
                ' set the flag to true, we found it
                bToggle = True
                Exit Do
            End If
        End If
    Loop

    If Not bToggle Then
        ' couldnt find the record. exit with a -1, and return to
        ' the record we were at
        mlCurrentRecord = lOldPosition
        'reposition the record pointer
        GetRow True
        FindLast = -1
        Exit Function
    End If

    ' if we are here, then the record was found.
    FindLast = mlCurrentRecord

    Select Case FindMethod
        Case MYSQL_FIND.MY_FIND_AND_RETURN_POSITION_ONLY
            mlCurrentRecord = lOldPosition
            'reposition the record pointer
            GetRow True
        Case MYSQL_FIND.MY_FIND_GOTO_RECORD_AND_RETURN_POSITION
            ' we are already on the record we found ...
    End Select
End Function

Public Function FindNext(ByVal vField As Variant, ByVal vData As Variant, Optional ByVal FindMethod As MYSQL_FIND = MY_FIND_GOTO_RECORD_AND_RETURN_POSITION) As Long
Attribute FindNext.VB_Description = "Find the next row where the given field matches the given value."
'---------------------------------------------------------------------------------------
' Procedure : FindNext
' DateTime  : 02/18/04 20:21
' Author    : Robert Rowe (originally by Eric Grau)
' Purpose   : Finds the next record containing specified data in the given field.
'             Returns the record number if a match if found otherwise -1 is returned.
'             The code that correctly excludes deleted records and the error traps
'             added by Robert Rowe.
'---------------------------------------------------------------------------------------

Dim lOldPosition As Long
Dim lCount As Long
Dim bToggle As Boolean

    If myState = MY_RS_CLOSED Then
        Err.Raise 40005, "FindNext", E40005 '"Recordset Closed."
        FindNext = -1
        Exit Function
    End If
    
    If RecordCount = 0 Then 'No records or they are all deleted
        ' get out of the function, since we have no records
        FindNext = -1
        Exit Function
    End If

    If IsNumeric(vField) Then
        ' check if it's out of bounds
        If vField >= 0 And vField < Me.FieldCount Then bToggle = True
    Else
        ' check if the field name exists
        For lCount = 0 To Me.FieldCount - 1
            If UCase(Trim(Me.Fields(lCount).Name)) = UCase(Trim(vField)) Then
                bToggle = True
                Exit For
            End If
        Next lCount
    End If

    If Not bToggle Then
        ' get out of the function since the field doesnt exist.
        Err.Raise 40006, "FindNext", E40006 '"Invalid Field Specified."
        FindNext = -1
        Exit Function
    End If

    If FindMethod = MYSQL_FIND.MY_FIND_GOTO_RECORD_AND_RETURN_POSITION Then
        'auto udpate as needed ...
        CheckAutoUpdate
    End If

    ' get the current record in case we are going to return to it
    lOldPosition = mlCurrentRecord

    ' reset this variable
    bToggle = False

    Do Until mlCurrentRecord = mlRecordCount
        ' go to next row.
        mlCurrentRecord = mlCurrentRecord + 1
        'reposition the record pointer
        GetRow True
        If Me.Fields(vField).Value = vData Then
            'Check if found row is deleted
            If Not IsRowDeleted(mlCurrentRecord) Then
                ' set the flag to true, we found it
                bToggle = True
                Exit Do
            End If
        End If
    Loop

    If Not bToggle Then
        ' couldnt find the record. exit with a -1, and return to
        ' the record we were at
        mlCurrentRecord = lOldPosition
        'reposition the record pointer
        GetRow True
        FindNext = -1
        Exit Function
    End If

    ' if we are here, then the record was found.
    FindNext = mlCurrentRecord

    Select Case FindMethod
        Case MYSQL_FIND.MY_FIND_AND_RETURN_POSITION_ONLY
            mlCurrentRecord = lOldPosition
            'reposition the record pointer
            GetRow True
        Case MYSQL_FIND.MY_FIND_GOTO_RECORD_AND_RETURN_POSITION
            ' we are already on the record we found ...
    End Select
End Function

Public Function FindPrevious(ByVal vField As Variant, ByVal vData As Variant, Optional ByVal FindMethod As MYSQL_FIND = MY_FIND_GOTO_RECORD_AND_RETURN_POSITION) As Long
Attribute FindPrevious.VB_Description = "Find the previous row where the given field matches the given value."
'---------------------------------------------------------------------------------------
' Procedure : FindPrevious
' DateTime  : 02/18/04 20:21
' Author    : Robert Rowe (originally by Eric Grau)
' Purpose   : Finds the previous record containing specified data in the given field.
'             Returns the record number if a match if found otherwise -1 is returned.
'             The code that correctly excludes deleted records and the error traps
'             added by Robert Rowe.
'---------------------------------------------------------------------------------------

Dim lOldPosition As Long
Dim lCount As Long
Dim bToggle As Boolean

    If myState = MY_RS_CLOSED Then
        Err.Raise 40005, "FindPrevious", E40005 '"Recordset Closed."
        FindPrevious = -1
        Exit Function
    End If
    
    If RecordCount = 0 Then 'No records or they are all deleted
        ' get out of the function, since we have no records
        FindPrevious = -1
        Exit Function
    End If

    If IsNumeric(vField) Then
        ' check if it's out of bounds
        If vField >= 0 And vField < Me.FieldCount Then bToggle = True
    Else
        ' check if the field name exists
        For lCount = 0 To Me.FieldCount - 1
            If UCase(Trim(Me.Fields(lCount).Name)) = UCase(Trim(vField)) Then
                bToggle = True
                Exit For
            End If
        Next lCount
    End If

    If Not bToggle Then
        ' get out of the function since the field doesnt exist.
        Err.Raise 40006, "FindPrevious", E40006 '"Invalid Field Specified."
        FindPrevious = -1
        Exit Function
    End If

    If FindMethod = MYSQL_FIND.MY_FIND_GOTO_RECORD_AND_RETURN_POSITION Then
        'auto udpate as needed ...
        CheckAutoUpdate
    End If

    ' get the current record in case we are going to return to it
    lOldPosition = mlCurrentRecord

    ' reset this variable
    bToggle = False

    Do Until mlCurrentRecord = 1
        ' go to next row.
        mlCurrentRecord = mlCurrentRecord - 1
        'reposition the record pointer
        GetRow True
        If Me.Fields(vField).Value = vData Then
            'Check if found row is deleted
            If Not IsRowDeleted(mlCurrentRecord) Then
                ' set the flag to true, we found it
                bToggle = True
                Exit Do
            End If
        End If
    Loop

    If Not bToggle Then
        ' couldnt find the record. exit with a -1, and return to
        ' the record we were at
        mlCurrentRecord = lOldPosition
        'reposition the record pointer
        GetRow True
        FindPrevious = -1
        Exit Function
    End If

    ' if we are here, then the record was found.
    FindPrevious = mlCurrentRecord

    Select Case FindMethod
    Case MYSQL_FIND.MY_FIND_AND_RETURN_POSITION_ONLY
        mlCurrentRecord = lOldPosition
        'reposition the record pointer
        GetRow True
    Case MYSQL_FIND.MY_FIND_GOTO_RECORD_AND_RETURN_POSITION
        ' we are already on the record we found ...
    End Select
End Function

Public Function RsToHtmlTable() As String
Attribute RsToHtmlTable.VB_Description = "Creates an HTML table based on the data in the record set."
'---------------------------------------------------------------------------------------
' Procedure : RsToHtmlTable
' DateTime  : ???
' Author    : Eric Grau with modifications by Robert Rowe
' Purpose   : Creates an HTML table based on the data in the record set.
'             Error trap added by Robert Rowe.
'---------------------------------------------------------------------------------------

Dim lOldPosition As Long
Dim lCount As Long
Dim v As Variant

    If myState = MY_RS_CLOSED Then
        Err.Raise 40005, "RsToHtmlTable", E40005 '"Recordset Closed."
        Exit Function
    End If
    
    'no data
    If mlRecordCount = 0 Then
        RsToHtmlTable = ""
        Exit Function
    End If

    ' get the current record ... we are going to return to it
    lOldPosition = mlCurrentRecord
    
    ' set this to 0 so our do/loop functions correctly
    mlCurrentRecord = 0

    'loop thru records to build html table
    Do Until mlCurrentRecord >= mlRecordCount
        ' go to next row.
        mlCurrentRecord = mlCurrentRecord + 1
        If Not IsRowDeleted(mlCurrentRecord) Then
            'reposition the record pointer
            GetRow True
    
            RsToHtmlTable = RsToHtmlTable & "<tr>"
            For lCount = 0 To (mlFieldCount - 1)
                v = Me.Fields(lCount).Value
                If IsNull(v) Then
                    v = "&nbsp;"
                ElseIf Len(Trim(v)) = 0 Then
                    v = "&nbsp;"
                End If
                RsToHtmlTable = RsToHtmlTable & "<td>" & v & "</td>"
            Next lCount
            RsToHtmlTable = RsToHtmlTable & "</tr>" & vbCrLf
        End If
    Loop
    RsToHtmlTable = "<table border=1>" & RsToHtmlTable & "</table>" & vbCrLf

    'return to the record we were at
    mlCurrentRecord = lOldPosition
    
    'reposition the record pointer
    GetRow True
End Function

Public Function AddNew(Optional FieldList As Variant, Optional Values As Variant) As Long
Attribute AddNew.VB_Description = "Enables add mode."
'---------------------------------------------------------------------------------------
' Procedure : AddNew
' DateTime  : 02/17/04 21:04 (original date unknown)
' Author    : Robert Rowe (originally by Eric Grau)
' Purpose   : Sets AddNew mode on. If provided with a FieldList and Values then an
'             Insert is attempted without turning AddNew mode on. The error traps and
'             the support for providing Field and Value arrays were added by Robert Rowe.
'---------------------------------------------------------------------------------------

Dim FieldCount As Long, ValueCount As Long
Dim i As Long
Dim FieldString As String, ValueString As String
Dim TableName As String, SQL As String
On Error GoTo eh

    AddNew = -1
    
    'can't add if rs is closed
    If myState = MY_RS_CLOSED Then
        Err.Raise 40005, "AddNew", E40005 '"Recordset Closed."
        Exit Function
    End If
    
    'shouldn't be here if any edit is in progress
    If Me.EditMode <> MY_EDIT_NONE Then
        Err.Raise 40007, "AddNew", E40007 '"Add/Edit in progress. Call CancelUpdate first."
        Exit Function
    End If
    
    'can't add if there's more than 1 table in rs
    If Me.TableCount > 1 Then
        Err.Raise 40008, "AddNew", E40008 '"Cannot Add, Update or Delete  if source query is based on multiple tables."
        Exit Function
    End If
    
    If Not IsMissing(FieldList) And Not IsMissing(Values) Then
        FieldCount = SafeBoundsCheck(FieldList)
        ValueCount = SafeBoundsCheck(Values)
        
        'get the name of the table
        TableName = Me.Fields(0).Table
        
        'gotta have a table to do the add
        If Len(TableName) = 0 Then
            Err.Raise 40011, "DoUpdate", E40011 '"Could not identify the table to update."
            Exit Function
        End If
        
        FieldString = "": ValueString = ""
        If FieldCount = ValueCount Then
            If FieldCount = 0 Then 'They aren't arrays
                If Len(Trim$(FieldList)) = 0 Then
                    'We are missing a field
                    Err.Raise 40006, "AddNew", E40006 'Invalid Field Specified.
                    Exit Function
                Else
                    SQL = "INSERT INTO " & TableName & " (" & FieldList & ") VALUES(" & Values & ")"
                    MyCONN.Execute SQL, AddNew
                End If
            Else 'They are arrays
                'Build Field and Value strings
                For i = LBound(FieldList) To UBound(FieldList)
                    FieldString = FieldString & IIf(Len(FieldString) > 0, ",", "") & FieldList(i)
                    ValueString = ValueString & IIf(Len(ValueString) > 0, ",", "") & IIf(IsNull(Values(i)), "NULL", Values(i))
                Next i
                
                SQL = "INSERT INTO " & TableName & " (" & FieldString & ") VALUES(" & ValueString & ")"
                MyCONN.Execute SQL, AddNew
            End If
        Else
            Err.Raise 40025, "AddNew", E40025 'Mismatched Number of Elements. FieldList and Values must contain the same number of elements.
            Exit Function
        End If
    Else
        'set edit mode to "add"
        MyEditMode = MY_EDIT_ADD
        'if we haven't done any add's yet, make a new update object for adding
        If MyAdd Is Nothing Then Set MyAdd = New MYSQL_UPDATE
        'create new add record so we can start "letting" fields
        MyAdd.Add
        AddNew = 0
    End If
Exit Function
eh:
    AddNew = -1
    Err.Raise Err.Number, "AddNew", "Error While Executing AddNew" & vbCrLf & Err.Description
End Function

Public Function Update(Optional FieldList As Variant, Optional Values As Variant) As Long
Attribute Update.VB_Description = "Updates any changes to the current row or adds a new record."
'---------------------------------------------------------------------------------------
' Procedure : Update
' DateTime  : 02/17/04 21:50
' Author    : Robert Rowe (originally by Eric Grau)
' Purpose   : Update modifications to the current table. If provided with a FieldList
'             and Values then the fields in the FieldList array are set to the values
'             in the Values array before the update occurs. The error traps and the
'             support for providing Field and Value arrays were added by Robert Rowe.
'---------------------------------------------------------------------------------------

Dim FieldCount As Long, ValueCount As Long
Dim i As Long
    
    Update = -1
    
    'can't add if rs is closed
    If myState = MY_RS_CLOSED Then
        Err.Raise 40005, "Update", E40005 '"Recordset Closed."
        Exit Function
    End If
    
    'can't add if there's more than 1 table in rs
    If Me.TableCount <> 1 Then
        Err.Raise 40008, "AddNew", E40008 '"Cannot Add, Update or Delete if source query is based on multiple tables."
        Exit Function
    End If
    
    If Not IsMissing(FieldList) And Not IsMissing(Values) Then
        FieldCount = SafeBoundsCheck(FieldList)
        ValueCount = SafeBoundsCheck(Values)
        
        If FieldCount = ValueCount Then
            If FieldCount = 0 Then 'They aren't arrays
                If Len(Trim$(FieldList)) = 0 Then
                    'We are missing a field
                    Err.Raise 40006, "AddNew", E40006 'Invalid Field Specified.
                    Exit Function
                Else
                    Me.Fields(FieldList) = Values
                End If
            Else 'They are arrays
                'Build Field and Value strings
                For i = LBound(FieldList) To UBound(FieldList)
                    Me.Fields(FieldList(i)) = Values(i)
                Next i
            End If
        Else
            Err.Raise 40025, "AddNew", E40025 'Mismatched Number of Elements. FieldList and Values must contain the same number of elements.
            Exit Function
        End If
    End If

    Select Case Me.EditMode
    Case MY_EDIT_IN_PROGRESS
        'can't update if we aren't pointing at a record
        If BOF Or EOF Then
            Err.Raise 40015, "Update", E40015 '"No Current Record. The requested operation requires a current record and either BOF or EOF are true."
            Exit Function
        End If
        
        'update the record
        Update = DoUpdate
        If Update > 0 Then
            'success ...
            SetUpdatedFlags
        End If
    Case MY_EDIT_ADD
        'add the record
        Update = DoAdd
        If Update < 1 Then
            'failure ...
            Me.MyAddObject.Remove Me.MyAddObject.Count
        End If
    Case Else
        Exit Function                                  'shouldn't be here
    End Select

    'done with edit ...
    Me.EditMode = MY_EDIT_NONE
End Function

Public Function Delete() As Long
Attribute Delete.VB_Description = "Delete the current row."
'---------------------------------------------------------------------------------------
' Procedure : Delete
' DateTime  : 02/18/04 20:05
' Author    : Robert Rowe (originally by Eric Grau)
' Purpose   : Delete the current row.  Returns -1 if a delete is not allowed otherwise
'             the return value is the number of records affected in the database.
'             Error traps added by Robert Rowe
'---------------------------------------------------------------------------------------

    Delete = -1
    
    'can't add if rs is closed
    If myState = MY_RS_CLOSED Then
        Err.Raise 40005, "Delete", E40005 '"Recordset Closed."
        Exit Function
    End If
    
    'can't delete if we aren't pointing at a record
    If BOF Or EOF Then
        Err.Raise 40015, "Delete", E40015 '"No Current Record. The requested operation requires a current record and either BOF or EOF are true."
        Exit Function
    End If
    
    'shouldn't be here if any edit is in progress
    If Me.EditMode <> MY_EDIT_NONE Then
        Err.Raise 40007, "Delete", E40007 '"Add/Edit in progress. Call CancelUpdate first."
        Exit Function
    End If
    
    'can't add if there's more than 1 table in rs
    If Me.TableCount <> 1 Then
        Err.Raise 40008, "Delete", E40008 '"Cannot Add, Update or Delete if source query is based on multiple tables."
        Exit Function
    End If

    'delete the record
    Delete = DoDelete

    If Delete > 0 Then
        'success ...
        If MyDelete Is Nothing Then Set MyDelete = New MYSQL_UPDATE
        'track the deleted record
        MyDelete.Add Me.AbsolutePosition
        
        'Upp the Deleted counter
        mlNumDeleted = mlNumDeleted + 1
    End If
End Function

Public Sub CancelUpdate()
Attribute CancelUpdate.VB_Description = "Cancels any edits in progress."
'---------------------------------------------------------------------------------------
' Procedure : CancelUpdate
' DateTime  : 02/18/04 19:59
' Author    : Robert Rowe (originally by Eric Grau)
' Purpose   : Acts as a public interface to the routine that actually cancels an update.
'             Error traps added by Robert Rowe.
'---------------------------------------------------------------------------------------

    If myState = MY_RS_CLOSED Then
        Err.Raise 40005, "CancelUpdate", E40005 '"Recordset Closed."
        Exit Sub
    End If
    
    'can't cancel if we aren't pointing at a record
    If BOF Or EOF Then
        Err.Raise 40015, "CancelUpdate", E40015 '"No Current Record. The requested operation requires a current record and either BOF or EOF are true."
        Exit Sub
    End If
    
    DoCancelUpdate
    
End Sub
Public Sub Requery()
Attribute Requery.VB_Description = "Requeries the record set."
'---------------------------------------------------------------------------------------
' Procedure : Requery
' DateTime  : ???
' Author    : Eric Grau with modifications by Robert Rowe
' Purpose   : Requeries the record set.
'             Code to close the recordset first added by Robert Rowe.
'---------------------------------------------------------------------------------------

    'Close existing recordset
    Me.CloseRecordset
        
    'execute the current query again ...
    MyCONN.InternalExecute Me, msQueryString
End Sub

Private Sub Class_Initialize()
    'create a new field object
    Set MyFIELDS = GetNewField
    'set the rs state to closed
    myState = MY_RS_CLOSED
End Sub

Private Sub Class_Terminate()
    'insures that the rs is closed
    RealClose
    'release my objects
    Set MyCONN = Nothing
    Set MyFIELDS = Nothing
End Sub

Private Sub RealClose()
'---------------------------------------------------------------------------------------
' Procedure : RealClose
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Routine that does the actual close of the recordset.
'---------------------------------------------------------------------------------------

    'release result set memory
    mysql_free_result mlMYSQL_RES

    'result my module wide variables
    mlMYSQL_RES = 0
    mlMYSQL_ROW = 0
    mlMYSQL_FIELD_LENGTHS = 0
    mlCurrentRecord = 0
    mlRecordCount = 0
    mlFieldCount = 0
    mlAffectedRecords = 0
    msQueryString = ""
    mlNumDeleted = 0

    'create a new field object
    Set MyFIELDS = GetNewField

    'release my edit objects
    Set MyAdd = Nothing
    Set MyReplace = Nothing
    Set MyUpdate = Nothing
    Set MyDelete = Nothing

    'rs is now closed
    myState = MY_RS_CLOSED
    'Debug.Print "MySQL RS closed."
End Sub

Private Sub CheckAutoUpdate()
'---------------------------------------------------------------------------------------
' Procedure : CheckAutoUpdate
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Called when moving off a record. Triggers either an update or a cancel
'             of the current update.
'---------------------------------------------------------------------------------------

    If MyCONN.AutoUpdate Then
        'auto update is enabled ...
        'update the record as needed
        Update
    Else
        'cancel any updated in progress
        DoCancelUpdate
    End If

    'done with edit ...
    MyEditMode = MY_EDIT_NONE
End Sub

Private Function GetNewField() As MYSQL_FIELD
'---------------------------------------------------------------------------------------
' Procedure : GetNewField
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Resets the current field object.
'---------------------------------------------------------------------------------------

Dim MyField As MYSQL_FIELD
    'create and return a new field object
    Set MyField = New MYSQL_FIELD
    Set MyField.rs = Me
    Set GetNewField = MyField
End Function

Private Sub GetRow(Optional bRetainEditObjects As Boolean)
'---------------------------------------------------------------------------------------
' Procedure : GetRow
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : The the current row from the MySQL data structures.
'---------------------------------------------------------------------------------------

Dim cSeekRow As Currency
    'auto update as needed ...
    If Not bRetainEditObjects Then CheckAutoUpdate

    'are we on a valid record ???
    If mlCurrentRecord > 0 And mlCurrentRecord <= mlRecordCount Then

        'adjust because currency fields have 4 fixed decimals
        ' 1.000 gets adjusted to 0.001
        cSeekRow = (mlCurrentRecord - 1) / (10 ^ 4)

        'move to the specified row ...
        mysql_data_seek mlMYSQL_RES, cSeekRow

        'get pointer to row ...
        mlMYSQL_ROW = mysql_fetch_row(mlMYSQL_RES)
        'Debug.Print "Pointer to mysql row = " & mlMYSQL_ROW

        'get pointer to array of field lengths ...
        mlMYSQL_FIELD_LENGTHS = mysql_fetch_lengths(mlMYSQL_RES)
        'Debug.Print "Pointer to array of column lengths = " & mlMYSQL_FIELD_LENGTHS
    Else
        'not on a valid record ...
        mlMYSQL_ROW = 0
        mlMYSQL_FIELD_LENGTHS = 0
    End If
End Sub

Private Function BuildTableArray() As Variant
'---------------------------------------------------------------------------------------
' Procedure : BuildTableArray
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Constructs an array of the tables referenced in the query used to open
'             the recordset
'---------------------------------------------------------------------------------------

Dim lCnt As Long
Dim lCnt2 As Long
Dim sTbl As String
Dim vTbl As Variant
Dim bInArr As Boolean

    'check the table name associated with each field ...
    'track unique table names in array ...
    For lCnt = 0 To (FieldCount - 1)
        sTbl = Trim(Fields(lCnt).Table)
        If Len(sTbl) > 0 Then
            If Not IsArray(vTbl) Then
                ReDim vTbl(0)
                vTbl(0) = sTbl
            Else
                bInArr = False
                lCnt2 = 0
                Do
                    If vTbl(lCnt2) = sTbl Then
                        bInArr = True
                        Exit Do
                    End If
                    lCnt2 = lCnt2 + 1
                Loop While lCnt2 <= UBound(vTbl)

                If Not bInArr Then
                    ReDim Preserve vTbl(0 To (UBound(vTbl) + 1))
                    vTbl(UBound(vTbl)) = sTbl
                End If
            End If
        End If
    Next lCnt

    'return the array
    BuildTableArray = vTbl
End Function

Private Function DoAdd() As Long
'---------------------------------------------------------------------------------------
' Procedure : DoAdd
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Inserts a new record. Called by the Update method.
'---------------------------------------------------------------------------------------

Dim SQL As String
Dim sTable As String
Dim sFld As String
Dim sVal As String

    'get the name of the table
    sTable = Me.Fields(0).Table
    
    'gotta have a table to do the add
    If Len(sTable) = 0 Then
        Err.Raise 40013, "DoAdd", E40013 '"Could not identify the table to Insert Into."
        Exit Function
    End If
    
    'get the field and value lists for the insert statement
    GetAddString sFld, sVal
    
    'gotta have a field and value list
    If Len(sFld) = 0 Or Len(sVal) = 0 Then
        Err.Raise 40014, "DoAdd", E40014 '"Cannot Insert new record. You must set the value of at least one field."
        Exit Function
    End If

    SQL = "INSERT INTO " & sTable & " (" & sFld & ") VALUES(" & sVal & ")"
    MyCONN.Execute SQL, DoAdd
End Function

Private Function DoUpdate() As Long
'---------------------------------------------------------------------------------------
' Procedure : DoUpdate
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Performs the actual update of a modified record. Called by the Update method.
'---------------------------------------------------------------------------------------

Dim SQL As String
Dim sTable As String
Dim sUpd As String
Dim sWhere As String

    'get the name of the table
    sTable = Me.Fields(0).Table
    
    'gotta have a table to do the add
    If Len(sTable) = 0 Then
        Err.Raise 40011, "DoUpdate", E40011 '"Could not identify the table to update."
        Exit Function
    End If
    
    'get the field list for the update statement
    GetUpdateString sUpd
    
    'gotta have an update list (leave quiet as this just means that nothing has changed)
    If Len(sUpd) = 0 Then Exit Function
    
    'make the where clause for this row
    BuildWhere sWhere
    
    'gotta have a where clause
    If Len(sWhere) = 0 Then
        Err.Raise 40012, "DoUpdate", E40012 '"Could not identify the record to update. Include the Primary Key in your query."
        Exit Function
    End If
    
    SQL = "UPDATE " & sTable & " SET " & sUpd & " WHERE " & sWhere
    MyCONN.Execute SQL, DoUpdate
End Function

Private Function DoDelete() As Long
'---------------------------------------------------------------------------------------
' Procedure : DoDelete
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Actually performs the deletion. Called by the Delete method.
'---------------------------------------------------------------------------------------

Dim SQL As String
Dim sTable As String
Dim sWhere As String

    'get the name of the table
    sTable = Me.Fields(0).Table
    'gotta have a table to do the add
    If Len(sTable) = 0 Then
        Err.Raise 40009, "DoDelete", E40009 '"Could not identify the table to delete from."
        Exit Function
    End If
    
    'make the where clause for this row
    BuildWhere sWhere
    'gotta have a where clause
    If Len(sWhere) = 0 Then
        Err.Raise 40010, "DoDelete", E40010 '"Could not identify the record to delete. Include the Primary Key in your query."
        Exit Function
    End If

    SQL = "DELETE FROM " & sTable & " WHERE " & sWhere
    'Debug.Print sql
    MyCONN.Execute SQL, DoDelete
End Function

Private Sub GetAddString(ByRef sFields As String, ByRef sValues As String)
'---------------------------------------------------------------------------------------
' Procedure : GetAddString
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Constructs the field and value lists for an Insert query based on the
'             fields of the recordset.
'---------------------------------------------------------------------------------------

Dim lCnt As Long
Dim uFld As New MYSQL_UPDATE_FIELD

    If Me.MyAddObject.Count < 1 Then Exit Sub 'no fields were "let" for this add ...

    'build the field and value lists based on the fields that were "let"
    For lCnt = 1 To Me.MyAddObject.Item(Me.MyAddObject.Count).Count
        Set uFld = Me.MyAddObject.Item(Me.MyAddObject.Count).Item(lCnt)
        sFields = sFields & IIf(Len(sFields) > 0, ",", "") & uFld.Name
        sValues = sValues & IIf(Len(sValues) > 0, ",", "") & DbTick(uFld.Value, Fields(uFld.Name).MySqlType)
    Next lCnt
End Sub

Private Sub GetUpdateString(ByRef sUpdate As String)
'---------------------------------------------------------------------------------------
' Procedure : GetUpdateString
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : builds the Set clause of an Update query. Called by DoUpdate.
'---------------------------------------------------------------------------------------

    Dim lCnt As Long
    Dim uFld As New MYSQL_UPDATE_FIELD

    If Me.MyUpdateObject.Count < 1 Then Exit Sub           'no fields were "let" for this update ...

    'build the field list based on the fields that were "let" and haven't been updated yet ...
    For lCnt = 1 To Me.MyUpdateObject.Item(mlUpdateRecord, True).Count
        Set uFld = Me.MyUpdateObject.Item(mlUpdateRecord, True).Item(lCnt)
        If uFld.Updated Then 'Not
            sUpdate = sUpdate _
                    & IIf(Len(sUpdate) > 0, ",", "") _
                    & uFld.Name & "=" & DbTick(uFld.Value, Me.Fields(uFld.Name).MySqlType)  'MyFIELDS.MySqlType)
        End If
    Next lCnt
End Sub

Private Sub BuildWhere(ByRef sWhere As String)
'---------------------------------------------------------------------------------------
' Procedure : BuildWhere
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Constructs a where clause used to identify the current record.
'             Used for updating the record.
'---------------------------------------------------------------------------------------

Dim lCnt As Long

    'build the where clause for each primary key field in the current record
    For lCnt = 0 To (FieldCount - 1)
        If (Fields(lCnt).Flags And MY_PRI_KEY_FLAG) Then
            sWhere = sWhere & IIf(Len(sWhere) > 0, " AND ", "") _
                    & Fields(lCnt).Name & IIf(IsNull(Fields(lCnt).Value), " IS ", "=") & DbTick(Fields(lCnt).Value, Fields(lCnt).MySqlType)
        End If
    Next lCnt
End Sub

Private Sub SetUpdatedFlags()
'---------------------------------------------------------------------------------------
' Procedure : SetUpdatedFlags
' DateTime  : ???
' Author    : Eric Grau
' Purpose   : Sets the status of all Update objects to updated. Called by Update.
'---------------------------------------------------------------------------------------

Dim lCnt As Long
    'Dim uFld As New MYSQL_UPDATE_FIELD

    If Me.MyUpdateObject.Count < 1 Then Exit Sub           'no fields were "let" for this update ...

    'set updated = true for each update field
    For lCnt = 1 To Me.MyUpdateObject.Item(mlUpdateRecord, True).Count
        Me.MyUpdateObject.Item(mlUpdateRecord, True).Item(lCnt).Updated = True
    Next lCnt
End Sub

Public Function DbTick(ByVal v As Variant, fType As MYSQL_FIELD_TYPES) As String
'---------------------------------------------------------------------------------------
' Procedure : DbTick
' DateTime  : 02/08/04 21:26
' Author    : Robert Rowe (originally by Eric Grau)
' Purpose   : Escapes and delimits a fields value to prepare it for a query
'---------------------------------------------------------------------------------------

    Dim sFormat As String

    If IsNull(v) Then
        DbTick = "NULL"
        Exit Function
    End If

    Select Case fType
        Case MYSQL_DATE:
            sFormat = "YYYY-MM-DD"
            DbTick = "0000-00-00"
        Case MYSQL_TIME:
            sFormat = "HH:MM:SS"
            DbTick = "00:00:00"
        Case MYSQL_DATETIME:
            sFormat = "YYYY-MM-DD HH:MM:SS"
            DbTick = "0000-00-00 00:00:00"
    End Select

    Select Case True
    Case Len(sFormat) > 0
        If IsDate(v) Then DbTick = Format(v, sFormat)
    Case fType = MYSQL_BLOB Or fType = MYSQL_MEDIUMBLOB Or fType = MYSQL_TINYBLOB
        DbTick = v
    Case Else
        'Changed to use MySQL's API escaping routines
        'Yields a slight performance gain and makes
        'the MySQL escaping routines available to
        'the developer when building ad hoc queries
        'Robert Rowe 02/07/04
        DbTick = MyCONN.Esc(v)
    End Select
    
'Original code
'    If Len(sFormat) > 0 Then
'        If IsDate(v) Then DbTick = Format(v, sFormat)
'    Else
'        DbTick = Replace(v, "\", "\\", , , vbTextCompare)  'this must be the first Replace
'        DbTick = Replace(DbTick, Chr(0), "\0", , , vbTextCompare)
'        DbTick = Replace(DbTick, Chr(8), "\b", , , vbTextCompare)
'        DbTick = Replace(DbTick, Chr(9), "\t", , , vbTextCompare)
'        DbTick = Replace(DbTick, "'", "\'", , , vbTextCompare)
'        DbTick = Replace(DbTick, """", "\""", , , vbTextCompare)
'        DbTick = Replace(DbTick, Chr(10), "\n", , , vbTextCompare)
'        DbTick = Replace(DbTick, Chr(13), "\r", , , vbTextCompare)
'    End If

    DbTick = "'" & DbTick & "'"
End Function

Public Function ReadBinaryFromFile(txtFile As String) As Variant
'---------------------------------------------------------------------------------------
' Procedure : ReadBinaryFromFile
' DateTime  : 02/08/04 21:19
' Author    : Robert Rowe
' Purpose   : Reads a binary file from disk and formats it such that the result
'             can be assigned to a BLOB field of a recordset
'---------------------------------------------------------------------------------------

Dim vtData() As Byte
Dim nfile As Integer
On Error GoTo eh
    nfile = FreeFile
    Open txtFile For Binary As nfile
    ReDim vtData(LOF(nfile))
    Get nfile, , vtData
    Close nfile
    ReadBinaryFromFile = BytesToString(vtData)
Exit Function
eh:
    ReadBinaryFromFile = False
    Err.Raise Err.Number, "ReadBinaryFromFile", "Error While Reading: " & txtFile & vbCrLf & Err.Description
End Function

Public Function SaveBinaryToFile(FileContents As Variant, FName As String) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : SaveBinaryToFile
' DateTime  : 02/08/04 21:17
' Author    : Robert Rowe
' Purpose   : Used to write the contents of a BLOB field to disk
'---------------------------------------------------------------------------------------

Dim vtData() As Byte
Dim FNum As Integer
On Error GoTo eh
    vtData() = ObjectToByteArray(FileContents)
    If IsArrayByteEmpty(vtData) Then
        SaveBinaryToFile = False
        Exit Function
    End If
    FNum = FreeFile
    Open FName For Binary Access Write As FNum
        Put #FNum, , vtData
    Close FNum
    SaveBinaryToFile = True
Exit Function
eh:
    SaveBinaryToFile = False
    Err.Raise Err.Number, "SaveBinaryToFile", "Error While Saving: " & FName & vbCrLf & Err.Description
End Function
Private Function BytesToString(vtData() As Byte) As String
'---------------------------------------------------------------------------------------
' Procedure : BytesToString
' DateTime  : 02/08/04 21:20
' Author    : Robert Rowe
' Purpose   : Reformats data read from a binary file to make it acceptable for
'             assignment to a BLOB field (used by ReadBinaryFromFile)
'---------------------------------------------------------------------------------------

Dim i As Long
Dim theTemp As String
    theTemp = String(UBound(vtData) + 1, " ")
    For i = 0 To UBound(vtData)
        Mid(theTemp, i + 1, 1) = Chr(vtData(i))
    Next i
    BytesToString = theTemp
End Function

Private Function ObjectToByteArray(obj As Variant) As Byte()
'---------------------------------------------------------------------------------------
' Procedure : ObjectToByteArray
' DateTime  : 02/08/04 21:22
' Author    : Robert Rowe
' Purpose   : Converts a variant of type object to a byte array. Necessary for saving a
'             BLOB field to a file. (used by SaveBinaryToFile)
'---------------------------------------------------------------------------------------

Dim i As Long
Dim vtData() As Byte
Dim objLength As Long
Dim v As Variant
    v = obj
    objLength = Len(v) - 1 'The -1 is to strip the extra terminating NULL
    If objLength < 0 Then Exit Function 'Field was blank
    ReDim vtData(0 To objLength - 1)
    For i = 0 To objLength - 1
        vtData(i) = Asc(Mid$(v, i + 1, 1))
    Next i
    ObjectToByteArray = vtData
End Function

