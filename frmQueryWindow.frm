VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmQueryWindow 
   AutoRedraw      =   -1  'True
   Caption         =   "<UNTITLED>"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7875
   Icon            =   "frmQueryWindow.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   7875
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7440
      Top             =   6480
   End
   Begin VB.PictureBox picResizeBar 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      FillColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   70
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   75
      ScaleWidth      =   7815
      TabIndex        =   2
      Top             =   4450
      Width           =   7815
   End
   Begin VB.TextBox txtResultsPane 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFEAEB&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "frmQueryWindow.frx":014A
      Top             =   4560
      Width           =   7815
   End
   Begin RichTextLib.RichTextBox rtbQueryPane 
      Height          =   4455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7858
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmQueryWindow.frx":0163
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmQueryWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mmyLocalConnection      As MYSQL_CONNECTION '// Local Active connection
Private blnIsActiveConnection   As Boolean          '// Keep track if it's an empty query or connected to a DB
Private lngCurrentTopLine       As Long
Private lngCurrentBottomLine    As Long

Private strDatabaseName         As String
Private strTableNames           As String
Private strFieldNames           As String

Private Sub Form_Activate()
    '///////////////////////////////////////////////////////////////////////////////
    '// When the form is activated, update the database combo box and the treeview
    '//  to show the database objects associated with the query window being selected
    '///////////////////////////////////////////////////////////////////////////////
    Dim strCaption      As String       '// Caption of the active window
    Dim intDBNameStart  As Integer      '// Used for determining the place of a character in the caption

    On Error GoTo Window_Activate_Errors
        strCaption = Me.Caption
        '// Get the first half of the database name by looking for the first backslash (\)
        intDBNameStart = InStr(1, UCase(strCaption), UCase(gstrHostName) & "\")
        If intDBNameStart > 0 Then
            strCaption = Mid(strCaption, intDBNameStart + Len(gstrHostName) + 1, Len(strCaption) - (intDBNameStart + Len(gstrHostName) + 1))
            
            '// Get the actual database name by getting everything returned from above before the dash (-)
            intDBNameStart = InStr(1, strCaption, "-")
            If intDBNameStart > 0 Then
                strCaption = Trim(Left(strCaption, intDBNameStart - 1))
                
                '// Update the combo boxes and treeviews
                frmMain.cboDatabases.ComboItems(strCaption).Selected = True
                Call frmMain.subUpdate_DatabaseLists
            End If
        End If
    On Error GoTo 0
    
Window_Activate_Errors:
    If Err.Number <> 0 Then Call subLog_Errors(Err.Number, Err.Description, "frmQueryWindow::Form_Activate")
End Sub

Private Sub Form_Load()
    Dim i       As Integer  '// Standard incrimentor

    On Error GoTo Window_Load_Errors
        Me.Height = 5670
        Me.Width = 9765
                
        blnIsActiveConnection = False
        If Trim(frmMain.cboDatabases.Text) <> "" Then
            Set mmyLocalConnection = New MYSQL_CONNECTION   '// Keep an open active connection for each query form when a database has been selected
            mmyLocalConnection.OpenConnection gstrHostName, gstrUserName, gstrPassword, Trim(frmMain.cboDatabases.Text)
            
            If mmyLocalConnection.State = MY_CONN_OPEN Then '// Ensure the connection is open
                blnIsActiveConnection = True
                
                gintConnections = gintConnections + 1       '// Keep tally of all the open connections
                '// Show the user how many open database connections they have
                frmMain.StatusBar1.Panels(2).Text = "Connections: " & gintConnections
                
                '// Set initial query data to a generic select statement
                With rtbQueryPane
                    .SelColor = gastrColorValues(cnsKEYWORD)
                    .SelBold = True
                    .SelText = "SELECT"
                    .SelColor = gastrColorValues(cnsGENERAL)
                    .SelBold = False
                    .SelText = " * "
                    .SelColor = gastrColorValues(cnsKEYWORD)
                    .SelBold = True
                    .SelText = "FROM "
                    .SelColor = gastrColorValues(cnsTABLE)
                    .SelBold = False
                    .SelText = Trim(frmMain.treeObjects.Nodes(2).Text) & ";"
                End With
            End If
            
            strDatabaseName = frmMain.treeObjects.Nodes(1).Text
            For i = 2 To frmMain.treeObjects.Nodes.Count
                If Left(frmMain.treeObjects.Nodes(i).Key, 5) = "TABLE" Then
                    strTableNames = strTableNames & "|" & frmMain.treeObjects.Nodes(i).Text
                ElseIf Left(frmMain.treeObjects.Nodes(i).Key, 5) = "FIELD" Then
                    strFieldNames = strFieldNames & "|" & frmMain.treeObjects.Nodes(i).Text
                End If
            Next i
        Else
            rtbQueryPane.SelColor = RGB(0, 0, 240)          '// base color
            rtbQueryPane.SelText = "This is the query pane" '// generic results text
        End If
    On Error GoTo 0

Window_Load_Errors:
    If Err.Number <> 0 Then Call subLog_Errors(Err.Number, Err.Description, "frmQueryWindow::Form_Load")
End Sub

Public Sub subRunQuery()
    '///////////////////////////////////////////////////////////////////////////////
    '// Run the query specified in the query window
    '///////////////////////////////////////////////////////////////////////////////
    Dim myRecordset         As MYSQL_RS     '// mySQl Recordset Object
    Dim strFieldValue       As String * 30  '// Fixed width field for output
    Dim strWriteValue       As String       '// Will contain all the information for output
    Dim i                   As Integer      '// Standard Incrimentor
    
    Dim astrSQLCommands()   As String       '// Array of SQL Commands
    Dim blnMultipleCommands As Boolean      '// Keep track if given more than one SQL Command
    Dim intQueryCount       As Integer      '// Keep track of the current query being executed
    
    On Error GoTo RunQuery_Errors
        frmMain.StatusBar1.Panels(1).Text = "Executing Query -- Please Wait..."
        DoEvents
    
        Set myRecordset = New MYSQL_RS
        txtResultsPane.Text = ""
        
        '// Seperate multiple query statements, this will still work even if there is only one query statement
        '//  and it doesn't have to have a semi-colon as the array will be returned with a ubound of 0
        astrSQLCommands = Split(rtbQueryPane.Text, ";")
                
        For intQueryCount = LBound(astrSQLCommands) To UBound(astrSQLCommands)
            '// Halt on blank query
            If Trim(astrSQLCommands(intQueryCount)) = "" Then Exit For
            
            '// Open a recordset object based off the query
            Set myRecordset = mmyLocalConnection.Execute(Trim(astrSQLCommands(intQueryCount)))
            
            If Not myRecordset.EOF Then         '// Check for End of Recordset
                If myRecordset.RecordCount > 10000 Then
                    If MsgBox("The current query will return more than 10,000.  It is advised that you limit your queries to 10,000 or less with the LIMIT 10000 command.  To continue click 'Yes' to cancel query execution click 'No'", vbYesNo, "WARNING: Large Query...") = vbNo Then
                        GoTo RunQuery_Errors
                    End If
                End If
                
                '// Skip down a line to print headings on a seperate line from the previous query results
                If intQueryCount > LBound(astrSQLCommands) Then txtResultsPane.SelText = vbCrLf & vbCrLf
                
                '// Display the names of the database fields
                For i = 0 To myRecordset.FieldCount - 1
                    strFieldValue = myRecordset.Fields(i).Name
                    txtResultsPane.SelText = strFieldValue
                Next i
                txtResultsPane.SelText = vbCrLf
                
                For i = 0 To myRecordset.FieldCount - 1
                    strFieldValue = "------------------- "
                    txtResultsPane.SelText = strFieldValue
                Next i
                txtResultsPane.SelText = vbCrLf
        
                While Not myRecordset.EOF
                    '// Check to see if user cancels the query execution
                    If gblnCancel Then
                        '// Offer the user the chance to reconsider canceling the execution
                        If MsgBox("Are you sure you wish to cancel the current query?", vbQuestion + vbYesNo, "Cancel...") = vbYes Then
                            GoTo RunQuery_Cancel
                        Else
                            gblnCancel = False
                        End If
                    End If
                    DoEvents
                    
                    '// Display the values of each field under their corresponding field name
                    For i = 0 To myRecordset.FieldCount - 1
                        strFieldValue = IIf(IsNull(Trim(myRecordset.Fields(i).Value)), "", Trim(myRecordset.Fields(i).Value))
                        txtResultsPane.SelText = strFieldValue
                    Next i
                    txtResultsPane.SelText = vbCrLf
                    
                    myRecordset.MoveNext        '// Advance to next record
                Wend
                
                txtResultsPane.SelText = vbCrLf & "-------------------" & vbCrLf
            End If
            txtResultsPane.SelText = myRecordset.AffectedRecords & " record(s) affected"
        Next intQueryCount
    On Error GoTo 0
    
RunQuery_Errors:
    If Err.Number <> 0 Then
        If mmyLocalConnection.Error.Number <> 0 Then    '// We want any errors returned from mySQL to show up in the results field
            txtResultsPane.Text = mmyLocalConnection.Error.Description
        Else                                            '// All other errors are logged normally
            Call subLog_Errors(Err.Number, Err.Description, "modGeneral::subRunQuery")
        End If
    End If
    
    On Error Resume Next
        myRecordset.CloseRecordset
        Set myRecordset = Nothing
    On Error GoTo 0
    frmMain.StatusBar1.Panels(1).Text = ""
    Exit Sub
    
RunQuery_Cancel:
    gblnCancel = False
    txtResultsPane.Text = "Action Canceled..."
    GoTo RunQuery_Errors
End Sub

Private Sub Form_Resize()
    '///////////////////////////////////////////////////////////////////////////////
    '// Keep current ratios on items in the form when it's resized
    '///////////////////////////////////////////////////////////////////////////////
    On Error Resume Next
        rtbQueryPane.Height = Me.Height - 2910
        rtbQueryPane.Width = Me.Width - 120
        
        txtResultsPane.Top = rtbQueryPane.Height + 105
        txtResultsPane.Height = Me.Height - txtResultsPane.Top - 400
        txtResultsPane.Width = rtbQueryPane.Width
        
        picResizeBar.Top = txtResultsPane.Top - 110
        picResizeBar.Width = rtbQueryPane.Width
    On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '///////////////////////////////////////////////////////////////////////////////
    '// Close the active connection and decriment the connections the user is aware of
    '///////////////////////////////////////////////////////////////////////////////
    On Error Resume Next
        If blnIsActiveConnection Then
            gintConnections = gintConnections - 1
            frmMain.StatusBar1.Panels(2).Text = "Connections: " & gintConnections
        End If
    
        mmyLocalConnection.CloseConnection
        Set mmyLocalConnection = Nothing
    On Error GoTo 0
End Sub

Private Sub picResizeBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '///////////////////////////////////////////////////////////////////////////////
    '// Give the ability to resize the query and results panel based off the resize
    '//  bar.
    '///////////////////////////////////////////////////////////////////////////////
    On Error Resume Next
        If Button = vbLeftButton Then
            picResizeBar.Top = picResizeBar.Top + y
            txtResultsPane.Top = txtResultsPane.Top + y
            txtResultsPane.Height = Me.Height - txtResultsPane.Top - 400
            
            rtbQueryPane.Height = rtbQueryPane.Height + y
        End If
    On Error GoTo 0
End Sub

Private Sub rtbQueryPane_Change()
    Dim lngLineCount        As Long

    On Error GoTo PaneChange_Errors
        '///////////////////////////////////////////////////////////////////////////////
        '// Keep track if the query has been changed to allow the user who may think
        '// he/she has saved his/her query to be alarmed if otherwise
        '///////////////////////////////////////////////////////////////////////////////
        If Right(Me.Caption, 1) <> "*" Then Me.Caption = Me.Caption & "*"
        Call subCheck_TextSyntax
    On Error GoTo 0
    
PaneChange_Errors:
    If Err.Number <> 0 Then Call subLog_Errors(Err.Number, Err.Description, "frmQueryWindow::rtbQueryPane_Change")
End Sub

Private Sub subCheck_TextSyntax()
    '///////////////////////////////////////////////////////////////////////////////
    '// Color code things to go along with SQL Syntax.  Only known problem with this
    '//  function is dealing with color coding after you copy/cut and paste
    '///////////////////////////////////////////////////////////////////////////////
    Dim lngCursorPosition   As Long '// Cursor position before code performs selection
    Dim lngStart            As Long '// The beginning of the current word
    Dim lngEnd              As Long '// The end of the current word
        
    On Error GoTo Check_Syntax_Errors
        With rtbQueryPane
            If Len(.Text) = 0 Then Exit Sub
            
            lngCursorPosition = .SelStart       '// Keep track of the original cursor position
            lngStart = .SelStart                '// Get the current position so we can find the beginning of the word
            
            '// Decriment the start variable until we have reached a white space character.
            '// This will tell us the beginning of the word
            If lngStart <> 0 Then
                Do While Trim(Mid(.Text, lngStart, 1)) <> "" And Trim(Mid(.Text, lngStart, 1)) <> ";" _
                    And Trim(Mid(.Text, lngStart, 1)) <> vbLf And Trim(Mid(.Text, lngStart, 1)) <> vbCr _
                    And Trim(Mid(.Text, lngStart, 1)) <> "(" And Trim(Mid(.Text, lngStart, 1)) <> ")" _
                    And Trim(Mid(.Text, lngStart, 1)) <> "," And Trim(Mid(.Text, lngStart, 1)) <> "="
                    
                        lngStart = lngStart - 1
                        If lngStart = 0 Then Exit Do
                Loop
            End If
            lngStart = lngStart + 1             '// Don't include the white space character
                
            lngEnd = 0
            '// Incriment the end variable until we find a white space character or until we have
            '//  reached the end of the text to search.
            While Trim(Mid(.Text, lngStart + lngEnd, 1)) <> "" And Trim(Mid(.Text, lngStart + lngEnd, 1)) <> ";" And Trim(Mid(.Text, lngStart + lngEnd, 1)) <> vbLf And Trim(Mid(.Text, lngStart + lngEnd, 1)) <> vbCr And lngStart + lngEnd <= Len(.Text)
                lngEnd = lngEnd + 1
            Wend
            
            '// Select the word
            .SelStart = lngStart - 1
            .SelLength = lngEnd
                      
            '// Set the color of the word.  If it is a keyword, then set the color
            '//  according to what is set below.
            If InStr(1, UCase(gstrKeyWords), UCase(Mid(.Text, lngStart, lngEnd))) Then
                .SelColor = gastrColorValues(cnsKEYWORD)
                .SelBold = True
            '// Table names are set to a different color
            ElseIf InStr(1, UCase(strTableNames), UCase(Mid(.Text, lngStart, lngEnd))) Then
                .SelColor = gastrColorValues(cnsTABLE)
                .SelBold = False
            '// Field names are set to a different color
            ElseIf InStr(1, UCase(strFieldNames), UCase(Mid(.Text, lngStart, lngEnd))) Then
                .SelColor = gastrColorValues(cnsFIELD)
                .SelBold = False
            Else
                .SelColor = gastrColorValues(cnsGENERAL)
                .SelBold = False
            End If
            
            .SelStart = lngCursorPosition       '// Put the cursor back at its original position
            .SelColor = RGB(0, 0, 0)            '// Set the color back to default black
            .SelBold = False                    '// Set the bold face back to default (off)
        End With
    On Error GoTo 0
    
Check_Syntax_Errors:
    If Err.Number <> 0 Then Call subLog_Errors(Err.Number, Err.Description, "frmQueryWindow::subCheck_TextSyntax")
End Sub
