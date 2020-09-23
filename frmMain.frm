VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "mySQL Query Analyzer..."
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10500
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cmnDLG 
      Left            =   6720
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar barMenu 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New Query Window"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open Query"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save Query"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print Query"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Execute"
            Object.ToolTipText     =   "Execute Query"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop Query Execution"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   5000
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageCombo cboDatabases 
         Height          =   330
         Left            =   3840
         TabIndex        =   9
         Top             =   0
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         ImageList       =   "ImageList1"
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7320
      Top             =   6120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":014A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":02A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":03FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0558
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":06B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":080C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0966
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0AC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D74
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0ECE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1028
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1182
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1436
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1590
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1844
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   7035
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10724
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Text            =   "Connections: 0"
            TextSave        =   "Connections: 0"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picObjects 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6675
      Left            =   0
      ScaleHeight     =   6675
      ScaleWidth      =   3780
      TabIndex        =   0
      Top             =   360
      Width           =   3780
      Begin MSComctlLib.ImageCombo cboObjects 
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         ImageList       =   "ImageList1"
      End
      Begin TabDlg.SSTab tabBrowser 
         Height          =   5775
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   10186
         _Version        =   393216
         TabOrientation  =   1
         Style           =   1
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         TabCaption(0)   =   "Objects"
         TabPicture(0)   =   "frmMain.frx":1E12
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "treeObjects"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin MSComctlLib.TreeView treeObjects 
            Height          =   5175
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   9128
            _Version        =   393217
            Indentation     =   441
            LineStyle       =   1
            Sorted          =   -1  'True
            Style           =   7
            ImageList       =   "ImageList1"
            Appearance      =   1
         End
      End
      Begin VB.PictureBox picSeperatorBar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6375
         Left            =   3720
         MousePointer    =   9  'Size W E
         ScaleHeight     =   6375
         ScaleWidth      =   75
         TabIndex        =   3
         Top             =   50
         Width           =   75
      End
      Begin VB.CommandButton cmdHideObjects 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   2
         Top             =   50
         Width           =   255
      End
      Begin VB.Label lblFieldHeading 
         AutoSize        =   -1  'True
         Caption         =   "&Object Browser"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         X1              =   4320
         X2              =   0
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   4320
         X2              =   0
         Y1              =   10
         Y2              =   10
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New Query"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open Query"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Query"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print Query"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConnectToServer 
         Caption         =   "Connect To Server"
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuObjectBrowser 
         Caption         =   "Object Browser"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu mnuQuery 
      Caption         =   "&Query"
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFindNext 
         Caption         =   "Find Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuQuerySep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExecute 
         Caption         =   "Execute"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuStopExecution 
         Caption         =   "Halt Execution"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuCascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu mnuTileHorizontal 
         Caption         =   "Tile Horizontal"
      End
      Begin VB.Menu mnuTileVertical 
         Caption         =   "Tile Vertical"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub barMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "New"                  '// Open a new query window
            Call subNew_Query
            
        Case "Open"                 '// Open a .SQL file into a query window
            Call subOpen_Query
                    
        Case "Save"                 '// Save the current active query form
            Call subSave_Query
            
        Case "Print"                '// Print the current active query form
            Call mnuPrint_Click
        
        Case "Cut"                  '// Cut to clipboard
            Call mnuCut_Click
        
        Case "Copy"                 '// Copy to clipboard
            Call mnuCopy_Click
        
        Case "Paste"                '// Paste from clipboard
            Call mnuPaste_Click
        
        Case "Find"                 '// Find text in query (or result)
            frmFind.Show 1
        
        Case "Execute"              '// Execute SQL query
            Call ActiveForm.subRunQuery
        
        Case "Stop"                 '// Stop Execution of query
            gblnCancel = True
    End Select
End Sub

Public Sub subNew_Query(Optional strTitle As String = "")
    Dim frmNewQuery         As New frmQueryWindow
        
    On Error GoTo New_Query_Errors
        Load frmNewQuery                    '// Create a new instance of the child form
        
        With frmNewQuery
            If .WindowState = vbNormal Then
                .Left = .Left + 200
                .Top = .Top + 200
            End If
            
            '// If program does not pass in a caption title, create one
            If Trim(strTitle) = "" Then
                If Trim(cboDatabases.Text) <> "" Then
                    .Caption = "Query - " & UCase(gstrHostName) & "\" & Trim(frmMain.cboDatabases.Text) & " - Untitled*"
                Else
                    '// Allow for the user to not have to connect to a DB (this shouldn't happen but allow it)
                    '//  They will not be able to run this query until it is connected to a database
                    '//  unless they specify the database name in the query
                    .Caption = "Query - NOT CONNECTED - Untitled*"
                End If
            Else
                .Caption = "Query - " & strTitle
            End If
            .Show
        End With
    On Error GoTo 0
    
New_Query_Errors:
    If Err.Number <> 0 Then Call subLog_Errors(Err.Number, Err.Description, "frmMain::subNew_Query")
End Sub

Private Sub subSave_Query(Optional blnSaveAs As Boolean = False)
    Dim strFileName_Save    As String   '// SQL Filename
    Dim strSFOName          As String   '// SFO Filename
    Dim fleQuery            As Integer  '// File Handle
    
    On Error GoTo Save_Cancel
        '// Open a save as dialog box the program specifically sends it as save as
        '//  or if it hasn't ever been saved before (thus not having a filename)
        If blnSaveAs Or ActiveForm.Tag = "" Then
            With cmnDLG
                .CancelError = True
                .DialogTitle = "Save Query As..."
                .Filter = "All SQL Files (*.SQL)|*.SQL|"
                .InitDir = App.Path & "\Save"
                .ShowSave
                
                strFileName_Save = .FileName
            End With
        Else
            strFileName_Save = ActiveForm.Tag
        End If
        '// Get the filename for the header (SFO file)
        strSFOName = Left(strFileName_Save, Len(strFileName_Save) - 4) & ".SFO"
    On Error GoTo Save_Errors
        
        '// Save the header information first
        fleQuery = FreeFile                         '// Get a file handle
        Open strSFOName For Output As #fleQuery     '// If the file exists overwrite it
        '// Store database name
        Print #fleQuery, "#DB#" & Trim(cboDatabases.Text)
        '// Store server host name
        Print #fleQuery, "#HS#" & gstrHostName
        Close #fleQuery                             '// Close the filehandle
        
        '// Now save the SQL file using the Rich Text Box's Save command
        ActiveForm.rtbQueryPane.SaveFile strFileName_Save
        
        '// Update the caption to remove the * so the user (and program) knows it's been saved
        ActiveForm.Caption = "Query - " & UCase(gstrHostName) & "\" & Trim(cboDatabases.Text) & " - " & strFileName_Save
        '// Update the save filename information
        ActiveForm.Tag = strFileName_Save
    On Error GoTo 0
    
Save_Cancel:
    Exit Sub
    
Save_Errors:
    If Err.Number <> 0 Then Call subLog_Errors(Err.Number, Err.Description, "frmMain::subSave_Query")
    Close #fleQuery
End Sub

Private Sub subOpen_Query()
    Dim strFileName_Open    As String   '// SQL Filename
    Dim strSFOName          As String   '// SFO Filename
    Dim fleQuery            As Integer  '// File Handle
    Dim strLineIn           As String   '// Contains all information from 1 line (read until CRLF)
    
    On Error GoTo Open_Cancel
        With cmnDLG
            .CancelError = True             '// If the user clicks cancel it will cause an error so we can jump out of the sub
            .DialogTitle = "Open Query..."  '// Dialog Caption
            .Filter = "All SQL Files (*.SQL)|*.SQL|"
            .InitDir = App.Path & "\Save"   '// Initial directory to search
            .ShowOpen                       '// Open the "Open File" Dialog
            
            strFileName_Open = .FileName    '// Get user inputted file name
            
            '// Remove the .SQL extension at the end and replace it with a .SFO extension, this is the header information file
            strSFOName = Left(strFileName_Open, Len(strFileName_Open) - 4) & ".SFO"
        End With
    On Error GoTo Open_Errors
    
        If Dir(strSFOName) = "" Then MsgBox "Error! Header information missing cannot open file!", vbCritical + vbOKOnly, "Error...": Exit Sub
        fleQuery = FreeFile
        Open strSFOName For Input As #fleQuery
        Line Input #fleQuery, strLineIn
        
        '// Get the database name
        If Left(Trim(strLineIn), 4) <> "#DB#" Then MsgBox "Error! File structure not understood.  Cannot open file!", vbExclamation + vbOKOnly, "Error...": Exit Sub
        cboDatabases.Text = Right(Trim(strLineIn), Len(Trim(strLineIn)) - 4)
        Call cboDatabases_Click
                
        '// Get the host name
        Line Input #fleQuery, strLineIn
        If Left(Trim(strLineIn), 4) <> "#HS#" Then MsgBox "Error! File structure not understood.  Cannot open file!", vbExclamation + vbOKOnly, "Error...": Exit Sub
        '// In a later version, add multiple host recognition
        If gstrHostName <> Right(Trim(strLineIn), Len(Trim(strLineIn)) - 4) Then
            MsgBox "The current version of myAnalyzer only supports one host.  The file you have selected is connected to a differnet host. " & _
                "Please select a file that is connected to this host.", vbCritical + vbOKOnly, "Cannot Connect"
        End If
        Close #fleQuery
        
        '// Open the query (display the hostname and database name as well as the filename)
        Call subNew_Query(gstrHostName & "\" & Trim(cboDatabases.Text) & " - " & strFileName_Open)
        ActiveForm.rtbQueryPane.LoadFile strFileName_Open
        ActiveForm.Caption = Left(ActiveForm.Caption, Len(ActiveForm.Caption) - 1)
        '// Store the filename in the Child Form's tag property for Save purposes (if it's not there is will assume a Save As)
        ActiveForm.Tag = strFileName_Open
        
    On Error GoTo 0
    
Open_Cancel:
    Exit Sub
    
Open_Errors:
    If Err.Number <> 0 Then Call subLog_Errors(Err.Number, Err.Description, "frmMain::subOpen_Query")
    Close #fleQuery
End Sub

Private Sub cboDatabases_Click()
    Call subUpdate_DatabaseLists
End Sub

Public Sub subUpdate_DatabaseLists()
    '// Update the object combo with the current database
    If Trim(cboDatabases.Text) <> "" Then
        cboObjects.ComboItems(cboDatabases.Text).Selected = True
        Call cboObjects_Click
    End If
End Sub

Private Sub cboObjects_Click()
    '///////////////////////////////////////////////////////////////////////////////
    '// When a database is selected from the combo box, populate the tree view with
    '//  the tables and fields from that database for usage in the query window.
    '///////////////////////////////////////////////////////////////////////////////
    Dim myRS_Tables     As MYSQL_RS     '// mySQL Recordset for Table Information
    Dim myRS_Fields     As MYSQL_RS     '// mySQL Recordset for Field Information
    Dim i               As Integer      '// Standard Incrimentor
    
    On Error GoTo Object_Click_Errors
        If Trim(cboObjects.Text) <> "" Then
            '// Check connection state to ensure an active database connection
            If gmyConnection.State = MY_CONN_CLOSED Then MsgBox "Error! Database connection has been severed.  Please close this program and restart.", vbCritical + vbOKOnly, "Error...": Exit Sub
            '// Ensure that the database exists on this server (just in case)
            If Not gmyConnection.SelectDb(Trim(cboObjects.Text)) Then MsgBox "Error! The database entered could not be accessed!", vbExclamation + vbOKOnly, "Error...": Exit Sub
        
            treeObjects.Nodes.Clear     '// Clear all information from the treeview
            '// Add the database name as the parent
            treeObjects.Nodes.Add , , "DB_PARENT", Trim(cboObjects.Text), 1
            
            Set myRS_Tables = New MYSQL_RS  '//-|--Initialize recordset objects
            Set myRS_Fields = New MYSQL_RS  '//-|
            
            Set myRS_Tables = gmyConnection.Show(MY_SHOW_TABLES) '// Get table information from the database
            While Not myRS_Tables.EOF
                '// Store the table definition under its corresponding Database parent
                treeObjects.Nodes.Add "DB_PARENT", tvwChild, "TABLE" & Trim(myRS_Tables.Fields(0).Value), Trim(myRS_Tables.Fields(0).Value), 12
                                
                '// Run a query that returns only 1 record to get the field names
                Set myRS_Fields = gmyConnection.Execute("SELECT * FROM " & Trim(myRS_Tables.Fields(0).Value) & " LIMIT 1")
                If Not myRS_Fields.EOF Then
                    For i = 0 To myRS_Fields.FieldCount - 1
                        '// Store all field definitions under the current table definition node
                        treeObjects.Nodes.Add "TABLE" & Trim(myRS_Tables.Fields(0).Value), tvwChild, "FIELD" & Trim(myRS_Fields.Fields(i).Name) & "OFTAB" & Trim(myRS_Tables.Fields(0).Value), Trim(myRS_Fields.Fields(i).Name), 20
                    Next i
                End If
                myRS_Tables.MoveNext        '// Get next record
            Wend
            treeObjects.Nodes(1).Expanded = True
        End If
    On Error GoTo 0
    
Object_Click_Errors:
    If Err.Number <> 0 Then Call subLog_Errors(Err.Number, Err.Description, "frmMain::cboObjects_Click")
    'Resume Next
End Sub

Private Sub cmdHideObjects_Click()
    picObjects.Visible = False              '// Hide the object browser
End Sub

Private Sub MDIForm_Load()
    Dim frmNewQuery     As New frmQueryWindow

    On Error GoTo Form_Load_Errors
        treeObjects.Top = 50
        treeObjects.Left = 50
    On Error GoTo 0
    
Form_Load_Errors:
    If Err.Number <> 0 Then Call subLog_Errors(Err.Number, Err.Description, "frmMain::MDIForm_Load")
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    On Error Resume Next
        gmyConnection.CloseConnection
        Set gmyConnection = Nothing
    On Error GoTo 0
    End
End Sub

Private Sub mnuCopy_Click()
    Clipboard.Clear
    Clipboard.SetText ActiveForm.ActiveControl.SelText
End Sub

Private Sub mnuCut_Click()
    Clipboard.Clear
    Clipboard.SetText ActiveForm.ActiveControl.SelText
    ActiveForm.ActiveControl.SelText = ""
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub mnuCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuConnectToServer_Click()
    Dim frmChild        As Form
        
    For Each frmChild In Forms
        If Not (TypeOf frmChild Is MDIForm) Then
            Unload frmChild
        End If
    Next
    frmServerConnect.Show
End Sub

Private Sub mnuExecute_Click()
    Call ActiveForm.subRunQuery
End Sub

Private Sub mnuExit_Click()
    Unload Me
    End
End Sub

Private Sub mnuFind_Click()
    frmFind.Show 1
End Sub

Private Sub mnuFindNext_Click()
    Call subFindNext
End Sub

Private Sub mnuNew_Click()
    Call subNew_Query
End Sub

Private Sub mnuObjectBrowser_Click()
    If picObjects.Visible Then
        picObjects.Visible = False
        mnuObjectBrowser.Checked = False
    Else
        picObjects.Visible = True
        mnuObjectBrowser.Checked = True
    End If
End Sub

Private Sub mnuOpen_Click()
    Call subOpen_Query
End Sub

Private Sub mnuPaste_Click()
    ActiveForm.rtbQueryPane.SelText = Clipboard.GetText
End Sub

Private Sub mnuOptions_Click()
    frmOptions.Show 1
End Sub

Private Sub mnuPrint_Click()
    frmPrint.Show 1
End Sub

Private Sub mnuSave_Click()
    Call subSave_Query
End Sub

Private Sub mnuStopExecution_Click()
    gblnCancel = True
End Sub

Private Sub mnuTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub picObjects_Resize()
    '///////////////////////////////////////////////////////////////////////////////
    '// Resize everything inside of the Object Browser area to fit to proper ratio
    '///////////////////////////////////////////////////////////////////////////////
    On Error Resume Next
        picSeperatorBar.Height = picObjects.Height
        picSeperatorBar.Left = picObjects.Width - picSeperatorBar.Width
        
        Line1.X1 = picObjects.Width
        Line2.X1 = picObjects.Width
        
        tabBrowser.Height = picObjects.Height - 355 - StatusBar1.Height
        tabBrowser.Width = picObjects.Width - 285
        
        treeObjects.Width = tabBrowser.Width - 100
        treeObjects.Height = tabBrowser.Height - 400
        
        cboObjects.Width = tabBrowser.Width
        cmdHideObjects.Left = (cboObjects.Left + cboObjects.Width) - cmdHideObjects.Width
    On Error GoTo 0
End Sub

Private Sub picSeperatorBar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '///////////////////////////////////////////////////////////////////////////////
    '// Allow the user to resize the object browser
    '///////////////////////////////////////////////////////////////////////////////
    On Error Resume Next
        If Button = vbLeftButton Then
            picObjects.Width = picObjects.Width + x
            picSeperatorBar.Left = picObjects.Width - picSeperatorBar.Width
        End If
    On Error GoTo 0
End Sub

Private Sub treeObjects_DblClick()
    On Error GoTo tree_Click_Errors
        If treeObjects.SelectedItem.Text <> "" Then
            '// Dbl Clicking the treeview causes it to expand and retract, prevent this by reversing it
            '//  and only allowing a dbl click to move information into the query window pane
            treeObjects.SelectedItem.Expanded = IIf(treeObjects.SelectedItem.Expanded, False, True)
            
            '// Put the information into the active MDI child's query pane
            ActiveForm.rtbQueryPane.SelText = treeObjects.SelectedItem.Text
        End If
    On Error GoTo 0
    
tree_Click_Errors:
    If Err.Number <> 0 Then Call subLog_Errors(Err.Number, Err.Description, "frmMain::treeObjects_DblClick")
End Sub
