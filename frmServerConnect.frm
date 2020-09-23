VERSION 5.00
Begin VB.Form frmServerConnect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Connect to Server..."
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmServerConnect.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Default         =   -1  'True
      Height          =   375
      Left            =   1868
      TabIndex        =   8
      Top             =   1703
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2948
      TabIndex        =   7
      Top             =   1703
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   188
      TabIndex        =   0
      Top             =   143
      Width           =   3735
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtHostName 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblFieldHeading 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   990
         Width           =   1005
      End
      Begin VB.Label lblFieldHeading 
         AutoSize        =   -1  'True
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   630
         Width           =   1035
      End
      Begin VB.Label lblFieldHeading 
         AutoSize        =   -1  'True
         Caption         =   "Host:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   510
      End
   End
End
Attribute VB_Name = "frmServerConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdConnect_Click()
    Dim myRS            As MYSQL_RS '// Recordset object
    Dim strWorker       As String   '// Work Horse variable
    Dim astrDatabases() As String   '// Will contain the list of databases
    Dim i               As Integer  '// Standard incrimentor
    
    Dim strTMPHostName  As String   '// --|
    Dim strTMPUserName  As String   '//   | -- We want to keep current server settings if the user enters
    Dim strTMPPassword  As String   '// --|     bad information for a new server
    Dim myTMPConnection As MYSQL_CONNECTION
    
    On Error GoTo Connect_Errors
        strTMPHostName = Trim(txtHostName.Text)
        strTMPUserName = Trim(txtUsername.Text)
        strTMPPassword = Trim(txtPassword.Text)
            
        Set myTMPConnection = New MYSQL_CONNECTION
        myTMPConnection.OpenConnection strTMPHostName, strTMPUserName, strTMPPassword
        
        '// Is there a valid database connection
        If myTMPConnection.State = MY_CONN_OPEN Then
            gstrHostName = Trim(txtHostName.Text)   '// --|
            gstrUserName = Trim(txtUsername.Text)   '//   | -- We have connected so set the global variables
            gstrPassword = Trim(txtPassword.Text)   '// --|     to their new values
            
            Set gmyConnection = New MYSQL_CONNECTION
            gmyConnection.OpenConnection gstrHostName, gstrUserName, gstrPassword
        
            Set myRS = New MYSQL_RS
            Set myRS = gmyConnection.Show(MY_SHOW_DATABASES)    '// Get a list of all the database on this server
            
            If Not myRS.EOF Then
                frmMain.cboObjects.ComboItems.Clear             '// --| -- Clear all items from the database combo boxes
                frmMain.cboDatabases.ComboItems.Clear           '// --|
            
                strWorker = Trim(myRS.GetString(, ""))
                If strWorker <> "" Then
                    astrDatabases = Split(strWorker, vbCrLf)
                    For i = 0 To UBound(astrDatabases)
                        If Trim(astrDatabases(i)) <> "" And LCase(Trim(astrDatabases(i))) <> "mysql" And LCase(Trim(astrDatabases(i))) <> "temp" Then
                            frmMain.cboObjects.ComboItems.Add , Trim(astrDatabases(i)), Trim(astrDatabases(i)), 1
                            frmMain.cboDatabases.ComboItems.Add , Trim(astrDatabases(i)), Trim(astrDatabases(i)), 1
                        End If
                    Next i
                End If
            End If
            
            myRS.CloseRecordset
            Set myRS = Nothing
            
            frmMain.cboDatabases.ComboItems(1).Selected = True
            Call frmMain.subUpdate_DatabaseLists
            Call frmMain.subNew_Query
        End If
        Unload Me
    On Error GoTo 0
    
Connect_Errors:
    If Err.Number <> 0 Then Call subLog_Errors(Err.Number, Err.Description, "frmServerConnect::cmdConnect_Click")
End Sub

Private Sub Form_Load()
    txtHostName.Text = Trim(gstrHostName)
    txtUsername.Text = Trim(gstrUserName)
    txtPassword.Text = Trim(gstrPassword)
End Sub
