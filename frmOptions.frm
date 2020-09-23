VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IDE Options..."
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   4920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgColors 
      Left            =   120
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "Query Colors"
      Height          =   3975
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   4215
      Begin VB.Frame Frame2 
         Caption         =   "Update Key Words"
         Height          =   1935
         Left            =   240
         TabIndex        =   12
         Top             =   1920
         Width           =   3735
         Begin VB.CommandButton cmdRemove 
            Caption         =   "Remove"
            Height          =   375
            Left            =   1200
            TabIndex        =   18
            Top             =   1440
            Width           =   855
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   1200
            TabIndex        =   17
            Top             =   960
            Width           =   855
         End
         Begin VB.TextBox txtKeyWord_toAdd 
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   1935
         End
         Begin VB.ListBox lstKeyWords 
            Height          =   1425
            ItemData        =   "frmOptions.frx":27A2
            Left            =   2160
            List            =   "frmOptions.frx":27A4
            TabIndex        =   13
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lblFieldHeading 
            AutoSize        =   -1  'True
            Caption         =   "Key Word to Add"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Label lblFieldHeading 
         AutoSize        =   -1  'True
         Caption         =   "Click color to change"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   16
         Top             =   360
         Width           =   1785
      End
      Begin VB.Line Line2 
         X1              =   3120
         X2              =   3120
         Y1              =   1320
         Y2              =   2040
      End
      Begin VB.Line Line1 
         X1              =   2040
         X2              =   3120
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label lblGeneralColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1560
         TabIndex        =   11
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblFieldHeading 
         AutoSize        =   -1  'True
         Caption         =   "General Text :"
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   10
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label lblKeyColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1560
         TabIndex        =   9
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblFieldColor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1560
         TabIndex        =   8
         Top             =   960
         Width           =   495
      End
      Begin VB.Label lblTableColor 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1560
         TabIndex        =   7
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblFieldHeading 
         AutoSize        =   -1  'True
         Caption         =   "Key Words :"
         Height          =   195
         Index           =   2
         Left            =   450
         TabIndex        =   6
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label lblFieldHeading 
         AutoSize        =   -1  'True
         Caption         =   "Field Name :"
         Height          =   195
         Index           =   1
         Left            =   435
         TabIndex        =   5
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label lblFieldHeading 
         AutoSize        =   -1  'True
         Caption         =   "Table Name :"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   720
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "Okay"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   8281
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Options"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    If Trim(txtKeyWord_toAdd.Text) = "" Then Exit Sub       '// Don't add blanks
    lstKeyWords.AddItem UCase(Trim(txtKeyWord_toAdd.Text))  '// Add the new item to the list
    txtKeyWord_toAdd.Text = ""                              '// Clear the text box
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOkay_Click()
    Dim i               As Integer  '// Standard incrimentor
    Dim fle             As Integer  '// File handle
    Dim strTMPKeyWords  As String   '// Temporary list of key words

    On Error GoTo cmdOkay_Errors
        Call subWrite_Option("TABLECOLOR", CStr(lblTableColor.BackColor))
        Call subWrite_Option("FIELDCOLOR", CStr(lblFieldColor.BackColor))
        Call subWrite_Option("KEYCOLOR", CStr(lblKeyColor.BackColor))
        Call subWrite_Option("GENERALCOLOR", CStr(lblGeneralColor.BackColor))
        
        gastrColorValues(cnsTABLE) = CStr(lblTableColor.BackColor)
        gastrColorValues(cnsFIELD) = CStr(lblFieldColor.BackColor)
        gastrColorValues(cnsKEYWORD) = CStr(lblKeyColor.BackColor)
        gastrColorValues(cnsGENERAL) = CStr(lblGeneralColor.BackColor)
        
        For i = 0 To lstKeyWords.ListCount - 1
            strTMPKeyWords = strTMPKeyWords & "|" & Trim(lstKeyWords.List(i))
        Next i
        strTMPKeyWords = strTMPKeyWords & "|"
                
        fle = FreeFile
        Open App.Path & "\KeyWords.dat" For Output As #fle
        Print #fle, strTMPKeyWords
        Close #fle
        
        gstrKeyWords = strTMPKeyWords
    On Error GoTo 0
    Unload Me
    
cmdOkay_Errors:
    If Err.Number <> 0 Then Call subLog_Errors(Err.Number, Err.Description, "frmOptions::cmdOkay_Click")
End Sub

Private Sub cmdRemove_Click()
    If lstKeyWords.ListIndex >= 0 Then
        lstKeyWords.RemoveItem lstKeyWords.ListIndex        '// Remove the selected item from the list
    End If
End Sub

Private Sub Form_Load()
    Dim astrKeyWords()      As String   '// Array containing the key words
    Dim i                   As Integer  '// Standard incrimentor

    On Error GoTo OptionsForm_Load_Errors
        lblTableColor.BackColor = gastrColorValues(cnsTABLE)
        lblFieldColor.BackColor = gastrColorValues(cnsFIELD)
        lblKeyColor.BackColor = gastrColorValues(cnsKEYWORD)
        lblGeneralColor.BackColor = gastrColorValues(cnsGENERAL)
        
        astrKeyWords = Split(gstrKeyWords, "|")
        For i = LBound(astrKeyWords) To UBound(astrKeyWords)
            If Trim(astrKeyWords(i)) <> "" Then lstKeyWords.AddItem Trim(astrKeyWords(i))
        Next i
    On Error GoTo 0
        
OptionsForm_Load_Errors:
    If Err.Number <> 0 Then Call subLog_Errors(Err.Number, Err.Description, "frmOptions::Form_Load")
End Sub

Private Sub lblFieldColor_Click()
    On Error GoTo Exit_FieldColor
        dlgColors.ShowColor
        lblFieldColor.BackColor = dlgColors.Color
    On Error GoTo 0
    
Exit_FieldColor:
    
End Sub

Private Sub lblGeneralColor_Click()
    On Error GoTo Exit_GeneralColor
        dlgColors.ShowColor
        lblGeneralColor.BackColor = dlgColors.Color
    On Error GoTo 0
    
Exit_GeneralColor:

End Sub

Private Sub lblKeyColor_Click()
    On Error GoTo Exit_KeyColor
        dlgColors.ShowColor
        lblKeyColor.BackColor = dlgColors.Color
    On Error GoTo 0
    
Exit_KeyColor:

End Sub

Private Sub lblTableColor_Click()
    On Error GoTo Exit_TableColor
        dlgColors.ShowColor
        lblTableColor.BackColor = dlgColors.Color
    On Error GoTo 0
    
Exit_TableColor:
    
End Sub
