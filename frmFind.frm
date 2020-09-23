VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find..."
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frameFindIn 
      Caption         =   "Find In"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1920
      TabIndex        =   5
      Top             =   611
      Width           =   2175
      Begin VB.OptionButton optResultsWindow 
         Caption         =   "Results Window"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton optQueryWindow 
         Caption         =   "Query Window"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
   End
   Begin VB.CheckBox chkMatchCase 
      Caption         =   "Match Case"
      Height          =   255
      Left            =   143
      TabIndex        =   4
      Top             =   859
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4223
      TabIndex        =   3
      Top             =   971
      Width           =   1215
   End
   Begin VB.CommandButton cmdFindNext 
      Caption         =   "Find Next"
      Default         =   -1  'True
      Height          =   495
      Left            =   4223
      TabIndex        =   2
      Top             =   251
      Width           =   1215
   End
   Begin VB.TextBox txtFind 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1103
      TabIndex        =   1
      Top             =   244
      Width           =   3015
   End
   Begin VB.Label lblFieldHeading 
      AutoSize        =   -1  'True
      Caption         =   "Find what:"
      Height          =   195
      Left            =   143
      TabIndex        =   0
      Top             =   289
      Width           =   885
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkMatchCase_Click()
    mlngLastFoundPlace = 1      '// Reset the variable so we can start from the beginning
End Sub

Private Sub Option1_Click()
    mlngLastFoundPlace = 1      '// Reset the variable so we can start from the beginning
End Sub

Private Sub Form_Load()
    mlngLastFoundPlace = 1      '// Reset the variable so we can start from the beginning
End Sub

Private Sub optQueryWindow_Click()
    mlngLastFoundPlace = 1      '// Reset the variable so we can start from the beginning
End Sub

Private Sub txtFind_Change()
    mlngLastFoundPlace = 1      '// Reset the variable so we can start from the beginning
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFindNext_Click()
    Call subFindNext            '// Call the public search routine
End Sub

