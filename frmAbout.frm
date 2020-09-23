VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About myAnalyzer..."
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOkay 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   495
      Left            =   7523
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5966
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Program Information"
      Height          =   2895
      Left            =   203
      TabIndex        =   3
      Top             =   3003
      Width           =   8415
      Begin VB.Label lblAdditionalInfo 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":0442
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   8175
      End
      Begin VB.Label lblProgramInformation 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":053C
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   8175
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Program Copyright 2004 - Jimmy C. Broadhead, Jr.  All rights reserved."
      Height          =   210
      Left            =   210
      TabIndex        =   2
      Top             =   2643
      Width           =   7095
   End
   Begin VB.Label lblProgramVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Program Version: 1.0.0"
      Height          =   210
      Left            =   210
      TabIndex        =   1
      Top             =   2403
      Width           =   2325
   End
   Begin VB.Label lblProgramName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Program Name: myAnalyzer"
      Height          =   210
      Left            =   210
      TabIndex        =   0
      Top             =   2163
      Width           =   2790
   End
   Begin VB.Image Image1 
      Height          =   2010
      Left            =   923
      Picture         =   "frmAbout.frx":083B
      Top             =   123
      Width           =   6660
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOkay_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lblProgramVersion.Caption = "Program Version: " & App.Major & "." & App.Minor & "." & App.Revision
End Sub
