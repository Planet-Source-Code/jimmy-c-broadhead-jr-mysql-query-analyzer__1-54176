VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print..."
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   300
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   780
      Width           =   975
   End
   Begin VB.Frame frameFindIn 
      Caption         =   "Print What"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   2175
      Begin VB.OptionButton optQueryWindow 
         Caption         =   "Query Window"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optResultsWindow 
         Caption         =   "Results Window"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1935
      End
   End
   Begin MSComDlg.CommonDialog cmnPrint 
      Left            =   3000
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    '///////////////////////////////////////////////////////////////////////////////
    '// Nothing fancy here, it might be nice to come back at a later time and
    '// fancy it up, but this will suffice for now.
    '///////////////////////////////////////////////////////////////////////////////
    Dim strTextToPrint      As String
    
    On Error GoTo Print_Cancel
        cmnPrint.ShowPrinter    '// Get printer settings from user
    On Error GoTo Print_Errors
        If optQueryWindow Then  '// Get the text from either query pane or the resultspane, whichever is requested
            strTextToPrint = frmMain.ActiveForm.rtbQueryPane.Text
        Else
            strTextToPrint = frmMain.ActiveForm.txtResultsPane.Text
        End If
        
        '// Set printer items as per the Printer Common Dialog box
        Printer.Copies = cmnPrint.Copies
        Printer.Orientation = cmnPrint.Orientation
        
        '// Set the font to Arial Size 10
        Printer.Font = "Arial"
        Printer.FontSize = 10
        
        '// Set the text on the printer and print
        Printer.Print strTextToPrint
        Printer.EndDoc
        Unload Me
    On Error GoTo 0
    
Print_Cancel:
    Exit Sub
    
Print_Errors:
    If Err.Number <> 0 Then Call subLog_Errors(Err.Number, Err.Description, "frmMain::mnuPrint_Click")
End Sub
