Attribute VB_Name = "modGeneral"
Option Explicit

'###############################################################################
'###############################################################################
'
' MyAnalyzer - Query Analyzer for mySQL Database
' Copyright (C) 2004,2005 Jimmy C. Broadhead, Jr.
' Send Questions/Comments to trisight@yahoo.com
'
'###############################################################################
'###############################################################################
'
' This program utilizes a library to interface with a mySQL database
' (VBMySQLDirect.dll) written by Robert Rowe.
'
' The class files used in this program:
'   MYSQL_CONNECTION.Cls
'   MYSQL_ERR.Cls
'   MYSQL_FIELD.Cls
'   MYSQL_RS.Cls
'   MYSQL_UPDATE.Cls
'   MYSQL_UPDATE_FIELD.Cls
'   MYSQL_UPDATE_FIELDS.Cls
' And the module modGeneral.bas are all the source code that makes up
' the VBMySQLDirect.dll written by Robert Rowe.  I do not take responsibility
' or credit for these classes and modules.  My program utilizes them for the
' purpose of accessing a mySQL database without the need of going through the ODBC
' I have included them as classes and module because some places (planetsourcecode)
' will not allow you to send additional DLLs.  Not only that people new to programming
' in Visual Basic may become discouraged at receiving errors when they don't load it
' and VB can't find the proper reference.
'
' To obtain the latest version of Mr. Rowe's DLL and/or current source visit:
' http://www.vbmysql.com/projects/vbmysqldirect/Download.php
'
'###############################################################################
'###############################################################################


'///////////////////////////////////////////////////////////////////////////////
'// General API Calls
'///////////////////////////////////////////////////////////////////////////////
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'// API Call to read INI files
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'///////////////////////////////////////////////////////////////////////////////

Public Const EM_GETLINECOUNT = &HBA        '// Total Line Count
Public Const EM_GETFIRSTVISIBLELINE = &HCE '// First Visible Line

Public gmyConnection    As MYSQL_CONNECTION '// Open a main connection, do not keep this in tally
Public gintConnections  As Integer          '// Keep tally of all open query connections

Public gstrHostName         As String       '// Current database host
Public gstrUserName         As String       '// Current database username
Public gstrPassword         As String       '// Current database password

Public mlngLastFoundPlace   As Long         '// Keeps track of the last place an item was found (for Find Next purposes)
Public gblnCancel           As Boolean      '// See if user cancels an action
Public gastrColorValues(1 To 4) As String   '// Query color options
Public gstrKeyWords         As String       '// Will contain the list of key words to compare against

'//-----------------------------------|
Public Const cnsTABLE = 1           ' |
Public Const cnsFIELD = 2           ' | -- Constants used for the gastrColorValues variable
Public Const cnsKEYWORD = 3         ' |     for query color options
Public Const cnsGENERAL = 4         ' |
'//-----------------------------------|

Public Sub Main()
    Dim astrConnection()    As String   '// Connection Object
    Dim fle                 As Integer  '// File handle
    
    On Error GoTo Main_Errors
        If Dir(App.Path & "\Save", vbDirectory) = "" Then MkDir App.Path & "\Save"
    
        '// Get query color options
        gastrColorValues(cnsTABLE) = funcGet_Option("TABLECOLOR")
        gastrColorValues(cnsFIELD) = funcGet_Option("FIELDCOLOR")
        gastrColorValues(cnsKEYWORD) = funcGet_Option("KEYCOLOR")
        gastrColorValues(cnsGENERAL) = funcGet_Option("GENERALCOLOR")
        
        If Trim(gastrColorValues(cnsTABLE)) = "" Then gastrColorValues(cnsTABLE) = RGB(200, 0, 0)
        If Trim(gastrColorValues(cnsFIELD)) = "" Then gastrColorValues(cnsFIELD) = RGB(0, 0, 240)
        If Trim(gastrColorValues(cnsKEYWORD)) = "" Then gastrColorValues(cnsKEYWORD) = RGB(200, 0, 255)
        If Trim(gastrColorValues(cnsGENERAL)) = "" Then gastrColorValues(cnsGENERAL) = RGB(0, 0, 0)
        
        If Dir(App.Path & "\KeyWords.dat") <> "" Then
            fle = FreeFile
            Open App.Path & "\KeyWords.dat" For Input As #fle
            Line Input #fle, gstrKeyWords
            Close #fle
        Else
            gstrKeyWords = "SELECT|FROM|WHERE|INSERT|INTO|DELETE|UPDATE|VALUES|SET|LIMIT|GROUP BY|HAVING|"
        End If
    
        '// Check for command line arguments (connection information)
        If Command$ <> "" Then
            Dim strWorker           As String
            Dim astrDatabases()     As String
            Dim i                   As Integer
            Dim MyRS                As MYSQL_RS
        
            astrConnection = Split(Command$, ";")
            gstrHostName = astrConnection(0)
            gstrUserName = astrConnection(1)
            gstrPassword = astrConnection(2)
                           
            If Trim(gstrHostName) <> "" And Trim(gstrUserName) <> "" Then
                Set gmyConnection = New MYSQL_CONNECTION
                gmyConnection.OpenConnection gstrHostName, gstrUserName, gstrPassword
                If gmyConnection.State = MY_CONN_OPEN Then
                    Set MyRS = New MYSQL_RS
                    Set MyRS = gmyConnection.Show(MY_SHOW_DATABASES)
                    
                    If Not MyRS.EOF Then
                        strWorker = Trim(MyRS.GetString(, ""))
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
                    
                    MyRS.CloseRecordset
                    Set MyRS = Nothing
                    
                    frmMain.cboDatabases.ComboItems(1).Selected = True
                    Call frmMain.subUpdate_DatabaseLists
                End If
            End If
        End If
        Call frmMain.subNew_Query
        gblnCancel = False
    On Error GoTo 0
    
    frmMain.Show
    
Main_Errors:
    If Err.Number <> 0 Then Call subLog_Errors(Err.Number, Err.Description, "modGeneral::Main")
End Sub

Public Sub subFindNext()
    Dim lngPositionFound        As Long
    Dim objTextFieldToSearch    As Object

    On Error GoTo FindNext_Errors
        If Trim(frmFind.txtFind.Text) = "" Then Exit Sub
    
        If frmFind.optQueryWindow Then
            Set objTextFieldToSearch = frmMain.ActiveForm.rtbQueryPane
        Else
            Set objTextFieldToSearch = frmMain.ActiveForm.txtResultsPane
        End If
        
        If frmFind.chkMatchCase Then
            lngPositionFound = InStr(mlngLastFoundPlace, objTextFieldToSearch.Text, Trim(frmFind.txtFind.Text))
        Else
            lngPositionFound = InStr(mlngLastFoundPlace, UCase(objTextFieldToSearch.Text), UCase(Trim(frmFind.txtFind.Text)))
        End If
        
        If lngPositionFound > 0 Then
            mlngLastFoundPlace = lngPositionFound + 1
            
            objTextFieldToSearch.SelStart = lngPositionFound - 1
            objTextFieldToSearch.SelLength = Len(Trim(frmFind.txtFind.Text))
            frmFind.Hide
            objTextFieldToSearch.SetFocus
        Else
            If mlngLastFoundPlace = 1 Then
                MsgBox "Search text NOT found!", vbInformation + vbOKOnly, "No Results..."
            Else
                MsgBox "No more instances found for search text!", vbInformation + vbOKOnly, "No More Results..."
            End If
        End If
    On Error GoTo 0
    
FindNext_Errors:
    If Err.Number <> 0 Then Call subLog_Errors(Err.Number, Err.Description, "modGeneral::subFindNext")
End Sub

Public Function funcGet_Option(strOptionKey As String) As String
    Dim strReturn   As String   '// Option to return
    Dim x           As Variant
    
    On Error GoTo getOption_Errors
        strReturn = String(255, Chr(0))
        x = GetPrivateProfileString("myAnalyzer", strOptionKey, "", strReturn, Len(strReturn), App.Path & "\Options.ini")
        funcGet_Option = Trim(Replace(strReturn, Chr(0), ""))
    On Error GoTo 0
    
getOption_Errors:
    If Err.Number <> 0 Then Call subLog_Errors(Err.Number, Err.Description, "frmOptions::funcGet_Option")
End Function

Public Sub subWrite_Option(strOptionKey As String, strOptionValue As String)
    Dim x       As Variant
    
    On Error GoTo writeOption_Errors
        x = WritePrivateProfileString("myAnalyzer", strOptionKey, strOptionValue, App.Path & "\Options.ini")
    On Error GoTo 0
    
writeOption_Errors:
    If Err.Number <> 0 Then Call subLog_Errors(Err.Number, Err.Description, "frmOptions::subWrite_Option")
End Sub

Public Sub subLog_Errors(lngErrorNumber As Long, _
                         strErrorDescription As String, _
                         strErrorOrigination As String)
    '/////////////////////////////////////////////////////////////////////////////////
    '// Error handling routine
    '/////////////////////////////////////////////////////////////////////////////////
    On Error Resume Next
        Dim fle     As Integer
        
        fle = FreeFile
        Open App.Path & "\Error.log" For Append As #fle
        Print #fle, "--------------------------------------------------"
        Print #fle, "Errors Occurred on : " & Format(Now(), "MM/DD/YYYY HH:MM:SS")
        Print #fle, ""
        Print #fle, "Error Number: " & lngErrorNumber
        Print #fle, "Error Description: " & strErrorDescription
        Print #fle, ""
        Print #fle, "Error Occurred in: " & strErrorOrigination
        
        Close #fle
        MsgBox "Errors have occurred:" & vbCrLf & lngErrorNumber & ":" & strErrorDescription
    On Error GoTo 0
End Sub
