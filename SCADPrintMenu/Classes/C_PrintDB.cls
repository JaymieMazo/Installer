VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_PrintDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===========================================================
'=================Ardie G.===10/16/2007===REVISED CONNECTION===
'===========================================================
Option Explicit

Private adoConnection      As ADODB.Connection
Private adoConPrintMenu    As ADODB.Connection
'''''''''Ardie'''
Dim strServerName As String, _
    strDatabase     As String, _
    strDatabase2     As String, _
    strUserName     As String, _
    strPassWord     As String

Private blnConnected       As Boolean
Private Declare Function GetPrivateProfileString Lib "KERNEL32" Alias _
    "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
''''''''''''''''''
Public Enum SqlServerDB
     WarehouseManagement = 1
     PrintMenuDb = 2
End Enum

'Public Sub SQLServer(ByRef enmServer As SqlServerDB)
'''''''''Ardie'''
Public Sub SQLServer(ByRef enmServer As SqlServerDB, Optional ByVal strIniFile As String)

    On Error GoTo lnError
    '--- Update program first
    If App.PrevInstance = True Then Exit Sub
    'On Error GoTo lnError
    
    'If blnConnected Then Exit Sub
    Call ReadServerIni(strIniFile)
    
    Select Case enmServer
          Case 1:  '--- WarehouseManagement
              '--- Connection for WarehouseManagement
                Set adoConnection = New ADODB.Connection
                With adoConnection
                    .CursorLocation = adUseClient
                    .ConnectionTimeout = 60
                    .ConnectionString = "Provider=SQLOLEDB;Data Source=" & strServerName & _
                            ";Initial Catalog=" & strDatabase & _
                            ";User ID=" & strUserName & _
                            ";" & IIf(IsNull(strPassWord), vbNullString, "Password=" & strPassWord) '& _

                    .Open
                End With
         Case 2: '--- PrintMenu
               '--- Connection for PrintMenu
                Set adoConPrintMenu = New ADODB.Connection
                With adoConPrintMenu
                    .CursorLocation = adUseClient
                    .ConnectionTimeout = 60
                    .ConnectionString = "Provider=SQLOLEDB;Data Source=" & strServerName & _
                            ";Initial Catalog=" & strDatabase2 & _
                            ";User ID=" & strUserName & _
                            ";" & IIf(IsNull(strPassWord), vbNullString, "Password=" & strPassWord) '& _

                    .Open
                End With
    
    End Select
    blnConnected = True
    Exit Sub

lnError:
     blnConnected = False
    MsgBox Err.Number & "-" & Err.Description, vbCritical, pstrMessage
End Sub

'    Select Case enmServer
'          Case 1:  '--- WarehouseManagement
'               '--- Connection for WarehouseManagement
'               Set adoConnection = New ADODB.Connection
'               With adoConnection
'                   .CursorLocation = adUseClient
'                   .ConnectionTimeout = 60
'                   .ConnectionString = "Provider=SQLOLEDB;Data Source=MISSQL;OLE DB Services = -2;User ID=sa;Password=administrator;" _
'                                       & " Initial Catalog=WarehouseManagement;"
'                   .Open
'               End With
'          Case 2: '--- PrintMenu
'               '--- Connection for PrintMenu
'               Set adoConPrintMenu = New ADODB.Connection
'               With adoConPrintMenu
'                   .CursorLocation = adUseClient
'                   .ConnectionTimeout = 60
'                   .ConnectionString = "Provider=SQLOLEDB;Data Source=MISSQL;OLE DB Services = -2;User ID=sa;Password=administrator;" _
'                                       & " Initial Catalog=PrintMenu;"
'                   .Open
'               End With
'    End Select
'    blnConnected = True
'    Exit Sub
'
'lnError:
'     blnConnected = False
'    MsgBox Err.Number & "-" & Err.Description, vbCritical, pstrMessage
'End Sub

Property Get DB_WHMgt() As Object
     Set DB_WHMgt = adoConnection
End Property

Property Get DB_PrintMenu() As Object
     Set DB_PrintMenu = adoConPrintMenu
End Property

Property Get IsConnected() As Boolean
     IsConnected = blnConnected
End Property

Public Function GetRecordSet(ByVal strQuery As String, Optional ByVal IsDefaultConn As Boolean = True) As ADODB.Recordset
    Dim adoRecordSet As ADODB.Recordset
    
    Set adoRecordSet = New ADODB.Recordset
    
    On Error GoTo lnError
    
    If IsDefaultConn Then
        adoConnection.CommandTimeout = 0
        adoRecordSet.Open strQuery, adoConnection, adOpenDynamic, adLockReadOnly
    Else
        adoConPrintMenu.CommandTimeout = 0
        adoRecordSet.Open strQuery, adoConPrintMenu, adOpenDynamic, adLockOptimistic
    End If
    Set GetRecordSet = adoRecordSet
    Exit Function
lnError:
    MsgBox strQuery & vbCrLf & vbCrLf & Err.Number & "-" & Err.Description, vbCritical, pstrMessage
End Function

Public Function ServerDate() As String
    Dim adoServerDate   As ADODB.Recordset
    Dim strQuery        As String
    
    strQuery = "SELECT GETDATE()"
    
    Set adoServerDate = GetRecordSet(strQuery)
    
    ServerDate = Format(adoServerDate.Fields(0).Value, "yyyy/mm/dd")
End Function

Public Sub DisConnectPrintMenu()
     If Not adoConPrintMenu Is Nothing Then adoConPrintMenu.Close
     Set adoConPrintMenu = Nothing
End Sub

Private Sub Class_Terminate()
    adoConnection.Close
    Set adoConnection = Nothing
    Set adoConPrintMenu = Nothing
End Sub
'''''''''Ardie'''
Private Sub ReadServerIni(ByVal strFileIni As String)
    strServerName = ReadINI("SQLServer", "ServerName", strFileIni)
    strDatabase = ReadINI("SQLServer", "Database", strFileIni)
    strDatabase2 = ReadINI("SQLServer", "Database2", strFileIni)
    strUserName = ReadINI("SQLServer", "UserName", strFileIni)
    strPassWord = ReadINI("SQLServer", "PassWord", strFileIni)
    
End Sub
'''''''''Ardie'''
Function ReadINI(strSection As String, strKeyName As String, _
        strINIFileName As String) As String
     Dim strRet As String
    
    On Local Error Resume Next
    strRet = String(255, Chr(0))
    ReadINI = Left(strRet, GetPrivateProfileString(strSection, ByVal strKeyName, "", strRet, Len(strRet), strINIFileName))

End Function



