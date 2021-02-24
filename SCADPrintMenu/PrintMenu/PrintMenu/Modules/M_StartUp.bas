Attribute VB_Name = "M_StartUp"
Option Explicit

Public clsPrintMenu           As C_PrintMenu
Public clsDB                  As C_PrintDB

'Public pobjPGFormTools         As Object
'Public pobjPGDB                As Object
Public Const pstrMessage      As String = "System Message"
Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)

Sub Main()
    On Error GoTo lnError
    '--- check if application is already open
    If App.PrevInstance = True Then MsgBox "System already open!", vbInformation, "System": Exit Sub
    F_Splash.Show
    DoEvents
    Sleep 1000
    '--- Connection for WarehouseManagement
    Set clsDB = New C_PrintDB
    
    Set clsPrintMenu = New C_PrintMenu
    
    'Call clsDB.SQLServer(WarehouseManagement)
    '''''''''Ardie
    Call clsDB.SQLServer(WarehouseManagement, App.Path & "\Print.ini")
    '''''''''
    If clsDB.IsConnected Then F_PrintMainMenu.Show
    
    Unload F_Splash
    Exit Sub
lnError:
    MsgBox Err.Number & "-" & Err.Description, vbCritical, pstrMessage
End Sub

Public Function pfstrQt(ByVal strValue As String) As String
    pfstrQt = Replace(strValue, "'", "''")
    pfstrQt = "'" & pfstrQt & "'"
End Function

Public Function pfvarIs_Null(ByVal varNull _
                             , Optional ByVal blnStringType As Boolean = True) As Variant
     '--- if not null, dont change
     If varNull <> "" And Not IsNull(varNull) Then
          pfvarIs_Null = varNull
          Exit Function
     End If
     Select Case blnStringType
        Case True
            pfvarIs_Null = ""
        Case Else
            pfvarIs_Null = 0
     End Select
End Function

Public Sub psubShowStatMsg(ByVal strMessage As String, Optional bytPanel As Byte = 3)
     With F_PrintMainMenu
          DoEvents
          .StatusBar1.Panels(bytPanel).Text = strMessage
     End With
End Sub

Public Sub psubHideStatMsg(Optional ByVal bytPanel As Byte = 3)
     With F_PrintMainMenu
          DoEvents
          .StatusBar1.Panels(bytPanel).Text = ""
     End With
End Sub


