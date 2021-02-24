VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C_GetDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function GetRecordSet(ByVal strQuery As String) As Object
    Dim adoRecordset As ADODB.Recordset
    
    Set adoRecordset = New ADODB.Recordset
    On Error GoTo lnError
    
    adoRecordset.Open strQuery, cn, adOpenDynamic, adLockReadOnly
    Set GetRecordSet = adoRecordset
    
    Exit Function
lnError:
     MsgBox strQuery & vbCrLf & vbCrLf & Err.Number & "-" & Err.Description, vbCritical, "System Error"
End Function

Public Function pfLoadData(ByVal strTypeName As String)
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " Select"
    strSQL = strSQL & " *"
    strSQL = strSQL & " FROM"
    strSQL = strSQL & " DailyReportView"
    strSQL = strSQL & " WHERE"
    strSQL = strSQL & " TypeName = '" & strTypeName & "'"
    Set pfLoadData = clsGetData.GetRecordSet(strSQL)
End Function


'*******DIITO KA NA KIM






