VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C09_MaterialConsumption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function fGetMaterialConsumption(strDateFrom As String _
                                         , strDateTo As String, strDivision As String, ParamArray strItemId()) As Object
Dim strSQL As String

strSQL = ""
strSQL = strSQL & " SELECT TransactionView.ItemId,Items.ItemTypeId,Items.Description"
strSQL = strSQL & " FROM TransactionView"
strSQL = strSQL & " INNER JOIN Items ON"
strSQL = strSQL & " TransactionView.ItemId=Items.ItemId"
strSQL = strSQL & " INNER JOIN Divisions"
strSQL = strSQL & " ON TransactionView.DivisionId = Divisions.DivisionId"
strSQL = strSQL & " WHERE TransactionView.TransactedDate>='" & strDateFrom & "'"
strSQL = strSQL & " AND TransactionView.TransactedDate<='" & strDateTo & "'"
strSQL = strSQL & " AND Divisions.Description='" & strDivision & "'"

If strItemId(0) <> "" Then
    strSQL = strSQL & " AND TransactionView.ItemId Between '" & strItemId(0) & "' And '" & strItemId(1) & "'"
End If

'If strDivision = "MAINTENANCE" Then
'    strSQL = strSQL & " AND TransactionView.DivisionId=5"
'ElseIf strDivision = "MATERIAL CONTROL" Then
'    strSQL = strSQL & " AND TransactionView.DivisionId=6"
'ElseIf strDivision = "HRD" Then
'    strSQL = strSQL & " AND TransactionView.DivisionId=7"
'ElseIf strDivision = "HRD-C" Then
'    strSQL = strSQL & " AND TransactionView.DivisionId=8"
'ElseIf strDivision = "YUKA" Then
'    strSQL = strSQL & " AND TransactionView.DivisionId=9"
'End If


strSQL = strSQL & " GROUP BY TransactionView.ItemId,Items.ItemTypeId,Items.Description"
strSQL = strSQL & " ORDER BY TransactionView.ItemId"
    
Set fGetMaterialConsumption = clsDB.GetRecordSet(strSQL)
End Function

Public Function fGetMatConsumption( _
                 strDateFrom As String, strDateTo As String, _
                 strItemId As String) As Long
Dim strSQL As String
Dim rstQty  As Recordset

strSQL = ""
strSQL = strSQL & " SELECT SUM(Qty) as Qty FROM TransactionView"
strSQL = strSQL & " WHERE TransactedDate>='" & strDateFrom & "'"
strSQL = strSQL & " AND TransactedDate<='" & strDateTo & "'"
strSQL = strSQL & " AND ItemId='" & strItemId & "'"
strSQL = strSQL & " AND TransactionTypeId=3"
strSQL = strSQL & " AND QTY>0"

Set rstQty = clsDB.GetRecordSet(strSQL)

If IsNull(rstQty.Fields("Qty").Value) Then
    fGetMatConsumption = 0
Else
    fGetMatConsumption = rstQty.Fields("Qty").Value
End If
End Function
                

