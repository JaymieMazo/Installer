VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C03_InventoryRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim strSQLInvRecord As String

Public Function GetInventoryRecord(ByVal strItemId As String, ByVal dteFrom As Date, _
                                   ByVal dteUntil As Date, ByVal strDivision As String) As Object
                                   
      strSQLInvRecord = " SELECT " _
                      & "       TransactedDate, TransId, ItemId, Qty, TransactionTypeId, DivisionId, Remarks " _
                      & " FROM  InventoryWHSRecord " _
                      & " WHERE ItemId = " & pfstrQt(strItemId) _
                      & "       and TransactedDate >= " & pfstrQt(Format(dteFrom, "yyyy/mm/dd")) _
                      & "       and TransactedDate <= " & pfstrQt(Format(dteUntil, "yyyy/mm/dd")) _
                      & "       and DivisionId = " & clsPrintMenu.pfstrDivisionID(strDivision) _
                      & " ORDER BY TransactedDate asc, TransactionTypeId desc"
     Set GetInventoryRecord = clsDB.GetRecordSet(strSQLInvRecord)
End Function

Public Sub SaveToInvRecordWT(ParamArray varValues())
     Dim strFields As String
     
     strFields = "(SeqNo,TransDate,TransID,InQty,OutQty, ACQty, Balance,Remarks)"
     Call clsDB.DB_PrintMenu.Execute("BEGIN TRANSACTION")
     strSQLInvRecord = "INSERT INTO " _
                         & " PrintMenu..InventoryRecord " & strFields _
                    & " VALUES (" _
                         & varValues(0) & "," & pfstrQt(varValues(1)) & "," & pfstrQt(varValues(2)) & "," & varValues(3) _
                         & "," & varValues(4) & "," & varValues(5) & "," & varValues(6) & "," & pfstrQt(varValues(7)) & ")"
                    
     Call clsDB.DB_PrintMenu.Execute(strSQLInvRecord)
     Call clsDB.DB_PrintMenu.Execute("COMMIT TRANSACTION")
End Sub

Public Function SQLInventoryRecord() As String
     SQLInventoryRecord = " SELECT " _
                        & "   TransDate, TransId, Division, InQty, OutQty, AcQty, Balance, Remarks " _
                        & " FROM InventoryRecord"
End Function

Public Function TotalOut(ByVal strItemId As String, ByVal strDivision As String, ByVal strDate As String) As Double
     Dim rsTotalOut As Object
     
     strSQLInvRecord = "SELECT  SUM(ItemShipmentDetails.Qty) AS Qty " _
                    & " FROM  ItemShipmentHeaders INNER JOIN " _
                    & "       ItemShipmentDetails ON " _
                    & "       ItemShipmentHeaders.ItemShipmentId = ItemShipmentDetails.ItemShipmentId " _
                    & " GROUP BY ItemShipmentDetails.ItemId, ItemShipmentHeaders.ShippedDate, " _
                    & "       ItemShipmentHeaders.DivisionId " _
                    & " HAVING ItemShipmentDetails.ItemId = " & pfstrQt(strItemId) _
                    & "       and ItemShipmentHeaders.ShippedDate = " & pfstrQt(strDate) _
                    & "       and ItemShipmentHeaders.DivisionId = " & clsPrintMenu.pfstrDivisionID(strDivision)
                
     Set rsTotalOut = clsDB.GetRecordSet(strSQLInvRecord)
     If Not rsTotalOut.EOF Then
          TotalOut = pfvarIs_Null(rsTotalOut.Fields("Qty").Value, False)
     Else
          TotalOut = 0
     End If
     Set rsTotalOut = Nothing
End Function

Public Function TotalIn(ByVal strItemId As String, ByVal strDate As String) As Double
     Dim rsTotalIn As Object
     
     'Call clsDB.DB_PrintMenu.Execute("BEGIN TRANSACTION")
     strSQLInvRecord = " SELECT SUM(DeliveryDetails.QtyOk) AS  Qty " _
                     & " FROM  DeliveryHeaders  " _
                     & " INNER JOIN  DeliveryDetails " _
                     & " ON DeliveryHeaders.DeliveryId = DeliveryDetails.DeliveryId " _
                     & "      and DeliveryHeaders.InvoiceNo = DeliveryDetails.InvoiceNo " _
                     & "      and DeliveryHeaders.SupplierId = DeliveryDetails.SupplierId " _
                     & " INNER JOIN InvoiceDetails ON DeliveryDetails.InvoiceNo = InvoiceDetails.InvoiceNo " _
                     & "      and DeliveryDetails.SupplierId = InvoiceDetails.SupplierId " _
                     & "      and DeliveryDetails.InvoiceDetailSeq = InvoiceDetails.InvoiceDetailSeq " _
                     & " INNER JOIN InvoiceHeaders ON InvoiceDetails.InvoiceNo = InvoiceHeaders.InvoiceNo " _
                     & "      and InvoiceDetails.SupplierId = InvoiceHeaders.SupplierId " _
                     & " WHERE InvoiceDetails.ItemId = " & pfstrQt(strItemId) _
                     & "      and InvoiceHeaders.ReceivedDate = " & pfstrQt(strDate)
     Set rsTotalIn = clsDB.GetRecordSet(strSQLInvRecord)
     If Not rsTotalIn.EOF Then
          TotalIn = pfvarIs_Null(rsTotalIn.Fields("QTY").Value, False)
     Else
          TotalIn = 0
     End If
     Set rsTotalIn = Nothing
     'Call clsDB.DB_PrintMenu.Execute("COMMIT TRANSACTION")
End Function

Public Sub DeleteInvRecordsWT()
     Call clsDB.DB_PrintMenu.Execute("BEGIN TRANSACTION")
     Call clsDB.DB_PrintMenu.Execute("DELETE from PrintMenu..InventoryRecord")
     Call clsDB.DB_PrintMenu.Execute("COMMIT TRANSACTION")
End Sub
'Public Function pfGetIQCQty(ByVal strItemId, strDivision As String)
'Dim strSQL As String
'
'strSQL = " "
'strSQL = strSQL & " SELECT "
'strSQL = strSQL & " POInvoiceSearchView.QtyReceived "
'strSQL = strSQL & " FROM POInvoiceSearchView "
'strSQL = strSQL & " INNER JOIN Divisions "
'strSQL = strSQL & " ON POInvoiceSearchView.DivisionId=Divisions.DivisionId "
'strSQL = strSQL & " WHERE POInvoiceSearchView.ItemId='" & strItemId & "'"
'
'Set pfGetIQCQty = clsDB.GetRecordSet(strSQL)
'
'End Function

