VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C07_InvoiceSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub GetInvoiceSearch(ByRef hflexGrid As Object, ByVal intSupCateg As Integer, ByVal strItemId As String _
                          , ByVal strSupplier As String, ByVal strDivision As String, ByVal strPONo As String _
                          , ByVal strInvoiceNo As String, ByVal dtpODFrom As Object, ByVal dtpODTo As Object _
                          , ByVal dtpInvFrom As Object, ByVal dtpInvTo As Object, ByVal strItemTypeId As String _
                          , Optional ByVal blnTotal As Boolean = False, Optional ByVal blnFTotal As Boolean = False _
                          , Optional ByVal blnComplete As Boolean = False)
     Dim strSQLQuery     As String, _
         strFields       As String
     Dim rsInvoiceSearch As Object

On Error GoTo lnErrHandler
     If blnTotal Then
     '---Total received per invoice
          strSQLQuery = "SELECT " _
                      & "       PoNo, SupplierName, POSeq, ItemTypeId,ItemId, Description, " _
                      & "       OrderQty,ConvertingCoefficient, OrderDate, ETA,FTD, InvoiceNo, " _
                      & "       ReceivedDate, InvoiceQty, SUM(QtyReceived) AS QtyReceived, " _
                      & "       SUM(QtyOk) AS QtyOk, ReceivedAllInvoices, Maker " _
                      & " FROM POInvoiceSearchView "
     ElseIf blnFTotal Then
     '---Over all total delivery
          strSQLQuery = "SELECT " _
                      & "      PoNo, SupplierName, POSeq,ItemTypeId, ItemId, Description, " _
                      & "      OrderQty,ConvertingCoefficient, OrderDate, ETA, FTD, " _
                      & "      SUM(QtyReceived) AS QtyReceived, " _
                      & "      SUM(QtyOk) AS QtyOk, ReceivedAllInvoices, Maker " _
                      & " FROM POInvoiceSearchView "
     Else
          strSQLQuery = "SELECT * from POInvoiceSearchView "
     End If
     strSQLQuery = strSQLQuery & " WHERE DivisionId = " & clsPrintMenu.pfstrDivisionID(strDivision)
     '---Search by ItemId
     If Trim$(strItemId) <> "" Then _
          strSQLQuery = strSQLQuery & " and ItemId = " & pfstrQt(strItemId)
     '---Search by Supplier
     If strSupplier <> "" Then _
          strSQLQuery = strSQLQuery & " and SupplierName = " & pfstrQt(strSupplier)
     '---Search by PO No
     If Trim$(strPONo) <> "" Then _
          strSQLQuery = strSQLQuery & " and PONo = " & pfstrQt(strPONo)
     '---Search by Invoice NO
     If Trim$(strInvoiceNo) <> "" Then _
          strSQLQuery = strSQLQuery & " and InvoiceNo like " & pfstrQt(strInvoiceNo & "%")
     '----Search by ItemTypeId
     If Trim$(strItemTypeId) <> "" Then _
                strSQLQuery = strSQLQuery & " and ItemTypeId =" & pfstrQt(strItemTypeId)
     '-----Search by Supplier Category
     If intSupCateg <= 1 Then _
            strSQLQuery = strSQLQuery & " and ImportSupplier=" & intSupCateg & ""
            
     If Not IsNull(dtpODFrom.Value) Then _
               strSQLQuery = strSQLQuery & _
                              " and OrderDate >= " & pfstrQt(Format(dtpODFrom.Value, "yyyy/mm/dd"))
     If Not IsNull(dtpODTo.Value) Then _
               strSQLQuery = strSQLQuery & _
                              " and OrderDate < " & pfstrQt(DateAdd("d", 1, Format(dtpODTo.Value, "yyyy/mm/dd")))
     If Not IsNull(dtpInvFrom.Value) Then _
               strSQLQuery = strSQLQuery & _
                              " and DeliveryDate >= " & pfstrQt(Format(dtpInvFrom.Value, "yyyy/mm/dd"))
     If Not IsNull(dtpInvTo.Value) Then _
               strSQLQuery = strSQLQuery & _
                              " and DeliveryDate < " & pfstrQt(DateAdd("d", 1, Format(dtpInvTo.Value, "yyyy/mm/dd")))
     If blnComplete Then _
               strSQLQuery = strSQLQuery & " and ReceivedAllInvoices = 1"
               
     '---Total of Received Qty per Invoice
     If blnTotal Then
          strSQLQuery = strSQLQuery & _
                    " GROUP BY PoNo, SupplierName, POSeq, ItemId, Description, OrderQty, OrderDate, ETA,FTD, InvoiceNo, " _
                    & "       ReceivedDate, InvoiceQty, ReceivedAllInvoices, Maker,ItemTypeId,ConvertingCoefficient"
     ElseIf blnFTotal Then
          strSQLQuery = strSQLQuery & _
                    " GROUP BY PoNo, SupplierName, POSeq, ItemId, Description, OrderQty, OrderDate, ETA, FTD," _
                    & "       ReceivedAllInvoices, Maker,ItemTypeId,ConvertingCoefficient"
     End If
     
     strSQLQuery = strSQLQuery & " ORDER BY PONo, POSeq"
     
     Set rsInvoiceSearch = clsDB.GetRecordSet(strSQLQuery)
     Set hflexGrid.DataSource = rsInvoiceSearch
     Set rsInvoiceSearch = Nothing
     Exit Sub
     
lnErrHandler:
     MsgBox Err.Description, vbCritical
End Sub
