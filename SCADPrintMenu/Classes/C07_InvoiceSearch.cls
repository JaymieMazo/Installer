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
                          , ByVal DTEncFrom As Object, ByVal DTEncTo As Object _
                          , ByVal DTInvoiceFrom As Object, ByVal DTInvoiceTo As Object _
                          , ByVal dtETDFrom As Object, ByVal dtETDTo As Object _
                          , ByVal dtActFrom As Object, ByVal dtActTo As Object _
                          , Optional ByVal blnTotal As Boolean = False, Optional ByVal blnFTotal As Boolean = False _
                          , Optional ByVal blnComplete As Boolean = False, Optional ByVal blnProdPurchase As Boolean = False)

     Dim strSQLQuery     As String, _
         strFields       As String
     Dim rsInvoiceSearch As Object

On Error GoTo lnErrHandler
     If blnTotal Then
     '---Total received per invoice
     
        If Not IsNull(dtActFrom.Value) Or Not IsNull(dtActTo.Value) Then
           strSQLQuery = "SELECT " _
                         & "       POInvoiceSearchView.PoNo, POInvoiceSearchView.SupplierName, POInvoiceSearchView.POSeq, " _
                         & "       POInvoiceSearchView.ItemTypeId,POInvoiceSearchView.ItemId, POInvoiceSearchView.Description, " _
                         & "       POInvoiceSearchView.OrderQty,POInvoiceSearchView.ConvertingCoefficient, " _
                         & "       POInvoiceSearchView.OrderDate , POInvoiceSearchView.ETA, POInvoiceSearchView.FTD, POInvoiceSearchView.InvoiceNo, " _
                         & "       POInvoiceSearchView.ReceivedDate, POInvoiceSearchView.InvoiceQty, " _
                         & "       SUM(POInvoiceSearchView.QtyReceived) AS QtyReceived, " _
                         & "       SUM(POInvoiceSearchView.QtyOk) AS QtyOk, POInvoiceSearchView.ReceivedAllInvoices, POInvoiceSearchView.Maker, " _
                         & "       POInvoiceSearchView.StaffName,IH.DeliveryDate as ActualETD " _
                         & "       FROM POInvoiceSearchView " _
                         & "       INNER JOIN PoDetailsView ON POInvoiceSearchView.PoNo = PoDetailsView.Pono "
        Else
           strSQLQuery = "SELECT " _
                         & "       POInvoiceSearchView.PoNo, POInvoiceSearchView.SupplierName, POInvoiceSearchView.POSeq, " _
                         & "       POInvoiceSearchView.ItemTypeId,POInvoiceSearchView.ItemId, POInvoiceSearchView.Description, " _
                         & "       POInvoiceSearchView.OrderQty,POInvoiceSearchView.ConvertingCoefficient, " _
                         & "       POInvoiceSearchView.OrderDate , POInvoiceSearchView.ETA, POInvoiceSearchView.FTD, POInvoiceSearchView.InvoiceNo, " _
                         & "       POInvoiceSearchView.ReceivedDate, POInvoiceSearchView.InvoiceQty, " _
                         & "       SUM(POInvoiceSearchView.QtyReceived) AS QtyReceived, " _
                         & "       SUM(POInvoiceSearchView.QtyOk) AS QtyOk, POInvoiceSearchView.ReceivedAllInvoices, POInvoiceSearchView.Maker, " _
                         & "       POInvoiceSearchView.StaffName" _
                         & "       FROM POInvoiceSearchView " _
                         & "       INNER JOIN PoDetailsView ON POInvoiceSearchView.PoNo = PoDetailsView.Pono "
        End If
          
     ElseIf blnFTotal Then
     '---Over all total delivery
        If Not IsNull(dtActFrom.Value) Or Not IsNull(dtActTo.Value) Then
           strSQLQuery = "SELECT " _
                         & "      POInvoiceSearchView.PoNo, POInvoiceSearchView.SupplierName, POInvoiceSearchView.POSeq," _
                         & "      POInvoiceSearchView.ItemTypeId, POInvoiceSearchView.ItemId, POInvoiceSearchView.Description, " _
                         & "      POInvoiceSearchView.OrderQty,POInvoiceSearchView.ConvertingCoefficient, POInvoiceSearchView.OrderDate, " _
                         & "      POInvoiceSearchView.ETA, POInvoiceSearchView.FTD, " _
                         & "      SUM(POInvoiceSearchView.QtyReceived) AS QtyReceived, " _
                         & "      SUM(POInvoiceSearchView.QtyOk) AS QtyOk, POInvoiceSearchView.ReceivedAllInvoices, POInvoiceSearchView.Maker, " _
                         & "      POInvoiceSearchView.StaffName,IH.DeliveryDate as ActualETD " _
                         & "      FROM POInvoiceSearchView " _
                         & "      INNER JOIN PoDetailsView ON POInvoiceSearchView.PoNo = PoDetailsView.Pono "
        
        Else
           strSQLQuery = "SELECT " _
                         & "      POInvoiceSearchView.PoNo, POInvoiceSearchView.SupplierName, POInvoiceSearchView.POSeq," _
                         & "      POInvoiceSearchView.ItemTypeId, POInvoiceSearchView.ItemId, POInvoiceSearchView.Description, " _
                         & "      POInvoiceSearchView.OrderQty,POInvoiceSearchView.ConvertingCoefficient, POInvoiceSearchView.OrderDate, " _
                         & "      POInvoiceSearchView.ETA, POInvoiceSearchView.FTD, " _
                         & "      SUM(POInvoiceSearchView.QtyReceived) AS QtyReceived, " _
                         & "      SUM(POInvoiceSearchView.QtyOk) AS QtyOk, POInvoiceSearchView.ReceivedAllInvoices, POInvoiceSearchView.Maker, " _
                         & "      POInvoiceSearchView.StaffName" _
                         & "      FROM POInvoiceSearchView " _
                         & "      INNER JOIN PoDetailsView ON POInvoiceSearchView.PoNo = PoDetailsView.Pono "
        
        End If
     Else
         ' strSQLQuery = "SELECT * from POInvoiceSearchView " Gerald 20141115
        If Not IsNull(dtActFrom.Value) Or Not IsNull(dtActTo.Value) Then
           strSQLQuery = " SELECT " _
                           & "      POInvoiceSearchView.PoNo,POInvoiceSearchView.SupplierId,POInvoiceSearchView.SupplierName, " _
                           & "      POInvoiceSearchView.POSeq,POInvoiceSearchView.ItemTypeId,POInvoiceSearchView.ItemId," _
                           & "      POInvoiceSearchView.Description,POInvoiceSearchView.OrderQty," _
                           & "      POInvoiceSearchView.ConvertingCoefficient,POInvoiceSearchView.CurrencyUnit," _
                           & "      POInvoiceSearchView.UnitPrice,POInvoiceSearchView.OrderDate, " _
                           & "      PoDetailsView.EtdDate as ETD, " _
                           & "      POInvoiceSearchView.ETA,POInvoiceSearchView.FTD as FTR, " _
                           & "      POInvoiceSearchView.InvoiceNo,POInvoiceSearchView.InvoiceQty, " _
                           & "      POInvoiceSearchView.ReceivedDate,POInvoiceSearchView.DeliveryId, " _
                           & "      POInvoiceSearchView.DeliveryDate,POInvoiceSearchView.RegisteredDate as EncodedDate,POInvoiceSearchView.QtyReceived, " _
                           & "      POInvoiceSearchView.QtyOk,POInvoiceSearchView.ReceivedAllInvoices, " _
                           & "      POInvoiceSearchView.DivisionId,POInvoiceSearchView.Maker,POInvoiceSearchView.ImportSupplier," _
                           & "      POInvoiceSearchView.PaymentRequestId, " _
                           & "      POInvoiceSearchView.Amount, " _
                           & "      POInvoiceSearchView.EntryDate, " _
                           & "      POInvoiceSearchView.DueDate, " _
                           & "      POInvoiceSearchView.StaffName,IH.DeliveryDate as ActualETD " _
                           & "      From POInvoiceSearchView Inner join PoDetailsView " _
                           & "      on POInvoiceSearchView.PoNo = PoDetailsView.PoNo"
        Else
           strSQLQuery = " SELECT " _
                           & "      POInvoiceSearchView.PoNo,POInvoiceSearchView.SupplierId,POInvoiceSearchView.SupplierName, " _
                           & "      POInvoiceSearchView.POSeq,POInvoiceSearchView.ItemTypeId,POInvoiceSearchView.ItemId," _
                           & "      POInvoiceSearchView.Description,POInvoiceSearchView.OrderQty," _
                           & "      POInvoiceSearchView.ConvertingCoefficient,POInvoiceSearchView.CurrencyUnit," _
                           & "      POInvoiceSearchView.UnitPrice,POInvoiceSearchView.OrderDate, " _
                           & "      PoDetailsView.EtdDate as ETD, " _
                           & "      POInvoiceSearchView.ETA,POInvoiceSearchView.FTD as FTR, " _
                           & "      POInvoiceSearchView.InvoiceNo,POInvoiceSearchView.InvoiceQty, " _
                           & "      POInvoiceSearchView.ReceivedDate,POInvoiceSearchView.DeliveryId, " _
                           & "      POInvoiceSearchView.DeliveryDate,POInvoiceSearchView.RegisteredDate as EncodedDate,POInvoiceSearchView.QtyReceived, " _
                           & "      POInvoiceSearchView.QtyOk,POInvoiceSearchView.ReceivedAllInvoices, " _
                           & "      POInvoiceSearchView.DivisionId,POInvoiceSearchView.Maker,POInvoiceSearchView.ImportSupplier," _
                           & "      POInvoiceSearchView.PaymentRequestId, " _
                           & "      POInvoiceSearchView.Amount, " _
                           & "      POInvoiceSearchView.EntryDate, " _
                           & "      POInvoiceSearchView.DueDate, " _
                           & "      POInvoiceSearchView.StaffName" _
                           & "      From POInvoiceSearchView Inner join PoDetailsView " _
                           & "      on POInvoiceSearchView.PoNo = PoDetailsView.PoNo"
        End If
                        
     End If
     If Not IsNull(dtActFrom.Value) Or Not IsNull(dtActTo.Value) Then _
               strSQLQuery = strSQLQuery & _
                              " INNER JOIN InvoiceHeaders IH ON POInvoiceSearchView.InvoiceNo = IH.InvoiceNo "
     
     
     strSQLQuery = strSQLQuery & " WHERE POInvoiceSearchView.DivisionId = " & clsPrintMenu.pfstrDivisionID(strDivision)
     '---Search by ItemId
     If Trim$(strItemId) <> "" Then _
          strSQLQuery = strSQLQuery & " and POInvoiceSearchView.ItemId = " & pfstrQt(strItemId)
     '---Search by Supplier
     If strSupplier <> "" Then _
          strSQLQuery = strSQLQuery & " and POInvoiceSearchView.SupplierName = " & pfstrQt(strSupplier)
     '---Search by PO No
     If Trim$(strPONo) <> "" Then _
          strSQLQuery = strSQLQuery & " and POInvoiceSearchView.PONo = " & pfstrQt(strPONo)
     '---Search by Invoice NO
     If Trim$(strInvoiceNo) <> "" Then _
          strSQLQuery = strSQLQuery & " and POInvoiceSearchView.InvoiceNo like " & pfstrQt(strInvoiceNo & "%")
     '----Search by ItemTypeId
     If Trim$(strItemTypeId) <> "" Then _
                strSQLQuery = strSQLQuery & " and POInvoiceSearchView.ItemTypeId =" & pfstrQt(strItemTypeId)
     '-----Search by Supplier Category
     If intSupCateg <= 1 Then _
            strSQLQuery = strSQLQuery & " and POInvoiceSearchView.ImportSupplier=" & intSupCateg & ""
            
     If Not IsNull(dtpODFrom.Value) Then _
               strSQLQuery = strSQLQuery & _
                              " and POInvoiceSearchView.OrderDate >= " & pfstrQt(Format(dtpODFrom.Value, "yyyy/mm/dd"))
     If Not IsNull(dtpODTo.Value) Then _
               strSQLQuery = strSQLQuery & _
                              " and POInvoiceSearchView.OrderDate < " & pfstrQt(DateAdd("d", 1, Format(dtpODTo.Value, "yyyy/mm/dd")))
     If Not IsNull(dtpInvFrom.Value) Then _
               strSQLQuery = strSQLQuery & _
                              " and POInvoiceSearchView.DeliveryDate >= " & pfstrQt(Format(dtpInvFrom.Value, "yyyy/mm/dd"))
     If Not IsNull(dtpInvTo.Value) Then _
               strSQLQuery = strSQLQuery & _
                              " and POInvoiceSearchView.DeliveryDate < " & pfstrQt(DateAdd("d", 1, Format(dtpInvTo.Value, "yyyy/mm/dd")))
    '------------Leng 20160629
    If Not IsNull(DTEncFrom.Value) Then _
               strSQLQuery = strSQLQuery & _
                              " and POInvoiceSearchView.RegisteredDate >= " & pfstrQt(Format(DTEncFrom.Value, "yyyy/mm/dd"))
     If Not IsNull(DTEncTo.Value) Then _
               strSQLQuery = strSQLQuery & _
                              " and POInvoiceSearchView.RegisteredDate < " & pfstrQt(DateAdd("d", 1, Format(DTEncTo.Value, "yyyy/mm/dd")))
      '------------Leng 20170816
    If Not IsNull(DTInvoiceFrom.Value) Then _
               strSQLQuery = strSQLQuery & _
                              " and POInvoiceSearchView.Invoicedate >= " & pfstrQt(Format(DTInvoiceFrom.Value, "yyyy/mm/dd"))
     If Not IsNull(DTInvoiceTo.Value) Then _
               strSQLQuery = strSQLQuery & _
                              " and POInvoiceSearchView.Invoicedate < " & pfstrQt(DateAdd("d", 1, Format(DTInvoiceTo.Value, "yyyy/mm/dd")))
    '-----------Ye 20180113
    If Not IsNull(dtETDFrom.Value) Then _
               strSQLQuery = strSQLQuery & _
                              " and PoDetailsView.EtdDate >= " & pfstrQt(Format(dtETDFrom.Value, "yyyy/mm/dd"))
     If Not IsNull(dtETDTo.Value) Then _
               strSQLQuery = strSQLQuery & _
                              " and PoDetailsView.EtdDate < " & pfstrQt(DateAdd("d", 1, Format(dtETDTo.Value, "yyyy/mm/dd")))
    If Not IsNull(dtActFrom.Value) Then _
               strSQLQuery = strSQLQuery & _
                              " and IH.DeliveryDate >= " & pfstrQt(Format(dtActFrom.Value, "yyyy/mm/dd"))
     If Not IsNull(dtActTo.Value) Then _
               strSQLQuery = strSQLQuery & _
                              " and IH.DeliveryDate <= " & pfstrQt(Format(dtActTo.Value, "yyyy/mm/dd"))
'    //----------------------------------
     If blnComplete Then _
               strSQLQuery = strSQLQuery & " and POInvoiceSearchView.ReceivedAllInvoices = 1 "
               
    If blnProdPurchase Then _
               strSQLQuery = strSQLQuery & " and POInvoiceSearchView.IsProductionPurchase = 1 "
               
     '---Total of Received Qty per Invoice
     If blnTotal Then
         If Not IsNull(dtActFrom.Value) Or Not IsNull(dtActTo.Value) Then
            strSQLQuery = strSQLQuery & _
                    " GROUP BY POInvoiceSearchView.PoNo, POInvoiceSearchView.SupplierName, POInvoiceSearchView.POSeq, " _
                    & "       POInvoiceSearchView.ItemId, POInvoiceSearchView.Description, POInvoiceSearchView.OrderQty, " _
                    & "       POInvoiceSearchView.OrderDate, POInvoiceSearchView.ETA,POInvoiceSearchView.FTD, POInvoiceSearchView.InvoiceNo, " _
                    & "       POInvoiceSearchView.ReceivedDate, POInvoiceSearchView.InvoiceQty, POInvoiceSearchView.ReceivedAllInvoices, " _
                    & "       POInvoiceSearchView.Maker,POInvoiceSearchView.ItemTypeId,POInvoiceSearchView.ConvertingCoefficient," _
                    & "       POInvoiceSearchView.StaffName,IH.DeliveryDate"
         
         Else
            strSQLQuery = strSQLQuery & _
                    " GROUP BY POInvoiceSearchView.PoNo, POInvoiceSearchView.SupplierName, POInvoiceSearchView.POSeq, " _
                    & "       POInvoiceSearchView.ItemId, POInvoiceSearchView.Description, POInvoiceSearchView.OrderQty, " _
                    & "       POInvoiceSearchView.OrderDate, POInvoiceSearchView.ETA,POInvoiceSearchView.FTD, POInvoiceSearchView.InvoiceNo, " _
                    & "       POInvoiceSearchView.ReceivedDate, POInvoiceSearchView.InvoiceQty, POInvoiceSearchView.ReceivedAllInvoices, " _
                    & "       POInvoiceSearchView.Maker,POInvoiceSearchView.ItemTypeId,POInvoiceSearchView.ConvertingCoefficient," _
                    & "       POInvoiceSearchView.StaffName"
         End If
         
     ElseIf blnFTotal Then
        If Not IsNull(dtActFrom.Value) Or Not IsNull(dtActTo.Value) Then
             strSQLQuery = strSQLQuery & _
                    " GROUP BY POInvoiceSearchView.PoNo, POInvoiceSearchView.SupplierName, POInvoiceSearchView.POSeq, " _
                    & "       POInvoiceSearchView.ItemId,POInvoiceSearchView.Description,POInvoiceSearchView.OrderQty, " _
                    & "       POInvoiceSearchView.OrderDate,POInvoiceSearchView.ETA,POInvoiceSearchView.FTD," _
                    & "       POInvoiceSearchView.ReceivedAllInvoices, POInvoiceSearchView.Maker,POInvoiceSearchView.ItemTypeId," _
                    & "       POInvoiceSearchView.ConvertingCoefficient,POInvoiceSearchView.StaffName,IH.DeliveryDate"
                    
        
        Else
            strSQLQuery = strSQLQuery & _
                    " GROUP BY POInvoiceSearchView.PoNo, POInvoiceSearchView.SupplierName, POInvoiceSearchView.POSeq, " _
                    & "       POInvoiceSearchView.ItemId,POInvoiceSearchView.Description,POInvoiceSearchView.OrderQty, " _
                    & "       POInvoiceSearchView.OrderDate,POInvoiceSearchView.ETA,POInvoiceSearchView.FTD," _
                    & "       POInvoiceSearchView.ReceivedAllInvoices, POInvoiceSearchView.Maker,POInvoiceSearchView.ItemTypeId," _
                    & "       POInvoiceSearchView.ConvertingCoefficient,POInvoiceSearchView.StaffName"
                    
        End If
         
     Else
        If Not IsNull(dtActFrom.Value) Or Not IsNull(dtActTo.Value) Then
            strSQLQuery = strSQLQuery & _
                    "GROUP BY  POInvoiceSearchView.PoNo,POInvoiceSearchView.SupplierId,POInvoiceSearchView.SupplierName, " _
                        & " POInvoiceSearchView.POSeq,POInvoiceSearchView.ItemTypeId,POInvoiceSearchView.ItemId, " _
                        & " POInvoiceSearchView.Description,POInvoiceSearchView.OrderQty,POInvoiceSearchView.ConvertingCoefficient, " _
                        & " POInvoiceSearchView.CurrencyUnit,POInvoiceSearchView.UnitPrice,POInvoiceSearchView.OrderDate, " _
                        & " POInvoiceSearchView.ETA,PoDetailsView.EtdDate,POInvoiceSearchView.FTD,POInvoiceSearchView.InvoiceNo," _
                        & " POInvoiceSearchView.InvoiceQty,POInvoiceSearchView.ReceivedDate,POInvoiceSearchView.DeliveryId, " _
                        & " POInvoiceSearchView.DeliveryDate,POInvoiceSearchView.RegisteredDate,POInvoiceSearchView.QtyReceived,POInvoiceSearchView.QtyOk, " _
                        & " POInvoiceSearchView.ReceivedAllInvoices,POInvoiceSearchView.DivisionId,POInvoiceSearchView.Maker, " _
                        & " POInvoiceSearchView.ImportSupplier,POInvoiceSearchView.PaymentRequestId,POInvoiceSearchView.Amount, " _
                        & " POInvoiceSearchView.EntryDate,POInvoiceSearchView.DueDate,POInvoiceSearchView.StaffName,PoDetailsView.EtdDate,IH.DeliveryDate"
        
        Else
            strSQLQuery = strSQLQuery & _
                    "GROUP BY  POInvoiceSearchView.PoNo,POInvoiceSearchView.SupplierId,POInvoiceSearchView.SupplierName, " _
                        & " POInvoiceSearchView.POSeq,POInvoiceSearchView.ItemTypeId,POInvoiceSearchView.ItemId, " _
                        & " POInvoiceSearchView.Description,POInvoiceSearchView.OrderQty,POInvoiceSearchView.ConvertingCoefficient, " _
                        & " POInvoiceSearchView.CurrencyUnit,POInvoiceSearchView.UnitPrice,POInvoiceSearchView.OrderDate, " _
                        & " POInvoiceSearchView.ETA,PoDetailsView.EtdDate,POInvoiceSearchView.FTD,POInvoiceSearchView.InvoiceNo," _
                        & " POInvoiceSearchView.InvoiceQty,POInvoiceSearchView.ReceivedDate,POInvoiceSearchView.DeliveryId, " _
                        & " POInvoiceSearchView.DeliveryDate,POInvoiceSearchView.RegisteredDate,POInvoiceSearchView.QtyReceived,POInvoiceSearchView.QtyOk, " _
                        & " POInvoiceSearchView.ReceivedAllInvoices,POInvoiceSearchView.DivisionId,POInvoiceSearchView.Maker, " _
                        & " POInvoiceSearchView.ImportSupplier,POInvoiceSearchView.PaymentRequestId,POInvoiceSearchView.Amount, " _
                        & " POInvoiceSearchView.EntryDate,POInvoiceSearchView.DueDate,POInvoiceSearchView.StaffName,PoDetailsView.EtdDate"
        End If
            'Gerald 20141115
                   
     End If
     
     strSQLQuery = strSQLQuery & "  ORDER BY POInvoiceSearchView.PONo, POInvoiceSearchView.POSeq"
     
     Set rsInvoiceSearch = clsDB.GetRecordSet(strSQLQuery)
     Set hflexGrid.DataSource = rsInvoiceSearch
     
        If rsInvoiceSearch.EOF Then
            MsgBox "No Data Found", vbInformation, "SYSTEM MESSAGE"
        End If
     Set rsInvoiceSearch = Nothing
     Exit Sub
     
lnErrHandler:
     MsgBox Err.Description, vbCritical
End Sub
